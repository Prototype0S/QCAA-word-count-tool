"""HTML preview builder.

The template lives in preview_template.html alongside this module.
Python does simple {{placeholder}} substitution with str.replace —
no str.format, no brace escaping.
"""
from __future__ import annotations

import html
import json
import os
from pathlib import Path

from ..helpers import heading_level_from_style
from ..model import Block, BlockType, Decision, Document, Span, SpanType
from .ansi import interleaved_items
from .stats import compute_statistics


# ---------------------------------------------------------------------------
# Template loading
# ---------------------------------------------------------------------------

def _load_template() -> str:
    path = Path(__file__).with_name("preview_template.html")
    return path.read_text(encoding="utf-8")


_TEMPLATE_CACHE: str | None = None


def _template() -> str:
    global _TEMPLATE_CACHE
    if _TEMPLATE_CACHE is None:
        _TEMPLATE_CACHE = _load_template()
    return _TEMPLATE_CACHE


def _fill(template: str, values: dict[str, str]) -> str:
    for key, value in values.items():
        template = template.replace("{{" + key + "}}", value)
    return template


# ---------------------------------------------------------------------------
# Span IDs
#
# Every toggleable span needs a stable, unique identifier so the browser
# can toggle *this* instance of "the" rather than every "the" on the page.
#
# The ID is (block_index, char_start), which is exactly the identity of
# the span within the document. That means the ID is also what we can
# safely round-trip through qcaa_overrides.json for --apply-overrides —
# the same span gets the same ID on every render.
# ---------------------------------------------------------------------------

def _span_id(span: Span) -> str:
    return f"{span.source.block_index}:{span.source.char_start}"


# ---------------------------------------------------------------------------
# Span rendering
# ---------------------------------------------------------------------------

def _span_html(span: Span) -> str:
    text = html.escape(span.text)
    dec = span.decision_record.decision
    rule = span.decision_record.rule_id or ""
    reason = span.decision_record.reason or ""
    dec_attr = html.escape(dec.value if hasattr(dec, "value") else str(dec))
    rule_attr = html.escape(rule)
    reason_attr = html.escape(reason)
    text_attr = html.escape(span.text)
    origin_attr = "1" if dec == Decision.COUNT else "0"
    sid = _span_id(span)
    meta = (
        f' data-span-id="{sid}"'
        f' data-decision="{dec_attr}"'
        f' data-rule="{rule_attr}"'
        f' data-reason="{reason_attr}"'
        f' data-text="{text_attr}"'
        f' data-original-counted="{origin_attr}"'
    )

    if dec == Decision.COUNT:
        if span.span_type == SpanType.UNIT:
            return f'<span class="unit span-toggle"{meta}>{text}</span>'
        return f'<span class="count span-toggle"{meta}>{text}</span>'
    if dec == Decision.EXCLUDE:
        if rule == "I-CITATION":
            return f'<span class="citation span-toggle"{meta}>{text}</span>'
        return f'<span class="exclude span-toggle"{meta}>{text}</span>'
    if dec in (Decision.FLAG, Decision.UNDECIDED):
        return f'<span class="flag span-toggle"{meta}>{text}</span>'
    return text


def _block_html(block: Block) -> str:
    if block.block_type == BlockType.HEADING:
        level = heading_level_from_style(block.style_name) or 1
        return f'<p class="heading h{min(level, 3)}">{html.escape(block.text.strip())}</p>'
    if block.region_record.decision == Decision.EXCLUDE:
        return f'<p class="region-ex">{html.escape(block.text.strip())}</p>'
    if not block.spans:
        return f'<p>{html.escape(block.text.strip())}</p>'
    parts = []
    last = 0
    text = block.text
    for span in block.spans:
        s, e = span.source.char_start, span.source.char_end
        if s > last:
            parts.append(html.escape(text[last:s]))
        parts.append(_span_html(span))
        last = e
    if last < len(text):
        parts.append(html.escape(text[last:]))
    return f"<p>{''.join(parts)}</p>"


def _table_html(table) -> str:
    rows = []
    for r_idx, row in enumerate(table.rows):
        cells = []
        for cell in row:
            if not cell.spans:
                cells.append("<td></td>")
                continue
            parts = []
            last = 0
            text = cell.text
            for span in cell.spans:
                s, e = span.source.char_start, span.source.char_end
                if s > last:
                    parts.append(html.escape(text[last:s]))
                parts.append(_span_html(span))
                last = e
            if last < len(text):
                parts.append(html.escape(text[last:]))
            tag = "th" if r_idx == table.header_row_index else "td"
            cells.append(f"<{tag}>{''.join(parts)}</{tag}>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    answer = table.answer.value if table.answer else "unresolved"
    answer_badge = {
        "count_all": "badge-good",
        "exclude_all": "badge-bad",
        "headers_only": "badge-warn",
        "unresolved": "badge-warn",
    }.get(answer, "badge-warn")
    answer_label = {
        "count_all": "counted whole",
        "exclude_all": "excluded whole",
        "headers_only": "headers only",
        "unresolved": "unresolved",
    }.get(answer, answer)
    return (
        f'<table>'
        f'<caption class="caption">'
        f'Table {table.table_index + 1} '
        f'<span class="badge {answer_badge}">{answer_label}</span>'
        f'</caption>'
        f'{"".join(rows)}'
        f'</table>'
    )


# ---------------------------------------------------------------------------
# Statistics block (unchanged from before — kept for completeness)
# ---------------------------------------------------------------------------

def _build_statistics_html(stats: dict, result: dict) -> str:
    sections = stats["sections"]
    total_counted = result["total"] or 1

    cards = [
        ("Word count", str(result["total"]), "counted words"),
        ("Sections", str(len(sections)), "headings + their content"),
        ("Unique words", str(stats["unique_words"]),
         f'{stats["richness_pct"]:.1f}% vocabulary richness'),
        ("Longest paragraph", str(stats["longest_paragraph_words"]),
         "words in the longest counted paragraph"),
    ]
    cards_html = "".join(
        f'<div class="stat-card">'
        f'<div class="label">{html.escape(label)}</div>'
        f'<div class="value">{html.escape(value)}</div>'
        f'<div class="sub">{html.escape(sub)}</div>'
        f'</div>'
        for label, value, sub in cards
    )

    palette = [
        "#0969da", "#1a7f37", "#8250df", "#9a6700", "#cf222e",
        "#0d9488", "#d97706", "#7c3aed", "#db2777", "#0284c7",
    ]
    bar_segments = []
    for i, section in enumerate(sections):
        pct = (section["words"] / total_counted) * 100
        if pct <= 0:
            continue
        colour = palette[i % len(palette)]
        bar_segments.append(
            f'<div style="background:{colour};width:{pct:.2f}%" '
            f'title="{html.escape(section["heading"])}: '
            f'{section["words"]} words">'
            f'{section["words"] if pct >= 6 else ""}</div>'
        )
    bar_html = (
        f'<div class="section-bar">{"".join(bar_segments)}</div>'
        if bar_segments else ""
    )

    section_items = []
    for i, section in enumerate(sections):
        pct = (section["words"] / total_counted) * 100
        colour = palette[i % len(palette)]
        tables_note = (
            f' <span style="color:var(--grey);font-size:.8rem;">'
            f'· {section["tables"]} table(s)</span>'
            if section["tables"] else ""
        )
        section_items.append(
            f'<li>'
            f'<span class="name">'
            f'<span style="display:inline-block;width:10px;height:10px;'
            f'background:{colour};border-radius:2px;margin-right:.4rem;'
            f'vertical-align:middle;"></span>'
            f'{html.escape(section["heading"])}{tables_note}'
            f'</span>'
            f'<span class="count">{section["words"]}'
            f'<span class="pct">{pct:.1f}%</span></span>'
            f'</li>'
        )
    section_list = (
        f'<ul class="section-list">{"".join(section_items)}</ul>'
        if section_items else "<p><em>No countable sections.</em></p>"
    )

    longest = stats.get("longest_paragraph_preview") or "(none)"
    return (
        f'<div class="stat-cards">{cards_html}</div>'
        f'{bar_html}'
        f'<h3 style="font-size:1rem;margin:1rem 0 .5rem 0;">'
        f'Words per section</h3>'
        f'{section_list}'
        f'<p style="color:var(--grey);font-size:.8rem;margin-top:1rem;">'
        f'Longest paragraph preview: <em>{html.escape(longest)}</em></p>'
    )


# ---------------------------------------------------------------------------
# Summary text
# ---------------------------------------------------------------------------

def _build_summary_text(document: Document, result: dict) -> str:
    lines = [
        f"QCAA word count: {result['total']}",
        f"File: {document.filename}",
    ]
    if result.get("tables"):
        lines.append(f"Tables processed: {len(result['tables'])}")
    cit = result.get("citations", {})
    if cit.get("auto") or cit.get("manual"):
        lines.append(
            f"Citations excluded: {cit.get('auto', 0) + cit.get('manual', 0)}"
        )

    stats = compute_statistics(document, result)
    sections = stats.get("sections", [])
    if sections:
        total_counted = result["total"] or 1
        lines.append("")
        lines.append("Words per section:")
        longest = min(max(len(s["heading"]) for s in sections), 40)
        for s in sections:
            heading = s["heading"]
            if len(heading) > longest:
                heading = heading[: longest - 1] + "…"
            pct = (s["words"] / total_counted) * 100
            tables_note = f"  ({s['tables']} table(s))" if s["tables"] else ""
            lines.append(
                f"  {heading:<{longest}}  {s['words']:>5}  ({pct:>5.1f}%){tables_note}"
            )

    if result["buckets"]:
        lines.append("")
        lines.append("Excluded:")
        for key in sorted(result["buckets"]):
            lines.append(f"  {key}: {result['buckets'][key]} word(s)")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Top-level builders
# ---------------------------------------------------------------------------

def build_html_preview(document: Document, result: dict) -> str:
    body = []
    for kind, obj in interleaved_items(document):
        if kind == "block":
            body.append(_block_html(obj))
        else:
            body.append(_table_html(obj))
    body_html = "\n".join(body) or "<p><em>(No content)</em></p>"

    stats_rows = []
    for key, count in sorted(result["buckets"].items()):
        stats_rows.append(
            f'<dt>{html.escape(key)}</dt><dd>{count} word(s)</dd>'
        )
    if result.get("math_objects"):
        stats_rows.append(
            f'<dt>Math objects</dt><dd>{result["math_objects"]} (excluded)</dd>'
        )
    if document.drawings:
        stats_rows.append(
            f'<dt>Drawings</dt><dd>{document.drawings} '
            f'({document.drawings_with_text} contain text)</dd>'
        )
    if document.text_boxes:
        stats_rows.append(
            f'<dt>Text boxes</dt><dd>{len(document.text_boxes)}</dd>'
        )

    banner_html = ""
    stated = result.get("stated_count")
    if stated and stated != result["total"]:
        delta = result["total"] - stated
        banner_html = (
            f'<div class="mismatch-banner">'
            f'⚠ Title page says "Word Count: {stated}" &middot; '
            f'Tool calculates <strong>{result["total"]}</strong> '
            f'(differs by {abs(delta)})'
            f'</div>'
        )

    stats = compute_statistics(document, result)
    statistics_html = _build_statistics_html(stats, result)

    additional = ""
    if result.get("tables"):
        rows = []
        for t in result["tables"]:
            chosen_class = (
                "table-chosen-count" if t["answer"] == "count_all"
                else "table-chosen-exclude"
            )
            answer_label = {
                "count_all": "counted whole",
                "exclude_all": "excluded whole",
                "headers_only": "headers only",
                "unresolved": "unresolved",
            }.get(t["answer"], t["answer"])
            rows.append(
                f'<tr><td>Table {t["index"]}</td>'
                f'<td>{html.escape(t["suggestion"])}</td>'
                f'<td class="{chosen_class}">{html.escape(answer_label)}</td>'
                f'<td>{t["counted"]}</td>'
                f'<td>{t["excluded"]}</td></tr>'
            )
        additional += (
            "<div class='panel'><h2>Tables "
            "<span class='badge badge-info'>per-table breakdown</span></h2>"
            "<table><tr><th>#</th><th>Suggestion</th><th>Chosen</th>"
            "<th>Counted</th><th>Excluded</th></tr>"
            + "".join(rows) + "</table></div>"
        )

    cit = result.get("citations", {})
    if cit.get("auto") or cit.get("manual"):
        additional += (
            f"<div class='panel'><h2>Citations "
            f"<span class='badge badge-info'>APA 7 auto-detected</span></h2>"
            f"<p>Auto-detected: <strong>{cit.get('auto', 0)}</strong> "
            f"({cit.get('auto_words', 0)} words)<br>"
            f"Manually added: <strong>{cit.get('manual', 0)}</strong> "
            f"({cit.get('manual_words', 0)} words)</p></div>"
        )

    summary_text = _build_summary_text(document, result)
    summary_json = json.dumps(summary_text).replace("</", "<\\/")

    return _fill(_template(), {
        "filename": html.escape(document.filename or ""),
        "total": str(result["total"]),
        "banner_html": banner_html,
        "stats_rows": "".join(stats_rows),
        "statistics_html": statistics_html,
        "body_html": body_html,
        "additional_panels": additional,
        "summary_json": summary_json,
        "original_total": str(result["total"]),
    })


def write_html_preview(document: Document, result: dict,
                       output_path: str | None = None,
                       open_browser: bool = True) -> str:
    html_text = build_html_preview(document, result)
    if output_path is None:
        base = os.path.splitext(document.filename or "preview")[0]
        output_path = f"{base}_qcaa_preview.html"
    Path(output_path).write_text(html_text, encoding="utf-8")
    if open_browser:
        try:
            import webbrowser
            webbrowser.open(f"file://{os.path.abspath(output_path)}")
        except Exception:
            pass
    return output_path