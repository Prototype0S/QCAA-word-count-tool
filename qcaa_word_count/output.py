"""All output: ANSI preview, HTML preview, decision log, plain report."""
from __future__ import annotations

import csv
import html
import io
import json
import os
import webbrowser
from pathlib import Path

from .helpers import heading_level_from_style
from .model import Block, BlockType, Decision, Document, Region, Span, SpanType

# ---------------------------------------------------------------------------
# ANSI colour helpers
# ---------------------------------------------------------------------------

import os as _os

_COLOURS_ENABLED = True


def _enable_windows_ansi() -> bool:
    if _os.name != "nt":
        return True
    try:
        import ctypes
        k = ctypes.windll.kernel32
        k.SetConsoleMode(k.GetStdHandle(-11), 7)
        return True
    except Exception:
        return False


if not _enable_windows_ansi():
    _COLOURS_ENABLED = False

_R = "\033[0m"; _B = "\033[1m"
_GRN = "\033[32m"; _RED = "\033[31m"; _YEL = "\033[33m"
_CYN = "\033[36m"; _MAG = "\033[35m"; _GRY = "\033[90m"


def _c(t: str, *codes: str) -> str:
    return "".join(codes) + t + _R if _COLOURS_ENABLED else t


def green(t):   return _c(t, _GRN)
def red(t):     return _c(t, _RED)
def yellow(t):  return _c(t, _YEL)
def cyan(t):    return _c(t, _CYN)
def grey(t):    return _c(t, _GRY)
def magenta(t): return _c(t, _MAG)
def bold(t):    return _c(t, _B)
def bold_green(t): return _c(t, _B, _GRN)


def span_to_ansi(span: Span) -> str:
    dec = span.decision_record.decision
    rule = span.decision_record.rule_id or ""
    if dec == Decision.COUNT:
        return green(span.text)
    if dec == Decision.EXCLUDE:
        return magenta(span.text) if rule == "I-CITATION" else red(span.text)
    if dec in (Decision.FLAG, Decision.UNDECIDED):
        return yellow(span.text)
    return span.text


def _interleaved_items(document):
    items = []
    for b in document.blocks:
        items.append((b.block_index, 1, "block", b))
    for t in document.tables:
        items.append((t.block_index, 0, "table", t))
    items.sort(key=lambda x: (x[0], x[1]))
    for _, _, kind, obj in items:
        yield kind, obj


def print_ansi_preview(document: Document) -> None:
    print("=" * 78)
    print(bold("COLOUR-CODED PREVIEW"))
    print(grey("Green = counted | Magenta = citation | Red = excluded | "
               "Yellow = flagged | Grey = region | Cyan = heading"))
    print("=" * 78)
    print()

    for kind, obj in _interleaved_items(document):
        if kind == "block":
            block = obj
            if not block.spans:
                continue
            if block.block_type == BlockType.HEADING:
                print()
                print(cyan(bold(block.text.strip())))
                continue
            if block.region_record.decision == Decision.EXCLUDE:
                print(grey(block.text.strip()))
                continue
            print("".join(span_to_ansi(s) for s in block.spans))
        else:
            table = obj
            print()
            answer = table.answer.value if table.answer else "unresolved"
            print(cyan(bold(f"Table {table.table_index + 1} — {answer}")))
            for row in table.rows:
                row_strs = [" ".join(span_to_ansi(s) for s in cell.spans)
                            for cell in row]
                print(" | ".join(row_strs))

    print()
    print("=" * 78)


def build_html_preview(document: Document, result: dict) -> str:
    body = []
    for kind, obj in _interleaved_items(document):
        if kind == "block":
            body.append(_block_html(obj))
        else:
            body.append(_table_html(obj))
    body_html = "\n".join(body) or "<p><em>(No content)</em></p>"

    stats_rows = []
    for key, count in sorted(result["buckets"].items()):
        stats_rows.append(f"<dt>{html.escape(key)}</dt><dd>{count} word(s)</dd>")
    if result.get("math_objects"):
        stats_rows.append(
            f"<dt>Math objects</dt><dd>{result['math_objects']} (excluded)</dd>"
        )
    if document.drawings:
        stats_rows.append(
            f"<dt>Drawings</dt><dd>{document.drawings} "
            f"({document.drawings_with_text} contain text)</dd>"
        )
    if document.text_boxes:
        stats_rows.append(f"<dt>Text boxes</dt><dd>{len(document.text_boxes)}</dd>")

    additional = ""
    if result.get("tables"):
        rows = []
        for t in result["tables"]:
            rows.append(
                f"<tr><td>Table {t['index']}</td><td>{html.escape(t['suggestion'])}</td>"
                f"<td>{html.escape(t['answer'])}</td>"
                f"<td>{t['counted']}</td><td>{t['excluded']}</td></tr>"
            )
        additional += (
            "<div class='panel'><h2>Tables</h2><table>"
            "<tr><th>#</th><th>Suggestion</th><th>Chosen</th>"
            "<th>Counted</th><th>Excluded</th></tr>"
            + "".join(rows) + "</table></div>"
        )

    cit = result.get("citations", {})
    if cit.get("auto") or cit.get("manual"):
        additional += (
            f"<div class='panel'><h2>Citations</h2>"
            f"<p>Auto-detected (APA 7): <strong>{cit.get('auto', 0)}</strong> "
            f"({cit.get('auto_words', 0)} words)<br>"
            f"Manually added: <strong>{cit.get('manual', 0)}</strong> "
            f"({cit.get('manual_words', 0)} words)</p></div>"
        )

    return _HTML_TEMPLATE.format(
        filename=html.escape(document.filename or ""),
        total=result["total"],
        stats_rows="".join(stats_rows),
        body_html=body_html,
        additional_panels=additional,
    )

# ---------------------------------------------------------------------------
# Plain report
# ---------------------------------------------------------------------------

def print_report(document: Document, result: dict) -> None:
    print()
    print("=" * 60)
    print(bold("QCAA WORD COUNT REPORT"))
    print(f"File: {document.filename}")
    print("=" * 60)
    print()
    print(f"  {bold_green('FINAL WORD COUNT:')} {bold_green(str(result['total']))}")
    print()

    if result["buckets"]:
        print(bold("Excluded / flagged (not counted):"))
        samples = result.get("bucket_samples", {})
        for key in sorted(result["buckets"]):
            col = yellow if (key.startswith("flagged:") or "unresolved" in key) else red
            count = result["buckets"][key]
            sample_str = ""
            if samples.get(key):
                parts = [f'"{t}"' for t in samples[key]]
                if count > len(samples[key]):
                    parts.append("...")
                sample_str = "  " + grey("(" + ", ".join(parts) + ")")
            print(f"  {col(key):<40} {count:>5} word(s){sample_str}")
        print()

    if result.get("tables"):
        print(bold("Tables:"))
        for t in result["tables"]:
            label = {
                "count_all": "counted whole",
                "exclude_all": "excluded whole",
                "headers_only": "headers only",
                "unresolved": "unresolved",
            }.get(t["answer"], t["answer"])
            line = (f"  Table {t['index']}: suggested={t['suggestion']:<12} "
                    f"chosen={label:<15} "
                    f"counted={t['counted']:<4} excluded={t['excluded']}")
            if t["answer"] == "count_all":
                print(green(line))
            elif t["answer"] in ("exclude_all", "headers_only"):
                print(red(line))
            else:
                print(yellow(line))
        print()

    cit = result.get("citations", {})
    if cit.get("auto") or cit.get("manual"):
        print(bold("In-text citations:"))
        if cit["auto"]:
            print(green(f"  Auto-detected (APA 7):   {cit['auto']:>4} "
                       f"({cit['auto_words']} words)"))
        if cit["manual"]:
            print(green(f"  Manually added:          {cit['manual']:>4} "
                       f"({cit['manual_words']} words)"))
        print()

    fn = result.get("footnotes", {})
    if fn.get("counted") or fn.get("excluded"):
        print(bold("Footnotes / endnotes:"))
        if fn["counted"]:
            print(green(f"  Counted (commentary):     {fn['counted']:>4} "
                       f"({fn['counted_words']} words)"))
        if fn["excluded"]:
            print(red(f"  Excluded (bibliographic): {fn['excluded']:>4} "
                     f"({fn['excluded_words']} words)"))
        print()

    if result.get("unresolved_footnotes"):
        print(yellow(f"  {result['unresolved_footnotes']} footnote(s) "
                     f"unresolved and NOT counted."))

    if result.get("math_objects"):
        print(red(f"Math objects (OMML):       {result['math_objects']:>4} "
                 f"(excluded)"))

    if document.drawings:
        msg = f"Drawings/shapes:          {document.drawings:>4}"
        if document.drawings_with_text:
            msg += f" ({document.drawings_with_text} contain text — not extracted)"
        print(grey(msg))
    if document.text_boxes:
        print(grey(f"Text boxes:               {len(document.text_boxes):>4}"))

    print()
    print(grey("Citation detection uses APA 7 patterns. Review the report and"))
    print(grey("use manual entry if any were missed."))


# ---------------------------------------------------------------------------
# HTML preview
# ---------------------------------------------------------------------------

_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>QCAA word count preview — {filename}</title>
<style>
:root {{
  --green:#1a7f37; --red:#cf222e; --magenta:#8250df; --yellow:#9a6700;
  --grey:#6e7781; --cyan:#0969da; --bg:#ffffff; --card:#f6f8fa;
  --border:#d0d7de;
}}
body {{
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
  max-width:1100px; margin:2rem auto; padding:0 1.5rem;
  color:#1f2328; line-height:1.6;
}}
h1 {{ font-size:1.6rem; margin-bottom:.25rem; }}
.subtitle {{ color:var(--grey); margin-bottom:1.5rem; }}
.legend {{ display:flex; flex-wrap:wrap; gap:.75rem; margin-bottom:1.5rem; font-size:.9rem; }}
.legend span {{ display:inline-flex; align-items:center; gap:.4rem;
  padding:.2rem .5rem; border-radius:4px; background:var(--card);
  border:1px solid var(--border); }}
.swatch {{ display:inline-block; width:12px; height:12px; border-radius:2px; }}
.panel {{ background:var(--card); border:1px solid var(--border);
  border-radius:6px; padding:1rem 1.25rem; margin-bottom:1.5rem; }}
.panel h2 {{ font-size:1.1rem; margin:0 0 .5rem 0; }}
.doc-body {{ font-family:"Source Serif Pro",Georgia,"Times New Roman",serif;
  font-size:1.05rem; white-space:pre-wrap; word-wrap:break-word; }}
.doc-body p {{ margin:0 0 .6rem 0; }}
.doc-body .heading {{ color:var(--cyan); font-weight:600; font-size:1.15rem;
  margin-top:1.2rem; }}
.doc-body .heading.h1 {{ font-size:1.4rem; }}
.doc-body .heading.h2 {{ font-size:1.2rem; }}
.count {{ color:var(--green); }}
.exclude {{ color:var(--red); text-decoration:line-through;
  text-decoration-color:rgba(207,34,46,.4); }}
.citation {{ color:var(--magenta); text-decoration:line-through;
  text-decoration-color:rgba(130,80,223,.4); }}
.flag {{ color:var(--yellow); background:rgba(154,103,0,.1);
  border-bottom:1px dashed var(--yellow); }}
.region-ex {{ color:var(--grey); font-style:italic; }}
.unit {{ color:var(--green); border-bottom:1px dotted var(--green); }}
table {{ border-collapse:collapse; margin:.5rem 0 1rem 0; width:100%; font-size:.95rem; }}
td, th {{ border:1px solid var(--border); padding:.35rem .6rem;
  vertical-align:top; white-space:pre-wrap; }}
th {{ background:var(--card); text-align:left; }}
.stats-grid {{ display:grid; grid-template-columns:max-content 1fr;
  gap:.25rem 1rem; font-size:.95rem; }}
.stats-grid dt {{ color:var(--grey); }}
.stats-grid dd {{ margin:0; }}
details {{ margin-top:.5rem; }}
summary {{ cursor:pointer; color:var(--cyan); }}
.final-count {{ font-size:2rem; font-weight:600; color:var(--green);
  margin:.25rem 0 1rem 0; }}
</style></head><body>

<h1>QCAA word count preview</h1>
<div class="subtitle">{filename}</div>

<div class="legend">
<span><span class="swatch" style="background:var(--green)"></span>counted</span>
<span><span class="swatch" style="background:var(--red)"></span>excluded</span>
<span><span class="swatch" style="background:var(--magenta)"></span>citation</span>
<span><span class="swatch" style="background:var(--yellow)"></span>flagged</span>
<span><span class="swatch" style="background:var(--grey)"></span>region-excluded</span>
<span><span class="swatch" style="background:var(--cyan)"></span>heading</span>
</div>

<div class="panel">
<h2>Final word count</h2>
<div class="final-count">{total}</div>
<dl class="stats-grid">{stats_rows}</dl>
</div>

<div class="panel">
<h2>Document with decisions applied</h2>
<div class="doc-body">{body_html}</div>
</div>

{additional_panels}

<details>
<summary>Notes</summary>
<p>Green text counts. Red text was excluded (numbers, symbols,
equations). Magenta text was excluded as a citation. Yellow text was
flagged and left undecided. Grey text is from an excluded region.
Cyan text is a heading.</p>
<p>Citation detection is APA 7 only. Other styles may be counted as
ordinary words.</p>
</details>

</body></html>
"""


def _span_html(span: Span) -> str:
    text = html.escape(span.text)
    dec = span.decision_record.decision
    rule = span.decision_record.rule_id or ""
    if dec == Decision.COUNT:
        if span.span_type == SpanType.UNIT:
            return f'<span class="unit" title="unit (counted)">{text}</span>'
        return f'<span class="count" title="counted">{text}</span>'
    if dec == Decision.EXCLUDE:
        if rule == "I-CITATION":
            return f'<span class="citation" title="citation (excluded)">{text}</span>'
        return f'<span class="exclude" title="{html.escape(rule)} (excluded)">{text}</span>'
    if dec in (Decision.FLAG, Decision.UNDECIDED):
        return f'<span class="flag" title="{html.escape(rule)} (flagged)">{text}</span>'
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
    cap = (f"<caption style='text-align:left;color:#6e7781;font-size:.85rem;'>"
           f"Table {table.table_index + 1} — {answer}</caption>")
    return f"<table>{cap}{''.join(rows)}</table>"


def write_html_preview(document: Document, result: dict,
                       output_path: str | None = None,
                       open_browser: bool = False) -> str:
    html_text = build_html_preview(document, result)
    if output_path is None:
        base = os.path.splitext(document.filename or "preview")[0]
        output_path = f"{base}_qcaa_preview.html"
    Path(output_path).write_text(html_text, encoding="utf-8")
    if open_browser:
        try:
            webbrowser.open(f"file://{os.path.abspath(output_path)}")
        except Exception:
            pass
    return output_path


# ---------------------------------------------------------------------------
# Decision log
# ---------------------------------------------------------------------------

def export_decisions(document: Document, result: dict, format: str = "json") -> str:
    payload = {
        "filename": document.filename,
        "total": result["total"],
        "buckets": result["buckets"],
        "blocks": [
            {
                "index": b.block_index,
                "type": b.block_type.value,
                "region": b.region.value,
                "region_decision": b.region_record.decision.value,
                "region_rule": b.region_record.rule_id,
                "text": b.text,
                "spans": [
                    {
                        "text": s.text,
                        "type": s.span_type.value,
                        "decision": s.decision_record.decision.value,
                        "rule": s.decision_record.rule_id,
                        "reason": s.decision_record.reason,
                        "tags": sorted(s.tags),
                    }
                    for s in b.spans
                ],
            }
            for b in document.blocks
        ],
        "tables": [
            {
                "index": t.table_index + 1,
                "suggestion": t.suggestion.value,
                "answer": t.answer.value if t.answer else None,
                "header_row": t.header_row_index,
            }
            for t in document.tables
        ],
        "citations": [
            {"text": c.text, "occurrences": c.occurrences,
             "source": c.source, "pattern": c.pattern_name}
            for c in document.citations
        ],
        "footnotes": [
            {"id": f.footnote_id, "kind": f.kind, "text": f.text,
             "suggestion": f.suggestion, "answer": f.answer}
            for f in document.footnotes
        ],
    }
    if format == "json":
        return json.dumps(payload, indent=2, ensure_ascii=False)
    if format == "csv":
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(["block_index", "span_text", "span_type",
                    "decision", "rule", "reason"])
        for b in document.blocks:
            for s in b.spans:
                w.writerow([
                    b.block_index, s.text, s.span_type.value,
                    s.decision_record.decision.value,
                    s.decision_record.rule_id,
                    s.decision_record.reason or "",
                ])
        return buf.getvalue()
    return ""