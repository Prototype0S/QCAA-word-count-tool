"""All output: ANSI preview, HTML preview, decision log, plain report."""
from __future__ import annotations

import csv
import html
import io
import json
import os
import re
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

_R = "\033[0m"; _B = "\033[1m"; _DIM = "\033[2m"
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
def dim(t):     return _c(t, _DIM)
def bold_green(t): return _c(t, _B, _GRN)
def bold_cyan(t):  return _c(t, _B, _CYN)


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


# ---------------------------------------------------------------------------
# Interleaved iteration
# ---------------------------------------------------------------------------

def _interleaved_items(document):
    items = []
    for b in document.blocks:
        items.append((b.block_index, 1, "block", b))
    for t in document.tables:
        items.append((t.block_index, 0, "table", t))
    items.sort(key=lambda x: (x[0], x[1]))
    for _, _, kind, obj in items:
        yield kind, obj


# ---------------------------------------------------------------------------
# ANSI preview
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Plain-text report
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
# Statistics
# ---------------------------------------------------------------------------

def compute_statistics(document: Document, result: dict) -> dict:
    sections: list[dict] = []
    current = {"heading": "(start of document)", "words": 0, "tables": 0}

    counted_by_index: dict[int, int] = {}
    for block in document.blocks:
        counted_by_index[block.block_index] = sum(
            1 for s in block.spans
            if s.decision_record.decision == Decision.COUNT
        )

    table_words_by_position: dict[int, int] = {}
    for table in document.tables:
        n = sum(
            1 for cell in table.iter_cells() for s in cell.spans
            if s.decision_record.decision == Decision.COUNT
        )
        table_words_by_position[table.block_index] = (
            table_words_by_position.get(table.block_index, 0) + n
        )

    for kind, obj in _interleaved_items(document):
        if kind == "block":
            block = obj
            if block.block_type == BlockType.HEADING:
                if current["words"] or current["tables"] or current["heading"] != "(start of document)":
                    sections.append(current)
                current = {
                    "heading": block.text.strip() or "(untitled)",
                    "words": counted_by_index.get(block.block_index, 0),
                    "tables": 0,
                }
            else:
                current["words"] += counted_by_index.get(block.block_index, 0)
        else:
            table = obj
            current["words"] += table_words_by_position.get(table.block_index, 0)
            current["tables"] += 1

    if current["words"] or current["tables"] or current["heading"] != "(start of document)":
        sections.append(current)

    sections = [s for s in sections if s["words"] or s["tables"]]

    counted_words: list[str] = []
    for block in document.blocks:
        if block.region_record.decision != Decision.COUNT:
            continue
        for span in block.spans:
            if span.decision_record.decision == Decision.COUNT:
                counted_words.append(span.text)

    alpha_words = [w.strip(".,;:!?()[]{}\"'").lower() for w in counted_words]
    alpha_words = [w for w in alpha_words if w and any(c.isalpha() for c in w)]
    unique = len(set(alpha_words))
    total_alpha = len(alpha_words)
    richness = (unique / total_alpha * 100) if total_alpha else 0.0

    longest_block = None
    longest_count = 0
    for block in document.blocks:
        n = counted_by_index.get(block.block_index, 0)
        if n > longest_count:
            longest_count = n
            longest_block = block

    return {
        "sections": sections,
        "unique_words": unique,
        "total_alpha_words": total_alpha,
        "richness_pct": richness,
        "longest_paragraph_words": longest_count,
        "longest_paragraph_preview": (
            longest_block.text.strip()[:120] + "…"
            if longest_block and len(longest_block.text.strip()) > 120
            else (longest_block.text.strip() if longest_block else "")
        ),
    }


# ---------------------------------------------------------------------------
# HTML preview — fully coloured, with interactivity
# ---------------------------------------------------------------------------

_HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>QCAA word count preview — {filename}</title>
<style>
:root {{
  --green:#1a7f37; --green-bg:#e8f5ec;
  --red:#cf222e; --red-bg:#fdeeee;
  --magenta:#8250df; --magenta-bg:#f3eefc;
  --yellow:#9a6700; --yellow-bg:#fff8e5;
  --grey:#6e7781; --grey-bg:#f0f2f4;
  --cyan:#0969da; --cyan-bg:#eaf2fd;
  --bg:#ffffff; --card:#f6f8fa;
  --border:#d0d7de; --text:#1f2328;
  --shadow: 0 1px 3px rgba(0,0,0,.06);
}}
body.dark {{
  --green:#4ade80; --green-bg:#12321e;
  --red:#f87171; --red-bg:#3a1414;
  --magenta:#c084fc; --magenta-bg:#2a1845;
  --yellow:#fbbf24; --yellow-bg:#3a2d0a;
  --grey:#9ca3af; --grey-bg:#1f2227;
  --cyan:#60a5fa; --cyan-bg:#0c1e3a;
  --bg:#0d1117; --card:#161b22;
  --border:#30363d; --text:#e6edf3;
  --shadow: 0 1px 3px rgba(0,0,0,.4);
}}
* {{ box-sizing: border-box; }}
body {{
  font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,
              Helvetica,Arial,sans-serif;
  max-width:1280px; margin:2rem auto; padding:0 1.5rem;
  color:var(--text); line-height:1.6; background:var(--bg);
  transition: background .2s, color .2s;
}}
h1 {{ font-size:1.7rem; margin:0 0 .25rem 0; }}
.subtitle {{ color:var(--grey); margin-bottom:1.5rem; font-size:.95rem; }}

.toolbar {{
  display:flex; flex-wrap:wrap; gap:.5rem;
  margin-bottom:1.5rem;
}}
.toolbar button {{
  background:var(--card); border:1px solid var(--border);
  color:var(--text); font-family:inherit; font-size:.85rem;
  padding:.4rem .85rem; border-radius:6px; cursor:pointer;
  transition: background .15s;
}}
.toolbar button:hover {{ background:var(--cyan-bg); border-color:var(--cyan); color:var(--cyan); }}
.toolbar button.copied {{ background:var(--green-bg); border-color:var(--green); color:var(--green); }}

.legend {{
  display:flex; flex-wrap:wrap; gap:.5rem; margin-bottom:1.5rem;
  font-size:.85rem;
}}
.legend span {{
  display:inline-flex; align-items:center; gap:.4rem;
  padding:.25rem .6rem; border-radius:999px;
  background:var(--card); border:1px solid var(--border);
  font-weight:500;
}}
.legend .count {{ background:var(--green-bg); border-color:var(--green); color:var(--green); }}
.legend .exclude {{ background:var(--red-bg); border-color:var(--red); color:var(--red); }}
.legend .citation {{ background:var(--magenta-bg); border-color:var(--magenta); color:var(--magenta); }}
.legend .flag {{ background:var(--yellow-bg); border-color:var(--yellow); color:var(--yellow); }}
.legend .region-ex {{ background:var(--grey-bg); border-color:var(--grey); color:var(--grey); }}
.legend .heading {{ background:var(--cyan-bg); border-color:var(--cyan); color:var(--cyan); }}

.panel {{
  background:var(--card); border:1px solid var(--border);
  border-radius:8px; padding:1.25rem 1.5rem; margin-bottom:1.5rem;
  box-shadow:var(--shadow);
}}
.panel h2 {{
  font-size:1.15rem; margin:0 0 .75rem 0;
  display:flex; align-items:baseline; gap:.75rem;
}}
.panel h2 .badge {{
  font-size:.75rem; padding:.15rem .5rem; border-radius:999px;
  font-weight:600; letter-spacing:.02em;
}}
.badge-good {{ background:var(--green-bg); color:var(--green); }}
.badge-bad  {{ background:var(--red-bg); color:var(--red); }}
.badge-info {{ background:var(--cyan-bg); color:var(--cyan); }}
.badge-warn {{ background:var(--yellow-bg); color:var(--yellow); }}

.doc-body {{
  font-family:"Source Serif Pro",Georgia,"Times New Roman",serif;
  font-size:1.05rem; word-wrap:break-word;
}}
.doc-body p {{ margin:0 0 .7rem 0; }}
.doc-body .heading {{
  color:var(--cyan); font-weight:700; font-size:1.18rem;
  margin:1.4rem 0 .5rem 0;
  padding-left:.5rem; border-left:4px solid var(--cyan);
}}
.doc-body .heading.h1 {{ font-size:1.5rem; }}
.doc-body .heading.h2 {{ font-size:1.28rem; }}
.doc-body .heading.h3 {{ font-size:1.16rem; }}

.count     {{ color:var(--green); cursor:pointer; }}
.exclude   {{ color:var(--red); cursor:pointer;
             text-decoration:line-through;
             text-decoration-color:rgba(207,34,46,.35);
             background:var(--red-bg); border-radius:2px; padding:0 1px; }}
.citation  {{ color:var(--magenta); cursor:pointer;
             text-decoration:line-through;
             text-decoration-color:rgba(130,80,223,.35);
             background:var(--magenta-bg); border-radius:2px; padding:0 1px; }}
.flag      {{ color:var(--yellow); cursor:pointer;
             background:var(--yellow-bg);
             border-bottom:1px dashed var(--yellow); }}
.region-ex {{ color:var(--grey); font-style:italic;
             background:var(--grey-bg); padding:2px 4px; border-radius:2px; }}
.unit      {{ color:var(--green); border-bottom:1px dotted var(--green); cursor:pointer; }}

.count:hover, .exclude:hover, .citation:hover, .flag:hover, .unit:hover {{
  outline:2px solid currentColor; outline-offset:1px; border-radius:2px;
}}

table {{ border-collapse:collapse; margin:.75rem 0 1.25rem 0;
        width:100%; font-size:.92rem; }}
td, th {{ border:1px solid var(--border); padding:.45rem .7rem;
         vertical-align:top; white-space:pre-wrap; }}
th {{ background:var(--cyan-bg); color:var(--cyan); text-align:left; }}
tr:nth-child(even) td {{ background:var(--card); }}

.caption {{
  color:var(--grey); font-size:.85rem; font-style:italic;
  margin:.15rem 0 .75rem 0;
}}

.stats-grid {{ display:grid; grid-template-columns:max-content 1fr;
              gap:.35rem 1rem; font-size:.95rem; }}
.stats-grid dt {{ color:var(--red); font-weight:500; }}
.stats-grid dd {{ margin:0; }}
.stats-grid dt.total-row {{ color:var(--green); font-weight:700; }}
.stats-grid dd.total-row {{ color:var(--green); font-weight:700; font-size:1.1rem; }}

details {{ margin-top:.75rem; }}
summary {{ cursor:pointer; color:var(--cyan); font-weight:500; }}
.final-count {{
  font-size:2.6rem; font-weight:700; color:var(--green);
  margin:0 0 .5rem 0; line-height:1;
}}
.final-count small {{ font-size:.9rem; color:var(--grey);
  font-weight:400; margin-left:.5rem; }}

.table-chosen-count {{ color:var(--green); font-weight:600; }}
.table-chosen-exclude {{ color:var(--red); font-weight:600; }}

.stat-cards {{
  display:grid; grid-template-columns:repeat(auto-fit, minmax(170px, 1fr));
  gap:.75rem; margin-bottom:1rem;
}}
.stat-card {{
  background:var(--bg); border:1px solid var(--border);
  border-radius:6px; padding:.75rem 1rem;
}}
.stat-card .label {{
  font-size:.72rem; text-transform:uppercase; letter-spacing:.05em;
  color:var(--grey); font-weight:600;
}}
.stat-card .value {{
  font-size:1.5rem; font-weight:700; color:var(--cyan); line-height:1.2;
}}
.stat-card .sub {{
  font-size:.8rem; color:var(--grey); margin-top:.15rem;
}}

.section-bar {{
  display:flex; height:24px; border-radius:4px; overflow:hidden;
  background:var(--grey-bg); margin-bottom:1rem;
}}
.section-bar > div {{
  height:100%; display:flex; align-items:center; justify-content:center;
  color:#fff; font-size:.7rem; font-weight:600; white-space:nowrap;
  overflow:hidden;
}}

.section-list {{ list-style:none; padding:0; margin:0; }}
.section-list li {{
  display:flex; justify-content:space-between;
  padding:.35rem .5rem; border-bottom:1px solid var(--border);
  font-size:.9rem;
}}
.section-list li:last-child {{ border-bottom:none; }}
.section-list .name {{ color:var(--text); }}
.section-list .count {{ color:var(--green); font-weight:600; }}
.section-list .pct {{ color:var(--grey); font-size:.8rem; margin-left:.4rem; }}

.span-tooltip {{
  position:absolute; z-index:1000;
  background:var(--card); border:1px solid var(--border);
  border-radius:8px; padding:.85rem 1rem;
  box-shadow:0 4px 12px rgba(0,0,0,.25);
  max-width:340px; font-size:.85rem;
  display:none;
}}
.span-tooltip.visible {{ display:block; }}
.span-tooltip .st-text {{
  font-weight:600; margin-bottom:.4rem;
  padding:.25rem .4rem; border-radius:3px;
  word-break:break-all;
}}
.span-tooltip .st-row {{
  display:grid; grid-template-columns:max-content 1fr;
  gap:.15rem .75rem; margin-top:.35rem;
}}
.span-tooltip .st-row dt {{ color:var(--grey); font-weight:500; }}
.span-tooltip .st-row dd {{ margin:0; word-break:break-word; }}
.span-tooltip .st-close {{
  float:right; cursor:pointer; color:var(--grey);
  font-size:1.1rem; line-height:1; padding:0 .25rem;
}}
.span-tooltip .st-close:hover {{ color:var(--red); }}
</style></head><body>

<h1>QCAA word count preview</h1>
<div class="subtitle">{filename} — generated by qcaa_word_count.py v5</div>

<div class="toolbar">
  <button id="theme-btn" type="button">🌙 Dark mode</button>
  <button id="copy-btn" type="button">📋 Copy summary</button>
  <button type="button" onclick="window.print()">🖨️ Print / Save PDF</button>
</div>

<div class="legend">
<span class="count">counted</span>
<span class="exclude">excluded</span>
<span class="citation">citation</span>
<span class="flag">flagged</span>
<span class="region-ex">region-excluded</span>
<span class="heading">heading</span>
</div>

<div class="panel">
  <h2><span class="badge badge-good">FINAL WORD COUNT</span></h2>
  <div class="final-count">{total}<small>words</small></div>
  <dl class="stats-grid">{stats_rows}</dl>
</div>

<div class="panel">
  <h2>Statistics <span class="badge badge-info">writing breakdown</span></h2>
  {statistics_html}
</div>

<div class="panel">
  <h2>Document <span class="badge badge-info">colour-coded preview</span></h2>
  <div class="doc-body">{body_html}</div>
</div>

{additional_panels}

<details>
<summary>Notes on how this was produced</summary>
<p>Green text counts. Red text was excluded (numbers, symbols,
equations). Magenta text was excluded as a citation. Yellow text was
flagged and left undecided. Grey text is from an excluded region.
Cyan text is a heading.</p>
<p><strong>Click any coloured word</strong> to see the rule that was
applied to it.</p>
<p><strong>Citation detection is APA 7 only.</strong> Other styles may
be counted as ordinary words. Use <code>python qcaa_word_count.py "your_file_name.docx" --report json</code> to audit
every decision.</p>
</details>

<div class="span-tooltip" id="tooltip">
  <span class="st-close" onclick="hideTooltip()">×</span>
  <div class="st-text" id="tooltip-text"></div>
  <dl class="st-row">
    <dt>Decision</dt><dd id="tooltip-decision"></dd>
    <dt>Rule</dt><dd id="tooltip-rule"></dd>
    <dt>Reason</dt><dd id="tooltip-reason"></dd>
  </dl>
</div>

<script>
(function() {{
  const themeBtn = document.getElementById('theme-btn');
  const saved = localStorage.getItem('qcaa-theme');
  if (saved === 'dark') {{
    document.body.classList.add('dark');
    themeBtn.textContent = '☀️ Light mode';
  }}
  themeBtn.addEventListener('click', () => {{
    const dark = document.body.classList.toggle('dark');
    themeBtn.textContent = dark ? '☀️ Light mode' : '🌙 Dark mode';
    localStorage.setItem('qcaa-theme', dark ? 'dark' : 'light');
  }});

  const copyBtn = document.getElementById('copy-btn');
  const summary = {summary_json};
  copyBtn.addEventListener('click', async () => {{
    try {{
      await navigator.clipboard.writeText(summary);
      copyBtn.textContent = '✅ Copied!';
      copyBtn.classList.add('copied');
      setTimeout(() => {{
        copyBtn.textContent = '📋 Copy summary';
        copyBtn.classList.remove('copied');
      }}, 1500);
    }} catch (e) {{
      copyBtn.textContent = '❌ Copy failed';
    }}
  }});

  const tooltip = document.getElementById('tooltip');
  const tooltipText = document.getElementById('tooltip-text');
  const tooltipDecision = document.getElementById('tooltip-decision');
  const tooltipRule = document.getElementById('tooltip-rule');
  const tooltipReason = document.getElementById('tooltip-reason');

  document.addEventListener('click', (e) => {{
    const t = e.target;
    if (!t.matches || !t.matches('span.count, span.exclude, span.citation, span.flag, span.unit')) {{
      return;
    }}
    if (!t.hasAttribute('data-rule')) return;

    tooltipText.textContent = '"' + t.textContent + '"';
    tooltipDecision.textContent = t.getAttribute('data-decision') || '(none)';
    tooltipRule.textContent = t.getAttribute('data-rule') || '(none)';
    tooltipReason.textContent = t.getAttribute('data-reason') || '(none)';
    tooltipText.className = 'st-text ' + t.className.split(' ')[0];

    const rect = t.getBoundingClientRect();
    const top = window.scrollY + rect.bottom + 6;
    const left = Math.min(
      window.scrollX + rect.left,
      window.scrollX + window.innerWidth - 360
    );
    tooltip.style.top = top + 'px';
    tooltip.style.left = left + 'px';
    tooltip.classList.add('visible');
    e.stopPropagation();
  }});

  document.addEventListener('click', (e) => {{
    if (!tooltip.contains(e.target)) {{
      tooltip.classList.remove('visible');
    }}
  }}, true);

  window.hideTooltip = function() {{
    tooltip.classList.remove('visible');
  }};
  document.addEventListener('keydown', (e) => {{
    if (e.key === 'Escape') hideTooltip();
  }});
}})();
</script>

</body></html>
"""


def _span_html(span: Span) -> str:
    text = html.escape(span.text)
    dec = span.decision_record.decision
    rule = span.decision_record.rule_id or ""
    reason = span.decision_record.reason or ""
    dec_attr = html.escape(dec.value if hasattr(dec, "value") else str(dec))
    rule_attr = html.escape(rule)
    reason_attr = html.escape(reason)
    meta = (f' data-decision="{dec_attr}" data-rule="{rule_attr}" '
            f'data-reason="{reason_attr}"')

    if dec == Decision.COUNT:
        if span.span_type == SpanType.UNIT:
            return f'<span class="unit"{meta}>{text}</span>'
        return f'<span class="count"{meta}>{text}</span>'
    if dec == Decision.EXCLUDE:
        if rule == "I-CITATION":
            return f'<span class="citation"{meta}>{text}</span>'
        return f'<span class="exclude"{meta}>{text}</span>'
    if dec in (Decision.FLAG, Decision.UNDECIDED):
        return f'<span class="flag"{meta}>{text}</span>'
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
    if result["buckets"]:
        lines.append("")
        lines.append("Excluded:")
        for key in sorted(result["buckets"]):
            lines.append(f"  {key}: {result['buckets'][key]} word(s)")
    return "\n".join(lines)


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

    return _HTML_TEMPLATE.format(
        filename=html.escape(document.filename or ""),
        total=result["total"],
        stats_rows="".join(stats_rows),
        statistics_html=statistics_html,
        body_html=body_html,
        additional_panels=additional,
        summary_json=json.dumps(summary_text),
    )


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
            webbrowser.open(f"file://{os.path.abspath(output_path)}")
        except Exception:
            pass
    return output_path


# ---------------------------------------------------------------------------
# Decision log export
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