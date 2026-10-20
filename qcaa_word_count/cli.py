"""Command-line entry point."""
from __future__ import annotations

import argparse
import os
import sys

from .captions import detect_captions
from .citations import apply_citation_exclusions, collect_citations
from .count import count_words
from .footnotes import apply_footnote_suggestions, read_footnotes_from_docx
from .helpers import classify_span, is_equation_heavy_block
from .ingest import ingest
from .interactive import resolve_all
from .model import Decision, DecisionRecord
from .output import (
    export_decisions,
    print_ansi_preview,
    print_report,
    write_html_preview,
)
from .regions import apply_regional_decisions, detect_regions
from .tables import classify_table


QCAA_RULES_TEXT = """\
QCAA WORD-LENGTH RULES
----------------------
INCLUDED:
  - words in the text of the response
  - title, headings and subheadings
  - tables/figures/diagrams containing information other than raw
    or processed data (the whole item counts)
  - quotations
  - footnotes/endnotes unless purely bibliographical
  - abbreviations, initialisms, units, chemical formulas

EXCLUDED:
  - title page, contents, abstract
  - visual elements (by-lines, banners, captions, call-outs)
  - raw or processed data in tables, figures, diagrams
  - numbers, symbols, equations, calculations
  - bibliography / reference list
  - appendixes (only if supplementary-only)
  - page numbers
  - in-text citations
  - blank pages
"""


def _apply_span_decisions(document) -> None:
    """Assign span-level decisions for all body blocks."""
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            for span in block.spans:
                span.decision_record = DecisionRecord(
                    decision=Decision.EXCLUDE,
                    rule_id="R-REGION",
                    confidence=1.0,
                    reason=f"Inherits exclusion from region '{block.region.value}'",
                    source="auto",
                )
            continue
        equation_heavy = is_equation_heavy_block(block)
        for span in block.spans:
            span_type, rec = classify_span(
                span.text, equation_heavy=equation_heavy,
            )
            span.span_type = span_type
            span.decision_record = rec


def run(path: str, *, auto_accept: bool = False,
        open_browser: bool = False,
        export_format: str | None = None,
        export_path: str | None = None,
        html_path: str | None = None) -> int:
    if not os.path.isfile(path):
        print(f"Couldn't find a file at: {path}")
        return 2
    if not path.lower().endswith(".docx"):
        print(f"'{path}' doesn't look like a .docx file.")
        return 2

    print()
    print("Colour guide:")
    print("  green   = counted")
    print("  red     = excluded (numbers, symbols, equations)")
    print("  magenta = excluded (citations)")
    print("  grey    = excluded region")
    print("  yellow  = flagged")
    print("  cyan    = headings")
    print()

    print(f"Loading: {path}")
    document = ingest(path)
    print(f"  {len(document.blocks)} paragraph(s), "
          f"{len(document.tables)} table(s)")
    math_total = sum(b.math_object_count for b in document.blocks)
    if math_total:
        print(f"  {math_total} math object(s)")
    if document.drawings:
        print(f"  {document.drawings} drawing(s) "
              f"({document.drawings_with_text} contain text)")
    if document.text_boxes:
        print(f"  {len(document.text_boxes)} text box(es)")

    for table in document.tables:
        table.suggestion, table.suggestion_confidence = classify_table(table)

    detect_regions(document)
    apply_regional_decisions(document)
    _apply_span_decisions(document)
    detect_captions(document)

    document.footnotes = read_footnotes_from_docx(path)
    apply_footnote_suggestions(document)
    document.citations = collect_citations(document)
    apply_citation_exclusions(document)

    resolve_all(document, auto_accept=auto_accept)
    result = count_words(document)

    print()
    print_ansi_preview(document)
    print()

    html_file = write_html_preview(
        document, result, output_path=html_path, open_browser=open_browser,
    )
    print(f"HTML preview written to: {html_file}")
    if not open_browser:
        print("  (Use --open to open it in your browser.)")
    print()

    print_report(document, result)

    if export_format:
        out = export_decisions(document, result, format=export_format)
        if export_path is None:
            base = os.path.splitext(path)[0]
            export_path = f"{base}_qcaa_decisions.{export_format}"
        with open(export_path, "w", encoding="utf-8") as f:
            f.write(out)
        print(f"\nDecision log written to: {export_path}")

    return 0


def _pick_file() -> str | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    path = filedialog.askopenfilename(
        title="Select your Word document",
        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
    )
    root.destroy()
    return path or None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="qcaa_word_count.py",
        description="Estimate the QCAA word count of a .docx response.",
    )
    parser.add_argument("file", nargs="?",
                        help=".docx file (omit to use file picker)")
    parser.add_argument("--explain", action="store_true",
                        help="print the QCAA rules and exit")
    parser.add_argument("--default", action="store_true", dest="auto_accept",
                        help="Accept all suggestions without prompting")
    parser.add_argument("--open", action="store_true", dest="open_browser",
                        help="Open the HTML preview in your browser")
    parser.add_argument("--html-path",
                        help="Path for the HTML preview")
    parser.add_argument("--report", choices=["json", "csv"],
                        help="Export a decision log")
    parser.add_argument("--report-path",
                        help="Path for the decision log")
    args = parser.parse_args(argv)

    if args.explain:
        print(QCAA_RULES_TEXT)
        return 0

    if args.file:
        path = args.file
    else:
        path = _pick_file() or ""
        if not path:
            print("No file selected.")
            return 1

    return run(
        path,
        auto_accept=args.auto_accept,
        open_browser=args.open_browser,
        export_format=args.report,
        export_path=args.report_path,
        html_path=args.html_path,
    )