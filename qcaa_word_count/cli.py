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
    bold, bold_cyan, bold_green, cyan, dim, green, grey, magenta, red, yellow,
    export_decisions, print_ansi_preview, print_report, write_html_preview,
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


def _banner() -> None:
    print()
    print(bold_cyan("╔══════════════════════════════════════════════════════════════════╗"))
    print(bold_cyan("║                    QCAA WORD COUNT TOOL                          ║"))
    print(bold_cyan("╚══════════════════════════════════════════════════════════════════╝"))
    print()
    print("This tool estimates the QCAA word count of a .docx response.")
    print("It applies the official QCAA rules and lets you review every")
    print("ambiguous decision.")
    print()
    print(bold("How do you want to run it?"))
    print()
    print(f"  {yellow('[1]')} {green('Include everything')}  {dim('(fastest)')}")
    print(f"      Accept every suggestion. Numbers, symbols, equations and")
    print(f"      citations are still excluded automatically. Tables, captions,")
    print(f"      footnotes and appendixes are all counted. Good for a quick")
    print(f"      upper-bound estimate.")
    print()
    print(f"  {yellow('[2]')} {green('Review each decision')}  {dim('(most accurate)')}")
    print(f"      You'll be prompted about every table, caption, footnote and")
    print(f"      appendix so you can make the call for each one. Takes a")
    print(f"      couple of minutes longer but gives the most accurate count.")
    print()
    print(f"  {yellow('[3]')} What are the QCAA rules?")
    print(f"  {yellow('[4]')} Exit.")
    print()
    print(grey("Both modes generate an HTML preview and open it in your browser"))
    print(grey("once the count is done."))
    print()


def _startup_choice() -> str | None:
    _banner()
    while True:
        try:
            raw = input(f"{bold('Your choice')} {yellow('[1]')}: ").strip()
        except (EOFError, KeyboardInterrupt):
            return None
        if not raw:
            return "include_all"
        if raw == "1":
            return "include_all"
        if raw == "2":
            return "interactive"
        if raw == "3":
            print()
            print(QCAA_RULES_TEXT)
            print()
            try:
                input(grey("Press Enter to return to the menu..."))
            except (EOFError, KeyboardInterrupt):
                return None
            _banner()
            continue
        if raw == "4":
            return None
        print(red("Please type a number between 1 and 4."))


def _apply_span_decisions(document) -> None:
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


def _prepare(path: str):
    document = ingest(path)
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

    return document


def _print_summary(document) -> None:
    math_total = sum(b.math_object_count for b in document.blocks)
    line = f"  {len(document.blocks)} paragraph(s), {len(document.tables)} table(s)"
    print(green(line))
    if math_total:
        print(grey(f"  {math_total} math object(s)"))
    if document.drawings:
        print(grey(f"  {document.drawings} drawing(s) "
                   f"({document.drawings_with_text} contain text)"))
    if document.text_boxes:
        print(grey(f"  {len(document.text_boxes)} text box(es)"))


def run(path: str, *, preset: str | None = None,
        open_browser: bool = True,
        export_format: str | None = None,
        export_path: str | None = None,
        html_path: str | None = None) -> int:
    if not os.path.isfile(path):
        print(red(f"Couldn't find a file at: {path}"))
        return 2
    if not path.lower().endswith(".docx"):
        print(red(f"'{path}' doesn't look like a .docx file."))
        return 2

    print()
    print(f"{bold('Loading:')} {cyan(path)}")
    document = _prepare(path)
    _print_summary(document)

    auto_accept = (preset == "include_all")
    if preset is not None:
        label = {
            "include_all": green("Include everything (auto-accept all)"),
            "interactive": green("Interactive review"),
        }.get(preset, green(preset))
        print()
        print(f"{bold('Mode:')} {label}")
        print()
        resolve_all(document, auto_accept=auto_accept)
    else:
        resolve_all(document, auto_accept=False)

    result = count_words(document)

    print()
    print_ansi_preview(document)
    print()

    html_file = write_html_preview(
        document, result, output_path=html_path, open_browser=open_browser,
    )
    print(f"{bold('HTML preview written to:')} {cyan(html_file)}")
    if not open_browser:
        print(grey("  (Use --open to open it in your browser.)"))
    print()

    print_report(document, result)

    if export_format:
        out = export_decisions(document, result, format=export_format)
        if export_path is None:
            base = os.path.splitext(path)[0]
            export_path = f"{base}_qcaa_decisions.{export_format}"
        with open(export_path, "w", encoding="utf-8") as f:
            f.write(out)
        print()
        print(f"{bold('Decision log written to:')} {cyan(export_path)}")

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
    parser.add_argument("--include-all", action="store_true",
                        help="Accept all suggestions (same as preset 1)")
    parser.add_argument("--review", action="store_true",
                        help="Interactive review (same as preset 2)")
    parser.add_argument("--no-open", action="store_true",
                        help="Don't open the HTML preview automatically")
    parser.add_argument("--open", action="store_true",
                        help="Open the HTML preview (default anyway)")
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
            print(red("No file selected."))
            return 1

    if args.include_all:
        preset = "include_all"
    elif args.review:
        preset = "interactive"
    else:
        preset = _startup_choice()
        if preset is None:
            return 0

    open_browser = not args.no_open

    return run(
        path,
        preset=preset,
        open_browser=open_browser,
        export_format=args.report,
        export_path=args.report_path,
        html_path=args.html_path,
    )