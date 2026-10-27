"""Command-line entry point."""
from __future__ import annotations

import argparse
import os
import sys

from .interactive import resolve_all
from .output import bold, bold_cyan, cyan, dim, green, grey, red, yellow
from .pipeline import run
from .tutorial import run_tutorial_if_first_time


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

NOTE ON UNITS:
  "5 kg" counts one word (the unit); "5kg" counts zero (number+unit
  is treated as one number token). Both are the literal reading of
  the QCAA rule, not a bug.
"""


def _banner() -> None:
    print()
    print(bold_cyan("╔══════════════════════════════════════════════════════════════════╗"))
    print(bold_cyan("║                    QCAA WORD COUNT TOOL                          ║"))
    print(bold_cyan("╚══════════════════════════════════════════════════════════════════╝"))
    print()
    print("This tool estimates the QCAA word count of a .docx response.")
    print()
    print(bold("How do you want to run it?"))
    print()
    print(f"  {yellow('[1]')} {green('Include everything')}  {dim('(fastest)')}")
    print(f"      Accept every suggestion. Numbers, symbols, equations and")
    print(f"      citations are still excluded. Tables, captions, footnotes")
    print(f"      and appendixes are all counted.")
    print()
    print(f"  {yellow('[2]')} {green('Review each decision')}  {dim('(most accurate)')}")
    print(f"      Prompted for every ambiguous case. Slower, most accurate.")
    print()
    print(f"  {yellow('[3]')} What are the QCAA rules?")
    print(f"  {yellow('[4]')} {dim('Re-run the tutorial')}")
    print(f"  {yellow('[5]')} Exit.")
    print()


def _startup_choice() -> str | None:
    _banner()
    while True:
        try:
            raw = input(f"{bold('Your choice')} {yellow('[1]')}: ").strip()
        except (EOFError, KeyboardInterrupt):
            return None
        if not raw or raw == "1":
            return "include-all"
        if raw == "2":
            return "review"
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
            run_tutorial_if_first_time(force=True)
            _banner()
            continue
        if raw == "5":
            return None
        print(red("Please type a number between 1 and 5."))


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

    # Preset modes (mutually exclusive)
    g = parser.add_mutually_exclusive_group()
    g.add_argument("--summary", action="store_true",
                   help="Print only the count, then exit (fastest)")
    g.add_argument("--preview-only", action="store_true",
                   help="Generate the HTML preview without prompting "
                        "or saving decisions")
    g.add_argument("--check", action="store_true",
                   help="Compare this run's count to the last saved run")
    g.add_argument("--review", action="store_true",
                   help="Interactive review (same as menu option 2)")
    g.add_argument("--include-all", action="store_true",
                   help="Accept all suggestions (same as menu option 1)")

    # State
    parser.add_argument("--no-save", action="store_true",
                        help="Don't save decisions after this run")
    parser.add_argument("--no-load", action="store_true",
                        help="Ignore any saved decisions file")
    parser.add_argument("--reset", action="store_true",
                        help="Delete saved decisions, then run normally")

    # Output
    parser.add_argument("--open", action="store_true",
                        help="Open the HTML preview in the browser (default)")
    parser.add_argument("--no-open", action="store_true",
                        help="Don't open the HTML preview")
    parser.add_argument("--html-path", help="Path for the HTML preview")
    parser.add_argument("--report", choices=["json", "csv"],
                        help="Export a full decision log")
    parser.add_argument("--report-path", help="Path for the decision log")
    parser.add_argument("--apply-overrides",
                        help="Path to qcaa_overrides.json from the HTML preview")
    # Info / meta
    parser.add_argument("--explain", action="store_true",
                        help="Print the QCAA rules and exit")
    parser.add_argument("--tutorial", action="store_true",
                        help="Re-run the first-run tutorial")
    args = parser.parse_args(argv)

    if args.explain:
        print(QCAA_RULES_TEXT)
        return 0

    if args.tutorial:
        run_tutorial_if_first_time(force=True)
        return 0
    if args.apply_overrides:
        try:
            with open(args.apply_overrides, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except (OSError, json.JSONDecodeError) as e:
            print(red(f"Couldn't read overrides file: {e}"))
            return 2
        overrides = payload.get("overrides", {})
        if overrides:
            print(grey(f"  Applying {len(overrides)} override(s) from "
                       f"{args.apply_overrides}"))
            # Defer to run() — pass through

    # First-run tutorial (only if no preset flag was given)
    preset_given = any([args.summary, args.preview_only, args.check,
                        args.review, args.include_all])
    if not preset_given:
        run_tutorial_if_first_time(force=False)

    if args.file:
        path = args.file
    else:
        path = _pick_file() or ""
        if not path:
            print(red("No file selected."))
            return 1

    # [FIX 1] Exit correctly when the user picks option 5.
    if args.summary:
        mode = "summary"
    elif args.preview_only:
        mode = "preview"
    elif args.check:
        mode = "check"
    elif args.review:
        mode = "review"
    elif args.include_all:
        mode = "include-all"
    else:
        chosen = _startup_choice()
        if chosen is None:
            return 0
        mode = chosen

    # [FIX 2] --open was previously a no-op. --no-open still wins if both given.
    if args.no_open:
        open_browser = False
    else:
        open_browser = True

    return run(
        path,
        mode=mode,
        auto_accept=False,
        open_browser=open_browser,
        save=not args.no_save,
        load=not args.no_load,
        reset=args.reset,
        show_diff=args.check,
        export_format=args.report,
        export_path=args.report_path,
        html_path=args.html_path,
        overrides=overrides if 'overrides' in locals() else None,
    )