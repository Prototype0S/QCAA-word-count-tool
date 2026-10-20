#!/usr/bin/env python3
"""
QCAA Word Count Tool (v5)
=========================

A Python tool that estimates the QCAA word count of a .docx response,
applying the official QCAA word-length rules and letting the user
review every ambiguous decision.

WHAT IT DOES
------------
Parses a Word document, applies QCAA rules to exclude the things that
don't count (title page, references, numbers, symbols, equations,
citations, visual elements), and reports the resulting word count.

Interactive review is built in: tables, captions, footnotes, text boxes
and appendixes are all prompted for. The user can accept every
suggestion with --default for batch use, or override each one.

The tool produces two previews of the same result:
  - an ANSI colour preview in the terminal (fast, glanceable)
  - an HTML rich-text preview in the browser (printable, auditable)

A full decision log is available in JSON or CSV for compliance or
testing.

HOW TO RUN
----------
    python qcaa_word_count.py                    # opens file picker
    python qcaa_word_count.py path/to/file.docx  # runs on a specific file
    python qcaa_word_count.py --explain          # shows QCAA rules
    python qcaa_word_count.py --help             # all options

REQUIRES
--------
    pip install python-docx

SOURCE LAYOUT
-------------
    qcaa_word_count.py            # this file — the entry point
    qcaa_word_count/model.py      # dataclasses and enums
    qcaa_word_count/helpers.py    # tokenization and span classification
    qcaa_word_count/ingest.py     # .docx -> Document
    qcaa_word_count/regions.py    # structure detection
    qcaa_word_count/tables.py     # table classification
    qcaa_word_count/citations.py  # APA 7 detection
    qcaa_word_count/captions.py   # caption detection
    qcaa_word_count/footnotes.py  # footnote extraction
    qcaa_word_count/interactive.py # user prompts
    qcaa_word_count/count.py      # walk the decision map
    qcaa_word_count/output.py     # ANSI + HTML + export
    qcaa_word_count/cli.py        # argparse + run()
    tests/test_pipeline.py        # end-to-end tests

CHANGELOG (this version)
------------------------
    - Splits the 1,500-line monolith into 12 focused modules
    - Removes the half-implemented --genre flag
    - Fixes caption detection to reject only prose-like continuations,
      not all captions containing "shows"/"illustrates"/etc.
    - Removes dead code in looks_like_caption
    - Stops auto-opening the browser; use --open to opt in
    - Adds tests covering each genre and edge case
    - Preserves ANSI preview as the default output
"""

import sys

from qcaa_word_count.cli import main

if __name__ == "__main__":
    sys.exit(main())