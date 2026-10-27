"""All output: ANSI preview, HTML preview, decision log, plain report.

Re-exports everything that used to live in output.py, so existing
`from .output import ...` imports keep working.

Directory layout:
    ansi.py    — colour helpers, print_ansi_preview
    report.py  — print_report
    stats.py   — compute_statistics
    export.py  — export_decisions
    html.py    — HTML template, build_html_preview, write_html_preview
"""
from __future__ import annotations

# Colour helpers — every module in the package relies on these
from .ansi import (
    _c, _COLOURS_ENABLED,
    green, red, yellow, cyan, grey, magenta, bold, dim,
    bold_green, bold_cyan,
    span_to_ansi, print_ansi_preview,
)
from .report import print_report
from .stats import compute_statistics
from .export import export_decisions
from .html import build_html_preview, write_html_preview

__all__ = [
    "green", "red", "yellow", "cyan", "grey", "magenta",
    "bold", "dim", "bold_green", "bold_cyan",
    "span_to_ansi", "print_ansi_preview", "print_report",
    "compute_statistics", "export_decisions",
    "build_html_preview", "write_html_preview",
]