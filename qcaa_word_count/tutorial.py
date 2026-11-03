"""First-run tutorial.

Shown once per installation. The marker lives in the user's home
directory (not next to the tool) so it survives moving the folder.
"""
from __future__ import annotations

import platform
from pathlib import Path

from .output import (
    bold, bold_cyan, cyan, green, grey, magenta, red, yellow,
)

_MARKER_DIR = Path.home() / ".qcaa_word_count"
_MARKER_FILE = _MARKER_DIR / "tutorial_seen"

_IS_WINDOWS = platform.system() == "Windows"
_IS_MAC = platform.system() == "Darwin"


def has_seen() -> bool:
    return _MARKER_FILE.exists()


def mark_seen() -> None:
    try:
        _MARKER_DIR.mkdir(parents=True, exist_ok=True)
        _MARKER_FILE.write_text("seen\n", encoding="utf-8")
    except OSError:
        pass


def _launcher_line() -> str:
    if _IS_WINDOWS:
        return f"double-click {cyan('install-windows.bat')}"
    if _IS_MAC:
        return f"double-click {cyan('install-mac.command')}"
    return f"run {cyan('python qcaa_word_count.py')}"


def _mac_first_run_note() -> str:
    if not _IS_MAC:
        return ""
    return f"""
  {yellow('First-run macOS warning:')} because this was downloaded, macOS
  will block the file the first time. Either:
    {bold('1.')} Double-click, click {cyan('Done')} (not "Move to Trash"),
       then go to {cyan('System Settings → Privacy & Security')} and click
       {cyan('Open Anyway')} next to this file.
    {bold('2.')} Or in Terminal: {cyan('xattr -cr install-mac.command')}
  After that, double-click works normally.
"""


TUTORIAL_TEXT = f"""
{bold_cyan('╔══════════════════════════════════════════════════════════════════╗')}
{bold_cyan('║                    WELCOME TO THE QCAA WORD COUNT TOOL           ║')}
{bold_cyan('╚══════════════════════════════════════════════════════════════════╝')}

This is a quick tour. It only appears once — after this, the tool
goes straight to the startup menu. You can re-run it any time with
{yellow('--tutorial')}.

{grey('─' * 66)}
{bold('WHAT IT DOES')}
{grey('─' * 66)}

Word's word count treats every visible word the same. QCAA doesn't.
This tool applies the QCAA rules — excluding title pages, references,
numbers, symbols, equations, citations, and visual elements — and
asks you about the ambiguous cases (tables, captions, footnotes).

Every word in the preview is colour-coded:

  {green('green')}     counted
  {red('red')}       excluded (numbers, symbols, equations)
  {magenta('magenta')}   excluded (citations)
  {yellow('yellow')}    flagged for your review
  {grey('grey')}      excluded region (title page, references, etc.)
  {cyan('cyan')}      heading

{grey('─' * 66)}
{bold('HOW TO LAUNCH IT')}
{grey('─' * 66)}

  {_launcher_line()}

  The launcher finds Python, installs the one dependency (python-docx),
  and starts the tool. On every later run it skips straight to the
  menu — no setup messages, no waiting.
{_mac_first_run_note()}
{grey('─' * 66)}
{bold('TWO WAYS TO RUN')}
{grey('─' * 66)}

  {yellow('[1]')} {green('Include everything')}  — fastest, accepts every suggestion.
      Still excludes numbers, symbols, equations, citations.
  {yellow('[2]')} {green('Review each decision')} — prompts for every ambiguous case.
      Slower but most accurate. Recommended for your first real doc.

Both modes generate an {cyan('HTML preview')} and open it in your browser.

{grey('─' * 66)}
{bold('THE HTML PREVIEW')}
{grey('─' * 66)}

  • {bold('Click')} any coloured word to see the rule behind it
  • {bold('Ctrl+click')} (⌘+click on Mac) toggles one word — click again
    to put it back
  • {bold('Select mode')} (press {yellow('S')}) lets you click multiple words
    then Count/Exclude them all at once
  • {bold('Ctrl+Z')} / {bold('Ctrl+Y')} to undo / redo any change
  • {bold('Low-confidence panel')} lists every word the tool wasn't sure
    about (below 70%), with Jump buttons to scroll to them
  • {bold('Browse all words')} shows every word if you want to audit
  • {bold('Dark mode')}, {bold('Copy summary')}, {bold('Print / Save PDF')}

{grey('─' * 66)}
{bold('SAVED DECISIONS')}
{grey('─' * 66)}

After your first run, decisions are saved next to the .docx as
{cyan('<filename>_qcaa_decisions.json')}. Next time you run the tool on
the same file, those decisions are restored automatically — you
won't be asked the same questions twice.

To start fresh, delete that JSON file, or pass {yellow('--reset')}.

{grey('─' * 66)}
{bold('TIPS FOR YOUR FIRST RUN')}
{grey('─' * 66)}

  1. Try {yellow('[2]')} {green('Review each decision')} once. It walks you through
     every ambiguous case, so you learn what the tool thinks about.
  2. Open the HTML preview and click through a few coloured words.
     The tooltip explains why each one got its decision.
  3. If your document has tables, use the "count whole / exclude whole /
     headers only" options carefully — this is where most counts differ.
  4. When you're happy, run again with {yellow('--check')} after edits to
     see whether your changes added or removed words.

{grey('─' * 66)}
{bold('GETTING HELP')}
{grey('─' * 66)}

  {yellow('--tutorial')}   Show this tour again
  {yellow('--explain')}    Print the full QCAA rules
  {yellow('--help')}       Show all command-line options

{grey('=' * 66)}
"""


def run_tutorial_if_first_time(force: bool = False) -> None:
    """Show the tutorial if it's the user's first time, or if forced."""
    if not force and has_seen():
        return

    # Box-drawing characters need UTF-8 on Windows cmd.exe.
    try:
        import sys
        if hasattr(sys.stdout, "reconfigure"):
            sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    print(TUTORIAL_TEXT)
    try:
        input(grey("Press Enter to continue to the startup menu..."))
    except (EOFError, KeyboardInterrupt):
        pass
    print()
    mark_seen()