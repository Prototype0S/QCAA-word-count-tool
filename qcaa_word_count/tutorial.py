"""First-run tutorial.

Shown once per installation. Marker is stored in a user-home directory,
not next to the tool, so it survives moving the folder.
"""
from __future__ import annotations

import os
from pathlib import Path

from .output import bold, bold_cyan, cyan, dim, green, grey, magenta, red, yellow

_MARKER_DIR = Path.home() / ".qcaa_word_count"
_MARKER_FILE = _MARKER_DIR / "tutorial_seen"


def has_seen() -> bool:
    return _MARKER_FILE.exists()


def mark_seen() -> None:
    try:
        _MARKER_DIR.mkdir(parents=True, exist_ok=True)
        _MARKER_FILE.write_text("seen\n", encoding="utf-8")
    except OSError:
        # If we can't write, we just show the tutorial again. Not fatal.
        pass


TUTORIAL_TEXT = f"""
{bold_cyan('╔══════════════════════════════════════════════════════════════════╗')}
{bold_cyan('║                    WELCOME TO THE QCAA WORD COUNT TOOL           ║')}
{bold_cyan('╚══════════════════════════════════════════════════════════════════╝')}

This is a quick tour. It only appears once — after this, the tool
goes straight to the startup menu.

{grey('─' * 66)}
{bold('WHAT IT DOES')}
{grey('─' * 66)}

Word's word count treats every visible word the same. QCAA doesn't.
This tool applies the QCAA rules — excluding title pages, references,
numbers, symbols, equations, citations, and visual elements — and
asks you about the ambiguous cases (tables, captions, footnotes).

Every word is colour-coded:

  {green('green')}     counted
  {red('red')}       excluded (numbers, symbols, equations)
  {magenta('magenta')}   excluded (citations)
  {yellow('yellow')}    flagged for your review
  {grey('grey')}      excluded region (title page, references, etc.)
  {cyan('cyan')}      heading

{grey('─' * 66)}
{bold('HOW TO RUN IT')}
{grey('─' * 66)}

  {yellow('[1]')} {green('Include everything')}  — fastest, counts tables/captions/footnotes,
      still excludes numbers, symbols, equations, citations.
  {yellow('[2]')} {green('Review each decision')} — slower, prompts for every ambiguous case.
      Most accurate. Recommended for your first run on a real document.

Both modes generate an {cyan('HTML preview')} and open it in your browser.

{grey('─' * 66)}
{bold('WHAT THE HTML PREVIEW DOES')}
{grey('─' * 66)}

  • Colour-codes the whole document, same scheme as the terminal
  • {bold('Click any coloured word')} to see the rule that was applied
  • {bold('Dark mode')} toggle (remembers your choice)
  • {bold('Copy summary')} — puts the count and section breakdown on your clipboard
  • {bold('Print / Save PDF')} — for attaching to your assignment
  • Per-section word counts, so you can see which sections are longest

{grey('─' * 66)}
{bold('SAVED DECISIONS')}
{grey('─' * 66)}

After your first run, the tool saves your decisions next to the .docx
as {cyan('<filename>_qcaa_decisions.json')}. Next time you run the tool
on the same file, those decisions are restored automatically — you
won't be asked the same questions twice.

If you want to start fresh, delete that JSON file, or pass {yellow('--reset')}.

{grey('─' * 66)}
{bold('COMMAND-LINE FLAGS')}
{grey('─' * 66)}

  {yellow('--summary')}        Just print the count, then exit. Fastest.
  {yellow('--preview-only')}   Show the HTML preview without committing to decisions.
  {yellow('--check')}          Compare this run's count to the last saved run.
  {yellow('--review')}         Interactive review (same as menu option 2).
  {yellow('--include-all')}    Accept all suggestions (same as menu option 1).
  {yellow('--no-save')}        Don't save decisions this run.
  {yellow('--no-load')}        Ignore any saved decisions file.
  {yellow('--reset')}          Delete the saved decisions file, then run normally.
  {yellow('--tutorial')}       Show this tutorial again.
  {yellow('--explain')}        Print the QCAA rules and exit.
  {yellow('--help')}           Show all options.

{grey('─' * 66)}
{bold('TIPS FOR YOUR FIRST RUN')}
{grey('─' * 66)}

  1. Try {yellow('[2]')} {green('Review each decision')} once. It walks you through every
     ambiguous case, which teaches you what the tool thinks about.
  2. Open the HTML preview and click through a few coloured words.
     The tooltip explains why each one got its decision.
  3. If your document has tables, use the "count whole / exclude whole /
     headers only" options carefully — this is where most counts differ.
  4. When you're happy with the count, run again with {yellow('--check')}
     after edits to see whether your changes added or removed words.

{grey('─' * 66)}
{bold('GETTING HELP')}
{grey('─' * 66)}

  {yellow('--tutorial')}   Show this tour again
  {yellow('--explain')}    Print the full QCAA rules
  {yellow('--help')}       Show all command-line options

If you're stuck, the README on GitHub has a troubleshooting section.

{grey('=' * 66)}
"""


def run_tutorial_if_first_time(force: bool = False) -> None:
    """Show the tutorial if it's the user's first time, or if forced."""
    if not force and has_seen():
        return
    print(TUTORIAL_TEXT)
    try:
        input(grey("Press Enter to continue to the startup menu..."))
    except (EOFError, KeyboardInterrupt):
        pass
    print()
    mark_seen()