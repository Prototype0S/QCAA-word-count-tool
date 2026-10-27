"""Interactive prompts for tables, captions, footnotes, citations."""
from __future__ import annotations

import sys
from typing import Callable

from .citations import (
    apply_citation_exclusions,
    count_occurrences_in_body,
    find_missed_candidates,
)
from .model import (
    Citation, Decision, DecisionRecord, Document, Region, TableAnswer,
)
from .output import bold, cyan, green, grey, magenta, red, yellow
from .tables import (
    apply_table_answer, suggest_table_answer, table_has_strong_default,
)

# [FIX 5] Resolver prompt injection — every resolver accepts a `prompt`
# callable. Default is `_menu`. Tests can pass a lambda to skip the
# terminal entirely.
PromptFn = Callable[[str, list[tuple[str, str]], str | None], str]



QCAA_TABLE_HELP = f"""
{grey('-' * 66)}
{cyan(bold('TABLES (QCAA rule)'))}
{grey('-' * 66)}
  Option 1 - {green('Count the whole table')} (contains information)
  Option 2 - {red('Exclude the whole table')} (raw data or calculations)
  Option 3 - {yellow('Count the header row only')}, exclude the data rows

{grey('Note: even when a table is "counted whole", numbers, symbols and')}
{grey('equations inside cells are still excluded by the span-level rules.')}
"""

QCAA_CAPTION_CAVEAT = f"""
{grey('=' * 66)}
{yellow(bold('HEADS UP: CAPTIONS ARE A CARVE-OUT'))}
{grey('=' * 66)}

The general rule is {green('"a word is a word"')} - if it's a word, it counts.

But QCAA explicitly excludes {red('visual elements')}:
by-lines, banners, captions and call-outs. Captions are the one place
where the general rule doesn't apply cleanly.

{grey('Common cases:')}
  {red('"Figure 1: Reaction rate vs. temperature"')}
      -> caption - excluded by QCAA

  {green('"Figure 1 shows the relationship between temperature and rate."')}
      -> body text - counted

{grey('But:')} in science reports and PSMTs, captions often identify
variables, conditions, or what a graph shows. Some markers count these
as information. If your school treats captions that way, override the
default - press {yellow('[1]')} to count.

{grey('When in doubt, ask your teacher.')}
{grey('=' * 66)}
"""

QCAA_CAPTION_HELP = f"""
{grey('-' * 66)}
{cyan(bold('CAPTIONS - quick reminder'))}
{grey('-' * 66)}
Captions are a carve-out from "a word is a word": QCAA excludes
{red('visual elements')} including captions. But some markers count
captions that identify variables or conditions.

  {green('Count')}   - if the caption carries information
  {red('Exclude')} - if it's a pure label

Full explanation and examples were printed before the first caption
prompt. {grey('When in doubt, ask your teacher.')}
{grey('-' * 66)}
"""


def _menu(question: str, options: list[tuple[str, str]],
          default_key: str | None = None) -> str:
    print()
    print(cyan(bold(f"== {question} ==")))
    for i, (key, label) in enumerate(options, start=1):
        marker = "  " + yellow("(suggested)") if key == default_key else ""
        print(f"  {yellow(f'[{i}]')} {label}{marker}")
    while True:
        default_hint = ""
        if default_key is not None:
            for i, (key, _) in enumerate(options, start=1):
                if key == default_key:
                    default_hint = f" {yellow('[' + str(i) + ']')}"
                    break
        try:
            raw = input(f"\n{bold('Your choice')}{default_hint}: ").strip()
        except EOFError:
            # EOF (piped input ended) = use default if available
            if default_key is not None:
                return default_key
            print(red("\nNo input."))
            sys.exit(1)
        except KeyboardInterrupt:
            # [FIX 6] Ctrl+C aborts the program, not just re-prompts
            print()
            print(grey("Aborted."))
            sys.exit(130)
        if not raw and default_key is not None:
            return default_key
        if not raw.isdigit():
            print(red("Please type a number."))
            continue
        idx = int(raw) - 1
        if 0 <= idx < len(options):
            return options[idx][0]
        print(red(f"Please type a number between 1 and {len(options)}."))


def _panel(title: str, subtitle: str = "") -> None:
    print()
    print(grey("-" * 66))
    line = cyan(bold(title))
    if subtitle:
        line += "  " + grey(subtitle)
    print(line)


def preview_table(table, max_rows: int = 4) -> str:
    lines = []
    for r_idx, row in enumerate(table.rows):
        if r_idx >= max_rows:
            lines.append(grey(f"... ({len(table.rows) - max_rows} more row(s))"))
            break
        cells = []
        for cell in row:
            t = cell.text.strip().replace("\n", " ")
            if len(t) > 20:
                t = t[:17] + "..."
            cells.append(t)
        lines.append(" | ".join(cells))
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Appendixes
# ---------------------------------------------------------------------------

def resolve_appendixes(document: Document, auto_accept: bool) -> bool:
    appendix_blocks = [
        b for b in document.blocks
        if b.region == Region.APPENDIX
        and b.region_record.decision == Decision.FLAG
    ]
    if not appendix_blocks:
        return auto_accept

    if auto_accept:
        for b in appendix_blocks:
            b.region_record.as_manual_override(
                Decision.COUNT, reason="Auto-accepted (--default)",
            )
        print(green("  Appendix(es) auto-counted."))
        return True

    _panel("Appendix region detected",
           f"{len(appendix_blocks)} block(s)")
    print(f'  Starts at: {grey(appendix_blocks[0].text[:80])}')
    choice = _menu(
        "How should this appendix be treated?",
        [
            ("exclude", f"{red('Exclude')} this appendix"),
            ("include", f"{green('Include')} this appendix"),
            ("all",     "Accept suggestions for all remaining prompts"),
        ],
        default_key="exclude",
    )
    if choice == "all":
        for b in appendix_blocks:
            b.region_record.as_manual_override(
                Decision.COUNT, reason="Auto-accepted (user chose Accept All)",
            )
        return True
    dec = Decision.EXCLUDE if choice == "exclude" else Decision.COUNT
    reason = ("User chose to exclude appendix" if choice == "exclude"
              else "User chose to include appendix")
    for b in appendix_blocks:
        b.region_record.as_manual_override(dec, reason=reason)
    return auto_accept


# ---------------------------------------------------------------------------
# Text boxes
# ---------------------------------------------------------------------------

def resolve_text_boxes(document: Document, auto_accept: bool) -> bool:
    boxes = [
        b for b in document.blocks
        if b.block_type.value == "text_box"
        and b.region_record.decision in (Decision.FLAG, Decision.UNDECIDED)
    ]
    if not boxes:
        return auto_accept
    if auto_accept:
        for b in boxes:
            b.region_record.as_manual_override(
                Decision.COUNT, reason="Auto-accepted text box",
            )
        print(green(f"  {len(boxes)} text box(es) auto-counted."))
        return True

    _panel(f"{len(boxes)} text box(es) detected")
    for b in boxes[:5]:
        print(f"  - {grey(b.text.strip()[:70])}")
    choice = _menu(
        "How should text box content be treated?",
        [
            ("count",   f"{green('Count')} (treat as prose)"),
            ("exclude", f"{red('Exclude')} (treat as visual)"),
        ],
        default_key="count",
    )
    dec = Decision.COUNT if choice == "count" else Decision.EXCLUDE
    reason = "User counted text boxes" if choice == "count" else "User excluded text boxes"
    for b in boxes:
        b.region_record.as_manual_override(dec, reason=reason)
    return auto_accept


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------

def resolve_tables(document: Document, auto_accept: bool) -> bool:
    if not document.tables:
        return auto_accept

    if auto_accept:
        for table in document.tables:
            if table.answer is not None:
                apply_table_answer(
                    table, table.answer,
                    reason="Reapplied saved decision",
                )
                continue
            table.answer = TableAnswer.COUNT_ALL
            apply_table_answer(
                table, TableAnswer.COUNT_ALL,
                reason="Auto-accepted (include all)",
            )
        print(green(f"  All {len(document.tables)} table(s) auto-counted."))
        return True

    _panel(f"{len(document.tables)} table(s) found.")
    for i, table in enumerate(document.tables, start=1):
        if table.answer is not None:
            apply_table_answer(
                table, table.answer,
                reason="Reapplied saved decision",
            )
            continue

        print()
        print(grey(f"--- Table {i} of {len(document.tables)} ---"))
        conf = table.suggestion_confidence
        if table.suggestion.value == "information":
            sugg = green(table.suggestion.value)
        elif table.suggestion.value in ("raw_data", "calculation"):
            sugg = red(table.suggestion.value)
        else:
            sugg = yellow(table.suggestion.value)
        print(f"Suggested: {sugg} {grey(f'(confidence {conf:.2f})')}")
        print()
        print(preview_table(table))
        print()
        print(f"  {bold(str(table.word_count))} {grey('word(s) in this table.')}")

        default_key = (
            suggest_table_answer(table).value if table_has_strong_default(table)
            else None
        )
        choice = _menu(
            "How should this table be counted?",
            [
                ("count_all",    f"{green('Count the whole table')} (information)"),
                ("exclude_all",  f"{red('Exclude the whole table')} (raw data or calculations)"),
                ("headers_only", f"{yellow('Count the header row only')}, exclude the data rows"),
                ("help",         "Explain the QCAA rule for tables"),
                ("all",          "Accept the tool's suggestions for remaining tables"),
            ],
            default_key=default_key,
        )
        if choice == "help":
            print(QCAA_TABLE_HELP)
            choice = _menu(
                "How should this table be counted?",
                [
                    ("count_all",    "Count the whole table"),
                    ("exclude_all",  "Exclude the whole table"),
                    ("headers_only", "Count the header row only"),
                    ("all",          "Accept the tool's suggestions for remaining tables"),
                ],
                default_key=default_key,
            )
        if choice == "all":
            auto_accept = True
            table.answer = TableAnswer.COUNT_ALL
            apply_table_answer(
                table, TableAnswer.COUNT_ALL,
                reason="Auto-accepted (user chose Accept All)",
            )
            continue
        table.answer = TableAnswer(choice)
        apply_table_answer(table, table.answer, reason=f"User chose '{choice}'")
    return auto_accept


# ---------------------------------------------------------------------------
# Captions
# ---------------------------------------------------------------------------

def resolve_captions(document: Document, auto_accept: bool) -> bool:
    caption_blocks = [b for b in document.blocks if b.is_visual]
    if not caption_blocks:
        return auto_accept

    if auto_accept:
        for block in caption_blocks:
            if block.region_record.decision in (Decision.COUNT, Decision.EXCLUDE):
                continue
            block.region_record.as_manual_override(
                Decision.COUNT, reason="Auto-accepted (--default)",
            )
        print(green(f"  {len(caption_blocks)} caption(s)/visual element(s) auto-counted."))
        return True

    _panel(f"{len(caption_blocks)} caption/visual element(s) detected.")
    print(QCAA_CAPTION_CAVEAT)

    for i, block in enumerate(caption_blocks, start=1):
        if block.region_record.decision in (Decision.COUNT, Decision.EXCLUDE):
            continue
        print()
        print(grey(f"--- Caption {i} of {len(caption_blocks)} ---"))
        print(f'  "{cyan(block.text.strip()[:120])}"')
        choice = _menu(
            "How should this caption be counted?",
            [
                ("count",   f"{green('Count')} it (a word is a word - overrides the default)"),
                ("exclude", f"{red('Exclude')} it (QCAA treats captions as visual elements)"),
                ("help",    "Explain the QCAA rule for captions"),
                ("all",     "Accept suggestions for all remaining captions"),
            ],
            default_key="count",
        )
        if choice == "help":
            print(QCAA_CAPTION_HELP)
            choice = _menu(
                "How should this caption be counted?",
                [
                    ("count",   "Count it"),
                    ("exclude", "Exclude it"),
                    ("all",     "Accept suggestions for all remaining captions"),
                ],
                default_key="count",
            )
        if choice == "all":
            auto_accept = True
            block.region_record.as_manual_override(
                Decision.COUNT, reason="Auto-accepted (user chose Accept All)",
            )
        elif choice == "exclude":
            block.region_record.as_manual_override(
                Decision.EXCLUDE, reason="User chose to exclude caption",
            )
        else:
            block.region_record.as_manual_override(
                Decision.COUNT, reason="User chose to count caption",
            )
    return auto_accept


# ---------------------------------------------------------------------------
# Footnotes
# ---------------------------------------------------------------------------

def resolve_footnotes(document: Document, auto_accept: bool) -> bool:
    if not document.footnotes:
        return auto_accept

    if auto_accept:
        for fn in document.footnotes:
            if fn.answer is not None:
                continue
            fn.answer = "count"
            fn.answer_record = DecisionRecord(
                decision=Decision.COUNT,
                rule_id="F-AUTO", confidence=0.6,
                reason="Auto-accepted", source="manual",
            )
        print(green(f"  {len(document.footnotes)} footnote(s) auto-counted."))
        return True

    _panel(f"{len(document.footnotes)} footnote/endnote(s) detected.")
    for i, fn in enumerate(document.footnotes, start=1):
        if fn.answer is not None:
            continue
        print()
        print(grey(f"--- Footnote {i} of {len(document.footnotes)} ({fn.kind}) ---"))
        print(f'  "{cyan(fn.text[:200])}"')
        default = None
        if fn.suggestion == "bibliographic":
            default = "exclude"
        elif fn.suggestion == "commentary":
            default = "count"
        choice = _menu(
            "How should this footnote be treated?",
            [
                ("exclude", f"{red('Exclude')} (purely a bibliographic reference)"),
                ("count",   f"{green('Count')} (contains commentary or content)"),
                ("all",     "Accept suggestions for all remaining footnotes"),
            ],
            default_key=default,
        )
        if choice == "all":
            auto_accept = True
            fn.answer = "count"
            fn.answer_record = DecisionRecord(
                decision=Decision.COUNT,
                rule_id="F-AUTO", confidence=0.6,
                reason="Auto-accepted (user chose Accept All)", source="manual",
            )
            continue
        fn.answer = choice  # type: ignore
        fn.answer_record = DecisionRecord(
            decision=Decision.EXCLUDE if choice == "exclude" else Decision.COUNT,
            rule_id="F-USER", confidence=1.0,
            reason=f"User chose '{choice}'", source="manual",
        )
    return auto_accept


# ---------------------------------------------------------------------------
# Citations
# ---------------------------------------------------------------------------

def resolve_citations(document: Document) -> None:
    _panel("Citation detection", "(APA 7)")
    auto = [c for c in document.citations if c.source == "auto"]
    if auto:
        print(f"Found {bold(green(str(len(auto))))} in-text citation(s):")
        for i, c in enumerate(auto, start=1):
            occ = "" if c.occurrences == 1 else f" {grey(f'(x{c.occurrences})')}"
            print(f"  {yellow(f'{i:>3}.')} {magenta(c.text)}{occ}")
    else:
        print(grey("No APA 7 in-text citations detected."))
    apply_citation_exclusions(document)

    missed = find_missed_candidates(document)
    if not missed:
        return

    print()
    print(grey("-" * 66))
    print(f"{yellow(str(len(missed)))} other parenthetical(s) contain a year but weren't")
    print("auto-excluded. Choose which (if any) to exclude.")
    print()
    print(grey("  Type a number to exclude. 'a' for all. Enter to finish."))
    print()

    extra: list[Citation] = []
    remaining = list(enumerate(missed, start=1))
    while True:
        if remaining:
            print(bold("Remaining:"))
            for i, m in remaining:
                print(f"  {yellow(f'[{i}]')} {m}")
            print()
        else:
            print(grey("  (No candidates remaining.)"))
            print()
        try:
            raw = input(f"{bold('>')} ").strip()
        except EOFError:
            break
        except KeyboardInterrupt:
            # [FIX 6] Ctrl+C aborts the program here too
            print()
            print(grey("Aborted."))
            sys.exit(130)
        if not raw:
            break
        if raw.lower() == "a":
            for _, m in remaining:
                count = count_occurrences_in_body(document, m)
                if count:
                    extra.append(Citation(text=m, occurrences=count, source="manual"))
            remaining = []
            print(green("  Excluded all remaining candidates."))
            continue
        if not raw.isdigit():
            print(red("  Please type a number, 'a', or Enter to finish."))
            continue
        choice = int(raw)
        match = next((m for idx, m in remaining if idx == choice), None)
        if match is None:
            print(red(f"  No candidate #{choice}."))
            continue
        count = count_occurrences_in_body(document, match)
        if count == 0:
            print(yellow(f"  Not found in body: {match}"))
        else:
            extra.append(Citation(text=match, occurrences=count, source="manual"))
            print(green(f"  Excluded {count} occurrence(s)."))
        remaining = [(i, m) for i, m in remaining if m != match]

    if extra:
        document.citations.extend(extra)
        apply_citation_exclusions(document)
        print(green(f"Added {len(extra)} citation(s)."))


# ---------------------------------------------------------------------------
# Top-level
# ---------------------------------------------------------------------------

def resolve_all(document: Document, auto_accept: bool = False) -> None:
    auto_accept = resolve_appendixes(document, auto_accept)
    auto_accept = resolve_text_boxes(document, auto_accept)
    auto_accept = resolve_tables(document, auto_accept)
    auto_accept = resolve_captions(document, auto_accept)
    auto_accept = resolve_footnotes(document, auto_accept)
    resolve_citations(document)