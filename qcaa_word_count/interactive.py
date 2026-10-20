"""All interactive prompts in one place."""
from __future__ import annotations

from .citations import (
    apply_citation_exclusions,
    count_occurrences_in_body,
    find_missed_candidates,
)
from .helpers import is_numeric_token  # noqa: F401  (kept for future use)
from .model import (
    Citation, Decision, DecisionRecord, Document, Region, TableAnswer,
)
from .tables import apply_table_answer, suggest_table_answer, table_has_strong_default


QCAA_TABLE_HELP = """
------------------------------------------------------------------
TABLES (QCAA rule)
------------------------------------------------------------------
  Option 1 — Count the whole table (contains information).
  Option 2 — Exclude the whole table (raw data or calculations).
  Option 3 — Count the header row only, exclude the data rows.
"""

QCAA_CAPTION_HELP = """
------------------------------------------------------------------
CAPTIONS AND VISUAL ELEMENTS (QCAA rule)
------------------------------------------------------------------
QCAA excludes captions and other visual elements of written genres.
But in science reports and PSMTs, captions often identify variables
and count. When in doubt, ask your teacher.
"""


def _menu(question: str, options: list[tuple[str, str]],
          default_key: str | None = None) -> str:
    print()
    print(f"== {question} ==")
    for i, (key, label) in enumerate(options, start=1):
        marker = "  (suggested)" if key == default_key else ""
        print(f"  [{i}] {label}{marker}")
    while True:
        default_hint = ""
        if default_key is not None:
            for i, (key, _) in enumerate(options, start=1):
                if key == default_key:
                    default_hint = f" [{i}]"
                    break
        try:
            raw = input(f"\nYour choice{default_hint}: ").strip()
        except (EOFError, KeyboardInterrupt):
            raw = ""
        if not raw and default_key is not None:
            return default_key
        if not raw.isdigit():
            print("Please type a number.")
            continue
        idx = int(raw) - 1
        if 0 <= idx < len(options):
            return options[idx][0]
        print(f"Please type a number between 1 and {len(options)}.")


def preview_table(table, max_rows: int = 4) -> str:
    lines = []
    for r_idx, row in enumerate(table.rows):
        if r_idx >= max_rows:
            lines.append(f"... ({len(table.rows) - max_rows} more row(s))")
            break
        cells = []
        for cell in row:
            t = cell.text.strip().replace("\n", " ")
            if len(t) > 20:
                t = t[:17] + "..."
            cells.append(t)
        lines.append(" | ".join(cells))
    return "\n".join(lines)


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
                Decision.EXCLUDE, reason="Auto-accepted (--default)",
            )
        return True

    print()
    print(f"Appendix region detected ({len(appendix_blocks)} block(s))")
    print(f"Starts at: {appendix_blocks[0].text[:80]}")
    choice = _menu(
        "How should this appendix be treated?",
        [
            ("exclude", "Exclude this appendix"),
            ("include", "Include this appendix"),
            ("all",     "Accept suggestions for all remaining prompts"),
        ],
        default_key="include",
    )
    if choice == "all":
        for b in appendix_blocks:
            b.region_record.as_manual_override(
                Decision.EXCLUDE, reason="Auto-accepted (user chose Accept All)",
            )
        return True
    dec = Decision.EXCLUDE if choice == "exclude" else Decision.COUNT
    reason = "User chose to exclude appendix" if choice == "exclude" else "User chose to include appendix"
    for b in appendix_blocks:
        b.region_record.as_manual_override(dec, reason=reason)
    return auto_accept


def resolve_text_boxes(document: Document, auto_accept: bool) -> bool:
    boxes = [
        b for b in document.blocks
        if b.block_type == b.block_type.TEXT_BOX  # type: ignore[attr-defined]
        and b.region_record.decision in (Decision.FLAG, Decision.UNDECIDED)
    ]
    if not boxes:
        return auto_accept
    if auto_accept:
        for b in boxes:
            b.region_record.as_manual_override(
                Decision.COUNT, reason="Auto-accepted text box",
            )
        return True
    print()
    print(f"{len(boxes)} text box(es) detected")
    for b in boxes[:5]:
        print(f"  - {b.text.strip()[:70]}")
    choice = _menu(
        "How should text box content be treated?",
        [
            ("count",   "Count (treat as prose)"),
            ("exclude", "Exclude (treat as visual)"),
        ],
        default_key="count",
    )
    dec = Decision.COUNT if choice == "count" else Decision.EXCLUDE
    reason = "User counted text boxes" if choice == "count" else "User excluded text boxes"
    for b in boxes:
        b.region_record.as_manual_override(dec, reason=reason)
    return auto_accept


def resolve_tables(document: Document, auto_accept: bool) -> bool:
    if not document.tables:
        return auto_accept
    print()
    print(f"{len(document.tables)} table(s) found.")
    for i, table in enumerate(document.tables, start=1):
        if table.answer is not None:
            continue
        if auto_accept:
            suggested = suggest_table_answer(table)
            table.answer = suggested
            apply_table_answer(table, suggested, reason="Auto-accepted")
            continue
        print()
        print(f"Table {i} of {len(document.tables)}")
        print(f"Suggested: {table.suggestion.value} "
              f"(confidence {table.suggestion_confidence:.2f})")
        print(preview_table(table))
        print(f"  {table.word_count} word(s) in this table.")
        default_key = suggest_table_answer(table).value if table_has_strong_default(table) else None
        choice = _menu(
            "How should this table be counted?",
            [
                ("count_all",    "Count the whole table (information)"),
                ("exclude_all",  "Exclude the whole table (raw data or calculations)"),
                ("headers_only", "Count the header row only, exclude the data rows"),
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
            suggested = suggest_table_answer(table)
            table.answer = suggested
            apply_table_answer(table, suggested, reason="Auto-accepted")
            continue
        table.answer = TableAnswer(choice)
        apply_table_answer(table, table.answer, reason=f"User chose '{choice}'")
    return auto_accept


def resolve_captions(document: Document, auto_accept: bool) -> bool:
    caption_blocks = [b for b in document.blocks if b.is_visual]
    if not caption_blocks:
        return auto_accept
    print()
    print(f"{len(caption_blocks)} caption/visual element(s) detected.")
    for i, block in enumerate(caption_blocks, start=1):
        if block.region_record.decision in (Decision.COUNT, Decision.EXCLUDE):
            continue
        if auto_accept:
            block.region_record.as_manual_override(
                Decision.EXCLUDE, reason="Auto-accepted (--default)",
            )
            continue
        print()
        print(f"Caption {i} of {len(caption_blocks)}")
        print(f'  "{block.text.strip()[:120]}"')
        choice = _menu(
            "How should this caption be counted?",
            [
                ("exclude", "Exclude it (QCAA's default)"),
                ("count",   "Count it (names variables or carries information)"),
                ("help",    "Explain the QCAA rule for captions"),
                ("all",     "Accept suggestions for all remaining captions"),
            ],
            default_key="include",
        )
        if choice == "help":
            print(QCAA_CAPTION_HELP)
            choice = _menu(
                "How should this caption be counted?",
                [
                    ("exclude", "Exclude it"),
                    ("count",   "Count it"),
                    ("all",     "Accept suggestions for all remaining captions"),
                ],
                default_key="include",
            )
        if choice == "all":
            auto_accept = True
            block.region_record.as_manual_override(
                Decision.EXCLUDE, reason="Auto-accepted (user chose Accept All)",
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


def resolve_footnotes(document: Document, auto_accept: bool) -> bool:
    if not document.footnotes:
        return auto_accept
    print()
    print(f"{len(document.footnotes)} footnote/endnote(s) detected.")
    for i, fn in enumerate(document.footnotes, start=1):
        if fn.answer is not None:
            continue
        if auto_accept:
            fn.answer = "exclude" if fn.suggestion == "bibliographic" else "count"
            fn.answer_record = DecisionRecord(
                decision=Decision.EXCLUDE if fn.answer == "exclude" else Decision.COUNT,
                rule_id="F-AUTO", confidence=0.6,
                reason="Auto-accepted", source="manual",
            )
            continue
        print()
        print(f"Footnote {i} of {len(document.footnotes)} ({fn.kind})")
        print(f'  "{fn.text[:200]}"')
        default = None
        if fn.suggestion == "bibliographic":
            default = "exclude"
        elif fn.suggestion == "commentary":
            default = "count"
        choice = _menu(
            "How should this footnote be treated?",
            [
                ("exclude", "Exclude (purely a bibliographic reference)"),
                ("count",   "Count (contains commentary or content)"),
                ("all",     "Accept suggestions for all remaining footnotes"),
            ],
            default_key=default,
        )
        if choice == "all":
            auto_accept = True
            suggested = "exclude" if fn.suggestion == "bibliographic" else "count"
            fn.answer = suggested
            fn.answer_record = DecisionRecord(
                decision=Decision.EXCLUDE if suggested == "exclude" else Decision.COUNT,
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


def resolve_citations(document: Document) -> None:
    print()
    print("Citation detection")
    auto = [c for c in document.citations if c.source == "auto"]
    if auto:
        print(f"Found {len(auto)} in-text citation(s):")
        for i, c in enumerate(auto, start=1):
            occ = "" if c.occurrences == 1 else f" (x{c.occurrences})"
            print(f"  {i:>3}. {c.text}{occ}")
    else:
        print("No APA 7 in-text citations detected.")
    apply_citation_exclusions(document)

    missed = find_missed_candidates(document)
    if not missed:
        return

    print()
    print(f"{len(missed)} other parenthetical(s) contain a year but weren't")
    print("auto-excluded. Choose which (if any) to exclude.")
    print("  Type a number to exclude. 'a' for all. Enter to finish.")

    extra: list[Citation] = []
    remaining = list(enumerate(missed, start=1))
    while True:
        if remaining:
            print("Remaining:")
            for i, m in remaining:
                print(f"  [{i}] {m}")
        else:
            print("  (No candidates remaining.)")
        try:
            raw = input("> ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not raw:
            break
        if raw.lower() == "a":
            for _, m in remaining:
                count = count_occurrences_in_body(document, m)
                if count:
                    extra.append(Citation(text=m, occurrences=count, source="manual"))
            remaining = []
            continue
        if not raw.isdigit():
            print("  Please type a number, 'a', or Enter to finish.")
            continue
        choice = int(raw)
        match = next((m for idx, m in remaining if idx == choice), None)
        if match is None:
            print(f"  No candidate #{choice}.")
            continue
        count = count_occurrences_in_body(document, match)
        if count == 0:
            print(f"  Not found in body: {match}")
        else:
            extra.append(Citation(text=match, occurrences=count, source="manual"))
            print(f"  Excluded {count} occurrence(s).")
        remaining = [(i, m) for i, m in remaining if m != match]

    if extra:
        document.citations.extend(extra)
        apply_citation_exclusions(document)
        print(f"Added {len(extra)} citation(s).")


def resolve_all(document: Document, auto_accept: bool = False) -> None:
    auto_accept = resolve_appendixes(document, auto_accept)
    auto_accept = resolve_text_boxes(document, auto_accept)
    auto_accept = resolve_tables(document, auto_accept)
    auto_accept = resolve_captions(document, auto_accept)
    auto_accept = resolve_footnotes(document, auto_accept)
    resolve_citations(document)