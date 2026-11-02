"""The main pipeline: ingest → decisions → interactive → count → output."""
from __future__ import annotations

import os

from .captions import detect_captions
from .citations import apply_citation_exclusions, collect_citations
from .count import count_words
from .diff import print_diff
from .footnotes import apply_footnote_suggestions, read_footnotes_from_docx
from .helpers import apply_span_decisions_for_block, classify_span, is_equation_heavy_block
from .ingest import ingest
from .interactive import resolve_all
from .model import Decision, DecisionRecord, Document, Region, TableAnswer
from .output import (
    bold, cyan, export_decisions, green, grey, print_ansi_preview,
    print_report, red, write_html_preview, yellow,
)
from .regions import (
    apply_regional_decisions, assign_table_regions, detect_regions,
)
from .state import default_save_path, load_decisions, save_decisions
from .stated import extract_stated_count
from .tables import classify_table


def _apply_span_decisions(document: Document) -> None:
    for block in document.blocks:
        apply_span_decisions_for_block(block)


def _apply_user_overrides(document: Document, overrides: dict) -> None:
    """Apply browser-toggled overrides from the interactive HTML preview.

    Handles both span-level overrides (individual words) and table-level
    overrides (whole tables toggled via the caption).

    Accepts any of these shapes:
      - new:    {"spans": {...}, "tables": {...}, "total": N}
      - legacy: {"overrides": <either of the above>}
      - legacy: {span_id: "count"|"exclude"}   (flat dict)
    """
    if not overrides:
        return

    # Unwrap the legacy {"overrides": {...}} envelope.
    if "overrides" in overrides and not ("spans" in overrides or "tables" in overrides):
        overrides = overrides["overrides"]

    if "spans" in overrides or "tables" in overrides:
        span_overrides = overrides.get("spans") or {}
        table_overrides = overrides.get("tables") or {}
    else:
        span_overrides = overrides
        table_overrides = {}

    # ---- Index every span by (block_index, char_start) ---------------
    spans_by_id: dict[str, object] = {}
    for block in document.blocks:
        for span in block.spans:
            sid = f"{span.source.block_index}:{span.source.char_start}"
            spans_by_id[sid] = span
    for table in document.tables:
        for cell in table.iter_cells():
            for span in cell.spans:
                sid = f"{span.source.block_index}:{span.source.char_start}"
                spans_by_id[sid] = span

    # ---- Span overrides ----------------------------------------------
    applied_spans = 0
    skipped_spans = 0
    for key, want in span_overrides.items():
        if want not in ("count", "exclude"):
            skipped_spans += 1
            continue
        span = spans_by_id.get(key)
        if span is None and ":" not in key:
            for block in document.blocks:
                if block.region_record.decision == Decision.EXCLUDE:
                    continue
                for s in block.spans:
                    if s.text == key:
                        span = s
                        break
                if span is not None:
                    break
        if span is None:
            skipped_spans += 1
            continue
        span.decision_record = DecisionRecord(
            decision=Decision.COUNT if want == "count" else Decision.EXCLUDE,
            rule_id="I-USER-OVERRIDE",
            confidence=1.0,
            reason="User toggle from HTML preview",
            source="manual",
        )
        applied_spans += 1

    # ---- Table overrides ---------------------------------------------
    applied_tables = 0
    skipped_tables = 0
    for key, answer in table_overrides.items():
        if answer not in ("count_all", "exclude_all"):
            skipped_tables += 1
            continue
        try:
            table_idx = int(key)
        except (ValueError, TypeError):
            skipped_tables += 1
            continue
        table = next((t for t in document.tables if t.table_index == table_idx),
                     None)
        if table is None:
            skipped_tables += 1
            continue

        for cell in table.iter_cells():
            equation_heavy = (
                is_equation_heavy_block(cell) if answer == "count_all" else False
            )
            for span in cell.spans:
                if span.decision_record.rule_id == "I-USER-OVERRIDE":
                    continue
                if answer == "exclude_all":
                    span.decision_record = DecisionRecord(
                        decision=Decision.EXCLUDE,
                        rule_id="I-USER-TABLE-OVERRIDE",
                        confidence=1.0,
                        reason="User table toggle from HTML preview",
                        source="manual",
                    )
                else:
                    span_type, rec = classify_span(
                        span.text, equation_heavy=equation_heavy,
                    )
                    span.span_type = span_type
                    span.decision_record = rec

        table.answer = (
            TableAnswer.EXCLUDE_ALL if answer == "exclude_all"
            else TableAnswer.COUNT_ALL
        )
        applied_tables += 1

    # ---- Citation exclusions must be re-applied ---------------------
    # count_all re-classifies cell spans, which can silently overwrite
    # I-CITATION decisions. The apply pass is safe to run multiple times
    # (it preserves I-USER-OVERRIDE and MANUAL spans), so we run it
    # unconditionally to catch any citations we just clobbered.
    if applied_tables:
        apply_citation_exclusions(document)

    total_applied = applied_spans + applied_tables
    if total_applied:
        parts = []
        if applied_spans:
            parts.append(f"{applied_spans} span")
        if applied_tables:
            parts.append(f"{applied_tables} table")
        print(grey(f"  Applied {total_applied} user override(s) "
                   f"({', '.join(parts)}) from HTML preview"))
    if skipped_spans or skipped_tables:
        skip_total = skipped_spans + skipped_tables
        print(grey(f"  ({skip_total} override(s) skipped — "
                   f"span no longer present)"))


def _prepare(path: str) -> Document:
    document = ingest(path)
    for table in document.tables:
        table.suggestion, table.suggestion_confidence = classify_table(table)
    detect_regions(document)
    assign_table_regions(document)
    apply_regional_decisions(document)
    _apply_span_decisions(document)
    detect_captions(document)
    document.footnotes = read_footnotes_from_docx(path)
    apply_footnote_suggestions(document)
    document.citations = collect_citations(document)
    apply_citation_exclusions(document)
    return document


def _print_summary_line(document: Document) -> None:
    math_total = sum(b.math_object_count for b in document.blocks)
    line = f"  {len(document.blocks)} paragraph(s), {len(document.tables)} table(s)"
    print(green(line))
    if math_total:
        print(grey(f"  {math_total} math object(s)"))
    if document.drawings:
        print(grey(f"  {document.drawings} drawing(s) "
                   f"({document.drawings_with_text} contain text)"))


def _print_region_debug(document: Document) -> None:
    if not os.environ.get("QCAA_DEBUG_REGIONS"):
        return
    print()
    print("Region assignments:")
    for i, b in enumerate(document.blocks):
        preview = b.text.strip()[:70].replace("\n", " ")
        print(f"  {i:>3}  {b.region.value:<12}  {preview}")
    print()


def run(path: str, *, mode: str = "interactive",
        auto_accept: bool = False,
        open_browser: bool = True,
        save: bool = True,
        load: bool = True,
        reset: bool = False,
        show_diff: bool = False,
        export_format: str | None = None,
        export_path: str | None = None,
        html_path: str | None = None,
        overrides: dict | None = None) -> int:
    """Run the tool. `mode` is one of:
        review, include-all, summary, preview, check
    """
    if not os.path.isfile(path):
        print(red(f"Couldn't find a file at: {path}"))
        return 2
    if not path.lower().endswith(".docx"):
        print(red(f"'{path}' doesn't look like a .docx file."))
        return 2

    save_path = default_save_path(path)

    if reset and os.path.exists(save_path):
        try:
            os.remove(save_path)
            print(grey(f"Removed saved decisions: {save_path}"))
        except OSError as e:
            print(yellow(f"Couldn't remove {save_path}: {e}"))

    print()
    print(f"{bold('Loading:')} {cyan(path)}")
    document = _prepare(path)

    if load:
        stats = load_decisions(document, save_path)
        if stats["loaded"]:
            restored = stats["restored"]
            mismatched = stats["mismatched"]
            if restored:
                msg = f"  Restored {restored} saved decision(s)"
                if mismatched:
                    msg += f", {mismatched} block(s) changed"
                print(grey(msg))

    _print_summary_line(document)
    _print_region_debug(document)

    # ---- Resolve decisions (all modes) ------------------------------
    if mode in ("summary", "preview"):
        resolve_all(document, auto_accept=True)
    elif mode == "include-all":
        print()
        print(f"{bold('Mode:')} {green('Include everything (auto-accept all)')}")
        print()
        resolve_all(document, auto_accept=True)
    else:
        resolve_all(document, auto_accept=False)

    if overrides:
        _apply_user_overrides(document, overrides)

    # ---- Early-exit: summary mode ----------------------------------
    if mode == "summary":
        result = count_words(document)
        _snapshot_sections(document, result)
        if save:
            save_decisions(document, path, total=result["total"])
        print(grey("  (Fast mode: footnotes counted as body. "
                   "Use --review for the QCAA-accurate count.)"))
        print()
        print(f"{document.filename}\t{result['total']}")
        return 0

    # ---- Early-exit: preview mode ----------------------------------
    if mode == "preview":
        result = count_words(document)
        _snapshot_sections(document, result)
        html_file = write_html_preview(
            document, result, output_path=html_path, open_browser=open_browser,
        )
        print()
        print(f"{bold('HTML preview written to:')} {cyan(html_file)}")
        print(grey("  (No decisions were saved. Use --review or --include-all "
                   "to commit.)"))
        return 0

    # ---- Normal modes -----------------------------------------------
    result = count_words(document)
    stated = extract_stated_count(document)
    result["stated_count"] = stated

    _snapshot_sections(document, result)

    if show_diff or mode == "check":
        sections = getattr(document, "_sections_snapshot", [])
        print_diff(save_path, sections, result["total"])

    if stated and stated != result["total"]:
        delta = result["total"] - stated
        pct = abs(delta) / stated * 100 if stated else 0
        print()
        print(yellow(
            f"  ⚠ Title page says \"Word Count: {stated}\"\n"
            f"    Tool calculates {result['total']} "
            f"(differs by {abs(delta):+d})"
        ))
        if pct > 20:
            print()
            print(yellow(
                f"  That's a {pct:.0f}% difference. That usually means a "
                f"whole region,\n"
                f"  table, or caption was excluded when it shouldn't be "
                f"(or vice versa).\n"
                f"  Run with --review to check the ambiguous decisions."
            ))

    total_blocks = len(document.blocks)
    excluded_blocks = sum(
        1 for b in document.blocks
        if b.region_record.decision == Decision.EXCLUDE
    )
    if total_blocks and excluded_blocks / total_blocks > 0.20:
        pct_blocks = excluded_blocks / total_blocks * 100
        print()
        print(yellow(
            f"  ⚠ {excluded_blocks} of {total_blocks} blocks "
            f"({pct_blocks:.0f}%) were excluded by region.\n"
            f"    If that looks wrong, check the 'Regions detected' "
            f"section of the report\n"
            f"    or run with --review."
        ))

    if save:
        save_decisions(document, path, total=result["total"])

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


def _snapshot_sections(document: Document, result: dict) -> None:
    from .output import compute_statistics
    stats = compute_statistics(document, result)
    document._sections_snapshot = stats.get("sections", [])