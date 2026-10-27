"""The main pipeline: ingest → decisions → interactive → count → output."""
from __future__ import annotations

import os
import sys

from .captions import detect_captions
from .citations import apply_citation_exclusions, collect_citations
from .count import count_words
from .diff import print_diff
from .footnotes import apply_footnote_suggestions, read_footnotes_from_docx
from .helpers import classify_span, is_equation_heavy_block
from .ingest import ingest
from .interactive import resolve_all
from .model import Decision, DecisionRecord, Document
from .output import (
    bold, bold_green, cyan, export_decisions, green, grey, print_ansi_preview,
    print_report, red, write_html_preview, yellow,
)
from .regions import apply_regional_decisions, detect_regions
from .state import default_save_path, load_decisions, save_decisions
from .stated import extract_stated_count
from .tables import classify_table


def _apply_span_decisions(document: Document) -> None:
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


def _apply_user_overrides(document: Document, overrides: dict) -> None:
    """Apply browser-toggled overrides from the interactive HTML preview.

    Overrides are keyed by span ID ("{block_index}:{char_start}") — the
    same identifier used in the HTML preview. That makes them
    *per-instance*: toggling one occurrence of "the" does not affect
    the other twenty-three.

    Falls back to text-keyed overrides for compatibility with any
    overrides file created before per-instance IDs existed.
    """
    # Build an index of every span in the document, keyed by span ID.
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

    applied = 0
    skipped = 0
    for key, want in overrides.items():
        if want not in ("count", "exclude"):
            skipped += 1
            continue

        # New format: "block_index:char_start"
        span = spans_by_id.get(key)

        # Legacy fallback: key is a text token (from old previews).
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
            skipped += 1
            continue

        if want == "count":
            span.decision_record = DecisionRecord(
                decision=Decision.COUNT,
                rule_id="I-USER-OVERRIDE",
                confidence=1.0,
                reason="User toggle from HTML preview",
                source="manual",
            )
        else:
            span.decision_record = DecisionRecord(
                decision=Decision.EXCLUDE,
                rule_id="I-USER-OVERRIDE",
                confidence=1.0,
                reason="User toggle from HTML preview",
                source="manual",
            )
        applied += 1

    if applied or skipped:
        msg = f"  Applied {applied} user override(s) from HTML preview"
        if skipped:
            msg += f" ({skipped} skipped — span no longer present)"
        print(grey(msg))

def _prepare(path: str) -> Document:
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
    """Set QCAA_DEBUG_REGIONS=1 to dump every block's region."""
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

    # ---- Reset --------------------------------------------------------
    if reset and os.path.exists(save_path):
        try:
            os.remove(save_path)
            print(grey(f"Removed saved decisions: {save_path}"))
        except OSError as e:
            print(yellow(f"Couldn't remove {save_path}: {e}"))

    # ---- Load --------------------------------------------------------
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

    # ---- Apply browser overrides (after saved decisions, before prompts)
    if overrides:
        _apply_user_overrides(document, overrides)

    # ---- Early-exit modes --------------------------------------------
    if mode == "summary":
        resolve_all(document, auto_accept=True)
        result = count_words(document)
        _snapshot_sections(document, result)
        if save:
            save_decisions(document, path, total=result["total"])
        # Caveat BEFORE the number so scripts still see it.
        print(grey("  (Fast mode: footnotes counted as body. "
                   "Use --review for the QCAA-accurate count.)"))
        print()
        print(f"{document.filename}\t{result['total']}")
        return 0

    if mode == "preview":
        resolve_all(document, auto_accept=True)
        result = count_words(document)
        _snapshot_sections(document, result)
        # --preview-only is non-committal: do NOT save decisions.
        html_file = write_html_preview(
            document, result, output_path=html_path, open_browser=open_browser,
        )
        print()
        print(f"{bold('HTML preview written to:')} {cyan(html_file)}")
        print(grey("  (No decisions were saved. Use --review or --include-all "
                   "to commit.)"))
        return 0

    # ---- Normal modes -------------------------------------------------
    if mode == "include-all":
        print()
        print(f"{bold('Mode:')} {green('Include everything (auto-accept all)')}")
        print()
        resolve_all(document, auto_accept=True)
    else:
        resolve_all(document, auto_accept=False)

    result = count_words(document)
    stated = extract_stated_count(document)
    result["stated_count"] = stated

    _snapshot_sections(document, result)

    # ---- Diff (optional) ---------------------------------------------
    if show_diff or mode == "check":
        sections = getattr(document, "_sections_snapshot", [])
        print_diff(save_path, sections, result["total"])

    # ---- Mismatch warning + region sanity check ----------------------
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

    # ---- Save (before output, so preview shows final state) ----------
    if save:
        save_decisions(document, path, total=result["total"])

    # ---- Preview / report -------------------------------------------
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
    """Store the per-section summary so `state.save_decisions` can persist it."""
    from .output import compute_statistics
    stats = compute_statistics(document, result)
    document._sections_snapshot = stats.get("sections", [])