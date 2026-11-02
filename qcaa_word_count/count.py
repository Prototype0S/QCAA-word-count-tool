"""Walk the decision map and produce a total."""
from __future__ import annotations

from collections import defaultdict

from .helpers import classify_span, tokenize
from .model import Decision, Document, Region, Table


def _table_is_excluded(document: Document, table: Table) -> bool:
    """True if the table lives in a region QCAA excludes wholesale.

    Uses ``table.region`` (assigned by ``regions.assign_table_regions``)
    instead of the old block-index comparison, which relied on the
    synthetic 100_000+ indices given to table cells.

    Appendixes are special: the region itself isn't excluded, but the
    resolver may have decided to exclude it. We read the appendix
    decision from the first appendix block (the resolver sets them all
    to the same value, so first is representative).
    """
    if table.region in (Region.TITLE_PAGE, Region.CONTENTS,
                        Region.ABSTRACT, Region.REFERENCES):
        return True
    if table.region == Region.APPENDIX:
        for block in document.blocks:
            if block.region == Region.APPENDIX:
                return block.region_record.decision == Decision.EXCLUDE
    return False


def count_words(document: Document) -> dict:
    buckets: dict[str, int] = defaultdict(int)
    bucket_samples: dict[str, list[str]] = defaultdict(list)
    total = 0
    table_summary = []
    citation_summary = {"auto": 0, "auto_words": 0, "manual": 0, "manual_words": 0}
    footnote_summary = {"counted": 0, "counted_words": 0,
                        "excluded": 0, "excluded_words": 0}
    total_math_objects = 0
    unresolved_footnotes = 0

    for block in document.blocks:
        total_math_objects += block.math_object_count
        dec = block.region_record.decision

        if dec == Decision.EXCLUDE:
            # Respect manual span overrides inside otherwise-excluded
            # blocks. If every span is auto-excluded, take the fast path
            # (whole block to a bucket). If any span was toggled to
            # count in the HTML preview, fall through and process
            # span-by-span so the toggled ones are honoured.
            has_manual_count = any(
                s.decision_record.source == "manual"
                and s.decision_record.decision == Decision.COUNT
                for s in block.spans
            )
            if not has_manual_count:
                rule = block.region_record.rule_id or "R-REGION"
                buckets[f"excluded:{rule}"] += len(block.spans)
                continue

        if dec in (Decision.FLAG, Decision.UNDECIDED):
            buckets[f"flagged:{block.region.value}"] += len(block.spans)
            continue

        for span in block.spans:
            if span.decision_record.decision == Decision.COUNT:
                total += 1
            else:
                key = f"excluded:{span.decision_record.rule_id or 'unknown'}"
                buckets[key] += 1
                if len(bucket_samples[key]) < 5:
                    bucket_samples[key].append(span.text)
                if span.decision_record.rule_id == "I-CITATION":
                    if "citation:manual" in span.tags:
                        citation_summary["manual_words"] += 1
                    else:
                        citation_summary["auto_words"] += 1

    for table in document.tables:
        if _table_is_excluded(document, table):
            n_cells = sum(len(c.spans) for c in table.iter_cells())
            table_summary.append({
                "index": table.table_index + 1,
                "suggestion": table.suggestion.value,
                "answer": "excluded (region)",
                "counted": 0,
                "excluded": n_cells,
            })
            buckets["table:excluded (in excluded region)"] += n_cells
            continue

        counted = excluded = 0
        for cell in table.iter_cells():
            total_math_objects += cell.math_object_count
            for span in cell.spans:
                if span.decision_record.decision == Decision.COUNT:
                    total += 1
                    counted += 1
                else:
                    excluded += 1
        answer_label = table.answer.value if table.answer else "unresolved"
        table_summary.append({
            "index": table.table_index + 1,
            "suggestion": table.suggestion.value,
            "answer": answer_label,
            "counted": counted,
            "excluded": excluded,
        })
        if excluded:
            buckets[f"table:{answer_label}"] += excluded

    for fn in document.footnotes:
        tokens = tokenize(fn.text)
        if fn.answer == "count":
            counted = 0
            for tok in tokens:
                _, rec = classify_span(tok)
                if rec.decision == Decision.COUNT:
                    counted += 1
                else:
                    key = f"excluded:footnote-{rec.rule_id or 'unknown'}"
                    buckets[key] += 1
                    if len(bucket_samples[key]) < 5:
                        bucket_samples[key].append(tok)
            total += counted
            footnote_summary["counted"] += 1
            footnote_summary["counted_words"] += counted
        elif fn.answer == "exclude":
            buckets["footnotes (bibliographic)"] += len(tokens)
            footnote_summary["excluded"] += 1
            footnote_summary["excluded_words"] += len(tokens)
        else:
            unresolved_footnotes += 1
            buckets["footnotes (unresolved — not counted)"] += len(tokens)

    citation_summary["auto"] = sum(1 for c in document.citations if c.source == "auto")
    citation_summary["manual"] = sum(1 for c in document.citations if c.source == "manual")

    return {
        "total": total,
        "buckets": dict(buckets),
        "bucket_samples": dict(bucket_samples),
        "tables": table_summary,
        "citations": citation_summary,
        "footnotes": footnote_summary,
        "math_objects": total_math_objects,
        "unresolved_footnotes": unresolved_footnotes,
        "stated_count": None,
    }