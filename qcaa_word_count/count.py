"""Walk the decision map and produce a total."""
from __future__ import annotations

from collections import defaultdict

from .model import Decision, Document, Table


def _table_is_in_excluded_region(document: Document, table: Table) -> bool:
    """True if the block immediately preceding this table is excluded.

    [FIX 5] Tables were previously counted regardless of region, so a
    table inside an excluded appendix (or title page) still added its
    words to the total. `table.block_index` is the block index that
    the *next* paragraph would have been given, so every block with
    `block_index < table.block_index` precedes the table.
    """
    prev = None
    for block in document.blocks:
        if block.block_index < table.block_index:
            prev = block
        else:
            break
    return prev is not None and prev.region_record.decision == Decision.EXCLUDE


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
        if block.region_record.decision == Decision.EXCLUDE:
            rule = block.region_record.rule_id or "R-REGION"
            buckets[f"excluded:{rule}"] += len(block.spans)
            continue
        if block.region_record.decision in (Decision.FLAG, Decision.UNDECIDED):
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

    for table in document.tables:
        # [FIX 5] Skip tables that live inside an excluded region.
        if _table_is_in_excluded_region(document, table):
            table_summary.append({
                "index": table.table_index + 1,
                "suggestion": table.suggestion.value,
                "answer": "excluded (region)",
                "counted": 0,
                "excluded": sum(len(c.spans) for c in table.iter_cells()),
            })
            buckets["table:excluded (in excluded region)"] += sum(
                len(c.spans) for c in table.iter_cells()
            )
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
        if fn.answer == "count":
            total += len(fn.text.split())
            footnote_summary["counted"] += 1
            footnote_summary["counted_words"] += len(fn.text.split())
        elif fn.answer == "exclude":
            buckets["footnotes (bibliographic)"] += len(fn.text.split())
            footnote_summary["excluded"] += 1
            footnote_summary["excluded_words"] += len(fn.text.split())
        else:
            unresolved_footnotes += 1
            buckets["footnotes (unresolved — not counted)"] += len(fn.text.split())

    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        for span in block.spans:
            if span.decision_record.rule_id == "I-CITATION":
                if "citation:manual" in span.tags:
                    citation_summary["manual_words"] += 1
                else:
                    citation_summary["auto_words"] += 1
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
        "stated_count": None,   # filled in by pipeline.extract_stated_count
    }