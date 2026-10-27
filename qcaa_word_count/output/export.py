"""Decision-log export (JSON and CSV)."""
from __future__ import annotations

import csv
import io
import json

from ..model import Document


def export_decisions(document: Document, result: dict, format: str = "json") -> str:
    payload = {
        "filename": document.filename,
        "total": result["total"],
        "buckets": result["buckets"],
        "blocks": [
            {
                "index": b.block_index,
                "type": b.block_type.value,
                "region": b.region.value,
                "region_decision": b.region_record.decision.value,
                "region_rule": b.region_record.rule_id,
                "text": b.text,
                "spans": [
                    {
                        "text": s.text,
                        "type": s.span_type.value,
                        "decision": s.decision_record.decision.value,
                        "rule": s.decision_record.rule_id,
                        "reason": s.decision_record.reason,
                        "tags": sorted(s.tags),
                    }
                    for s in b.spans
                ],
            }
            for b in document.blocks
        ],
        "tables": [
            {
                "index": t.table_index + 1,
                "suggestion": t.suggestion.value,
                "answer": t.answer.value if t.answer else None,
                "header_row": t.header_row_index,
            }
            for t in document.tables
        ],
        "citations": [
            {"text": c.text, "occurrences": c.occurrences,
             "source": c.source, "pattern": c.pattern_name}
            for c in document.citations
        ],
        "footnotes": [
            {"id": f.footnote_id, "kind": f.kind, "text": f.text,
             "suggestion": f.suggestion, "answer": f.answer}
            for f in document.footnotes
        ],
    }
    if format == "json":
        return json.dumps(payload, indent=2, ensure_ascii=False)
    if format == "csv":
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(["block_index", "span_text", "span_type",
                    "decision", "rule", "reason"])
        for b in document.blocks:
            for s in b.spans:
                w.writerow([
                    b.block_index, s.text, s.span_type.value,
                    s.decision_record.decision.value,
                    s.decision_record.rule_id,
                    s.decision_record.reason or "",
                ])
        return buf.getvalue()
    return ""