"""Save and load decisions across runs.

SAVE FORMAT
-----------
    {
      "version": 2,
      "blocks": [
        {
          "text_hash": "<sha1[:16]>",
          "region": "<Region.value>",
          "occurrence": <int>,
          "decision": "count" | "exclude",
          "rule_id": "...",
          "reason": "..."
        }, ...
      ],
      "tables": [...],
      "footnotes": [...],
      "citations_manual": [...],
      "sections": [...]
    }

Block identity is (text_hash, region, occurrence) where `occurrence`
disambiguates identical blocks in the same region (e.g. two "Figure 1"
captions in the body). Old saves (version 1, no region/occurrence)
still load, with a one-off message.
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Optional

from .helpers import apply_span_decisions_for_block
from .model import (
    Block, Citation, Decision, DecisionRecord, Document, Region, Table,
    TableAnswer,
)
from .tables import apply_table_answer

SAVE_VERSION = 2


def _hash_text(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8", errors="replace")).hexdigest()[:16]


def _table_hash(table: Table) -> str:
    chunks = []
    for row in table.rows:
        for cell in row:
            chunks.append(cell.text)
    return _hash_text("||".join(chunks))


def _footnote_hash(fn) -> str:
    return _hash_text(fn.text)


def default_save_path(docx_path: str) -> str:
    base = os.path.splitext(docx_path)[0]
    return f"{base}_qcaa_decisions.json"


def save_decisions(document: Document, docx_path: str,
                   save_path: Optional[str] = None,
                   total: Optional[int] = None) -> str:
    if save_path is None:
        save_path = default_save_path(docx_path)

    # Build (text_hash, region) -> occurrence index for manually-overridden
    # blocks, so load can restore the *right* identical block.
    from collections import defaultdict
    counters: dict[tuple[str, str], int] = defaultdict(int)

    blocks = []
    for b in document.blocks:
        if b.region_record.source != "manual":
            continue
        h = _hash_text(b.text)
        key = (h, b.region.value)
        occurrence = counters[key]
        counters[key] += 1
        blocks.append({
            "text_hash": h,
            "region": b.region.value,
            "occurrence": occurrence,
            "decision": b.region_record.decision.value,
            "rule_id": b.region_record.rule_id,
            "reason": b.region_record.reason,
        })

    tables = []
    for t in document.tables:
        if t.answer is None:
            continue
        tables.append({
            "text_hash": _table_hash(t),
            "answer": t.answer.value,
        })

    footnotes = []
    for fn in document.footnotes:
        if fn.answer is None:
            continue
        footnotes.append({
            "footnote_id": fn.footnote_id,
            "text_hash": _footnote_hash(fn),
            "answer": fn.answer,
        })

    citations_manual = [
        {"text": c.text, "occurrences": c.occurrences}
        for c in document.citations if c.source == "manual"
    ]

    payload = {
        "version": SAVE_VERSION,
        "filename": document.filename,
        "docx_mtime": _safe_mtime(docx_path),
        "total_at_save": total if total is not None else 0,
        "blocks": blocks,
        "tables": tables,
        "footnotes": footnotes,
        "citations_manual": citations_manual,
        "sections": getattr(document, "_sections_snapshot", []),
    }

    Path(save_path).write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return save_path


def load_decisions(document: Document, load_path: str) -> dict:
    stats = {"restored": 0, "skipped": 0, "mismatched": 0, "loaded": False}

    try:
        payload = json.loads(Path(load_path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return stats

    version = payload.get("version")
    if version not in (1, 2):
        return stats

    stats["loaded"] = True

    if version == 1:
        _load_v1_blocks(document, payload, stats)
    else:
        _load_v2_blocks(document, payload, stats)

    _load_tables(document, payload, stats)
    _load_footnotes(document, payload, stats)

    for saved in payload.get("citations_manual", []):
        document.citations.append(Citation(
            text=saved["text"],
            occurrences=saved.get("occurrences", 0),
            source="manual",
        ))

    return stats


def _apply_block_decision(block: Block, decision: Decision,
                          rule_id: str, reason: str) -> None:
    block.region_record = DecisionRecord(
        decision=decision,
        rule_id=rule_id or "MANUAL",
        confidence=1.0,
        reason=reason or "Restored from saved decisions",
        source="manual",
    )
    # Re-classify spans so the count reflects the restored decision.
    # Without this, a restored "count this References block" would
    # still have its spans marked R-REGION (excluded).
    apply_span_decisions_for_block(block)


def _load_v1_blocks(document: Document, payload: dict, stats: dict) -> None:
    """Legacy: hash-only matching. Same collision behaviour as before."""
    blocks_by_hash: dict[str, Block] = {}
    for b in document.blocks:
        h = _hash_text(b.text)
        if h not in blocks_by_hash:
            blocks_by_hash[h] = b

    for saved in payload.get("blocks", []):
        h = saved.get("text_hash")
        if not h or h not in blocks_by_hash:
            stats["skipped"] += 1
            continue
        block = blocks_by_hash.pop(h)
        try:
            decision = Decision(saved["decision"])
        except (KeyError, ValueError):
            stats["skipped"] += 1
            continue
        _apply_block_decision(
            block, decision, saved.get("rule_id"), saved.get("reason"),
        )
        stats["restored"] += 1


def _load_v2_blocks(document: Document, payload: dict, stats: dict) -> None:
    """New: match by (text_hash, region, occurrence)."""
    from collections import defaultdict
    blocks_by_key: dict[tuple[str, str], list[Block]] = defaultdict(list)
    for b in document.blocks:
        key = (_hash_text(b.text), b.region.value)
        blocks_by_key[key].append(b)

    for saved in payload.get("blocks", []):
        h = saved.get("text_hash")
        region = saved.get("region")
        if not h or not region:
            stats["skipped"] += 1
            continue
        key = (h, region)
        candidates = blocks_by_key.get(key, [])
        occurrence = saved.get("occurrence", 0)
        if occurrence >= len(candidates):
            stats["skipped"] += 1
            continue
        block = candidates[occurrence]
        try:
            decision = Decision(saved["decision"])
        except (KeyError, ValueError):
            stats["skipped"] += 1
            continue
        _apply_block_decision(
            block, decision, saved.get("rule_id"), saved.get("reason"),
        )
        stats["restored"] += 1


def _load_tables(document: Document, payload: dict, stats: dict) -> None:
    tables_by_hash: dict[str, Table] = {}
    for t in document.tables:
        h = _table_hash(t)
        if h not in tables_by_hash:
            tables_by_hash[h] = t

    for saved in payload.get("tables", []):
        h = saved.get("text_hash")
        if not h or h not in tables_by_hash:
            stats["skipped"] += 1
            continue
        table = tables_by_hash.pop(h)
        try:
            answer = TableAnswer(saved["answer"])
        except (KeyError, ValueError):
            stats["skipped"] += 1
            continue
        table.answer = answer
        apply_table_answer(
            table, answer,
            reason="Restored from saved decisions",
        )
        stats["restored"] += 1


def _load_footnotes(document: Document, payload: dict, stats: dict) -> None:
    fns_by_id = {fn.footnote_id: fn for fn in document.footnotes}
    for saved in payload.get("footnotes", []):
        fid = saved.get("footnote_id")
        if fid not in fns_by_id:
            stats["skipped"] += 1
            continue
        fn = fns_by_id[fid]
        if _footnote_hash(fn) != saved.get("text_hash"):
            stats["mismatched"] += 1
            continue
        answer = saved.get("answer")
        if answer not in ("count", "exclude"):
            stats["skipped"] += 1
            continue
        fn.answer = answer
        stats["restored"] += 1


def _safe_mtime(path: str) -> float:
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0