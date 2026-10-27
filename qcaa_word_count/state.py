"""Save and load decisions across runs.

[FIX 8] The old loader matched saved blocks by *block_index* and
verified with a text hash. That meant inserting a paragraph in the
middle of the document shifted every subsequent block's index, and
every decision after the insertion was silently discarded.

This version matches by *text hash alone*. A block whose text hasn't
changed gets its decision back, regardless of where it moved to.
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Optional

from .model import (
    Block, Citation, Decision, DecisionRecord, Document, Table, TableAnswer,
)
from .tables import apply_table_answer

SAVE_VERSION = 1


def _hash_text(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8", errors="replace")).hexdigest()[:16]


def _table_hash(table: Table) -> str:
    """Hash a table's full text so we can match it across edits."""
    chunks = []
    for row in table.rows:
        for cell in row:
            chunks.append(cell.text)
    return _hash_text("||".join(chunks))


def _footnote_hash(fn) -> str:
    return _hash_text(fn.text)


def default_save_path(docx_path: str) -> str:
    """The default JSON path next to the .docx file."""
    base = os.path.splitext(docx_path)[0]
    return f"{base}_qcaa_decisions.json"


def save_decisions(document: Document, docx_path: str,
                   save_path: Optional[str] = None,
                   total: Optional[int] = None) -> str:
    """Write the current decision state to JSON. Returns the path written."""
    if save_path is None:
        save_path = default_save_path(docx_path)

    blocks = []
    for b in document.blocks:
        # Only save blocks that carry a *non-default* decision.
        if b.region_record.source != "manual":
            continue
        blocks.append({
            "block_index": b.block_index,
            "text_hash": _hash_text(b.text),
            "decision": b.region_record.decision.value,
            "rule_id": b.region_record.rule_id,
            "reason": b.region_record.reason,
        })

    tables = []
    for t in document.tables:
        if t.answer is None:
            continue
        tables.append({
            "table_index": t.table_index,
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

    # Persist manually-added citations so they survive re-runs.
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
    """Restore decisions from JSON where they still match.

    [FIX 8] Blocks are matched by text hash, not by block_index, so
    inserting a paragraph earlier in the document doesn't discard
    every saved decision that came after it.

    Tables are also matched by hash (index is unreliable after edits).
    Footnotes are matched by ID + hash, because footnote IDs are
    stable within a document.
    """
    stats = {"restored": 0, "skipped": 0, "mismatched": 0, "loaded": False}

    try:
        payload = json.loads(Path(load_path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return stats

    if payload.get("version") != SAVE_VERSION:
        return stats

    stats["loaded"] = True

    # --- Blocks: match by text hash -----------------------------------
    # A document with two identical blocks (two "Introduction" headings,
    # two "Figure 1" captions) matches the saved entry to whichever one
    # is found first, then deletes it from the pool so the same saved
    # entry isn't applied twice.
    blocks_by_hash: dict[str, Block] = {}
    for b in document.blocks:
        h = _hash_text(b.text)
        # Keep the first occurrence of each hash.
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
        block.region_record = DecisionRecord(
            decision=decision,
            rule_id=saved.get("rule_id") or "MANUAL",
            confidence=1.0,
            reason=saved.get("reason") or "Restored from saved decisions",
            source="manual",
        )
        stats["restored"] += 1

    # --- Tables: match by text hash -----------------------------------
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

    # --- Footnotes: match by ID + hash --------------------------------
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

    # --- Manual citations --------------------------------------------
    for saved in payload.get("citations_manual", []):
        document.citations.append(Citation(
            text=saved["text"],
            occurrences=saved.get("occurrences", 0),
            source="manual",
        ))

    return stats


def _safe_mtime(path: str) -> float:
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0