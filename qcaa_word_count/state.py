"""Save and load decisions across runs.

The goal is: after the user answers every prompt once, a second run on
the same document (even after edits) should skip the prompts whose
decisions are still valid. We match by (block_index, text_hash) so
that inserting paragraphs before a table shifts the index, but the
table's text hash still matches.

IMPORTANT: when a table answer is restored, `apply_table_answer` must be
called so the table's *spans* get their decisions. Otherwise the count
drops because spans default to UNDECIDED.
"""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
from typing import Optional

from .model import (
    Block, Decision, DecisionRecord, Document, Table, TableAnswer,
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

    payload = {
        "version": SAVE_VERSION,
        "filename": document.filename,
        "docx_mtime": _safe_mtime(docx_path),
        "total_at_save": total if total is not None else 0,
        "blocks": blocks,
        "tables": tables,
        "footnotes": footnotes,
        "sections": getattr(document, "_sections_snapshot", []),
    }

    Path(save_path).write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return save_path


def load_decisions(document: Document, load_path: str) -> dict:
    """Restore decisions from JSON where they still match.

    Returns a dict with 'restored', 'skipped', 'mismatched', 'loaded'.
    """
    stats = {"restored": 0, "skipped": 0, "mismatched": 0, "loaded": False}

    try:
        payload = json.loads(Path(load_path).read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return stats

    if payload.get("version") != SAVE_VERSION:
        return stats

    stats["loaded"] = True

    # --- Blocks (paragraphs, captions, text boxes) ---------------------
    by_index = {b.block_index: b for b in document.blocks}
    for saved in payload.get("blocks", []):
        idx = saved.get("block_index")
        if idx not in by_index:
            stats["skipped"] += 1
            continue
        block = by_index[idx]
        if _hash_text(block.text) != saved.get("text_hash"):
            stats["mismatched"] += 1
            continue
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

    # --- Tables --------------------------------------------------------
    tables_by_idx = {t.table_index: t for t in document.tables}
    for saved in payload.get("tables", []):
        idx = saved.get("table_index")
        if idx not in tables_by_idx:
            stats["skipped"] += 1
            continue
        table = tables_by_idx[idx]
        if _table_hash(table) != saved.get("text_hash"):
            stats["mismatched"] += 1
            continue
        try:
            answer = TableAnswer(saved["answer"])
        except (KeyError, ValueError):
            stats["skipped"] += 1
            continue
        table.answer = answer
        # CRITICAL: re-apply so the table's spans get their decisions.
        # Restoring table.answer alone leaves the spans UNDECIDED, which
        # makes the table count as 0 in count_words.
        apply_table_answer(
            table, answer,
            reason="Restored from saved decisions",
        )
        stats["restored"] += 1

    # --- Footnotes -----------------------------------------------------
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

    return stats


def _safe_mtime(path: str) -> float:
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0.0