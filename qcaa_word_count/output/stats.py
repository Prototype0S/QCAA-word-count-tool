"""Per-section and vocabulary statistics."""
from __future__ import annotations

from ..model import BlockType, Decision, Document
from .ansi import interleaved_items


def compute_statistics(document: Document, result: dict) -> dict:
    sections: list[dict] = []
    current = {"heading": "(start of document)", "words": 0, "tables": 0}

    counted_by_index: dict[int, int] = {}
    for block in document.blocks:
        counted_by_index[block.block_index] = sum(
            1 for s in block.spans
            if s.decision_record.decision == Decision.COUNT
        )

    table_words_by_position: dict[int, int] = {}
    for table in document.tables:
        n = sum(
            1 for cell in table.iter_cells() for s in cell.spans
            if s.decision_record.decision == Decision.COUNT
        )
        table_words_by_position[table.block_index] = (
            table_words_by_position.get(table.block_index, 0) + n
        )

    for kind, obj in interleaved_items(document):
        if kind == "block":
            block = obj
            if block.block_type == BlockType.HEADING:
                if current["words"] or current["tables"] or current["heading"] != "(start of document)":
                    sections.append(current)
                current = {
                    "heading": block.text.strip() or "(untitled)",
                    "words": counted_by_index.get(block.block_index, 0),
                    "tables": 0,
                }
            else:
                current["words"] += counted_by_index.get(block.block_index, 0)
        else:
            table = obj
            current["words"] += table_words_by_position.get(table.block_index, 0)
            current["tables"] += 1

    if current["words"] or current["tables"] or current["heading"] != "(start of document)":
        sections.append(current)

    sections = [s for s in sections if s["words"] or s["tables"]]

    counted_words: list[str] = []
    for block in document.blocks:
        if block.region_record.decision != Decision.COUNT:
            continue
        for span in block.spans:
            if span.decision_record.decision == Decision.COUNT:
                counted_words.append(span.text)

    alpha_words = [w.strip(".,;:!?()[]{}\"'").lower() for w in counted_words]
    alpha_words = [w for w in alpha_words if w and any(c.isalpha() for c in w)]
    unique = len(set(alpha_words))
    total_alpha = len(alpha_words)
    richness = (unique / total_alpha * 100) if total_alpha else 0.0

    longest_block = None
    longest_count = 0
    for block in document.blocks:
        n = counted_by_index.get(block.block_index, 0)
        if n > longest_count:
            longest_count = n
            longest_block = block

    return {
        "sections": sections,
        "unique_words": unique,
        "total_alpha_words": total_alpha,
        "richness_pct": richness,
        "longest_paragraph_words": longest_count,
        "longest_paragraph_preview": (
            longest_block.text.strip()[:120] + "…"
            if longest_block and len(longest_block.text.strip()) > 120
            else (longest_block.text.strip() if longest_block else "")
        ),
    }