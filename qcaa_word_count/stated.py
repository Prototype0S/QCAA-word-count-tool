"""Extract a stated word count from the title page, if present."""
from __future__ import annotations

import re

from .model import Block, BlockType, Document, Region

_STATED_COUNT_PATTERNS = [
    re.compile(r"word\s*count\s*[:\-]\s*(\d[\d,]*)", re.IGNORECASE),
    re.compile(r"word\s*length\s*[:\-]\s*(\d[\d,]*)", re.IGNORECASE),
    re.compile(r"\b(\d[\d,]*)\s*words?\b", re.IGNORECASE),
]


def extract_stated_count(document: Document) -> int | None:
    """Return the number the title page claims, or None."""
    for block in document.blocks:
        if block.region != Region.TITLE_PAGE:
            continue
        if block.block_type == BlockType.HEADING:
            continue
        text = block.text
        for pattern in _STATED_COUNT_PATTERNS:
            m = pattern.search(text)
            if m:
                try:
                    return int(m.group(1).replace(",", ""))
                except ValueError:
                    continue
    return None