"""Structure detection: assign a Region to each block."""
from __future__ import annotations

import re

from .helpers import heading_level_from_style
from .model import Block, BlockType, Decision, DecisionRecord, Document, Region

SECTION_TRIGGERS = {
    "contents": "contents",
    "table of contents": "contents",
    "list of figures": "contents",
    "list of tables": "contents",
    "abstract": "abstract",
    "bibliography": "references",
    "reference list": "references",
    "references": "references",
    "list of references": "references",
    "works cited": "references",
    "appendix": "appendix",
    "appendices": "appendix",
    "annex": "appendix",
}


def detect_regions(document: Document) -> None:
    first_heading_idx = None
    for block in document.blocks:
        if block.block_type == BlockType.HEADING:
            first_heading_idx = block.block_index
            break

    current_region = Region.BODY
    current_level: int | None = None

    for block in document.blocks:
        if first_heading_idx is not None and block.block_index < first_heading_idx:
            block.region = Region.TITLE_PAGE
            continue

        if block.block_type == BlockType.HEADING:
            text = block.text.strip().lower().rstrip(":.-")
            level = heading_level_from_style(block.style_name) or 1
            if text in SECTION_TRIGGERS:
                current_region = Region(SECTION_TRIGGERS[text])
                current_level = level
            elif current_region not in (Region.BODY, Region.TITLE_PAGE):
                if current_level is not None and level <= current_level:
                    current_region = Region.BODY
                    current_level = None
            block.region = current_region
        else:
            block.region = current_region

    _detect_trailing_references(document)


_APA_AUTHOR_START_RE = re.compile(
    r"^\s*[A-Z][A-Za-z'’\-]+(?:\s*,\s*[A-Z]\.?)+\s*\(\s*(?:19|20)\d{2}[a-z]?\s*\)",
)
_URL_LIKE_RE = re.compile(
    r"(?:https?://|www\.|<https?://|doi\.org/|doi:\s*10\.)"
    r"|\.(?:org|com|edu|gov|net|ac\.uk|edu\.au)\b",
    re.IGNORECASE,
)


def _looks_like_reference(text: str) -> bool:
    text = text.strip()
    if not text:
        return False
    if _APA_AUTHOR_START_RE.match(text):
        return True
    if _URL_LIKE_RE.search(text):
        return True
    if re.search(r"\(\s*(?:19|20)\d{2}[a-z]?\s*\)", text) and \
       re.search(r"\b\w+\.(?:org|com|edu|gov|net|au|uk)\b", text, re.IGNORECASE):
        return True
    return False


def _detect_trailing_references(document: Document) -> None:
    body_blocks = [b for b in document.blocks if b.region == Region.BODY]
    if not body_blocks:
        return
    marked = []
    for block in body_blocks[-20:]:
        text = block.text.strip()
        if not text or block.block_type == BlockType.HEADING:
            continue
        if _looks_like_reference(text):
            marked.append(block)
    if not marked:
        return
    has_url = any(_URL_LIKE_RE.search(b.text) for b in marked)
    if len(marked) < 2 and not has_url:
        return
    for block in marked:
        block.region = Region.REFERENCES


def apply_regional_decisions(document: Document) -> None:
    """Assign a decision record to each block based on its region.

    Preserves manual overrides on re-runs.
    """
    for block in document.blocks:
        if block.region_record.source == "manual":
            continue

        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            block.region_record = DecisionRecord(
                decision=Decision.EXCLUDE,
                rule_id=f"R-{block.region.value.upper()}",
                confidence=0.9,
                reason=f"Region '{block.region.value}' is excluded by QCAA rules",
                source="auto",
            )
        elif block.region == Region.APPENDIX:
            block.region_record = DecisionRecord(
                decision=Decision.FLAG,
                rule_id="R-APPENDIX",
                confidence=0.6,
                reason="Appendixes excluded only if supplementary-only",
                source="auto",
            )
        elif block.block_type == BlockType.TEXT_BOX:
            block.region_record = DecisionRecord(
                decision=Decision.FLAG,
                rule_id="R-TEXTBOX",
                confidence=0.5,
                reason="Text box content — confirm whether it's prose or visual",
                source="auto",
            )
        else:
            block.region_record = DecisionRecord(
                decision=Decision.COUNT,
                rule_id="R-BODY",
                confidence=1.0,
                reason="Body content counts",
                source="auto",
            )