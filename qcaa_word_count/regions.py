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


# [FIX 8] A title page is identified by its *content*, not just by
# "comes before the first heading." A block of running prose means the
# title page has ended — even if there is no Word heading style.
# This stops documents that use bold body paragraphs as pseudo-headings
# from being entirely swallowed into TITLE_PAGE.
_SENTENCE_END_RE = re.compile(r"[.!?](?:\s|$)")
_STATED_COUNT_LINE_RE = re.compile(
    r"(?:word\s*count|word\s*length)\s*[:\-]\s*\d",
    re.IGNORECASE,
)


def _looks_like_body_prose(text: str) -> bool:
    """A block is prose if it contains two or more sentence endings,
    or one sentence ending plus enough words to be a real sentence.

    A title-page line like "A Business Pitch By Edward Carroll" has
    zero sentence endings. "What am I trying to make?" has one but is
    short. "I chose the name Hour Maker as it captures the core idea
    of the business: to help young people..." has multiple sentence
    endings and is long — that's body prose, and the title page ends.
    """
    t = text.strip()
    if not t:
        return False
    if _STATED_COUNT_LINE_RE.search(t):
        return False   # "Word count: 1500" is a title-page line
    sentences = len(_SENTENCE_END_RE.findall(t))
    words = len(t.split())
    if sentences >= 2:
        return True
    if sentences >= 1 and words >= 15:
        return True
    return False


def detect_regions(document: Document) -> None:
    # [FIX 8] Walk the opening blocks and stop the TITLE_PAGE region
    # at the first block that looks like body prose. A Word heading
    # still ends the title page (as before); so does substantial prose.
    current_region = Region.BODY
    current_level: int | None = None
    in_title_page = True

    for block in document.blocks:
        text = block.text.strip()

        # --- Opening title-page window ---------------------------------
        if in_title_page:
            if block.block_type == BlockType.HEADING:
                # A real Word heading ends the title page.
                in_title_page = False
                # fall through to heading handling below
            elif _looks_like_body_prose(text):
                # Running prose ends the title page.
                in_title_page = False
                block.region = Region.BODY
                current_region = Region.BODY
                continue
            else:
                block.region = Region.TITLE_PAGE
                continue

        # --- Section heading handling ----------------------------------
        if block.block_type == BlockType.HEADING:
            key = text.lower().rstrip(":.-")
            level = heading_level_from_style(block.style_name) or 1
            if key in SECTION_TRIGGERS:
                current_region = Region(SECTION_TRIGGERS[key])
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

# [FIX 6a] Extra bibliographic signals. A block containing a year
# *and* any of these is almost certainly a reference entry, even if
# it doesn't start with "Surname, X." (non-APA styles, edited books,
# web sources, etc.).
_REFERENCE_HINT_RE = re.compile(
    r"\b(?:Retrieved|Accessed|Available\s+(?:at|from)|Published|"
    r"Press|Journal|Vol\.?|Vols?\.?|Ed\.?|Eds?\.?|Trans\.?|"
    r"doi|ISBN|ISSN|Reprinted|Revised)\b"
    r"|https?://|www\.|doi\.org/",
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
    # [FIX 6a] year + bibliographic keyword => reference entry
    has_year = bool(re.search(r"\(\s*(?:19|20)\d{2}[a-z]?\s*\)", text))
    has_hint = bool(_REFERENCE_HINT_RE.search(text))
    if has_year and has_hint:
        return True
    return False


def _detect_trailing_references(document: Document) -> None:
    """Mark consecutive reference-like blocks at the end of the body.

    [FIX 6b] The old version only looked at the last 20 body blocks,
    so long bibliographies were under-detected. This version walks
    backwards and stops at the first non-reference once it has
    started — so a 40-entry reference list is caught in full.
    """
    body_blocks = [b for b in document.blocks if b.region == Region.BODY]
    if not body_blocks:
        return

    marked: list[Block] = []
    for block in reversed(body_blocks):
        text = block.text.strip()
        if not text:
            continue
        if block.block_type == BlockType.HEADING:
            # A heading at the very bottom: if we've already marked
            # entries, this is the "References" heading — stop.
            if marked:
                break
            continue
        if _looks_like_reference(text):
            marked.append(block)
        elif marked:
            # First non-reference after we've started: stop.
            break

    # A single stray reference-like block is not a bibliography.
    if len(marked) < 2:
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
            