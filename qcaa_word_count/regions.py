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
    # [FIX 8] Walk the opening blocks and stop the TITLE_PAGE region
    # at the first block that looks like body prose. A Word heading
    # still ends the title page (as before); so does substantial prose.
    #
    # [FIX 16] Look ahead one block. If a block would be classified as
    # TITLE_PAGE but the *next* block is prose, this block is the
    # opening line of the body, not a title-page line. A title page
    # cannot "resume" after prose, so anything directly before the
    # first prose block is prose too.
    current_region = Region.BODY
    current_level: int | None = None
    in_title_page = True

    # Precompute prose verdicts so the lookahead doesn't re-run the regex.
    prose_flags = [_looks_like_body_prose(b.text) for b in document.blocks]

    for idx, block in enumerate(document.blocks):
        text = block.text.strip()

        # --- Opening title-page window ---------------------------------
        if in_title_page:
            if block.block_type == BlockType.HEADING:
                in_title_page = False
                # fall through to heading handling below
            elif prose_flags[idx]:
                in_title_page = False
                block.region = Region.BODY
                current_region = Region.BODY
                continue
            else:
                # [FIX 16] Lookahead: opening prose block.
                next_is_prose = (
                    idx + 1 < len(document.blocks)
                    and document.blocks[idx + 1].block_type != BlockType.HEADING
                    and prose_flags[idx + 1]
                )
                if next_is_prose and any(c.isalpha() for c in text):
                    in_title_page = False
                    block.region = Region.BODY
                    current_region = Region.BODY
                    continue
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


_SENTENCE_END_RE = re.compile(r"[.!?](?:\s|$)")
_STATED_COUNT_LINE_RE = re.compile(
    r"(?:word\s*count|word\s*length)\s*[:\-]\s*\d",
    re.IGNORECASE,
)


def _looks_like_body_prose(text: str) -> bool:
    """A block is prose if it has two or more sentence endings, or one
    sentence ending and at least 15 words. Used to end the title-page
    window early when a document uses bold body paragraphs as pseudo-
    headings.
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


_APA_AUTHOR_START_RE = re.compile(
    r"^\s*[A-Z][A-Za-z'’\-]+(?:\s*,\s*[A-Z]\.?)+\s*\(\s*(?:19|20)\d{2}[a-z]?\s*\)",
)

# [FIX 12a] Student-format author (year) at the start of a block.
#   Australian Bureau of Statistics (2026). ...
#   Smith (2020). The title. ...
#   Smith and Jones (2020). ...
#   Smith, J., & Jones, B. (2020). ...
_REFERENCE_AUTHOR_YEAR_RE = re.compile(
    r"^\s*[A-Z][A-Za-z'’\-]+"
    r"(?:\s*,\s*[A-Z]\.?)?"
    r"(?:\s*,?\s*(?:&|and)\s+[A-Z][A-Za-z'’\-]+"
    r"(?:\s*,\s*[A-Z]\.?)?)?"
    r"(?:\s+[A-Z][A-Za-z'’\-]+){0,6}"
    r"\s*\(\s*(?:19|20)\d{2}[a-z]?\s*\)\s*\."
)

_URL_LIKE_RE = re.compile(
    r"(?:https?://|www\.|<https?://|doi\.org/|doi:\s*10\.)"
    r"|\.(?:org|com|edu|gov|net|ac\.uk|edu\.au)\b",
    re.IGNORECASE,
)

# Counts distinct URL *starts* rather than every fragment the broader
# _URL_LIKE_RE matches. One URL contributes 1 here, not 2.
_URL_START_RE = re.compile(r"https?://|www\.", re.IGNORECASE)

_REFERENCE_HINT_RE = re.compile(
    r"\b(?:Retrieved|Accessed|Available\s+(?:at|from)|Published|"
    r"Press|Journal|Vol\.?|Vols?\.?|Ed\.?|Eds?\.?|Trans\.?|"
    r"doi|ISBN|ISSN|Reprinted|Revised)\b"
    r"|https?://|www\.|doi\.org/",
    re.IGNORECASE,
)


def _looks_like_reference(text: str) -> bool:
    """Decide whether a block is a reference-list entry.

    Principle: when in doubt, treat as a reference. A missed reference
    entry costs 30-80 words; a wrongly-excluded prose paragraph costs
    maybe 20. Bias toward exclusion.

    [FIX 12c] The sentence-ending guard is *only* applied in the
    no-URL branch, and only when a year+hint pair is present. Reference
    entries routinely contain 4-6 periods (page ranges, "[Data set].",
    stray punctuation) so an unconditional guard rejects them.
    """
    text = text.strip()
    if not text:
        return False

    # --- URL present -------------------------------------------------
    if _URL_LIKE_RE.search(text):
        url_starts = len(_URL_START_RE.findall(text))
        sentences = len(_SENTENCE_END_RE.findall(text))
        # Two or more URLs — treat as reference unconditionally.
        # Covers the "pasted whole reference list as one paragraph" case.
        if url_starts >= 2:
            return True
        # One URL and a compact block — reference entry.
        # Allow up to 6 sentence endings so [Data set]. and stray
        # periods inside an entry don't disqualify it.
        if sentences <= 6:
            return True
        return False

    # --- No URL: rely on author/year patterns ------------------------
    if _APA_AUTHOR_START_RE.match(text):
        return True
    if _REFERENCE_AUTHOR_YEAR_RE.match(text):
        return True

    has_year = bool(re.search(r"\(\s*(?:19|20)\d{2}[a-z]?\s*\)", text))
    has_hint = bool(_REFERENCE_HINT_RE.search(text))
    if has_year and has_hint:
        # Year + bibliographic keyword, no URL. Apply a sentence
        # guard only here — without a URL this is a weaker signal.
        if len(_SENTENCE_END_RE.findall(text)) > 4:
            return False
        return True

    return False


def _detect_trailing_references(document: Document) -> None:
    """Mark consecutive reference-like blocks at the end of the body.

    Walks backwards and stops at the first non-reference once it has
    started, so a 40-entry reference list is caught in full. A single
    stray reference-like block is not marked (needs >= 2).
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
            if marked:
                break
            continue
        if _looks_like_reference(text):
            marked.append(block)
        elif marked:
            break

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