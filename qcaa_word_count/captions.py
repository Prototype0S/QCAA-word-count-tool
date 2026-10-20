"""Caption detection.

The QCAA rule: captions and visual elements of written genres are
excluded. But a prose sentence that merely names a figure ("Figure 3
shows...") is not a caption; it's body text.

The right rule is: a caption is a *short standalone block* that leads
with a figure/table label. It's not a caption if the label is
followed by a verb that turns the sentence into a continuation.
"""
from __future__ import annotations

import re

from .model import Block, BlockType, Decision, DecisionRecord, Document, Region

# A caption MUST start with a label. This is the strongest signal.
CAPTION_LABEL_RE = re.compile(
    r"^\s*(?:figure|fig\.?|table|tbl\.?|chart|graph|diagram|image|"
    r"photo|map|exhibit|plate)\s*[\dIVXivx]+(?:\.[\d]+)?[a-z]?\b",
    re.IGNORECASE,
)

SOURCE_RE = re.compile(
    r"^\s*source\s*[:.\u2013\u2014\-]\s*\S",
    re.IGNORECASE,
)

BYLINE_RE = re.compile(
    r"^\s*(?:by|words by|written by|illustration by|illustrated by|"
    r"photo by|images? by)\s+\S",
    re.IGNORECASE,
)

# Continuation verbs that turn "Figure 3 shows..." into prose.
# These only count when they appear *immediately after* the label,
# before any colon or em-dash that would separate label from body.
_CONTINUATION_VERBS = (
    "shows", "illustrates", "indicates", "demonstrates", "depicts",
    "compares", "summarises", "summarizes", "presents", "outlines",
    "explains", "details", "sets out",
)


def looks_like_caption(block: Block) -> tuple[bool, str]:
    """Return (is_caption, kind).

    kind is 'caption', 'source', 'by-line', or ''.
    """
    text = block.text.strip()
    if not text:
        return False, ""

    # Word's own "Caption" style is definitive
    if block.style_name and block.style_name.lower().startswith("caption"):
        return True, "caption"

    # Source and by-line are separate
    if SOURCE_RE.match(text):
        return True, "source"
    if BYLINE_RE.match(text):
        return True, "by-line"

    # Must lead with a caption label
    label_match = CAPTION_LABEL_RE.match(text)
    if not label_match:
        return False, ""

    # After the label, check for continuation verbs that would make
    # this prose ("Figure 3 shows ..."). Only reject if the verb
    # appears *before* a colon/em-dash — otherwise it's a label like
    # "Figure 3: Comparison, showing the effect of X" (a real caption).
    rest = text[label_match.end():]
    # Strip the separator if present
    sep_match = re.match(r"\s*[:.\u2013\u2014\-]\s*", rest)
    if sep_match:
        # There is a separator — this is a caption of the form
        # "Figure 3: ...". The label is separated from the body, so
        # continuation verbs in the body don't disqualify it.
        return True, "caption"

    # No separator — check for prose continuation
    rest_words = rest.lower().split()[:3]
    if any(w.rstrip(",.;:") in _CONTINUATION_VERBS for w in rest_words):
        return False, ""

    # No separator, no continuation verb — but it must still be
    # short to count as a caption (a caption is one line, not a
    # paragraph)
    if len(text.split()) > 25:
        return False, ""

    return True, "caption"


def detect_captions(document: Document) -> None:
    """Mark probable captions with an UNDECIDED region record.

    The interactive resolver then prompts the user to accept or reject
    each one. Nothing is auto-excluded here.
    """
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            continue
        if block.block_type in (BlockType.HEADING, BlockType.TEXT_BOX):
            continue

        is_cap, kind = looks_like_caption(block)
        if is_cap:
            block.is_visual = True
            block.region_record = DecisionRecord(
                decision=Decision.UNDECIDED,
                rule_id=f"R-{kind.upper().replace('-', '_')}",
                confidence=0.6,
                reason=f"Possible {kind}: QCAA treats visual elements as excluded",
                source="auto",
            )