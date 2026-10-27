"""APA 7 in-text citation detection and exclusion."""
from __future__ import annotations

import re

from .model import Citation, Decision, DecisionRecord, Document, Region, SpanType

# ---------------------------------------------------------------------------
# Patterns
# ---------------------------------------------------------------------------

# Ordered most-specific to least-specific. The claim-check in
# find_apa_citations relies on this order: a broader pattern will not
# claim a range that a narrower pattern already claimed.
APA_PATTERNS: list[tuple[str, re.Pattern]] = [
    # (Smith, 2020, p. 45) — with page numbers
    ("with page numbers", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+(?:\s+(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"(?:\s+et\s+al\.?)?\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?\s*\)"
    )),

    # [FIX 1] Multi-citation: (Smith, 2020; Jones, 2021)
    # Must appear BEFORE "simple", otherwise "simple" claims the first
    # half of the multi-citation and blocks the full match.
    ("multi-citation", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"(?:\s+et\s+al\.?)?\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?"
        r"(?:\s*;\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"(?:\s+et\s+al\.?)?\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?)+"
        r"\s*\)"
    )),

    # (Smith & Jones, 2020) — multi-author, comma before year
    ("multi-author", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s*(?:,|&|and)\s*[A-Z][A-Za-z'’\-]+)+"
        r"(?:\s*,?\s*&?\s*[A-Z][A-Za-z'’\-]+)?"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),

    # (Smith et al., 2020)
    ("et-al", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+\s+et\s+al\.?\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),

    # [FIX 2] (Smith, n.d.) — APA "no date" form.
    # Must appear BEFORE "simple" so it isn't shadowed by the year
    # requirement (which "simple" doesn't satisfy either, so ordering
    # here is mostly for readability, but placing it before "simple"
    # is defensive against future edits).
    ("no-date", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"(?:\s+et\s+al\.?)?\s*,\s*n\.?\s*d\.?\s*\)"
    )),

    # (Smith, 2020)
    ("simple", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),

    # [FIX 3c] Tolerant "Author. , year" form. Catches malformed
    # citations like "(Australian Bureau of Statistics. , 2026)"
    # which students paste from citation generators.
    ("simple-loose", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+(?:\s+[A-Z][A-Za-z'’\-]+){0,5}"
        r"\s*\.?\s*,\s*(?:19|20)\d{2}[a-z]?\s*\)"
    )),

    # (Australian Bureau of Statistics, 2026) — multi-word corporate
    ("corporate", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+(?:\s+[A-Z][A-Za-z'’\-]+){1,6}"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?\s*\)"
    )),

    # Smith (2020) — narrative form
    ("narrative", re.compile(
        r"\b[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:and|&)\s+[A-Z][A-Za-z'’\-]+|\s+et\s+al\.?)?"
        r"\s+\((?:19|20)\d{2}[a-z]?\)"
    )),
]

# [FIX 3] Words that look like author surnames but aren't.
_NON_AUTHOR_WORDS = frozenset({
    "figure", "fig", "figs", "table", "tbl", "tables",
    "section", "sect", "chapter", "ch", "part", "appendix", "app",
    "page", "pages", "p", "pp", "vol", "volume", "vols",
    "issue", "no", "number", "note", "notes", "see", "cf",
    "exhibit", "graph", "chart", "diagram", "image", "photo",
    "map", "plate", "equation", "eq", "eqs", "formula",
    "example", "ex", "source", "case", "step", "phase",
    "in", "of", "the", "a", "an", "and", "or", "by",
})


def _author_is_plausible(matched: str) -> bool:
    """Reject matches whose leading token is a non-author noun."""
    m = re.match(r"[\(\s]*([A-Za-z'’\-]+)", matched)
    if not m:
        return True
    return m.group(1).lower() not in _NON_AUTHOR_WORDS


def find_apa_citations(text: str) -> list[tuple[str, str]]:
    found = []
    claimed = []
    for name, pattern in APA_PATTERNS:
        for m in pattern.finditer(text):
            start, end = m.span()
            # A range that overlaps a claimed range is a shadow match
            # (usually a shorter pattern inside a longer one). Skip.
            if any(s < end and start < e for s, e in claimed):
                continue
            matched = m.group(0)
            if not _author_is_plausible(matched):
                continue
            claimed.append((start, end))
            found.append((matched, name))
    return found


def collect_citations(document: Document) -> list[Citation]:
    by_text: dict[str, Citation] = {}
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            continue
        for matched, pattern_name in find_apa_citations(block.text):
            if matched in by_text:
                by_text[matched].occurrences += 1
            else:
                by_text[matched] = Citation(
                    text=matched, occurrences=1, source="auto",
                    pattern_name=pattern_name,
                )
    return sorted(by_text.values(), key=lambda c: c.text.lower())


# Trailing punctuation that can cling to the last span of a citation.
_TRAILING_PUNCT = '.,;:!?)]}\'"\u2019\u201d'


def apply_citation_exclusions(document: Document) -> None:
    """Mark every span inside a detected citation as EXCLUDE.

    Uses case-insensitive search so a manual entry in a different case
    than the source still matches.
    """
    items = [(c.text, c.source) for c in document.citations if c.text]
    if not items:
        return

    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        block_text = block.text
        claimed_ranges = []
        for ctext, csource in items:
            for m in re.finditer(re.escape(ctext), block_text, re.IGNORECASE):
                claimed_ranges.append((m.start(), m.end(), csource))
        if not claimed_ranges:
            continue
        for span in block.spans:
            s_start = span.source.char_start
            s_end = span.source.char_end

            # Trim trailing punctuation before the containment check.
            trimmed_end = s_end
            while (trimmed_end > s_start
                   and block_text[trimmed_end - 1] in _TRAILING_PUNCT):
                trimmed_end -= 1
            if trimmed_end <= s_start:
                continue

            for c_start, c_end, c_source in claimed_ranges:
                if s_start >= c_start and trimmed_end <= c_end:
                    span.span_type = SpanType.CITATION
                    span.decision_record = DecisionRecord(
                        decision=Decision.EXCLUDE,
                        rule_id="I-CITATION",
                        confidence=0.9,
                        reason="In-text citation (excluded by QCAA rules)",
                        source="auto",
                    )
                    span.tags.discard("citation:auto")
                    span.tags.discard("citation:manual")
                    span.tags.add(f"citation:{c_source}")
                    break


def find_missed_candidates(document: Document) -> list[str]:
    """Parentheticals containing a 4-digit year that weren't auto-excluded.

    Deduplicated by case-insensitive value so the user isn't offered the
    same string twice.
    """
    already = {c.text.lower() for c in document.citations}
    candidates: list[str] = []
    seen: set[str] = set()
    loose = re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+[^()]{0,60}?(?:19|20)\d{2}[^()]{0,30}?\)"
    )
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            continue
        for m in loose.finditer(block.text):
            matched = m.group(0)
            key = matched.lower().strip()
            if key in already or key in seen:
                continue
            if not _author_is_plausible(matched):
                continue
            seen.add(key)
            candidates.append(matched)
    return candidates


def count_occurrences_in_body(document: Document, needle: str) -> int:
    """Count case-insensitive occurrences of a substring in countable body."""
    if not needle:
        return 0
    total = 0
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            continue
        total += sum(1 for _ in re.finditer(re.escape(needle), block.text, re.IGNORECASE))
    return total