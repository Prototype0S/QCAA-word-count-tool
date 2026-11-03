"""APA 7 in-text citation detection and exclusion.

Walks BOTH paragraph blocks AND table cell blocks. Covers all standard
APA 7 in-text forms.

CHANGELOG (this version)
------------------------
  - Adds ``no-date-multiword``: (Exposure Draft, n.d.), (Some Long Title,
    n.d.) etc. — previously counted as ordinary words because ``no-date``
    only allowed a single-word author.
  - Adds ``narrative-no-date``: Smith (n.d.), Smith et al. (n.d.) —
    previously counted because ``narrative`` required a year.
  - ``apply_citation_exclusions`` now handles spans that straddle two
    adjacent citations, e.g. "(Smith, 2020)(Jones, 2021)" tokenises to
    the single span "2020)(Jones" which the old fully-inside check
    couldn't exclude.
"""
from __future__ import annotations

import re

from .model import Citation, Decision, DecisionRecord, Document, Region, SpanType


_NARRATIVE_LEAD_BLACKLIST = (
    "According|As|However|Therefore|Thus|Moreover|Furthermore|"
    "Additionally|Although|While|When|Where|Since|Because|Despite|"
    "Besides|Also|In|Of|On|At|By|For|From|With|This|These|That|Those|It"
)


APA_PATTERNS: list[tuple[str, re.Pattern]] = [
    ("with page numbers", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s*,\s*[A-Z][A-Za-z'’\-]+)*"
        r"(?:\s*,?\s*(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"(?:\s+et\s+al\.?)?"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"\s*,\s*pp?\.?\s*(?:\d+|[ivxlc]+)"
        r"(?:\s*[-\u2013]\s*(?:\d+|[ivxlc]+))?"
        r"\s*\)"
    )),
    ("same-author-multi-year", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*(?:19|20)\d{2}[a-z]?)+"
        r"\s*\)"
    )),
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
    ("secondary-source", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s*,?\s*(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"\s*,?\s+as\s+cited\s+in\s+"
        r"[A-Z][A-Za-z'’\-]+"
        r"(?:\s*,?\s*(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"\s*\)"
    )),
    ("multi-author", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s*(?:,|&|and)\s*[A-Z][A-Za-z'’\-]+)+"
        r"(?:\s*,?\s*&?\s*[A-Z][A-Za-z'’\-]+)?"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),
    ("et-al", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+\s+et\s+al\.?\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),
    ("no-date", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"(?:\s+et\s+al\.?)?\s*,\s*n\.?\s*d\.?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),
    # Multi-word author + n.d. — "(Exposure Draft, n.d.)",
    # "(Working Title of a Report, n.d.)". The single-word "no-date"
    # pattern above doesn't match these.
    ("no-date-multiword", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s+[A-Za-z][A-Za-z'’\-]+){1,8}"
        r"\s*\.?\s*,\s*n\.?\s*d\.?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),
    ("simple", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),
    ("simple-loose", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+(?:\s+[A-Z][A-Za-z'’\-]+){0,5}"
        r"\s*\.?\s*,\s*(?:19|20)\d{2}[a-z]?\s*\)"
    )),
    ("corporate-loose", re.compile(
        r"\(\s*[A-Z][^()]{0,120}?"
        r",\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?"
        r"\s*\)"
    )),
    ("corporate", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+(?:\s+[A-Z][A-Za-z'’\-]+){1,6}"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?\s*\)"
    )),
    ("narrative-multiword", re.compile(
        rf"\b(?!(?:{_NARRATIVE_LEAD_BLACKLIST})\b)"
        r"[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:[A-Z][A-Za-z'’\-]+|[A-Z]\.|of|the|to|in|on|at|for|and|by|with|from|a|an)){1,12}"
        r"\s+\((?:n\.?\s*d\.?|(?:19|20)\d{2}[a-z]?)"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?"
        r"\)"
    )),
    ("narrative", re.compile(
        r"\b[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:and|&)\s+[A-Z][A-Za-z'’\-]+|\s+et\s+al\.?)?"
        r"\s+\((?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\)"
    )),
    # Narrative form with n.d. — "Smith (n.d.)", "Smith et al. (n.d.)".
    # The existing "narrative" pattern only matches a year, and
    # "narrative-multiword" requires >= 2 author words.
    ("narrative-no-date", re.compile(
        r"\b[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:and|&)\s+[A-Z][A-Za-z'’\-]+|\s+et\s+al\.?)?"
        r"\s+\(n\.?\s*d\.?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\)"
    )),
]


_NON_AUTHOR_WORDS = frozenset({
    "figure", "fig", "figs", "table", "tbl", "tables",
    "section", "sect", "chapter", "ch", "part", "appendix", "app",
    "page", "pages", "p", "pp", "vol", "volume", "vols",
    "issue", "no", "number", "note", "notes", "see", "cf",
    "exhibit", "graph", "chart", "diagram", "image", "photo",
    "map", "plate", "equation", "eq", "eqs", "formula",
    "example", "ex", "source", "case", "step", "phase",
    "in", "of", "the", "a", "an", "and", "or", "by",
    "after", "before", "during", "since", "until", "between",
    "around", "about", "because", "although", "while", "when",
    "where", "here", "there", "this", "that", "these", "those",
    "however", "moreover", "furthermore", "therefore", "thus",
    "according", "stated", "noted", "reported", "says", "said",
    "argued", "argues", "noting", "stating", "arguing",
    "mentioned", "mentions", "discussed", "discusses",
})


def _author_is_plausible(matched: str) -> bool:
    m = re.match(r"[\(\s]*([A-Za-z'’\-]+)", matched)
    if not m:
        return True
    first = m.group(1).lower()
    if first not in _NON_AUTHOR_WORDS:
        return True

    rest = matched.lstrip("( \t\n")[len(m.group(1)):].lstrip(" \t\n")
    rest = rest.lstrip(".,;:-\u2013\u2014").lstrip()
    m2 = re.match(r"([A-Za-z'’\-]+)", rest)
    if not m2:
        return False
    second = m2.group(1).lower()
    if second in _NON_AUTHOR_WORDS:
        return False
    if not m2.group(1)[0].isupper():
        return False
    return True


def _iter_citable_blocks(document: Document):
    """Yield every Block that can contain an in-text citation.

    Includes both paragraph blocks (non-excluded regions) and every
    table cell block, so citations inside tables are detected and
    excluded just like citations in body paragraphs.
    """
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            continue
        yield block

    for table in document.tables:
        for cell in table.iter_cells():
            if not cell.spans:
                continue
            yield cell


def find_apa_citations(text: str) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    claimed: list[tuple[int, int]] = []
    for name, pattern in APA_PATTERNS:
        for m in pattern.finditer(text):
            start, end = m.span()
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
    for block in _iter_citable_blocks(document):
        for matched, pattern_name in find_apa_citations(block.text):
            if matched in by_text:
                by_text[matched].occurrences += 1
            else:
                by_text[matched] = Citation(
                    text=matched, occurrences=1, source="auto",
                    pattern_name=pattern_name,
                )
    return sorted(by_text.values(), key=lambda c: c.text.lower())


_TRAILING_PUNCT = '.,;:!?)]}\'"\u2019\u201d'


def _all_alnum_in_citations(
    block_text: str, s_start: int, s_end: int,
    claimed: list[tuple[int, int, str]],
) -> tuple[bool, str]:
    """Return (True, source) if every alnum char in [s_start, s_end)
    is inside *some* citation range. Returns (False, '') otherwise.

    Handles adjacent citations glued into one token:
    "(Smith, 2020)(Jones, 2021)" tokenises to a single span
    "2020)(Jones" that straddles both citation boundaries. The old
    fully-inside check couldn't exclude it.
    """
    source = ""
    for i in range(s_start, s_end):
        if not block_text[i].isalnum():
            continue
        for cs, ce, src in claimed:
            if cs <= i < ce:
                if not source:
                    source = src
                break
        else:
            return False, ""
    return True, source


def apply_citation_exclusions(document: Document) -> None:
    """Mark every span inside a detected citation as EXCLUDE.

    Handles:
      - Spans fully inside one citation.
      - Spans that straddle two adjacent citations (every alnum char
        is in *some* citation), e.g. "2020)(Exposure".
      - Explicit user toggles (I-USER-OVERRIDE / MANUAL) are preserved.

    Re-runnable: calling it again after resolve_all or
    _apply_user_overrides reclassifies table cells is safe and
    restores any citations those passes wiped.
    """
    items = [(c.text, c.source) for c in document.citations if c.text]
    if not items:
        return

    for block in _iter_citable_blocks(document):
        block_text = block.text
        claimed_ranges: list[tuple[int, int, str]] = []
        for ctext, csource in items:
            for m in re.finditer(re.escape(ctext), block_text, re.IGNORECASE):
                claimed_ranges.append((m.start(), m.end(), csource))
        if not claimed_ranges:
            continue

        for span in block.spans:
            if span.decision_record.rule_id in ("I-USER-OVERRIDE", "MANUAL"):
                continue

            s_start = span.source.char_start
            s_end = span.source.char_end

            trimmed_end = s_end
            while (trimmed_end > s_start
                   and block_text[trimmed_end - 1] in _TRAILING_PUNCT):
                trimmed_end -= 1
            if trimmed_end <= s_start:
                continue

            # --- Pass 1: fully inside a single citation --------------
            matched_source: str | None = None
            for c_start, c_end, c_source in claimed_ranges:
                if s_start >= c_start and trimmed_end <= c_end:
                    matched_source = c_source
                    break

            # --- Pass 2: every alnum char inside *some* citation -----
            if matched_source is None:
                all_in, source = _all_alnum_in_citations(
                    block_text, s_start, trimmed_end, claimed_ranges,
                )
                if all_in:
                    matched_source = source or "auto"

            if matched_source is None:
                continue

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
            span.tags.add(f"citation:{matched_source}")


def find_missed_candidates(document: Document) -> list[str]:
    already = {c.text.lower() for c in document.citations}
    candidates: list[str] = []
    seen: set[str] = set()
    loose = re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+[^()]{0,60}?(?:19|20)\d{2}[^()]{0,30}?\)"
    )
    for block in _iter_citable_blocks(document):
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
    if not needle:
        return 0
    total = 0
    for block in _iter_citable_blocks(document):
        total += sum(1 for _ in re.finditer(re.escape(needle), block.text,
                                            re.IGNORECASE))
    return total