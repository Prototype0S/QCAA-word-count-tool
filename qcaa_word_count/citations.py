"""APA 7 in-text citation detection and exclusion.

Walks BOTH paragraph blocks AND table cell blocks. This was previously
missing — a citation inside a table (e.g. the "Already existing target
audience ... (Australian Bureau of Statistics, 2021)" bullet in a SWOT
table) was silently counted because collect_citations and
apply_citation_exclusions only iterated document.blocks.

Covers all standard APA 7 in-text forms:

    Parenthetical:
        (Smith, 2020)
        (Smith, 2020, p. 170) / (Smith, 2020, pp. 170-175)
        (Smith & Jones, 2020)
        (Smith, Jones, & Brown, 2020)
        (Smith, Jones, & Brown, 2020, p. 170)
        (Smith et al., 2020) / (Smith et al., 2020, p. 170)
        (Smith, n.d.) / (Smith, n.d., p. 170)
        (Smith, 2020a) / (Smith, 2020a, 2020b)
        (Smith, 2020; Jones, 2021)
        (Smith, 2020, p. 170; Jones, 2021, p. 45)
        (Smith, 2020, as cited in Jones, 2021)
        (Australian Bureau of Statistics, 2026)
        (A Guide to All Schools in Brisbane, Queensland - Australian Schools
            Directory, 2021)
        (Blue Cards for Legal Professionals - Qld.Gov.Au, 2026)

    Narrative:
        Smith (2020) / Smith (2020, p. 170)
        Smith and Jones (2020) / Smith et al. (2020)
        According to the AKG First Job to Future Career Report Australia
            (n.d.), 78% of young people...
        The Australian Bureau of Statistics (2026) reports...

KNOWN EDGE CASES

1. Reference-list entries that appear in *body* text may match the
   narrative pattern as "Association (2025)". In a real essay, reference
   entries live inside the References region, which is excluded
   wholesale, so this does not affect actual submissions.

2. Possessive narrative forms ("Smith's (2020) study") are not
   detected. Students typically write "Smith (2020) argues" instead.

3. Non-standard paraphrase titles in the body may not match the exact
   title in the reference list. Intentional — cross-referencing between
   body and reference is unreliable because students often shorten or
   rephrase the title in-text.
"""
from __future__ import annotations

import re

from .model import Citation, Decision, DecisionRecord, Document, Region, SpanType


# ---------------------------------------------------------------------------
# Prose connectors that may precede a narrative citation but are NOT
# part of the citation itself. Used as a negative lookahead so the
# narrative pattern starts at the author, not at "According to".
# ---------------------------------------------------------------------------

_NARRATIVE_LEAD_BLACKLIST = (
    "According|As|However|Therefore|Thus|Moreover|Furthermore|"
    "Additionally|Although|While|When|Where|Since|Because|Despite|"
    "Besides|Also|In|Of|On|At|By|For|From|With|This|These|That|Those|It"
)


# ---------------------------------------------------------------------------
# Patterns
#
# Order matters: more specific patterns first so they claim their range
# before broader patterns try to match a substring of it.
# ---------------------------------------------------------------------------

APA_PATTERNS: list[tuple[str, re.Pattern]] = [
    # (Smith, 2020, p. 170) / (Smith, Jones, & Brown, 2020, p. 170)
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

    # (Smith, 2020a, 2020b)
    ("same-author-multi-year", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*(?:19|20)\d{2}[a-z]?)+"
        r"\s*\)"
    )),

    # (Smith, 2020; Jones, 2021) / with optional page locators
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

    # (Smith, 2020, as cited in Jones, 2021)
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

    # (Smith & Jones, 2020) / (Smith, Jones, & Brown, 2020)
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

    # (Smith, n.d.) / (Smith, n.d., p. 170)
    ("no-date", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"(?:\s+et\s+al\.?)?\s*,\s*n\.?\s*d\.?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),

    # (Smith, 2020)
    ("simple", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),

    # (Smith. , 2020) — malformed from citation generators
    ("simple-loose", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+(?:\s+[A-Z][A-Za-z'’\-]+){0,5}"
        r"\s*\.?\s*,\s*(?:19|20)\d{2}[a-z]?\s*\)"
    )),

    # [FIX 18e] Corporate / title-style author with lowercase particles
    # and internal punctuation.
    ("corporate-loose", re.compile(
        r"\(\s*[A-Z][^()]{0,120}?"
        r",\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?"
        r"\s*\)"
    )),

    # (Australian Bureau of Statistics, 2026)
    ("corporate", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+(?:\s+[A-Z][A-Za-z'’\-]+){1,6}"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?\s*\)"
    )),

    # [FIX 19] Multi-word narrative: Org/Title (year) / Org/Title (n.d.)
    # Catches:
    #   "According to the AKG First Job to Future Career Report Australia
    #    (n.d.), 78%..."
    #   "The Australian Bureau of Statistics (2026) reports..."
    # The negative lookahead prevents "According to" from being swallowed.
    ("narrative-multiword", re.compile(
        rf"\b(?!(?:{_NARRATIVE_LEAD_BLACKLIST})\b)"
        r"[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:[A-Z][A-Za-z'’\-]+|[A-Z]\.|of|the|to|in|on|at|for|and|by|with|from|a|an)){1,12}"
        r"\s+\((?:n\.?\s*d\.?|(?:19|20)\d{2}[a-z]?)"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?"
        r"\)"
    )),

    # Smith (2020) / Smith (2020, p. 170) — single-word narrative
    ("narrative", re.compile(
        r"\b[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:and|&)\s+[A-Z][A-Za-z'’\-]+|\s+et\s+al\.?)?"
        r"\s+\((?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\)"
    )),
]


# ---------------------------------------------------------------------------
# Author-plausibility check
# ---------------------------------------------------------------------------

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
    """Reject matches whose leading token is a non-author noun.

    A match is a plausible author/title if:
      - the first word is not blacklisted, OR
      - the first word is blacklisted but the second word is a
        capitalised non-blacklisted word.
    """
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


# ---------------------------------------------------------------------------
# Citable-block iterator
#
# [FIX 20] THE BUG FIX. Previously collect_citations and
# apply_citation_exclusions only iterated document.blocks (paragraph
# blocks). A citation inside a table cell (SWOT table, data tables,
# comparison tables) was silently ignored — never detected, never
# excluded.
#
# _iter_citable_blocks yields BOTH paragraph blocks (whose region is
# not excluded) AND every table cell block. Table cells don't have a
# region_record in the same way; their exclusion is decided at the
# span level by apply_table_answer, so we yield them unconditionally
# and let the span-level logic handle it.
# ---------------------------------------------------------------------------

def _iter_citable_blocks(document: Document):
    """Yield every Block that can contain an in-text citation."""
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


# ---------------------------------------------------------------------------
# Pattern matching
# ---------------------------------------------------------------------------

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
    """Detect citations in paragraphs AND table cells."""
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


# ---------------------------------------------------------------------------
# Exclusion
# ---------------------------------------------------------------------------

_TRAILING_PUNCT = '.,;:!?)]}\'"\u2019\u201d'


def apply_citation_exclusions(document: Document) -> None:
    """Mark every span inside a detected citation as EXCLUDE.

    Walks paragraphs AND table cells. This runs multiple times in the
    pipeline (once during _prepare, again at the end of resolve_citations)
    so that any citation that got re-classified by apply_table_answer
    during resolve_tables is restored.
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
            s_start = span.source.char_start
            s_end = span.source.char_end

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


# ---------------------------------------------------------------------------
# Fallback candidate prompting
# ---------------------------------------------------------------------------

def find_missed_candidates(document: Document) -> list[str]:
    """Parentheticals containing a 4-digit year that weren't auto-excluded.

    Walks paragraphs AND table cells. Deduplicated by case-insensitive
    value so the user isn't offered the same string twice.
    """
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
    """Count case-insensitive occurrences of a substring in countable body.

    Walks paragraphs AND table cells so that manual entries for table
    citations register the correct occurrence count.
    """
    if not needle:
        return 0
    total = 0
    for block in _iter_citable_blocks(document):
        total += sum(1 for _ in re.finditer(re.escape(needle), block.text,
                                            re.IGNORECASE))
    return total