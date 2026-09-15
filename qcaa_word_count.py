#!/usr/bin/env python3
"""
QCAA Word Count Tool (v3, Option B + Tables + Citations)
=========================================================

Run with:
    python qcaa_word_count.py                    # opens file picker
    python qcaa_word_count.py path/to/file.docx  # runs on a specific file
    python qcaa_word_count.py --explain          # shows QCAA rules, exits

Requires:
    pip install python-docx

WHAT THIS VERSION DOES
  - File picker (or path from command line)
  - Parses paragraphs, headings, and tables from a .docx
  - Excludes title page / front matter (before the first heading)
  - Excludes contents / abstract / references sections
  - Asks about appendixes (they're only excluded if supplementary-only)
  - Counts headings and body prose
  - Excludes numbers and symbols (tokens with digits but no letters)
  - Classifies tables and asks the user about each one
  - Detects APA 7 in-text citations and offers a manual fallback
  - Prints a report with a table and citation breakdown

WHAT IT DOES NOT DO YET
  - Equations and calculations outside tables
  - Math objects (OMML) inside paragraphs or table cells
  - Footnotes / endnotes
  - Text boxes / SmartArt
  - Manual [[QCAA_EXCLUDE]] markers
  - Saved decisions between runs
  - Colour output or GUI review screen
  - Audit report export

CODE LAYOUT (each numbered section is a distinct pipeline stage)
  Section 1  — Data model (shapes of data)
  Section 2  — Helpers (small pure functions)
  Section 3  — Ingest (.docx -> Document)
  Section 4  — Structure detection (assign regions to blocks)
  Section 5  — Decisions (regional + span-level)
  Section 5b — Table classification
  Section 5c — Citation detection (APA 7)
  Section 6  — Interactive flag resolution
  Section 7  — Count (walk the decision map)
  Section 8  — Report
  Section 9  — File picker
  Section 10 — Entry point + --explain

If something breaks, the section label tells you which stage to look at.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional, Literal

try:
    import docx
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("This tool needs python-docx.\nInstall it with: pip install python-docx")


# ============================================================================
# SECTION 1 — DATA MODEL
# ============================================================================

class BlockType(str, Enum):
    PARAGRAPH   = "paragraph"
    HEADING     = "heading"
    LIST_ITEM   = "list_item"
    TABLE_CELL  = "table_cell"
    FOOTNOTE    = "footnote"
    ENDNOTE     = "endnote"
    CAPTION     = "caption"
    DRAWING     = "drawing"
    TEXT_BOX    = "text_box"


class SpanType(str, Enum):
    WORD         = "word"
    ABBREVIATION = "abbreviation"
    NUMBER       = "number"
    UNIT         = "unit"
    SYMBOL       = "symbol"
    EQUATION     = "equation"
    CALCULATION  = "calculation"
    CITATION     = "citation"
    QUOTATION    = "quotation"
    PUNCTUATION  = "punctuation"
    WHITESPACE   = "whitespace"
    UNKNOWN      = "unknown"


class Region(str, Enum):
    BODY            = "body"
    TITLE           = "title"
    TITLE_PAGE      = "title_page"
    CONTENTS        = "contents"
    ABSTRACT        = "abstract"
    REFERENCES      = "references"
    APPENDIX        = "appendix"
    VISUAL_TEXT     = "visual_text"
    VISUAL_NON_TEXT = "visual_non_text"
    BLANK           = "blank"
    UNKNOWN         = "unknown"


class Decision(str, Enum):
    COUNT     = "count"
    EXCLUDE   = "exclude"
    FLAG      = "flag"
    UNDECIDED = "undecided"


class TableClass(str, Enum):
    INFORMATION  = "information"
    CALCULATION  = "calculation"
    RAW_DATA     = "raw_data"
    AMBIGUOUS    = "ambiguous"


class TableAnswer(str, Enum):
    COUNT_ALL    = "count_all"
    EXCLUDE_ALL  = "exclude_all"
    HEADERS_ONLY = "headers_only"


@dataclass(frozen=True)
class SourceLocation:
    block_index: int
    char_start: int
    char_end: int
    run_index: Optional[int] = None

    def __post_init__(self) -> None:
        if self.char_end < self.char_start:
            raise ValueError(
                f"SourceLocation char_end ({self.char_end}) "
                f"< char_start ({self.char_start})"
            )


@dataclass
class DecisionRecord:
    decision: Decision = Decision.UNDECIDED
    rule_id: Optional[str] = None
    confidence: Optional[float] = None
    reason: Optional[str] = None
    source: Literal["auto", "manual", "default"] = "auto"

    auto_decision: Optional[Decision] = None
    auto_rule_id: Optional[str] = None
    auto_confidence: Optional[float] = None

    def as_manual_override(self, new_decision: Decision, reason: str = "") -> None:
        if self.source == "auto":
            self.auto_decision = self.decision
            self.auto_rule_id = self.rule_id
            self.auto_confidence = self.confidence
        self.decision = new_decision
        self.rule_id = "MANUAL"
        self.confidence = 1.0
        self.reason = reason or "Manually overridden by user"
        self.source = "manual"

    @property
    def is_resolved(self) -> bool:
        return self.decision in (Decision.COUNT, Decision.EXCLUDE)

    @property
    def needs_review(self) -> bool:
        return self.decision in (Decision.FLAG, Decision.UNDECIDED)


@dataclass
class Span:
    text: str
    source: SourceLocation
    span_type: SpanType = SpanType.UNKNOWN
    decision_record: DecisionRecord = field(default_factory=DecisionRecord)
    tags: set[str] = field(default_factory=set)

    @property
    def decision(self) -> Decision:
        return self.decision_record.decision

    @property
    def counts_as_word(self) -> bool:
        return self.decision_record.decision == Decision.COUNT


@dataclass
class Block:
    block_type: BlockType
    spans: list[Span] = field(default_factory=list)
    region: Region = Region.UNKNOWN
    region_record: DecisionRecord = field(default_factory=DecisionRecord)
    block_index: int = 0
    style_name: Optional[str] = None
    math_object_count: int = 0
    raw_text: str = ""     
    is_visual: bool = False
    @property
    def text(self) -> str:
        if self.raw_text:
            return self.raw_text
        return "".join(s.text for s in self.spans)

    @property
    def is_regionally_excluded(self) -> bool:
        return (
            self.region_record.decision == Decision.EXCLUDE
            and self.region_record.source != "manual"
        )

    def iter_word_spans(self):
        for s in self.spans:
            if s.counts_as_word:
                yield s


@dataclass
class Table:
    rows: list[list[Block]] = field(default_factory=list)
    table_index: int = 0
    block_index: int = 0

    suggestion: TableClass = TableClass.AMBIGUOUS
    suggestion_confidence: float = 0.0

    answer: Optional[TableAnswer] = None
    answer_record: DecisionRecord = field(default_factory=DecisionRecord)
    header_row_index: int = 0

    def iter_cells(self):
        for row in self.rows:
            for cell in row:
                yield cell

    @property
    def word_count(self) -> int:
        return sum(len(cell.spans) for cell in self.iter_cells())


@dataclass
class Citation:
    text: str
    occurrences: int = 0
    source: Literal["auto", "manual"] = "auto"
    pattern_name: Optional[str] = None

    @property
    def word_count(self) -> int:
        return len(tokenize(self.text))


@dataclass
class Document:
    blocks: list[Block] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    filename: Optional[str] = None
    profile: str = "default"

    citations: list[Citation] = field(default_factory=list)
    footnotes: list[Footnote] = field(default_factory=list)

    def all_spans(self):
        for b in self.blocks:
            for s in b.spans:
                yield s
        for t in self.tables:
            for c in t.iter_cells():
                for s in c.spans:
                    yield s

@dataclass
class Footnote:
    """A footnote or endnote extracted from the docx."""
    footnote_id: str
    text: str
    kind: Literal["footnote", "endnote"] = "footnote"

    suggestion: Literal["bibliographic", "commentary", "ambiguous"] = "ambiguous"
    suggestion_confidence: float = 0.0

    answer: Optional[Literal["exclude", "count"]] = None
    answer_record: DecisionRecord = field(default_factory=DecisionRecord)

    @property
    def word_count(self) -> int:
        return len(tokenize(self.text))
    
# ============================================================================
# SECTION 2 — HELPERS
# ============================================================================

def _is_inside_math(node) -> bool:
    """True if this XML node sits inside an m:oMath or m:oMathPara element."""
    for ancestor in node.iterancestors():
        if ancestor.tag in (qn("m:oMath"), qn("m:oMathPara")):
            return True
    return False


def _needs_space(prev_chunk: str, next_chunk: str) -> bool:
    """Word sometimes stores adjacent runs without a space between them.
    Returns True if we should insert a space when stitching chunks."""
    if not prev_chunk or not next_chunk:
        return False
    last = prev_chunk[-1]
    first = next_chunk[0]
    if last.isspace() or first.isspace():
        return False
    return last.isalnum() and first.isalnum()


def _extract_text_from_element(element) -> tuple[str, int]:
    """Walk an XML element and return (text, math_count).

    Handles w:t (text), w:tab, w:br/w:cr, and skips math objects.
    Includes text box content because w:txbxContent uses the same w:t
    elements internally.
    """
    parts: list[str] = []
    math_count = 0
    seen_math_roots: set[int] = set()

    for node in element.iter():
        tag = node.tag

        if tag in (qn("m:oMath"), qn("m:oMathPara")):
            node_id = id(node)
            if node_id not in seen_math_roots:
                seen_math_roots.add(node_id)
                math_count += 1
            continue

        if _is_inside_math(node):
            continue

        if tag == qn("w:t"):
            text = node.text or ""
            if parts and text and _needs_space(parts[-1], text):
                parts.append(" ")
            parts.append(text)
        elif tag == qn("w:tab"):
            parts.append("\t")
        elif tag in (qn("w:br"), qn("w:cr")):
            parts.append("\n")

    return "".join(parts), math_count


def paragraph_text(para) -> tuple[str, int]:
    """Reconstruct paragraph text from the XML, preserving spaces.
    Returns (text, math_object_count)."""
    return _extract_text_from_element(para._element)


def paragraph_text_boxes(para) -> list[tuple[str, int]]:
    """Extract text box content nested inside this paragraph.

    Word stores text-box content in <w:txbxContent> elements. These are
    children of drawings/shapes anchored in a paragraph. We extract each
    text box separately so it becomes its own block.
    """
    results: list[tuple[str, int]] = []
    for node in para._element.iter():
        if node.tag == qn("w:txbxContent"):
            text, math_count = _extract_text_from_element(node)
            if text.strip() or math_count:
                results.append((text, math_count))
    return results


def cell_text(cell) -> tuple[str, int]:
    """Reconstruct table cell text reliably. Returns (text, math_count)."""
    parts: list[str] = []
    total_math = 0
    for para in cell.paragraphs:
        t, m = paragraph_text(para)
        parts.append(t)
        total_math += m
    return "\n".join(parts), total_math


def tokenize(text: str) -> list[str]:
    return [t for t in re.findall(r"\S+", text) if any(c.isalnum() for c in t)]


def is_numeric_token(tok: str) -> bool:
    has_letter = any(c.isalpha() for c in tok)
    has_digit = any(c.isdigit() for c in tok)
    return has_digit and not has_letter


def heading_level_from_style(style_name: str) -> int | None:
    m = re.match(r"Heading (\d+)", style_name or "")
    return int(m.group(1)) if m else None


def make_spans_for_text(text: str, block_index: int) -> list[Span]:
    spans: list[Span] = []
    cursor = 0
    for token in re.findall(r"\S+", text):
        start = text.find(token, cursor)
        end = start + len(token)
        cursor = end
        spans.append(Span(
            text=token,
            source=SourceLocation(
                block_index=block_index,
                char_start=start,
                char_end=end,
            ),
        ))
    return spans


def describe_suggestion(table: Table) -> str:
    s = table.suggestion
    if s == TableClass.INFORMATION:
        return "this looks like a table of information → count the whole table"
    if s == TableClass.CALCULATION:
        return "this looks like calculation working → exclude the whole table"
    if s == TableClass.RAW_DATA:
        return "this looks like raw or processed data → exclude the whole table"
    return "the tool could not confidently classify this table"


def describe_confidence(conf: float) -> str:
    if conf >= 0.75:
        return "high confidence"
    if conf >= 0.5:
        return "medium confidence"
    return "low confidence"
# ============================================================================
# SECTION 3 — INGEST
# ============================================================================

def ingest(path: str) -> Document:
    doc = docx.Document(path)
    document = Document(filename=os.path.basename(path))

    block_index = 0

    for para in doc.paragraphs:
        text, math_count = paragraph_text(para)
        if text.strip() or math_count:
            style_name = (para.style.name or "") if para.style else ""
            level = heading_level_from_style(style_name)
            block_type = BlockType.HEADING if level else BlockType.PARAGRAPH

            block = Block(
                block_type=block_type,
                block_index=block_index,
                style_name=style_name,
            )
            block.raw_text = text
            block.spans = make_spans_for_text(text, block_index=block_index)
            block.math_object_count = math_count
            document.blocks.append(block)
            block_index += 1

        # Extract text boxes nested in this paragraph (if any)
        for tb_text, tb_math in paragraph_text_boxes(para):
            if not tb_text.strip() and not tb_math:
                continue
            block = Block(
                block_type=BlockType.TEXT_BOX,
                block_index=block_index,
            )
            block.raw_text = tb_text
            block.spans = make_spans_for_text(tb_text, block_index=block_index)
            block.math_object_count = tb_math
            document.blocks.append(block)
            block_index += 1

    for t_idx, raw_table in enumerate(doc.tables):
        table = Table(table_index=t_idx)

        for r_idx, raw_row in enumerate(raw_table.rows):
            row_blocks: list[Block] = []
            for c_idx, raw_cell in enumerate(raw_row.cells):
                text, math_count = cell_text(raw_cell)
                synthetic_index = 100_000 + t_idx * 1000 + r_idx * 100 + c_idx
                cell_block = Block(
                    block_type=BlockType.TABLE_CELL,
                    block_index=synthetic_index,
                )
                cell_block.raw_text = text
                cell_block.spans = make_spans_for_text(text, block_index=synthetic_index)
                cell_block.math_object_count = math_count
                row_blocks.append(cell_block)
            table.rows.append(row_blocks)

        document.tables.append(table)

    return document

# ============================================================================
# SECTION 4 — STRUCTURE DETECTION
# ============================================================================

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


# ============================================================================
# SECTION 5 — DECISIONS (paragraphs)
# ============================================================================

def apply_regional_decisions(document: Document) -> None:
    """Set the region_record on every paragraph block, based on its region."""
    for block in document.blocks:
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
        else:
            block.region_record = DecisionRecord(
                decision=Decision.COUNT,
                rule_id="R-BODY",
                confidence=1.0,
                reason="Body content counts",
                source="auto",
            )


# ============================================================================
# Math / equation character sets
# ============================================================================

# Strong indicators that a token is an equation or mathematical expression
EQUATION_CHARS = set("=×÷^*≤≥≈√∫Σ∑∴∵±∓∂∇∝∞")
SUPERSCRIPT_DIGITS = set("⁰¹²³⁴⁵⁶⁷⁸⁹")
SUBSCRIPT_DIGITS = set("₀₁₂₃₄₅₆₇₈₉")
GREEK_LETTERS = set("αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ")
MATH_SYMBOLS = EQUATION_CHARS | {"→", "←", "⇒", "⇔", "∈", "∉", "∪", "∩", "°", "′", "″"}


def is_symbol_token(tok: str) -> bool:
    """A token with no letters and no digits — e.g. '+', '=', '→', '%'."""
    return not any(c.isalnum() for c in tok)


def is_equation_token(tok: str) -> bool:
    """A token that looks like inline equation working.

    Catches:
      - Tokens with strong operators (=, ×, ÷, ^, etc.)
      - Tokens with superscripts or subscripts
      - Tokens with Greek letters
      - Algebraic patterns like "0.81a", "2a", "1.8b"
      - Weaker cases: digits mixed with +, -, /, if the letters don't
        spell a real word
    """
    if any(op in tok for op in MATH_SYMBOLS):
        return True
    if any(c in SUPERSCRIPT_DIGITS for c in tok):
        return True
    if any(c in SUBSCRIPT_DIGITS for c in tok):
        return True
    if any(c in GREEK_LETTERS for c in tok):
        return True

    # Algebraic pattern: digits followed by a single lowercase letter
    # e.g. "2a", "0.81a", "1.8b"
    if re.fullmatch(r"-?\d+(?:\.\d+)?[a-z]", tok):
        return True

    # Weak operators: +, -, /, digits — exclude if the letters
    # don't form a recognisable word
    if any(op in tok for op in "+-/"):
        if any(c.isdigit() for c in tok):
            letters = "".join(c for c in tok if c.isalpha())
            if len(letters) >= 3:
                return False
            return True

    return False

def _is_single_letter_variable(tok: str) -> bool:
    """A single lowercase letter (a, b, c, h, k, x, y, n, ...) that is
    likely a variable, not a word."""
    return len(tok) == 1 and tok.isalpha() and tok.islower()


def _contains_bracket(tok: str) -> bool:
    """A token that contains any bracket-like character."""
    return any(c in "()[]{}" for c in tok)


def _is_equation_heavy_block(block: Block, threshold: float = 0.4) -> bool:
    """True if a large enough fraction of this block's tokens look
    equation-like. Used to decide whether single letters and
    bracketed tokens are equation fragments rather than words.

    Threshold 0.4 means 40%+ of tokens must be equation/number/symbol
    before the equation-heavy rules kick in.
    """
    if not block.spans:
        return False

    mathy = 0
    for span in block.spans:
        tok = span.text
        if (is_equation_token(tok)
                or is_symbol_token(tok)
                or is_numeric_token(tok)):
            mathy += 1

    return (mathy / len(block.spans)) >= threshold


def classify_and_decide_spans(document: Document) -> None:
    """Assign a span_type and decision to every span in every paragraph block.

    Rule of thumb: if it's a word, it counts. QCAA then carves out
    exceptions for numbers, symbols, equations, calculations, and
    single-letter variables inside equation-heavy text.
    """
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            for span in block.spans:
                span.decision_record = DecisionRecord(
                    decision=Decision.EXCLUDE,
                    rule_id="R-REGION",
                    confidence=1.0,
                    reason=f"Inherits exclusion from region '{block.region.value}'",
                    source="auto",
                )
            continue

        # Compute equation density once per block.
        equation_heavy = _is_equation_heavy_block(block)

        for span in block.spans:
            tok = span.text

            # Rule 1: strong equation indicators
            if is_equation_token(tok):
                span.span_type = SpanType.EQUATION
                span.decision_record = DecisionRecord(
                    decision=Decision.EXCLUDE,
                    rule_id="I-EQUATION",
                    confidence=0.9,
                    reason="Equations and calculations are excluded by QCAA rules",
                    source="auto",
                )

            # Rule 2: symbols
            elif is_symbol_token(tok):
                span.span_type = SpanType.SYMBOL
                span.decision_record = DecisionRecord(
                    decision=Decision.EXCLUDE,
                    rule_id="I-SYMBOL",
                    confidence=0.95,
                    reason="Symbols are excluded by QCAA rules",
                    source="auto",
                )

            # Rule 3: numbers
            elif is_numeric_token(tok):
                span.span_type = SpanType.NUMBER
                span.decision_record = DecisionRecord(
                    decision=Decision.EXCLUDE,
                    rule_id="I-NUMBER",
                    confidence=0.95,
                    reason="Numbers are excluded by QCAA rules",
                    source="auto",
                )

            # Rule 4: single lowercase letter in equation-heavy block → variable
            elif equation_heavy and _is_single_letter_variable(tok):
                span.span_type = SpanType.EQUATION
                span.decision_record = DecisionRecord(
                    decision=Decision.EXCLUDE,
                    rule_id="I-VARIABLE",
                    confidence=0.75,
                    reason="Single-letter variable inside equation-heavy text",
                    source="auto",
                )

            # Rule 5: bracketed token in equation-heavy block → equation fragment
            elif equation_heavy and _contains_bracket(tok):
                span.span_type = SpanType.EQUATION
                span.decision_record = DecisionRecord(
                    decision=Decision.EXCLUDE,
                    rule_id="I-EQUATION-FRAGMENT",
                    confidence=0.75,
                    reason="Bracketed token inside equation-heavy text",
                    source="auto",
                )

            # Rule 6: otherwise it's a word
            else:
                span.span_type = SpanType.WORD
                span.decision_record = DecisionRecord(
                    decision=Decision.COUNT,
                    rule_id="I-WORD",
                    confidence=1.0,
                    reason="Ordinary word",
                    source="auto",
                )
# ============================================================================
# SECTION 5b — TABLE CLASSIFICATION
# ============================================================================

CALC_OPERATOR_CHARS = set("+-*/=×÷^<>≤≥≈√")
PROSE_RUN_RE = re.compile(r"\b[a-z]{2,}(?:\s+[a-z]{2,}){2,}\b")
CAPTION_PATTERN = re.compile(
    r"^\s*(figure|fig\.?|table|tbl\.?|chart|graph|diagram|image|photo|map|exhibit|plate)"
    r"\s*[\dIVXivx]+[a-z]?\s*[:.,\u2013\u2014\-]\s*\S",
    re.IGNORECASE,
)
SOURCE_PATTERN = re.compile(
    r"^\s*source\s*[:.\u2013\u2014\-]\s*\S",
    re.IGNORECASE,
)
BYLINE_PATTERN = re.compile(
    r"^\s*(by|words by|written by|illustration by|illustrated by|photo by|images? by)\s+\S",
    re.IGNORECASE,
)

def classify_table(table: Table) -> tuple[TableClass, float]:
    total_cells = 0
    prose_cells = 0
    operator_cells = 0
    total_tokens = 0
    numeric_tokens = 0
    total_equation_chars = 0

    for cell in table.iter_cells():
        text = cell.text
        if not text.strip():
            continue
        total_cells += 1

        if PROSE_RUN_RE.search(text.lower()):
            prose_cells += 1

        cell_has_operator = False
        for ch in text:
            if ch in CALC_OPERATOR_CHARS:
                total_equation_chars += 1
                cell_has_operator = True
        if cell_has_operator:
            operator_cells += 1

        for tok in tokenize(text):
            total_tokens += 1
            if is_numeric_token(tok):
                numeric_tokens += 1

    if total_cells == 0 or total_tokens == 0:
        return TableClass.AMBIGUOUS, 0.0

    prose_ratio = prose_cells / total_cells
    operator_ratio = operator_cells / total_cells
    numeric_ratio = numeric_tokens / total_tokens

    # Strong equation-character presence boosts calculation score.
    # Even if cells contain some prose ("Therefore mean = 159s"),
    # the "=" presence is a strong signal.
    equation_char_boost = min(total_equation_chars / 10.0, 0.5)

    scores = {
        TableClass.INFORMATION: prose_ratio,
        TableClass.CALCULATION: min(operator_ratio + equation_char_boost, 1.0),
        TableClass.RAW_DATA: numeric_ratio,
    }
    best_class = max(scores, key=lambda k: scores[k])
    best_score = scores[best_class]

    runner_up = sorted(scores.values(), reverse=True)[1]
    margin = best_score - runner_up

    confidence = max(0.0, min(1.0, (best_score * 0.6) + (margin * 0.4)))

    if best_score < 0.4:
        return TableClass.AMBIGUOUS, confidence

    return best_class, confidence

def suggest_table_answer(table: Table) -> TableAnswer:
    if table.suggestion == TableClass.INFORMATION:
        return TableAnswer.COUNT_ALL
    return TableAnswer.EXCLUDE_ALL


def table_has_strong_default(table: Table) -> bool:
    return (table.suggestion != TableClass.AMBIGUOUS
            and table.suggestion_confidence >= 0.5)

def looks_like_caption(block: Block) -> tuple[bool, str]:
    """Return (is_caption, kind). kind is 'caption', 'source', 'by-line', or ''."""
    text = block.text.strip()

    # Word style "Caption" is definitive
    if block.style_name and block.style_name.lower().startswith("caption"):
        return True, "caption"

    if CAPTION_PATTERN.match(text):
        return True, "caption"
    if SOURCE_PATTERN.match(text):
        return True, "source"
    if BYLINE_PATTERN.match(text):
        return True, "by-line"

    return False, ""
# ============================================================================
# SECTION 5c — CITATION DETECTION (APA 7)
# ============================================================================

APA_PATTERNS: list[tuple[str, re.Pattern]] = [

    ("with page numbers", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+(?:\s+(?:&|and)\s+[A-Z][A-Za-z'’\-]+)?"
        r"(?:\s+et\s+al\.?)?\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?\s*\)"
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

    ("simple", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+\s*,\s*(?:19|20)\d{2}[a-z]?"
        r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?\s*\)"
    )),

    ("corporate", re.compile(
        r"\(\s*[A-Z][A-Za-z'’\-]+(?:\s+[A-Z][A-Za-z'’\-]+){1,6}"
        r"\s*,\s*(?:19|20)\d{2}[a-z]?\s*\)"
    )),

    ("narrative", re.compile(
        r"\b[A-Z][A-Za-z'’\-]+"
        r"(?:\s+(?:and|&)\s+[A-Z][A-Za-z'’\-]+|\s+et\s+al\.?)?"
        r"\s+\((?:19|20)\d{2}[a-z]?\)"
    )),
]


def find_apa_citations(text: str) -> list[tuple[str, str]]:
    found: list[tuple[str, str]] = []
    claimed: list[tuple[int, int]] = []

    for name, pattern in APA_PATTERNS:
        for m in pattern.finditer(text):
            start, end = m.span()
            if any(s < end and start < e for s, e in claimed):
                continue
            claimed.append((start, end))
            found.append((m.group(0), name))

    return found


def collect_citations(document: Document) -> list[Citation]:
    by_text: dict[str, Citation] = {}

    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            continue

        text = block.text
        for matched, pattern_name in find_apa_citations(text):
            if matched in by_text:
                by_text[matched].occurrences += 1
            else:
                by_text[matched] = Citation(
                    text=matched,
                    occurrences=1,
                    source="auto",
                    pattern_name=pattern_name,
                )

    return sorted(by_text.values(), key=lambda c: c.text.lower())


def apply_citation_exclusions(document: Document) -> None:
    """Mark every span inside a detected citation as EXCLUDE.

    Tags each span with the source of the citation that claimed it
    ('citation:auto' or 'citation:manual') so the report can attribute
    exclusions correctly.
    """
    # Build a list of (citation_string, source) to search for
    citation_items = [(c.text, c.source) for c in document.citations if c.text]

    if not citation_items:
        return

    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue

        block_text = block.text

        # Find every (start, end, source) range in this block
        claimed_ranges: list[tuple[int, int, str]] = []
        for ctext, csource in citation_items:
            start = 0
            while True:
                idx = block_text.find(ctext, start)
                if idx == -1:
                    break
                claimed_ranges.append((idx, idx + len(ctext), csource))
                start = idx + len(ctext)

        if not claimed_ranges:
            continue

        for span in block.spans:
            s_start = span.source.char_start
            s_end = span.source.char_end
            for c_start, c_end, c_source in claimed_ranges:
                if s_start >= c_start and s_end <= c_end:
                    span.span_type = SpanType.CITATION
                    span.decision_record = DecisionRecord(
                        decision=Decision.EXCLUDE,
                        rule_id="I-CITATION",
                        confidence=0.9,
                        reason="In-text citation (excluded by QCAA rules)",
                        source="auto",
                    )
                    # Tag with the citation's source for reporting
                    span.tags.discard("citation:auto")
                    span.tags.discard("citation:manual")
                    span.tags.add(f"citation:{c_source}")
                    break
                
def find_missed_candidates(document: Document) -> list[str]:
    already = {c.text for c in document.citations}
    candidates: list[str] = []

    loose_year_paren = re.compile(r"\([^()]{1,80}?(?:19|20)\d{2}[^()]{0,40}?\)")

    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            continue

        text = block.text
        for m in loose_year_paren.finditer(text):
            matched = m.group(0)
            if matched in already:
                continue
            if matched in candidates:
                continue
            candidates.append(matched)

    return candidates

# ============================================================================
# SECTION 5d — CAPTION / VISUAL ELEMENT DETECTION
# ============================================================================

QCAA_CAPTION_HELP = """
------------------------------------------------------------------
CAPTIONS AND VISUAL ELEMENTS (QCAA rule)
------------------------------------------------------------------
QCAA excludes "visual elements associated with the written response"
-- by-lines, banners, captions and call-outs that are the visual
elements of written genres suitable for print or online publication
(literary article, blog, essay, column).

HOWEVER: in science reports and PSMTs, captions and figure labels
often identify variables, conditions, or what a graph shows. Markers
generally treat those as information, and they count.

When unsure, ask your teacher. The safest default is to exclude,
because that is what QCAA's rule says on its face.
------------------------------------------------------------------
"""


def detect_captions(document: Document) -> None:
    """Mark blocks that look like captions/by-lines/source lines."""
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            continue

        is_cap, kind = looks_like_caption(block)
        if is_cap:
            block.is_visual = True
            # Don't use FLAG here — FLAG is reserved for region-level
            # decisions (appendixes). Captions are their own category and
            # are resolved by resolve_captions_interactive().
            # Mark the region_record as "undecided" so count_words treats
            # it as a flag until the caption resolver runs.
            block.region_record = DecisionRecord(
                decision=Decision.UNDECIDED,
                rule_id=f"R-{kind.upper().replace('-', '_')}",
                confidence=0.6,
                reason=f"Possible {kind}: QCAA treats visual elements as excluded",
                source="auto",
            )


def resolve_captions_interactive(document: Document,
                                  auto_accept_remaining: bool = False) -> bool:
    """Prompt the user about each caption. Returns the updated
    auto_accept_remaining flag."""
    caption_blocks = [b for b in document.blocks if b.is_visual]

    if not caption_blocks:
        return auto_accept_remaining

    print()
    print("=" * 60)
    print(f"{len(caption_blocks)} caption/visual element(s) detected.")
    print("=" * 60)

    for i, block in enumerate(caption_blocks, start=1):
        if block.region_record.decision in (Decision.COUNT, Decision.EXCLUDE):
            continue  # already resolved

        if auto_accept_remaining:
            block.region_record.as_manual_override(
                Decision.EXCLUDE,
                reason="Auto-accepted suggestion (user chose Accept All)",
            )
            continue

        print()
        print("-" * 60)
        print(f"Caption {i} of {len(caption_blocks)}")
        snippet = block.text.strip()
        if len(snippet) > 120:
            snippet = snippet[:117] + "..."
        print(f'  "{snippet}"')
        print()
        print("  QCAA's rule: captions and other visual elements of written")
        print("  genres are excluded from the word count.")
        print()
        print("  But: in science reports and PSMTs, captions often identify")
        print("  variables or conditions, which some markers count as")
        print("  information.")

        choice = _menu(
            "How should this caption be counted?",
            [
                ("exclude", "Exclude it (QCAA's default: visual elements excluded)"),
                ("count",   "Count it (it names variables or carries information)"),
                ("help",    "Explain the QCAA rule for captions"),
                ("all",     "Accept suggestions for all remaining captions"),
            ],
            default_key="exclude",
        )

        if choice == "help":
            print(QCAA_CAPTION_HELP)
            choice = _menu(
                "How should this caption be counted?",
                [
                    ("exclude", "Exclude it (QCAA's default)"),
                    ("count",   "Count it (it names variables or carries information)"),
                    ("all",     "Accept suggestions for all remaining captions"),
                ],
                default_key="exclude",
            )

        if choice == "all":
            auto_accept_remaining = True
            block.region_record.as_manual_override(
                Decision.EXCLUDE,
                reason="Auto-accepted suggestion (user chose Accept All)",
            )
        elif choice == "exclude":
            block.region_record.as_manual_override(
                Decision.EXCLUDE,
                reason="User chose to exclude caption",
            )
        else:
            block.region_record.as_manual_override(
                Decision.COUNT,
                reason="User chose to count caption",
            )

    return auto_accept_remaining
# ============================================================================
# SECTION 5e — FOOTNOTE / ENDNOTE DETECTION
# ============================================================================

BIBLIOGRAPHIC_HINTS = (
    re.compile(r"(?:19|20)\d{2}"),                      # a year
    re.compile(r"\b(pp?\.|vol\.|no\.|ed\.|doi|isbn|issn)\b", re.IGNORECASE),
    re.compile(r"\b(press|journal|university|thesis|phd|edition|trans\.|"
               r"vols?|pub\.|publishing|http|www\.|accessed)\b", re.IGNORECASE),
    re.compile(r"\bet\s+al\.", re.IGNORECASE),
)


def classify_footnote(text: str) -> tuple[str, float]:
    """Classify a footnote as bibliographic vs commentary.

    Returns (kind, confidence). kind is 'bibliographic', 'commentary',
    or 'ambiguous'.
    """
    score = 0
    for pattern in BIBLIOGRAPHIC_HINTS:
        if pattern.search(text):
            score += 1

    # Commentary signals: prose-y sentences with personal commentary.
    # A footnote that is just a citation has almost no lowercase words
    # outside the citation. A commentary has multiple sentences.
    prose_sentences = len(re.findall(r"[a-z]{3,}\s+[a-z]{3,}\s+[a-z]{3,}", text))

    if score >= 2 and prose_sentences <= 1:
        return "bibliographic", 0.85
    if score <= 1 and prose_sentences >= 2:
        return "commentary", 0.75
    return "ambiguous", 0.4


def read_footnotes_from_docx(path: str) -> list[Footnote]:
    """Extract footnotes and endnotes from the docx package."""
    import zipfile
    from xml.etree import ElementTree as ET

    W_NS_LOCAL = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    footnotes: list[Footnote] = []

    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()

            for filename, tag, kind in (
                ("word/footnotes.xml", "footnote", "footnote"),
                ("word/endnotes.xml", "endnote", "endnote"),
            ):
                if filename not in names:
                    continue
                try:
                    root = ET.fromstring(z.read(filename))
                except ET.ParseError:
                    continue

                for note in root.iter(f"{W_NS_LOCAL}{tag}"):
                    ntype = note.get(f"{W_NS_LOCAL}type")
                    if ntype in ("separator", "continuationSeparator", "continuationNotice"):
                        continue
                    note_id = note.get(f"{W_NS_LOCAL}id") or ""
                    text = "".join(t.text or "" for t in note.iter(f"{W_NS_LOCAL}t"))
                    text = text.strip()
                    if not text:
                        continue
                    footnotes.append(Footnote(
                        footnote_id=note_id,
                        text=text,
                        kind=kind,  # type: ignore
                    ))
    except (zipfile.BadZipFile, OSError):
        pass

    return footnotes


def apply_footnote_suggestions(document: Document) -> None:
    """Apply the classifier's suggestion to each footnote."""
    for fn in document.footnotes:
        kind, conf = classify_footnote(fn.text)
        fn.suggestion = kind  # type: ignore
        fn.suggestion_confidence = conf


def resolve_footnotes_interactive(document: Document,
                                   auto_accept_remaining: bool = False) -> bool:
    """Prompt the user about each footnote. Returns updated auto flag."""
    if not document.footnotes:
        return auto_accept_remaining

    print()
    print("=" * 60)
    print(f"{len(document.footnotes)} footnote/endnote(s) detected.")
    print("=" * 60)
    print()
    print("QCAA: footnotes/endnotes are counted UNLESS they are purely")
    print("bibliographical (i.e. just a reference).")

    for i, fn in enumerate(document.footnotes, start=1):
        if fn.answer is not None:
            continue

        if auto_accept_remaining:
            fn.answer = "exclude" if fn.suggestion == "bibliographic" else "count"
            fn.answer_record = DecisionRecord(
                decision=Decision.EXCLUDE if fn.answer == "exclude" else Decision.COUNT,
                rule_id="F-AUTO",
                confidence=0.6,
                reason="Auto-accepted suggestion",
                source="manual",
            )
            continue

        print()
        print("-" * 60)
        print(f"Footnote {i} of {len(document.footnotes)} ({fn.kind})")
        snippet = fn.text if len(fn.text) <= 200 else fn.text[:197] + "..."
        print(f'  "{snippet}"')
        print()

        suggestion = fn.suggestion
        if suggestion == "bibliographic":
            print(f"  Suggested: bibliographic → exclude (confidence {fn.suggestion_confidence:.2f})")
            default = "exclude"
        elif suggestion == "commentary":
            print(f"  Suggested: commentary → count (confidence {fn.suggestion_confidence:.2f})")
            default = "count"
        else:
            print("  Suggested: could not confidently classify")
            default = None

        choice = _menu(
            "How should this footnote be treated?",
            [
                ("exclude", "Exclude (it's purely a bibliographic reference)"),
                ("count",   "Count (it contains commentary or content)"),
                ("help",    "Explain the QCAA rule for footnotes"),
                ("all",     "Accept suggestions for all remaining footnotes"),
            ],
            default_key=default,
        )

        if choice == "help":
            print()
            print("QCAA rule for footnotes/endnotes:")
            print("  - Footnotes and endnotes are INCLUDED in the word count,")
            print("    UNLESS they are used purely for bibliographical purposes.")
            print("  - 'Bibliographical' means it's just a citation or reference.")
            print("  - If it contains any commentary, clarification, or prose,")
            print("    it counts.")
            choice = _menu(
                "How should this footnote be treated?",
                [
                    ("exclude", "Exclude (bibliographic only)"),
                    ("count",   "Count (contains commentary or content)"),
                    ("all",     "Accept suggestions for all remaining footnotes"),
                ],
                default_key=default,
            )

        if choice == "all":
            auto_accept_remaining = True
            suggested = "exclude" if fn.suggestion == "bibliographic" else "count"
            fn.answer = suggested
            fn.answer_record = DecisionRecord(
                decision=Decision.EXCLUDE if suggested == "exclude" else Decision.COUNT,
                rule_id="F-AUTO",
                confidence=0.6,
                reason="Auto-accepted suggestion (user chose Accept All)",
                source="manual",
            )
            continue

        fn.answer = choice  # type: ignore
        fn.answer_record = DecisionRecord(
            decision=Decision.EXCLUDE if choice == "exclude" else Decision.COUNT,
            rule_id="F-USER",
            confidence=1.0,
            reason=f"User chose '{choice}'",
            source="manual",
        )

    return auto_accept_remaining


# ============================================================================
# SECTION 6 — INTERACTIVE FLAG RESOLUTION
# ============================================================================

QCAA_TABLE_HELP = """
------------------------------------------------------------------
TABLES (QCAA rule)
------------------------------------------------------------------
  Option 1 — Count the whole table.
             Use this when the table contains information other
             than raw or processed data: prose in cells, explanatory
             labels beyond column headers, annotations, etc.

  Option 2 — Exclude the whole table.
             Use this for tables that contain only raw data
             (individual measurements), processed data (means,
             totals, percentages), or calculation working.

  Option 3 — Count the header row only, exclude the data rows.
             A middle-ground reading: the column labels
             ("Temperature", "Trial 1 (s)") count as information,
             but the data does not.

  Option 4 — Show this message again.
------------------------------------------------------------------
"""


def preview_table(table: Table, max_rows: int = 4, max_width: int = 70) -> str:
    lines = []
    for r_idx, row in enumerate(table.rows):
        if r_idx >= max_rows:
            lines.append(f"... ({len(table.rows) - max_rows} more row(s))")
            break
        cells = []
        for cell in row:
            t = cell.text.strip().replace("\n", " ")
            if len(t) > 20:
                t = t[:17] + "..."
            cells.append(t)
        line = " | ".join(cells)
        if len(line) > max_width:
            line = line[:max_width - 3] + "..."
        lines.append(line)
    return "\n".join(lines)


def _menu(question: str, options: list[tuple[str, str]],
          default_key: Optional[str] = None) -> str:
    print()
    print(question)
    for i, (key, label) in enumerate(options, start=1):
        marker = " (suggested)" if key == default_key else ""
        print(f"  [{i}] {label}{marker}")

    while True:
        default_hint = ""
        if default_key is not None:
            for i, (key, _) in enumerate(options, start=1):
                if key == default_key:
                    default_hint = f" [{i}]"
                    break
        try:
            raw = input(f"\nYour choice{default_hint}: ").strip()
        except EOFError:
            raw = ""

        if not raw and default_key is not None:
            return default_key

        if not raw.isdigit():
            print("Please type a number.")
            continue

        idx = int(raw) - 1
        if 0 <= idx < len(options):
            return options[idx][0]

        print(f"Please type a number between 1 and {len(options)}.")


def resolve_flags_interactive(document: Document) -> None:
    auto_accept_remaining = False

    # --- 6a. Appendix / region flags ---------------------------------------
    # Prompt once per appendix region. All blocks in the same region get
    # the same decision.
    handled_regions: set[int] = set()

    for block in document.blocks:
        if block.region_record.decision != Decision.FLAG:
            continue
        if block.region != Region.APPENDIX:
            continue

        region_key = id(block.region_record)

        if region_key in handled_regions:
            continue

        # Find all FLAG blocks in APPENDIX region and treat them as one.
        region_blocks = [
            b for b in document.blocks
            if b.region == Region.APPENDIX
            and b.region_record.decision == Decision.FLAG
        ]
        if not region_blocks:
            continue

        first_block = region_blocks[0]

        if auto_accept_remaining:
            for b in region_blocks:
                b.region_record.as_manual_override(
                    Decision.EXCLUDE,
                    reason="Auto-accepted suggestion (user chose Accept All earlier)",
                )
                handled_regions.add(id(b.region_record))
            continue

        print()
        print("-" * 60)
        print(f"Appendix region detected ({len(region_blocks)} block(s))")
        snippet = first_block.text[:80]
        if len(first_block.text) > 80:
            snippet += "..."
        print(f'Starts at: "{snippet}"')
        print()
        print("Appendixes are only excluded if they contain supplementary")
        print("material that is NOT used as evidence when marking.")
        print()

        choice = _menu(
            "How should this appendix be treated?",
            [
                ("exclude", "Exclude this appendix from the count"),
                ("include", "Include this appendix in the count"),
                ("help",    "Explain the QCAA rule for appendixes"),
                ("all",     "Accept the tool's suggestions for all remaining prompts"),
            ],
            default_key="exclude",
        )

        if choice == "help":
            print()
            print("Appendixes (QCAA rule):")
            print("  Appendixes are excluded ONLY if they contain supplementary")
            print("  material that is not used as evidence when marking. If a")
            print("  marker would need to read it to assess your response, it")
            print("  counts. When in doubt, exclude it.")
            choice = _menu(
                "How should this appendix be treated?",
                [
                    ("exclude", "Exclude this appendix from the count"),
                    ("include", "Include this appendix in the count"),
                    ("all",     "Accept the tool's suggestions for all remaining prompts"),
                ],
                default_key="exclude",
            )

        if choice == "all":
            auto_accept_remaining = True
            for b in region_blocks:
                b.region_record.as_manual_override(
                    Decision.EXCLUDE,
                    reason="Auto-accepted suggestion (user chose Accept All)",
                )
                handled_regions.add(id(b.region_record))
        elif choice == "exclude":
            for b in region_blocks:
                b.region_record.as_manual_override(
                    Decision.EXCLUDE,
                    reason="User chose to exclude this appendix",
                )
                handled_regions.add(id(b.region_record))
        else:
            for b in region_blocks:
                b.region_record.as_manual_override(
                    Decision.COUNT,
                    reason="User chose to include this appendix",
                )
                handled_regions.add(id(b.region_record))

    if document.tables:
        print()
        print("=" * 60)
        print(f"{len(document.tables)} table(s) found. Each will be reviewed.")
        print("=" * 60)

        for i, table in enumerate(document.tables, start=1):
            if table.answer is not None:
                continue

            if auto_accept_remaining:
                suggested = suggest_table_answer(table)
                table.answer = suggested
                _apply_table_answer(table, suggested,
                                    reason="Auto-accepted suggestion")
                continue

            print()
            print("-" * 60)
            print(f"Table {i} of {len(document.tables)}")
            print(f"Suggested: {describe_suggestion(table)} "
                  f"({describe_confidence(table.suggestion_confidence)})")
            print()
            print(preview_table(table))
            print()
            print(f"  {table.word_count} word(s) in this table.")

            default_key = None
            if table_has_strong_default(table):
                default_key = suggest_table_answer(table).value

            choice = _menu(
                "How should this table be counted?",
                [
                    ("count_all",    "Count the whole table (it contains information)"),
                    ("exclude_all",  "Exclude the whole table (raw data or calculations only)"),
                    ("headers_only", "Count the header row only, exclude the data rows"),
                    ("help",         "Explain the QCAA rule for tables"),
                    ("all",          "Accept the tool's suggestions for all remaining tables"),
                ],
                default_key=default_key,
            )

            if choice == "help":
                print(QCAA_TABLE_HELP)
                choice = _menu(
                    "How should this table be counted?",
                    [
                        ("count_all",    "Count the whole table (it contains information)"),
                        ("exclude_all",  "Exclude the whole table (raw data or calculations only)"),
                        ("headers_only", "Count the header row only, exclude the data rows"),
                        ("all",          "Accept the tool's suggestions for all remaining tables"),
                    ],
                    default_key=default_key,
                )

            if choice == "all":
                auto_accept_remaining = True
                suggested = suggest_table_answer(table)
                table.answer = suggested
                _apply_table_answer(table, suggested,
                                    reason="Auto-accepted suggestion (user chose Accept All)")
                continue

            table.answer = TableAnswer(choice)
            _apply_table_answer(table, table.answer,
                                reason=f"User chose '{choice}'")

    # --- 6d. Captions -----------------------------------------------------
    auto_accept_remaining = resolve_captions_interactive(
        document, auto_accept_remaining
    )

    # --- 6e. Footnotes ----------------------------------------------------
    auto_accept_remaining = resolve_footnotes_interactive(
        document, auto_accept_remaining
    )

    # --- 6f. Citations ----------------------------------------------------
    _resolve_citations_interactive(document)


def _resolve_citations_interactive(document: Document) -> None:
    print()
    print("=" * 60)
    print("Citation detection")
    print("=" * 60)
    print()
    print("Note: this tool can only detect APA 7 style citations with good")
    print("accuracy. Other styles (MLA, Chicago, Harvard, numbered, etc.)")
    print("may not be detected correctly. Check the report at the end.")
    print()

    auto_citations = [c for c in document.citations if c.source == "auto"]

    if auto_citations:
        print(f"Found {len(auto_citations)} in-text citation(s):")
        print()
        for i, c in enumerate(auto_citations, start=1):
            occ = "" if c.occurrences == 1 else f"  (×{c.occurrences})"
            print(f"  {i:>3}. {c.text}{occ}")
        print()
        print("These will be excluded from the word count.")
    else:
        print("No APA 7 in-text citations detected.")
    print()

    apply_citation_exclusions(document)

    # -- Candidates the detector didn't auto-exclude -----------------------
    missed = find_missed_candidates(document)
    if not missed:
        return

    print("-" * 60)
    print(f"{len(missed)} other parenthetical(s) contain a 4-digit year but")
    print("were NOT auto-excluded. They might be citations, or might be")
    print("ordinary references (e.g. to a figure or table).")
    print()
    for i, m in enumerate(missed, start=1):
        print(f"  [{i}] {m}")
    print()
    print("Choose which ones to exclude:")
    print("  Type a number to exclude that one (you can repeat).")
    print("  Type 'a' to exclude all remaining.")
    print("  Press Enter to finish and count the rest normally.")
    print()

    extra: list[Citation] = []
    remaining = list(enumerate(missed, start=1))   # [(1, "..."), (2, "..."), ...]

    while True:
        # Re-print remaining candidates (or a note if none left)
        if remaining:
            print("Remaining:")
            for i, m in remaining:
                print(f"  [{i}] {m}")
            print()
        else:
            print("  (No candidates remaining.)")
            print()

        try:
            raw = input("> ").strip()
        except EOFError:
            break

        # Blank line = exit the loop
        if not raw:
            break

        if raw.lower() == "a":
            if not remaining:
                print("  No candidates to exclude.")
                continue
            for idx, m in remaining:
                count = _count_occurrences_in_body(document, m)
                if count:
                    extra.append(Citation(text=m, occurrences=count, source="manual"))
            remaining = []
            print("  Excluded all remaining candidates.")
            continue

        if not raw.isdigit():
            print("  Please type a number, 'a' for all, or Enter to finish.")
            continue

        choice = int(raw)
        match = next((m for idx, m in remaining if idx == choice), None)
        if match is None:
            print(f"  No candidate #{choice} in the remaining list.")
            continue

        count = _count_occurrences_in_body(document, match)
        if count == 0:
            print(f"  ✗ Not found in body: {match}")
        else:
            extra.append(Citation(text=match, occurrences=count, source="manual"))
            print(f"  ✓ Excluded {count} occurrence(s).")
        # Remove from the remaining pool
        remaining = [(i, m) for i, m in remaining if m != match]

    if extra:
        document.citations.extend(extra)
        apply_citation_exclusions(document)
        print()
        print(f"Added {len(extra)} citation(s).")
        
        
def _count_occurrences_in_body(document: Document, needle: str) -> int:
    """Count occurrences of a substring in countable body blocks.

    Tries exact match first, then falls back to case-insensitive match.
    This handles the common case where a user types a citation slightly
    differently from how it appears in the document.
    """
    if not needle:
        return 0

    total = 0
    needle_lower = needle.lower()
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        if block.region in (Region.TITLE_PAGE, Region.CONTENTS,
                            Region.ABSTRACT, Region.REFERENCES):
            continue
        text = block.text
        # Exact count
        n = text.count(needle)
        if n == 0:
            # Case-insensitive fallback
            n = text.lower().count(needle_lower)
        total += n
    return total


def _apply_table_answer(table: Table, answer: TableAnswer, reason: str) -> None:
    """Apply the user's (or auto-accepted) answer to this table.

    Rule of thumb (from my teacher): "if it's a word, it counts."
    QCAA then carves out exceptions — numbers, symbols, equations,
    calculations, and raw data in tables. So even when the user says
    "count the whole table", we still run the standard span-level
    classification: numbers, symbols, and equations inside the table
    are still excluded. The table-level answer only decides whether
    the table's *prose* content counts.
    """
    if answer == TableAnswer.COUNT_ALL:
        # Apply standard span classification to every cell.
        # Prose counts; numbers, symbols, and equations are excluded.
        for cell in table.iter_cells():
            for span in cell.spans:
                _classify_span_in_place(span, reason_prefix=reason)
        table.answer_record = DecisionRecord(
            decision=Decision.COUNT, rule_id="T-COUNT",
            confidence=1.0, reason=reason, source="manual",
        )

    elif answer == TableAnswer.EXCLUDE_ALL:
        for cell in table.iter_cells():
            for span in cell.spans:
                span.decision_record = DecisionRecord(
                    decision=Decision.EXCLUDE,
                    rule_id="T-EXCLUDE-ALL",
                    confidence=1.0,
                    reason=reason,
                    source="manual",
                )
        table.answer_record = DecisionRecord(
            decision=Decision.EXCLUDE, rule_id="T-EXCLUDE-ALL",
            confidence=1.0, reason=reason, source="manual",
        )

    elif answer == TableAnswer.HEADERS_ONLY:
        header_idx = table.header_row_index
        for r_idx, row in enumerate(table.rows):
            in_header = (r_idx == header_idx)
            for cell in row:
                for span in cell.spans:
                    if in_header:
                        # Header cells still go through standard rules
                        _classify_span_in_place(
                            span, reason_prefix=f"{reason} (header row)"
                        )
                    else:
                        span.decision_record = DecisionRecord(
                            decision=Decision.EXCLUDE,
                            rule_id="T-DATA-ROW",
                            confidence=1.0,
                            reason=f"{reason} (data row)",
                            source="manual",
                        )
        table.answer_record = DecisionRecord(
            decision=Decision.COUNT, rule_id="T-HEADERS-ONLY",
            confidence=1.0, reason=reason, source="manual",
        )


def _classify_span_in_place(span: Span, reason_prefix: str = "") -> None:
    """Classify a single span using the standard QCAA span rules.

    Used inside tables so that numbers, symbols, and equations inside
    a "counted" table are still excluded. Mutates the span in place.

    Rule of thumb: if it's a word, it counts; QCAA then carves out
    exceptions for numbers, symbols, equations, and calculations.
    """
    tok = span.text

    if is_equation_token(tok):
        span.span_type = SpanType.EQUATION
        span.decision_record = DecisionRecord(
            decision=Decision.EXCLUDE,
            rule_id="I-EQUATION",
            confidence=0.9,
            reason=f"{reason_prefix}: equation/calculation excluded by QCAA rules",
            source="manual",
        )
    elif is_symbol_token(tok):
        span.span_type = SpanType.SYMBOL
        span.decision_record = DecisionRecord(
            decision=Decision.EXCLUDE,
            rule_id="I-SYMBOL",
            confidence=0.95,
            reason=f"{reason_prefix}: symbol excluded by QCAA rules",
            source="manual",
        )
    elif is_numeric_token(tok):
        span.span_type = SpanType.NUMBER
        span.decision_record = DecisionRecord(
            decision=Decision.EXCLUDE,
            rule_id="I-NUMBER",
            confidence=0.95,
            reason=f"{reason_prefix}: number excluded by QCAA rules",
            source="manual",
        )
    else:
        span.span_type = SpanType.WORD
        span.decision_record = DecisionRecord(
            decision=Decision.COUNT,
            rule_id="I-WORD",
            confidence=1.0,
            reason=f"{reason_prefix}: counted as a word",
            source="manual",
        )
# ============================================================================
# SECTION 7 — COUNT
# ============================================================================
def count_words(document: Document) -> dict:
    buckets: dict[str, int] = defaultdict(int)
    total = 0
    table_summary: list[dict] = []
    citation_summary = {"auto": 0, "auto_words": 0,
                        "manual": 0, "manual_words": 0}
    total_math_objects = 0

    for block in document.blocks:
        total_math_objects += block.math_object_count

        if block.region_record.decision == Decision.EXCLUDE:
            buckets[f"region:{block.region.value}"] += len(block.spans)
            continue
        if block.region_record.decision in (Decision.FLAG, Decision.UNDECIDED):
            buckets[f"flagged:{block.region.value}"] += len(block.spans)
            continue

        for span in block.spans:
            if span.decision_record.decision == Decision.COUNT:
                total += 1
            else:
                rule = span.decision_record.rule_id or "unknown"
                buckets[f"excluded:{rule}"] += 1

    for table in document.tables:
        counted = 0
        excluded = 0
        for cell in table.iter_cells():
            total_math_objects += cell.math_object_count
            for span in cell.spans:
                if span.decision_record.decision == Decision.COUNT:
                    total += 1
                    counted += 1
                else:
                    excluded += 1

        answer_label = table.answer.value if table.answer else "unresolved"
        table_summary.append({
            "index": table.table_index + 1,
            "suggestion": table.suggestion.value,
            "answer": answer_label,
            "counted": counted,
            "excluded": excluded,
        })
        if excluded:
            buckets[f"table:{answer_label}"] += excluded
    # --- Footnotes / endnotes ---
    footnote_summary = {"counted": 0, "counted_words": 0,
                        "excluded": 0, "excluded_words": 0}

    for fn in document.footnotes:
        if fn.answer == "count" or (fn.answer is None and fn.suggestion == "commentary"):
            total += fn.word_count
            footnote_summary["counted"] += 1
            footnote_summary["counted_words"] += fn.word_count
        else:
            buckets["footnotes (bibliographic)"] += fn.word_count
            footnote_summary["excluded"] += 1
            footnote_summary["excluded_words"] += fn.word_count
    
    
    for block in document.blocks:
        if block.region_record.decision == Decision.EXCLUDE:
            continue
        for span in block.spans:
            if span.decision_record.rule_id == "I-CITATION":
                if "citation:manual" in span.tags:
                    citation_summary["manual_words"] += 1
                else:
                    citation_summary["auto_words"] += 1
    citation_summary["auto"] = sum(1 for c in document.citations if c.source == "auto")
    citation_summary["manual"] = sum(1 for c in document.citations if c.source == "manual")

    return {
        "total": total,
        "buckets": dict(buckets),
        "tables": table_summary,
        "citations": citation_summary,
        "math_objects": total_math_objects,
    }

# ============================================================================
# SECTION 8 — REPORT
# ============================================================================

def print_report(document: Document, result: dict) -> None:
    print()
    print("=" * 60)
    print("QCAA WORD COUNT REPORT")
    print(f"File: {document.filename}")
    print("=" * 60)
    print()
    print(f"  FINAL WORD COUNT: {result['total']}")
    print()

    if result["buckets"]:
        print("Excluded / flagged (not counted):")
        for key in sorted(result["buckets"]):
            print(f"  {key:<40} {result['buckets'][key]:>5} word(s)")
        print()

    if result.get("tables"):
        print("Tables:")
        for t in result["tables"]:
            answer_label = {
                "count_all": "counted whole",
                "exclude_all": "excluded whole",
                "headers_only": "headers only",
                "unresolved": "unresolved",
            }.get(t["answer"], t["answer"])
            print(f"  Table {t['index']}: "
                  f"suggested={t['suggestion']:<12} "
                  f"chosen={answer_label:<15} "
                  f"counted={t['counted']:<4} "
                  f"excluded={t['excluded']}")
        print()
        
        

    cit = result.get("citations", {})
    if cit.get("auto") or cit.get("manual"):
        print("In-text citations:")
        if cit["auto"]:
            print(f"  Auto-detected (APA 7 patterns):   {cit['auto']:>4} "
                  f"({cit['auto_words']} words)")
        if cit["manual"]:
            print(f"  Manually added by user:           {cit['manual']:>4} "
                  f"({cit['manual_words']} words)")
        print()

    fn = result.get("footnotes", {})
    if fn.get("counted") or fn.get("excluded"):
        print("Footnotes / endnotes:")
        if fn["counted"]:
            print(f"  Counted (commentary):             {fn['counted']:>4} "
                  f"({fn['counted_words']} words)")
        if fn["excluded"]:
            print(f"  Excluded (bibliographic):         {fn['excluded']:>4} "
                  f"({fn['excluded_words']} words)")
        print()

    if result.get("math_objects"):
        print(f"Math objects (OMML):              {result['math_objects']:>4} "
              f"(excluded, not counted)")
        print()

    print("Note: Citation detection uses APA 7 patterns. If your citations")
    print("use a different style (MLA, Chicago, Harvard, numbered, etc.),")
    print("they may be counted as ordinary words. Review the report and use")
    print("the manual citation entry if any were missed.")
    print()
    print("This version does not yet handle:")
    print("  - Equations outside tables or math objects (typed as plain text)")
    print("  - Text boxes and SmartArt")
    print("  - Manual [[QCAA_EXCLUDE]] markers")
    print()


# ============================================================================
# SECTION 9 — FILE PICKER
# ============================================================================

def pick_file() -> str | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None

    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    path = filedialog.askopenfilename(
        title="Select your Word document",
        filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
    )
    root.destroy()
    return path or None


# ============================================================================
# SECTION 10 — ENTRY POINT + --explain
# ============================================================================
QCAA_RULES_TEXT = """\
QCAA WORD-LENGTH RULES (official summary)
------------------------------------------

INCLUDED:
  - all words in the text of the response
  - title, headings and subheadings
  - tables, figures, maps and diagrams containing information other
    than raw or processed data (the whole item counts)
  - quotations
  - footnotes and endnotes (unless used for bibliographical purposes)
  - abbreviations, including initialisms (e.g. LPG), units of
    measurement (e.g. kg, m), and chemical formulas (e.g. KOH, HCl)

EXCLUDED:
  - title pages
  - contents pages
  - abstract
  - visual elements associated with the written response
    (e.g. by-lines, banners, captions and call-outs)
  - raw or processed data in tables, figures and diagrams
  - numbers, symbols, equations and calculations
  - bibliography / reference list
  - appendixes (only if supplementary-only)
  - page numbers
  - in-text citations
  - blank pages

HOW THIS TOOL HANDLES THEM:
  The tool applies the same rules regardless of the assessment type
  (student experiment, PSMT, extended response, essay, etc.). It does
  not need to be told what kind of document it is looking at.

  - Title page / contents / abstract / references: auto-excluded.
  - Appendixes: flagged; user confirms whether supplementary-only.
  - Tables: classified and prompted; user confirms each one.
  - Captions: flagged; user confirms each one.
  - Footnotes/endnotes: classified and prompted.
  - Citations: APA 7 patterns auto-detected; user can add missed ones.
  - Numbers, symbols, equations, math objects: auto-excluded.
  - Everything else: counted.

LIMITATIONS:
  - Citation detection is tuned for APA 7. Other styles may be missed.
  - Equations and math typed as images (not text) cannot be read.
  - Footnotes and text boxes are handled, but SmartArt is not.
"""

def run(path: str) -> int:

    if not os.path.isfile(path):
        print(f"Couldn't find a file at: {path}")
        return 2
    if not path.lower().endswith(".docx"):
        print(f"'{path}' doesn't look like a .docx file.")
        return 2
    
    print(f"Loading: {path}")
    document = ingest(path)
    print(f"  Found {len(document.blocks)} paragraph(s)")
    print(f"  Found {len(document.tables)} table(s)")
    math_total = sum(b.math_object_count for b in document.blocks)
    if math_total:
        print(f"  Found {math_total} math object(s) (equation editor)")
        for table in document.tables:
            table.suggestion, table.suggestion_confidence = classify_table(table)

    detect_regions(document)
    apply_regional_decisions(document)
    classify_and_decide_spans(document)

    # Caption detection (before interactive resolution so flags are ready)
    detect_captions(document)

    # Footnotes (read from docx, classify, ready to prompt)
    document.footnotes = read_footnotes_from_docx(path)
    apply_footnote_suggestions(document)

    # Citations
    document.citations = collect_citations(document)

    # Interactive resolution (tables, captions, footnotes, citations)
    resolve_flags_interactive(document)

    result = count_words(document)
    print_report(document, result)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="qcaa_word_count.py",
        description="Estimate the QCAA word count of a .docx response.",
        add_help=True,
    )
    parser.add_argument("file", nargs="?", help=".docx file (omit to use file picker)")
    parser.add_argument("--explain", action="store_true",
                        help="print the QCAA rules and exit")
    args = parser.parse_args(argv)

    if args.explain:
        print(QCAA_RULES_TEXT)
        return 0

    if args.file:
        path = args.file
    else:
        path = pick_file() or ""
        if not path:
            print("No file selected.")
            return 1

    return run(path)


if __name__ == "__main__":
    sys.exit(main())