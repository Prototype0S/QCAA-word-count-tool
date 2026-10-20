"""Data model: enums and dataclasses.

This is the shared shape of the pipeline. Nothing here does work; it
just defines what a Block, Span, Table, etc. look like so the other
modules can talk to each other.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Literal, Optional


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


@dataclass
class Footnote:
    footnote_id: str
    text: str
    kind: Literal["footnote", "endnote"] = "footnote"

    suggestion: Literal["bibliographic", "commentary", "ambiguous"] = "ambiguous"
    suggestion_confidence: float = 0.0

    answer: Optional[Literal["exclude", "count"]] = None
    answer_record: DecisionRecord = field(default_factory=DecisionRecord)


@dataclass
class Document:
    blocks: list[Block] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)
    filename: Optional[str] = None

    citations: list[Citation] = field(default_factory=list)
    footnotes: list[Footnote] = field(default_factory=list)
    text_boxes: list[Block] = field(default_factory=list)
    drawings: int = 0
    drawings_with_text: int = 0

    def all_spans(self):
        for b in self.blocks:
            for s in b.spans:
                yield s
        for t in self.tables:
            for c in t.iter_cells():
                for s in c.spans:
                    yield s