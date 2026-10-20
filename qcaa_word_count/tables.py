"""Table classification and answer application."""
from __future__ import annotations

import re

from .helpers import classify_span, is_equation_heavy_block, is_numeric_token, tokenize
from .model import (
    Decision, DecisionRecord, Table, TableAnswer, TableClass,
)

CALC_OPERATOR_CHARS = set("+-*/=×÷^<>≤≥≈√")
PROSE_RUN_RE = re.compile(r"\b[a-z]{2,}(?:\s+[a-z]{2,}){2,}\b")


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


def apply_table_answer(table: Table, answer: TableAnswer, reason: str) -> None:
    if answer == TableAnswer.COUNT_ALL:
        for cell in table.iter_cells():
            equation_heavy = is_equation_heavy_block(cell)
            for span in cell.spans:
                span_type, rec = classify_span(
                    span.text,
                    equation_heavy=equation_heavy,
                    reason_prefix=reason,
                )
                span.span_type = span_type
                span.decision_record = rec
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
                    reason=reason, source="manual",
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
                if in_header:
                    equation_heavy = is_equation_heavy_block(cell)
                    for span in cell.spans:
                        span_type, rec = classify_span(
                            span.text,
                            equation_heavy=equation_heavy,
                            reason_prefix=f"{reason} (header row)",
                        )
                        span.span_type = span_type
                        span.decision_record = rec
                else:
                    for span in cell.spans:
                        span.decision_record = DecisionRecord(
                            decision=Decision.EXCLUDE,
                            rule_id="T-DATA-ROW",
                            confidence=1.0,
                            reason=f"{reason} (data row)", source="manual",
                        )
        table.answer_record = DecisionRecord(
            decision=Decision.COUNT, rule_id="T-HEADERS-ONLY",
            confidence=1.0, reason=reason, source="manual",
        )