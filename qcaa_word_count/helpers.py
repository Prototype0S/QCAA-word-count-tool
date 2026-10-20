"""Text extraction, tokenization, and the unified span classifier.

The span classifier (classify_span) is the single source of truth for
QCAA span-level rules. Both the body path and the table path call it,
so behaviour can never drift between them.
"""
from __future__ import annotations

import re

from .model import Block, Decision, DecisionRecord, SpanType

# ---------------------------------------------------------------------------
# Word XML extraction
# ---------------------------------------------------------------------------

def _is_inside_math(node) -> bool:
    from docx.oxml.ns import qn
    for ancestor in node.iterancestors():
        if ancestor.tag in (qn("m:oMath"), qn("m:oMathPara")):
            return True
    return False


def _needs_space(prev_chunk: str, next_chunk: str) -> bool:
    if not prev_chunk or not next_chunk:
        return False
    last = prev_chunk[-1]
    first = next_chunk[0]
    if last.isspace() or first.isspace():
        return False
    return last.isalnum() and first.isalnum()


def _extract_text_from_element(element) -> tuple[str, int]:
    from docx.oxml.ns import qn
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
    return _extract_text_from_element(para._element)


def paragraph_text_boxes(para) -> list[tuple[str, int]]:
    from docx.oxml.ns import qn
    results: list[tuple[str, int]] = []
    for node in para._element.iter():
        if node.tag == qn("w:txbxContent"):
            text, math_count = _extract_text_from_element(node)
            if text.strip() or math_count:
                results.append((text, math_count))
    return results


def paragraph_is_list_item(para) -> bool:
    from docx.oxml.ns import qn
    try:
        pPr = para._element.find(qn("w:pPr"))
        if pPr is None:
            return False
        return pPr.find(qn("w:numPr")) is not None
    except Exception:
        return False


def cell_text(cell) -> tuple[str, int]:
    parts: list[str] = []
    total_math = 0
    for para in cell.paragraphs:
        t, m = paragraph_text(para)
        parts.append(t)
        total_math += m
    return "\n".join(parts), total_math


def heading_level_from_style(style_name: str | None) -> int | None:
    m = re.match(r"Heading (\d+)", style_name or "")
    return int(m.group(1)) if m else None


# ---------------------------------------------------------------------------
# Tokenization
# ---------------------------------------------------------------------------

def tokenize(text: str) -> list[str]:
    return [t for t in re.findall(r"\S+", text) if any(c.isalnum() for c in t)]


def make_spans_for_text(text: str, block_index: int):
    from .model import SourceLocation, Span
    spans = []
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


# ---------------------------------------------------------------------------
# Character sets
# ---------------------------------------------------------------------------

EQUATION_CHARS = set("=×÷^*≤≥≈√∫Σ∑∴∵±∓∂∇∝∞")
SUPERSCRIPT_DIGITS = set("⁰¹²³⁴⁵⁶⁷⁸⁹")
SUBSCRIPT_DIGITS = set("₀₁₂₃₄₅₆₇₈₉")
GREEK_LETTERS = set("αβγδεζηθικλμνξοπρστυφχψωΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ")
MATH_SYMBOLS = EQUATION_CHARS | {"→", "←", "⇒", "⇔", "∈", "∉", "∪", "∩", "°", "′", "″"}


# ---------------------------------------------------------------------------
# Token classification
# ---------------------------------------------------------------------------

_NUMBER_UNIT_RE = re.compile(r"^-?\d+(?:\.\d+)?[a-zA-Z][a-zA-Z0-9/%\^]*$")

_UNIT_ONLY_RE = re.compile(
    r"^(?:kg|g|mg|µg|μg|m|cm|mm|km|nm|µm|μm|s|ms|h|hr|hrs|min|mins|"
    r"L|mL|dL|cL|µL|μL|mol|mmol|µmol|μmol|K|°C|°F|N|J|kJ|MJ|W|kW|V|mV|"
    r"A|mA|Pa|kPa|MPa|Hz|kHz|MHz|rad|sr|cd|lm|lx|Bq|Gy|Sv|kat|bar|atm|psi)$"
)


def is_numeric_token(tok: str) -> bool:
    has_letter = any(c.isalpha() for c in tok)
    has_digit = any(c.isdigit() for c in tok)
    return has_digit and not has_letter


def is_number_unit_token(tok: str) -> bool:
    return bool(_NUMBER_UNIT_RE.match(tok))


def is_unit_token(tok: str) -> bool:
    return bool(_UNIT_ONLY_RE.match(tok))


def is_symbol_token(tok: str) -> bool:
    return not any(c.isalnum() for c in tok)


def is_equation_token(tok: str) -> bool:
    if any(op in tok for op in MATH_SYMBOLS):
        return True
    if any(c in SUPERSCRIPT_DIGITS for c in tok):
        return True
    if any(c in SUBSCRIPT_DIGITS for c in tok):
        return True
    if any(c in GREEK_LETTERS for c in tok):
        return True
    if re.fullmatch(r"-?\d+(?:\.\d+)?[a-z]", tok):
        return True
    if any(op in tok for op in "+-/"):
        if any(c.isdigit() for c in tok):
            letters = "".join(c for c in tok if c.isalpha())
            return len(letters) < 3
    return False


def _is_single_letter_variable(tok: str) -> bool:
    return len(tok) == 1 and tok.isalpha() and tok.islower()


def _contains_bracket(tok: str) -> bool:
    return any(c in "()[]{}" for c in tok)


def is_equation_heavy_block(block: Block, threshold: float = 0.4) -> bool:
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


# ---------------------------------------------------------------------------
# The unified span classifier
# ---------------------------------------------------------------------------

def classify_span(
    tok: str,
    *,
    equation_heavy: bool = False,
    reason_prefix: str = "",
) -> tuple[SpanType, DecisionRecord]:
    """Return (span_type, decision_record) for a single token.

    Order matters. Strong equation indicators first (so 5kg doesn't get
    caught as a number), then symbols, then numbers, then equation-
    heavy heuristics, then units (which count), then words.
    """
    def _rec(dec: Decision, rule: str, conf: float, why: str) -> DecisionRecord:
        return DecisionRecord(
            decision=dec,
            rule_id=rule,
            confidence=conf,
            reason=f"{reason_prefix}: {why}" if reason_prefix else why,
            source="auto" if not reason_prefix else "manual",
        )

    if is_number_unit_token(tok):
        return SpanType.NUMBER, _rec(
            Decision.EXCLUDE, "I-NUMBER-UNIT", 0.95,
            "number with unit excluded by QCAA rules",
        )

    if is_equation_token(tok):
        return SpanType.EQUATION, _rec(
            Decision.EXCLUDE, "I-EQUATION", 0.9,
            "equations and calculations are excluded by QCAA rules",
        )

    if is_symbol_token(tok):
        return SpanType.SYMBOL, _rec(
            Decision.EXCLUDE, "I-SYMBOL", 0.95,
            "symbols are excluded by QCAA rules",
        )

    if is_numeric_token(tok):
        return SpanType.NUMBER, _rec(
            Decision.EXCLUDE, "I-NUMBER", 0.95,
            "numbers are excluded by QCAA rules",
        )

    if equation_heavy and _is_single_letter_variable(tok):
        return SpanType.EQUATION, _rec(
            Decision.EXCLUDE, "I-VARIABLE", 0.75,
            "single-letter variable inside equation-heavy text",
        )

    if equation_heavy and _contains_bracket(tok):
        return SpanType.EQUATION, _rec(
            Decision.EXCLUDE, "I-EQUATION-FRAGMENT", 0.75,
            "bracketed token inside equation-heavy text",
        )

    if is_unit_token(tok):
        return SpanType.UNIT, _rec(
            Decision.COUNT, "I-UNIT", 0.9,
            "unit of measurement (counted)",
        )

    return SpanType.WORD, _rec(
        Decision.COUNT, "I-WORD", 1.0, "ordinary word",
    )