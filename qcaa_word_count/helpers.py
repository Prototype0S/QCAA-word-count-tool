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
# Temperature units
#
# "°C" / "°F" are two codepoints: degree sign + letter. Word sometimes
# stores them as the single-codepoint compatibility forms "℃" (U+2103)
# and "℉" (U+2109) instead — e.g. when text is pasted from a PDF or
# typed on macOS. Both forms must count.
#
# We normalise at comparison time (not in the tokenizer) because the
# span text is what the preview renders, and users should see the
# original glyphs.
# ---------------------------------------------------------------------------

_TEMPERATURE_UNITS = frozenset({
    "°C", "°F",           # two-codepoint form (common)
    "℃", "℉",            # single-codepoint compatibility form
    "C", "F",             # bare letter (only counts when preceded by a
                          # standalone degree; see classify_span)
})


def _is_temperature_unit(tok: str) -> bool:
    return tok in _TEMPERATURE_UNITS


# ---------------------------------------------------------------------------
# URLs
#
# A URL is one token (no spaces) containing "://", "www.", or a known
# top-level domain. It's not a word and not an equation — it's a symbol
# in a body paragraph (QCAA: "symbols are excluded"). In a References
# region it's excluded by region. Either way, 0 words.
# ---------------------------------------------------------------------------

_URL_RE = re.compile(
    r"^(?:"
    r"https?://\S+"
    r"|www\.\S+"
    r"|ftp://\S+"
    r"|<https?://[^>]+>"
    r")$",
    re.IGNORECASE,
)

# A token that *contains* a URL but isn't purely one (e.g. "(see
# https://example.com)"). Rare, but treat the same as a URL — 0 words.
_URL_SUBSTRING_RE = re.compile(
    r"https?://\S+|www\.\S+",
    re.IGNORECASE,
)


def is_url_token(tok: str) -> bool:
    return bool(_URL_RE.match(tok))


# ---------------------------------------------------------------------------
# Token classification
# ---------------------------------------------------------------------------

# Capture group 2 = the unit suffix, used to decide whether the
# number+unit token is a real measurement (5kg, 30m) or an ambiguous
# identifier (2020a). Note: the degree symbol is optional and can be
# the single-codepoint "℃"/"℉" form.
_NUMBER_UNIT_RE = re.compile(
    r"^-?(\d+(?:\.\d+)?)"
    r"(°?[a-zA-Z][a-zA-Z0-9/%\^]*"
    r"|[℃℉])$"
)

_UNIT_ONLY_RE = re.compile(
    r"^(?:kg|g|mg|µg|μg|m|cm|mm|km|nm|µm|μm|s|ms|h|hr|hrs|min|mins|"
    r"L|mL|dL|cL|µL|μL|mol|mmol|µmol|μmol|K|N|J|kJ|MJ|W|kW|V|mV|"
    r"A|mA|Pa|kPa|MPa|Hz|kHz|MHz|rad|sr|cd|lm|lx|Bq|Gy|Sv|kat|bar|atm|psi|"
    r"°C|°F|℃|℉)$"
)

_ORDINAL_RE = re.compile(r"^-?\d+(?:st|nd|rd|th)$", re.IGNORECASE)

# Common "number+letters" tokens that are words/adjectives, not quantities.
_TECH_ADJECTIVES = frozenset({
    "1d", "2d", "3d", "4d",
    "1k", "2k", "3k", "4k", "5k", "8k",
    "5g", "6g",
    "2fa", "3pl", "4pl",
})

_NUMBER_WORD_NUMBER_RE = re.compile(
    r"^\d+-(?:to|in|by|on|at|of|per)-\d+$", re.IGNORECASE
)


def is_numeric_token(tok: str) -> bool:
    has_letter = any(c.isalpha() for c in tok)
    has_digit = any(c.isdigit() for c in tok)
    return has_digit and not has_letter


def is_number_unit_token(tok: str) -> bool:
    if _ORDINAL_RE.match(tok):
        return False
    if tok.lower() in _TECH_ADJECTIVES:
        return False
    return bool(_NUMBER_UNIT_RE.match(tok))


def _is_measurement(tok: str) -> bool:
    m = _NUMBER_UNIT_RE.match(tok)
    if not m:
        return False
    unit_suffix = m.group(2)
    if unit_suffix in ("℃", "℉"):
        return True
    if f"°{unit_suffix}" in ("°C", "°F"):
        return True
    return bool(_UNIT_ONLY_RE.match(unit_suffix))


def is_unit_token(tok: str) -> bool:
    return bool(_UNIT_ONLY_RE.match(tok))


def is_symbol_token(tok: str) -> bool:
    return not any(c.isalnum() for c in tok)


def is_equation_token(tok: str) -> bool:
    if _NUMBER_WORD_NUMBER_RE.match(tok):
        return False

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


def is_equation_heavy_block(block: Block, threshold: float = 0.55) -> bool:
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

# ---------------------------------------------------------------------------
# Trailing punctuation
#
# Tokens carry their trailing punctuation ("°F.", "5kg,", "42.", ")Smith,").
# Every "is this token X?" check in classify_span wants the punctuation-
# stripped form, but the span text the preview renders must stay original.
# Strip here, once, and use `stripped` everywhere below.
# ---------------------------------------------------------------------------

_TRAILING_PUNCT = '.,;:!?)]}\'"\u2019\u201d\u00bb\u203a'


def _strip_trailing_punct(tok: str) -> str:
    stripped = tok.rstrip(_TRAILING_PUNCT)
    # "°F." -> "°F" ✓
    # "5kg," -> "5kg" ✓
    # "42." -> "42" ✓
    # ".." -> "" (empty; classify_span falls through to symbol check on
    #            the ORIGINAL token, which is correct for pure punct)
    return stripped


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

    Trailing punctuation (".", ",", ")", etc.) is stripped before
    matching so "°F." is treated the same as "°F", and "5kg," the same
    as "5kg". The span text itself is unchanged.
    """
    def _rec(dec: Decision, rule: str, conf: float, why: str) -> DecisionRecord:
        return DecisionRecord(
            decision=dec,
            rule_id=rule,
            confidence=conf,
            reason=f"{reason_prefix}: {why}" if reason_prefix else why,
            source="auto" if not reason_prefix else "manual",
        )

    stripped = _strip_trailing_punct(tok)
    # If the token is entirely punctuation, don't use the stripped form
    # (there's nothing left to classify). Fall through to the original.
    match_tok = stripped or tok

    # URLs are not words, not equations, not symbols. Exclude them
    # explicitly so the tooltip reason is honest.
    if is_url_token(match_tok):
        return SpanType.SYMBOL, _rec(
            Decision.EXCLUDE, "I-URL", 0.95,
            "URLs are excluded (not a word)",
        )

    # Temperature units in every form: "°C", "°F", "℃", "℉". The
    # single-codepoint forms can appear standalone ("100 ℃") or glued
    # to the number ("100℃", handled by is_number_unit_token below).
    # The stripped form catches "°F." and "℃.".
    if _is_temperature_unit(match_tok):
        return SpanType.UNIT, _rec(
            Decision.COUNT, "I-UNIT", 0.9,
            "temperature unit (counted)",
        )

    if is_number_unit_token(match_tok):
        if _is_measurement(match_tok):
            return SpanType.UNIT, _rec(
                Decision.COUNT, "I-MEASUREMENT", 0.9,
                "measurement (unit counts per QCAA rules)",
            )
        return SpanType.NUMBER, _rec(
            Decision.EXCLUDE, "I-NUMBER-UNIT", 0.95,
            "number with unrecognised suffix (treated as a number)",
        )

    if is_equation_token(match_tok):
        return SpanType.EQUATION, _rec(
            Decision.EXCLUDE, "I-EQUATION", 0.9,
            "equations and calculations are excluded by QCAA rules",
        )

    # For the symbol check, use the ORIGINAL token — if it's all
    # punctuation, we want it classified as a symbol even though the
    # stripped form is empty.
    if is_symbol_token(tok):
        return SpanType.SYMBOL, _rec(
            Decision.EXCLUDE, "I-SYMBOL", 0.95,
            "symbols are excluded by QCAA rules",
        )

    if is_numeric_token(match_tok):
        return SpanType.NUMBER, _rec(
            Decision.EXCLUDE, "I-NUMBER", 0.95,
            "numbers are excluded by QCAA rules",
        )

    if equation_heavy and _is_single_letter_variable(match_tok):
        return SpanType.EQUATION, _rec(
            Decision.EXCLUDE, "I-VARIABLE", 0.75,
            "single-letter variable inside equation-heavy text",
        )

    if equation_heavy and _contains_bracket(match_tok):
        inner = match_tok.strip("()[]{}")
        mathy_inner = (
            any(c.isdigit() for c in inner)
            or any(c in MATH_SYMBOLS or c in GREEK_LETTERS for c in inner)
            or (inner.isalpha() and len(inner) == 1)
        )
        if mathy_inner:
            return SpanType.EQUATION, _rec(
                Decision.EXCLUDE, "I-EQUATION-FRAGMENT", 0.75,
                "bracketed token inside equation-heavy text",
            )

    if is_unit_token(match_tok):
        return SpanType.UNIT, _rec(
            Decision.COUNT, "I-UNIT", 0.9,
            "unit of measurement (counted)",
        )

    # Criterion-style label (C6, A1, B2) — an abbreviation, still a word.
    if re.fullmatch(r"[A-Z]{1,3}\d{1,2}", match_tok):
        return SpanType.WORD, _rec(
            Decision.COUNT, "I-LABEL", 1.0,
            "label or abbreviation (counts as a word)",
        )

    return SpanType.WORD, _rec(
        Decision.COUNT, "I-WORD", 1.0, "ordinary word",
    )
# ---------------------------------------------------------------------------
# Block-level span application
# ---------------------------------------------------------------------------

def apply_span_decisions_for_block(block: Block) -> None:
    """Classify every span in ``block`` based on its ``region_record``.

    If the block's region is excluded, all spans are marked excluded.
    Otherwise, each span is classified individually.

    **Preserves manual overrides.** Re-running this after a state load
    will reclassify auto-decided spans but leave any user toggles alone.
    """
    if block.region_record.decision == Decision.EXCLUDE:
        for span in block.spans:
            if span.decision_record.source == "manual":
                continue
            span.decision_record = DecisionRecord(
                decision=Decision.EXCLUDE,
                rule_id="R-REGION",
                confidence=1.0,
                reason=f"Inherits exclusion from region '{block.region.value}'",
                source="auto",
            )
        return

    equation_heavy = is_equation_heavy_block(block)
    for span in block.spans:
        if span.decision_record.source == "manual":
            continue
        span_type, rec = classify_span(span.text, equation_heavy=equation_heavy)
        span.span_type = span_type
        span.decision_record = rec