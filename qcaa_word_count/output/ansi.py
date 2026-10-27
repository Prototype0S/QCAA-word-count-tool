"""ANSI colour helpers and the terminal preview."""
from __future__ import annotations

import os as _os

from ..model import BlockType, Decision, Document, Span

# ---------------------------------------------------------------------------
# Colour setup
# ---------------------------------------------------------------------------

_COLOURS_ENABLED = True


def _enable_windows_ansi() -> bool:
    if _os.name != "nt":
        return True
    try:
        import ctypes
        k = ctypes.windll.kernel32
        k.SetConsoleMode(k.GetStdHandle(-11), 7)
        return True
    except Exception:
        return False


if not _enable_windows_ansi():
    _COLOURS_ENABLED = False

_R = "\033[0m"; _B = "\033[1m"; _DIM = "\033[2m"
_GRN = "\033[32m"; _RED = "\033[31m"; _YEL = "\033[33m"
_CYN = "\033[36m"; _MAG = "\033[35m"; _GRY = "\033[90m"


def _c(t: str, *codes: str) -> str:
    return "".join(codes) + t + _R if _COLOURS_ENABLED else t


def green(t):      return _c(t, _GRN)
def red(t):        return _c(t, _RED)
def yellow(t):     return _c(t, _YEL)
def cyan(t):       return _c(t, _CYN)
def grey(t):       return _c(t, _GRY)
def magenta(t):    return _c(t, _MAG)
def bold(t):       return _c(t, _B)
def dim(t):        return _c(t, _DIM)
def bold_green(t): return _c(t, _B, _GRN)
def bold_cyan(t):  return _c(t, _B, _CYN)


# ---------------------------------------------------------------------------
# Span colouring
# ---------------------------------------------------------------------------

def span_to_ansi(span: Span) -> str:
    dec = span.decision_record.decision
    rule = span.decision_record.rule_id or ""
    if dec == Decision.COUNT:
        return green(span.text)
    if dec == Decision.EXCLUDE:
        return magenta(span.text) if rule == "I-CITATION" else red(span.text)
    if dec in (Decision.FLAG, Decision.UNDECIDED):
        return yellow(span.text)
    return span.text


# ---------------------------------------------------------------------------
# Interleaved iteration (blocks and tables in document order)
# ---------------------------------------------------------------------------

def interleaved_items(document: Document):
    """Yield ('block', Block) and ('table', Table) in document order.

    Used by both the ANSI preview and the HTML builder so they render
    the document in the same order Word shows it.
    """
    items = []
    for b in document.blocks:
        items.append((b.block_index, 1, "block", b))
    for t in document.tables:
        items.append((t.block_index, 0, "table", t))
    items.sort(key=lambda x: (x[0], x[1]))
    for _, _, kind, obj in items:
        yield kind, obj


# ---------------------------------------------------------------------------
# Terminal preview
# ---------------------------------------------------------------------------

def print_ansi_preview(document: Document) -> None:
    print("=" * 78)
    print(bold("COLOUR-CODED PREVIEW"))
    print(grey("Green = counted | Magenta = citation | Red = excluded | "
               "Yellow = flagged | Grey = region | Cyan = heading"))
    print("=" * 78)
    print()

    for kind, obj in interleaved_items(document):
        if kind == "block":
            block = obj
            if not block.spans:
                continue
            if block.block_type == BlockType.HEADING:
                print()
                print(cyan(bold(block.text.strip())))
                continue
            if block.region_record.decision == Decision.EXCLUDE:
                print(grey(block.text.strip()))
                continue
            print("".join(span_to_ansi(s) for s in block.spans))
        else:
            table = obj
            print()
            answer = table.answer.value if table.answer else "unresolved"
            print(cyan(bold(f"Table {table.table_index + 1} — {answer}")))
            for row in table.rows:
                row_strs = [" ".join(span_to_ansi(s) for s in cell.spans)
                            for cell in row]
                print(" | ".join(row_strs))

    print()
    print("=" * 78)