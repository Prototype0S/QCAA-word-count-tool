"""
QCAA Word Count Tool
--------------------
Run with:  python qcaa_word_count.py "path/to/your/file.docx"
(or just run it and paste the path when asked)

Requires:  pip install python-docx

WHAT THIS TOOL TRIES TO DO
  Mirrors the QCAA word-length rules as closely as an automated tool
  reasonably can, including:

  Included: body text, titles/headings/subheadings, quotations,
            abbreviations/units/chemical formulas (kg, KOH, LPG, 37kg),
            footnotes and endnotes (when not purely bibliographic),
            tables/figures/diagrams that contain anything beyond raw
            or processed data.

  Excluded: title pages, contents pages, abstract, bibliography/
            reference list, appendixes, page numbers, raw or processed
            data in tables/figures/diagrams, numbers, symbols, equations,
            calculations, in-text citations, visual elements of written
            genres (by-lines, banners, captions, call-outs), blank pages.

HOW IT ASKS YOU TO HELP
  Some categories genuinely require judgement (e.g. "does this table
  contain information other than raw data?"). The tool detects likely
  candidates and prompts you, showing enough context to answer.

WHAT IT STILL CANNOT DO RELIABLY (check by eye)
  - SmartArt / chart data / embedded diagram XML (SmartArt lives in a
    separate part of the .docx that python-docx doesn't reach; the tool
    flags the anchor paragraph but not the diagram's internal labels)
  - Citations in styles the pattern-matcher doesn't know (e.g. some
    custom Zotero styles, unusual footnote-based systems)
  - Text inside images (screenshots, scanned figures)
  - Field codes other than CITATION-family (some reference managers use
    proprietary field codes with unpredictable shapes)

MANUAL OVERRIDE
  Wrap any content in:
      [[QCAA_EXCLUDE_START]]  ...  [[QCAA_EXCLUDE_END]]
  typed as plain text, and it will be excluded regardless of category.
"""

import sys
import os
import re
import zipfile
from xml.etree import ElementTree as ET

import docx
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
WPS_NS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
VML_NS = "{urn:schemas-microsoft-com:vml}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

START_MARK = "[[QCAA_EXCLUDE_START]]"
END_MARK = "[[QCAA_EXCLUDE_END]]"

# Structural sections. Matched against the heading's trimmed, lowercased
# text as either an exact match or a startswith — NOT as a substring. So
# "References" and "References and Notes" trigger, but "References to the
# Divine in Paradise Lost" does not.
STRUCT_TRIGGERS = {
    "table of contents": "contents",
    "contents": "contents",
    "list of figures": "contents",
    "list of tables": "contents",
    "abstract": "abstract",
    "bibliography": "bibliography",
    "reference list": "bibliography",
    "references": "bibliography",
    "works cited": "bibliography",
    "appendix": "appendix",
    "appendices": "appendix",
    # Front-matter declarations. QCAA treats these as part of the title
    # page / front matter and excludes them. Adding them here means the
    # tool excludes everything under the heading without prompting.
    "declaration of authenticity": "frontmatter",
    "declaration": "frontmatter",
    "authenticity": "frontmatter",
    "academic integrity": "frontmatter",
    "student declaration": "frontmatter",
}

RAW_DATA_EXPLANATION = """
------------------------------------------------------------------
RAW vs PROCESSED DATA vs "OTHER INFORMATION" (QCAA rule)
------------------------------------------------------------------
QCAA excludes a table/figure/diagram from the word count ONLY if it
contains ONLY raw or processed data:

  Raw data       = results exactly as collected/measured, e.g. each
                   participant's individual score, each trial's
                   reading, unedited survey responses.
  Processed data = that same data after basic manipulation, e.g.
                   totals, means, percentages, a graph plotted
                   straight from the numbers. Still just numbers/
                   results — no written interpretation.

If the table/figure ALSO contains "information other than raw or
processed data" — written analysis or commentary sitting inside it,
qualitative/descriptive labels beyond simple column headers,
annotations explaining what the data means, categorisations you
made — then QCAA counts the WHOLE table/figure as included, not
just the extra text.

Quick test: if you deleted it and kept only a sentence like "see
Table 1 for full results", would the marker lose anything beyond raw
numbers? If yes, answer "n" (include it). If it's genuinely just
numbers/results, answer "y" (exclude).
------------------------------------------------------------------
"""

VISUAL_ELEMENT_EXPLANATION = """
------------------------------------------------------------------
VISUAL ELEMENTS OF WRITTEN GENRES (QCAA exclusion)
------------------------------------------------------------------
QCAA excludes visual elements that decorate or frame a written piece
rather than carry its argument — specifically by-lines, banners,
captions and call-outs of the kind you'd see in a literary article,
blog, essay or column.

Examples that SHOULD be excluded:
  - "Figure 3: Reaction rate vs. temperature"
  - "By Jane Smith, 12 March 2024"
  - A pull-quote banner restating a sentence
  - A call-out box repeating a key statistic

Examples that should NOT be excluded:
  - A sentence that happens to start with "Figure 3 shows..."
  - Ordinary body prose that names a figure inline

Say yes only if this line exists purely to label/announce something
else visually, not to carry meaning of its own.
------------------------------------------------------------------
"""


# ---------------------------------------------------------------- helpers --

def prompt_ynq(message, help_text=None):
    """Ask for y / n / ?. Loops until y or n. A second ? re-shows help
    rather than falling through to 'no'."""
    while True:
        ans = input(message).strip().lower()
        if ans in ("y", "yes"):
            return "y"
        if ans in ("n", "no"):
            return "n"
        if ans in ("?", "help", "h"):
            print(help_text if help_text else "Please enter y, n, or ?.")
            continue
        print("Please enter y, n, or ?.")


def prompt_ynmx(message, help_text=None):
    """Like prompt_ynq but also accepts 'm' for 'mixed'. Used where the
    right answer is genuinely 'some of these' (footnotes, endnotes)."""
    while True:
        ans = input(message).strip().lower()
        if ans in ("y", "yes"):
            return "y"
        if ans in ("n", "no"):
            return "n"
        if ans in ("m", "mixed"):
            return "m"
        if ans in ("?", "help", "h"):
            print(help_text if help_text else "Please enter y, n, m, or ?.")
            continue
        print("Please enter y, n, m, or ?.")

def prompt_dci(message, help_text=None):
    """Ask for d / c / i / ? and keep looping. Used for tables where the
    choices are: data-only, calculation-working, or information-that-counts."""
    while True:
        ans = input(message).strip().lower()
        if ans in ("d", "data"):
            return "d"
        if ans in ("c", "calc", "calculation"):
            return "c"
        if ans in ("i", "include", "info"):
            return "i"
        if ans in ("?", "help", "h"):
            print(help_text if help_text else "Please enter d, c, i, or ?.")
            continue
        print("Please enter d, c, i, or ?.")
        
        
def iter_block_items(document):
    """Top-level body blocks in order. Yields Paragraph or Table for each
    <w:p> or <w:tbl>. (Figures/drawings are paragraphs that contain a
    drawing child — see block_has_drawing.)"""
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, document)
        elif child.tag == qn('w:tbl'):
            yield Table(child, document)


def tokenize(text):
    return re.findall(r"\S+", text)


def is_numeric_symbol_token(tok):
    """Digits but no letters -> excluded (42, 3.14, 1,000, 37%).
    Letters present (kg, KOH, LPG, 37kg) -> left alone.
    Bare symbols like '%' or '#' with no digit aren't alnum, so they
    were never counted in the first place."""
    has_letter = any(c.isalpha() for c in tok)
    has_digit = any(c.isdigit() for c in tok)
    return has_digit and not has_letter


def word_count(text):
    return len([t for t in tokenize(text) if any(c.isalnum() for c in t)])


def numeric_symbol_count(text):
    return sum(1 for t in tokenize(text) if is_numeric_symbol_token(t))


def heading_level(paragraph):
    style_name = paragraph.style.name if paragraph.style else ""
    m = re.match(r"Heading (\d+)", style_name or "")
    if m:
        return int(m.group(1))
    return None


def effective_heading_level(paragraph, confirmed_headings):
    lvl = heading_level(paragraph)
    if lvl is not None:
        return lvl
    text = strip_markers(paragraph.text).strip()
    if text in confirmed_headings:
        return 1
    return None


def find_heading_candidates(document):
    """Short, uniformly bold or mostly-caps paragraphs that might be
    unstyled section headings. Conservative: EVERY non-whitespace run must
    be bold, or the whole line must be >70% capitals."""
    seen = set()
    ordered = []
    for block in iter_block_items(document):
        if not isinstance(block, Paragraph):
            continue
        if heading_level(block) is not None:
            continue
        text = strip_markers(block.text).strip()
        if not text or len(text) > 80 or text in seen:
            continue
        runs = [r for r in block.runs if r.text.strip()]
        if not runs:
            continue
        all_bold = all(getattr(r.font, "bold", False) for r in runs)
        letters = [c for c in text if c.isalpha()]
        is_capsy = (
            len(letters) >= 3
            and sum(c.isupper() for c in letters) / len(letters) > 0.7
        )
        if all_bold or is_capsy:
            seen.add(text)
            ordered.append(text)
    return ordered


def strip_markers(text):
    return text.replace(START_MARK, "").replace(END_MARK, "")


def _tc_text(tc):
    return "".join(t.text or "" for t in tc.iter(f"{W_NS}t"))


def table_text(table):
    """Cell text using iter_tcs(), which walks each underlying <w:tc>
    exactly once — merged cells aren't duplicated the way row.cells
    would duplicate them."""
    return "\n".join(_tc_text(tc) for tc in table._tbl.iter_tcs())


def paragraph_has_drawing(p):
    """True if this paragraph contains an image, shape, textbox, chart
    anchor, or SmartArt anchor. Checks <w:drawing> (modern), <w:pict>
    (legacy VML), and <w:object> (embedded OLE, e.g. an equation editor
    object)."""
    for tag in ("w:drawing", "w:pict", "w:object"):
        if p._p.find(f".//{qn(tag)}") is not None:
            return True
    return False


def paragraph_is_caption(p):
    """A paragraph using Word's built-in 'Caption' style, or any style
    whose name starts with 'Caption'."""
    style_name = p.style.name if p.style else ""
    return bool(style_name and style_name.lower().startswith("caption"))


# Caption-looking text NOT tagged with Word's Caption style (common in
# docs copied from Google Docs, or where the author typed the label by
# hand). Kept short and anchored to the typical start-of-caption prefixes.
CAPTION_TEXT_RE = re.compile(
    r"^\s*(figure|fig\.?|table|chart|graph|diagram|image|photo|map|exhibit)\s*"
    r"[\dIVXivx]+[a-z]?\s*[:.\u2013\u2014-]",
    re.IGNORECASE,
)

# By-line / banner shapes: "By Jane Smith", "Words by ...", "Illustration by ..."
BYLINE_RE = re.compile(
    r"^\s*(by|words by|illustration by|illustrated by|photo by|images? by)\s+\S+",
    re.IGNORECASE,
)


def get_footnote_endnote_texts(path):
    """Return two lists (footnotes, endnotes), each a list of note texts."""
    footnotes, endnotes = [], []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        for fname, tag, bucket in (
            ("word/footnotes.xml", "footnote", footnotes),
            ("word/endnotes.xml", "endnote", endnotes),
        ):
            if fname not in names:
                continue
            try:
                root = ET.fromstring(z.read(fname))
            except ET.ParseError:
                continue
            for note in root.iter(f"{W_NS}{tag}"):
                note_type = note.get(f"{W_NS}type")
                if note_type in ("separator", "continuationSeparator", "continuationNotice"):
                    continue
                t = "".join(n.text or "" for n in note.iter(f"{W_NS}t"))
                if t.strip():
                    bucket.append(t)
    return footnotes, endnotes


# Field codes we treat as citation-producing. Word's own CITATION field,
# plus the field codes Zotero and EndNote emit through the Word plugin.
CITATION_FIELD_CODES = ("CITATION", "ZOTERO_ITEM", "ZOTERO_BIBL", "ADDIN EN.CITE", "ADDIN ZOTERO_ITEM")


def _instr_is_citation(instr):
    up = instr.upper()
    return any(code in up for code in CITATION_FIELD_CODES)


def get_field_citation_texts(document_xml_bytes):
    """Extract the visible text of any field whose instruction looks like
    a citation-producing field (Word CITATION, Zotero, EndNote). Handles
    simple fields (<w:fldSimple>) and complex fields (<w:fldChar begin>...
    <w:instrText> ... <w:fldChar separate> ... <w:t> ... <w:fldChar end>).
    Uses a stack so nested fields don't clobber outer state. Accumulates
    instruction text across runs, since Word splits it freely."""
    try:
        root = ET.fromstring(document_xml_bytes)
    except ET.ParseError:
        return []
    texts = []

    # Simple fields
    for fs in root.iter(f"{W_NS}fldSimple"):
        instr = fs.get(f"{W_NS}instr", "") or ""
        if _instr_is_citation(instr):
            t = "".join(n.text or "" for n in fs.iter(f"{W_NS}t"))
            if t.strip():
                texts.append(t)

    # Complex fields
    stack = []
    for elem in root.iter():
        tag = elem.tag
        if tag == f"{W_NS}fldChar":
            fct = elem.get(f"{W_NS}fldCharType")
            if fct == "begin":
                stack.append({"instr": "", "collecting": False, "buf": ""})
            elif fct == "separate":
                if stack:
                    stack[-1]["collecting"] = True
            elif fct == "end":
                if stack:
                    cur = stack.pop()
                    if cur["collecting"] and _instr_is_citation(cur["instr"]) and cur["buf"].strip():
                        texts.append(cur["buf"])
        elif tag == f"{W_NS}instrText":
            if stack and not stack[-1]["collecting"]:
                stack[-1]["instr"] += elem.text or ""
        elif tag == f"{W_NS}t":
            if stack and stack[-1]["collecting"]:
                stack[-1]["buf"] += elem.text or ""

    return texts


def strip_substrings(text, substrings):
    out = text
    for s in substrings:
        if s and s in out:
            out = out.replace(s, " ")
    return out


# ---------------------------------------------------- citation patterns --
# Author-date with a comma before year, optional page tail, optional
# semicolon-separated second citation. Deliberately requires the comma so
# "(max-min)" and "(5\u00b0C)" don't false-match.
FULL_CITATION_RE = re.compile(
    r"\([^()]{1,120}?,\s*(?:19|20)\d{2}[a-z]?"
    r"(?:,\s*pp?\.?\s*\d+(?:[-\u2013]\d+)?)?"
    r"(?:\s*;\s*[^()]{1,120}?,\s*(?:19|20)\d{2}[a-z]?)*\)"
)
# Narrative citation: "Author (2020)". Matches any (year) parenthetical;
# user confirms each unique one.
BARE_YEAR_RE = re.compile(r"\((?:19|20)\d{2}[a-z]?\)")
# Bracketed numeric citations [1], [1,2], [1-3], [12].
VANCOUVER_RE = re.compile(r"\[\d{1,3}(?:\s*[,\-\u2013]\s*\d{1,3})*\]")

# Inline equation: single-letter (or Greek) variable, equals-like sign,
# then a right-hand side that contains at least one digit OR a
# multiplication/division/unit indicator. Symbolic forms like "n = m/M"
# and "v = d/t" match; plain "x = y" in prose does not.
EQUATION_RE = re.compile(
    r"(?<![\w.])"
    r"([A-Za-z\u0394\u03b4\u03b8\u03bb\u03bc\u03c3\u03c0])"
    r"\s*([=\u2243\u2248])\s*"
    r"([0-9][^\n.!?]{0,60}?|[A-Za-z0-9]+(?:\s*[/\u00d7*]\s*[A-Za-z0-9]+)+)"
    r"(?=[.!?\n]|$)"
)

# Continuation line of a worked calculation: starts with "=" then an
# arithmetic expression. "= 4.0 / 40.0" or "= 0.10 mol". These would
# otherwise be counted as body text.
CALC_CONTINUATION_RE = re.compile(
    r"^\s*=\s*[0-9][^\n.!?]{0,60}?(?=[.!?\n]|$)"
)

# A line that looks like a unit-bearing result: "= 0.10 mol", "= 5 kg m/s"
CALC_UNIT_RE = re.compile(
    r"^\s*=\s*[0-9.]+\s*[a-zA-Z]{1,6}(?:/[a-zA-Z]{1,6})?\s*$"
)


class ExclusionTracker:
    """Tracks whether the current block is inside a structural exclusion
    (contents/abstract/bibliography/appendix) or a manual marker region."""

    def __init__(self):
        self.in_struct = False
        self.struct_kind = None
        self.struct_level = None
        self.in_manual = False

    def visit_heading(self, lvl, text):
        if self.in_struct and lvl <= self.struct_level:
            self.in_struct = False
            self.struct_kind = None
            self.struct_level = None
        if not self.in_struct:
            lower = text.strip().lower().rstrip(":.-")
            for trigger, kind in STRUCT_TRIGGERS.items():
                if lower == trigger or lower.startswith(trigger + " ") or lower.startswith(trigger + ":"):
                    self.in_struct = True
                    self.struct_kind = kind
                    self.struct_level = lvl
                    return

    def visit_markers(self, has_start, has_end):
        block_manual = self.in_manual or has_start or has_end
        if has_start and not has_end:
            self.in_manual = True
        elif has_end:
            self.in_manual = False
        return block_manual

    @property
    def excluded(self):
        return self.in_struct or self.in_manual


def find_citation_candidates(document, confirmed_headings):
    """Collect (match, context) for typed-out in-text citations. Skips
    anything inside a struct/manual exclusion."""
    candidates = {}
    tracker = ExclusionTracker()
    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            raw = block.text
        else:
            raw = table_text(block)
        text = strip_markers(raw).replace("\n", " ")
        has_start = START_MARK in raw
        has_end = END_MARK in raw

        if isinstance(block, Paragraph):
            lvl = effective_heading_level(block, confirmed_headings)
            if lvl is not None:
                tracker.visit_heading(lvl, text)

        if tracker.visit_markers(has_start, has_end) or tracker.in_struct:
            continue

        for pattern in (FULL_CITATION_RE, BARE_YEAR_RE, VANCOUVER_RE):
            for m in pattern.finditer(text):
                full = m.group(0)
                if full not in candidates:
                    start = max(m.start() - 40, 0)
                    end = min(m.end() + 20, len(text))
                    candidates[full] = text[start:end]
    return candidates


def find_equation_candidates(document, confirmed_headings):
    """Collect inline-equation and calculation-line candidates as
    (match, context). Each is confirmed by the user before exclusion."""
    candidates = {}
    tracker = ExclusionTracker()
    for block in iter_block_items(document):
        if not isinstance(block, Paragraph):
            continue
        raw = block.text
        text = strip_markers(raw)
        has_start = START_MARK in raw
        has_end = END_MARK in raw

        lvl = effective_heading_level(block, confirmed_headings)
        if lvl is not None:
            tracker.visit_heading(lvl, text)

        if tracker.visit_markers(has_start, has_end) or tracker.in_struct:
            continue

        # Inline equations anywhere in the paragraph
        for m in EQUATION_RE.finditer(text):
            full = m.group(0).strip()
            if full and full not in candidates:
                start = max(m.start() - 30, 0)
                end = min(m.end() + 10, len(text))
                candidates[full] = text[start:end]

        # Whole-paragraph calculation continuations
        stripped = text.strip()
        if CALC_CONTINUATION_RE.match(stripped):
            key = stripped
            if key not in candidates:
                candidates[key] = stripped

    return candidates


def find_visual_element_candidates(document, confirmed_headings):
    """Find paragraphs likely to be visual elements of written genres:
    captions (styled or text-shaped) and by-lines. Only those adjacent
    to (or styled as) captions, to avoid false matches on ordinary prose
    that happens to start with 'Figure 3 shows...'."""
    candidates = {}
    tracker = ExclusionTracker()
    prev_was_drawing = False

    for block in iter_block_items(document):
        if not isinstance(block, Paragraph):
            prev_was_drawing = False
            continue
        raw = block.text
        text = strip_markers(raw).strip()
        has_start = START_MARK in raw
        has_end = END_MARK in raw

        lvl = effective_heading_level(block, confirmed_headings)
        if lvl is not None:
            tracker.visit_heading(lvl, text)

        if tracker.visit_markers(has_start, has_end) or tracker.in_struct:
            prev_was_drawing = paragraph_has_drawing(block)
            continue

        if not text:
            prev_was_drawing = paragraph_has_drawing(block)
            continue

        is_caption_style = paragraph_is_caption(block)
        looks_like_caption = bool(CAPTION_TEXT_RE.match(text))
        looks_like_byline = bool(BYLINE_RE.match(text))
        short_enough = len(text) <= 160

        # Caption styled or directly below a drawing -> very likely
        if (is_caption_style or prev_was_drawing) and short_enough and text not in candidates:
            candidates[text] = ("caption", text)
        elif looks_like_caption and short_enough and text not in candidates:
            candidates[text] = ("caption", text)
        elif looks_like_byline and short_enough and text not in candidates:
            candidates[text] = ("by-line", text)

        prev_was_drawing = paragraph_has_drawing(block)

    return candidates


def find_figure_candidates(document, confirmed_headings):
    """Paragraphs that contain a drawing/pict/object. We can't read the
    internal labels of a SmartArt or chart, but we can surface the anchor
    so the user can decide whether a nearby caption should be excluded
    (via the caption prompt) or whether the figure itself carries prose."""
    candidates = []
    tracker = ExclusionTracker()
    for block in iter_block_items(document):
        if not isinstance(block, Paragraph):
            continue
        raw = block.text
        text = strip_markers(raw)
        has_start = START_MARK in raw
        has_end = END_MARK in raw

        lvl = effective_heading_level(block, confirmed_headings)
        if lvl is not None:
            tracker.visit_heading(lvl, text)

        if tracker.visit_markers(has_start, has_end) or tracker.in_struct:
            continue

        if paragraph_has_drawing(block):
            # Any text directly inside the drawing's paragraph (a textbox
            # that isn't split into its own body paragraph) is captured.
            inner_text = " ".join(
                t.text or "" for t in block._p.iter(f"{W_NS}t")
            ).strip()
            candidates.append((len(candidates) + 1, inner_text))
    return candidates


class Node:
    __slots__ = ("level", "title", "own_words", "children", "excluded")

    def __init__(self, level, title, excluded=False):
        self.level = level
        self.title = title
        self.own_words = 0
        self.children = []
        self.excluded = excluded


def subtotal(node):
    return node.own_words + sum(subtotal(c) for c in node.children)


def print_tree(node, indent=0):
    for child in node.children:
        total = subtotal(child)
        tag = "  [excluded from count]" if child.excluded else ""
        label = child.title.strip() or "(untitled heading)"
        print(f"{'    ' * indent}- {label}: {total} word{'s' if total != 1 else ''}{tag}")
        print_tree(child, indent + 1)


def get_path_via_dialog():
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None
    try:
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
    except Exception:
        return None


def estimate_page_count(document):
    """Rough page count from declared page size and body flow. Sections
    with landscape orientation are counted as landscape pages. Not
    perfect — Word's actual pagination depends on fonts, tables, breaks
    and images that we can't reproduce — but a better signal than
    nothing."""
    try:
        section = document.sections[0]
        pg_sz = section._sectPr.find(qn('w:pgSz'))
        if pg_sz is None:
            return None
        w = int(pg_sz.get(qn('w:w'), 12240))
        h = int(pg_sz.get(qn('w:h'), 15840))
        landscape = w > h
        margins = section._sectPr.find(qn('w:pgMar'))
        if margins is not None:
            top = int(margins.get(qn('w:top'), 1440))
            bottom = int(margins.get(qn('w:bottom'), 1440))
            usable_h = h - top - bottom
        else:
            usable_h = h - 2880
        # Twips: 1440 per inch. Assume ~11pt line height at 1.15 spacing
        # for body text — that's ~253 twips per line. Usable height in
        # twips divided by line height gives an approximate line count.
        lines_per_page = max(usable_h // 253, 1)
        # Assume ~12 words per line for 11pt body on A4/Letter with 1" margins.
        words_per_page = lines_per_page * 12
        return (w, h, landscape, words_per_page)
    except Exception:
        return None


# ------------------------------------------------------------------ main --

def main():
    if len(sys.argv) > 1:
        path = sys.argv[1]
    else:
        path = get_path_via_dialog()
        if not path:
            path = input("Path to your .docx file: ").strip().strip('"')

    if not os.path.isfile(path):
        print(f"\nCouldn't find a file at: {path}")
        print("Check the path is correct and try again.")
        return

    try:
        document = docx.Document(path)
    except Exception:
        print(f"\nCouldn't open that as a Word document: {path}")
        print("Make sure it's a real .docx file. If it's a .doc, open it in Word")
        print("and use File > Save As > Word Document (.docx) first.")
        return

    # ---- unstyled heading confirmation ----
    confirmed_headings = set()
    heading_candidates = find_heading_candidates(document)
    if heading_candidates:
        print(
            f"\nFound {len(heading_candidates)} line(s) that might be unstyled section "
            "headings (bold or all-caps text, not using Word's Heading style)."
        )
        for text in heading_candidates:
            ans = prompt_ynq(
                f"\nIs \"{text}\" a section heading (like 'Results' or 'Method')? (y/n/?): ",
                help_text=(
                    "Say yes only if this line titles the section that follows it. "
                    "Say no if it's just emphasised text inside a sentence."
                ),
            )
            if ans == "y":
                confirmed_headings.add(text)

    # ---- citation fields (Word / Zotero / EndNote) ----
    with zipfile.ZipFile(path) as z:
        try:
            doc_xml = z.read("word/document.xml")
        except KeyError:
            doc_xml = b""
    citation_texts = get_field_citation_texts(doc_xml)

    # ---- typed-out citations ----
    manual_citation_candidates = find_citation_candidates(document, confirmed_headings)
    if manual_citation_candidates:
        print(f"\nFound {len(manual_citation_candidates)} possible typed-out in-text citation(s).")
        for cite, context in manual_citation_candidates.items():
            ans = prompt_ynq(
                f"\n...{context}...\nIs \"{cite}\" an in-text citation? (y/n/?): ",
                help_text="Confirm only if this is a source citation.",
            )
            if ans == "y":
                citation_texts.append(cite)

    citation_excluded = 0

    # ---- equations and calculations ----
    equation_texts = []
    equation_candidates = find_equation_candidates(document, confirmed_headings)
    if equation_candidates:
        print(f"\nFound {len(equation_candidates)} possible equation/calculation item(s).")
        for eq, context in equation_candidates.items():
            ans = prompt_ynq(
                f"\n...{context}...\nExclude \"{eq.strip()}\" as an equation/calculation? (y/n/?): ",
                help_text=(
                    "Yes if this is a mathematical expression or a worked-calculation "
                    "line (e.g. '= 4.0 / 40.0'). No if it's ordinary prose that happens "
                    "to contain an equals sign."
                ),
            )
            if ans == "y":
                equation_texts.append(eq)

    equation_excluded = 0

    # ---- visual elements (captions, by-lines) ----
    visual_candidates = find_visual_element_candidates(document, confirmed_headings)
    visual_texts = []
    if visual_candidates:
        print(f"\nFound {len(visual_candidates)} possible visual element(s) (captions/by-lines).")
        for text, (kind, _) in visual_candidates.items():
            ans = prompt_ynq(
                f"\n({kind}) \"{text}\"\nExclude this as a visual element? (y/n/?): ",
                help_text=VISUAL_ELEMENT_EXPLANATION,
            )
            if ans == "y":
                visual_texts.append(text)

    visual_excluded = 0

    # ---- figures (drawings / picts / embedded objects) ----
    figure_anchors = find_figure_candidates(document, confirmed_headings)
    figure_inner_texts = []
    if figure_anchors:
        print(f"\nFound {len(figure_anchors)} paragraph(s) containing an embedded figure/diagram.")
        print("(The tool can't read internal labels of SmartArt or chart XML, only text")
        print(" that lives directly in the surrounding paragraph — e.g. a textbox caption.)")
        for idx, inner in figure_anchors:
            if not inner.strip():
                continue
            ans = prompt_ynq(
                f"\nFigure/diagram {idx} has in-paragraph text: \"{inner[:120]}...\"\n"
                "Does this figure contain information other than raw/processed data? (y/n/?): ",
                help_text=RAW_DATA_EXPLANATION,
            )
            if ans == "n":
                # It has prose -> per QCAA, whole figure counted. Nothing
                # to subtract, but the text is already inside the total.
                pass
            else:
                figure_inner_texts.append(inner)

    figure_inner_excluded = 0

        # ---- title page / front-matter block detection ----
    # We want to find everything before the first REAL content heading and
    # offer to exclude it as front matter. "Real content" means a heading
    # that is not itself one of the front-matter triggers (Declaration,
    # Authenticity, etc.) — because those are part of the block we want
    # to exclude, not the start of the response.
    title_page_excluded = 0
    title_page_words = 0
    blocks = list(iter_block_items(document))

    def is_frontmatter_heading(text):
        lower = text.strip().lower().rstrip(":.-")
        for trigger in ("declaration of authenticity", "declaration",
                        "authenticity", "academic integrity", "student declaration"):
            if lower == trigger or lower.startswith(trigger + " ") or lower.startswith(trigger + ":"):
                return True
        return False

    first_content_idx = None
    for i, b in enumerate(blocks):
        if not isinstance(b, Paragraph):
            continue
        lvl = effective_heading_level(b, confirmed_headings)
        if lvl is None:
            continue
        heading_text = strip_markers(b.text).strip()
        # A front-matter heading is not the start of the response body.
        if is_frontmatter_heading(heading_text):
            continue
        first_content_idx = i
        break

    # Gather all text between the start of the document and the first real
    # content heading.
    frontmatter_lines = []
    if first_content_idx is None:
        # No real content heading at all — treat the whole doc as candidate.
        scan_end = len(blocks)
    else:
        scan_end = first_content_idx
    for b in blocks[:scan_end]:
        if isinstance(b, Paragraph):
            t = strip_markers(b.text).strip()
            if t:
                frontmatter_lines.append(t)

    frontmatter_words = sum(word_count(t) for t in frontmatter_lines)
    # Offer exclusion if the block looks like front matter: it's not the
    # bulk of the document, and it either contains a declaration heading
    # or is short enough to be a cover page.
    has_declaration = any(is_frontmatter_heading(t) for t in frontmatter_lines)
    if frontmatter_lines and (
        has_declaration
        or (frontmatter_words <= 250 and len(frontmatter_lines) <= 12)
    ):
        preview = " | ".join(frontmatter_lines[:4])
        more = f" ... and {len(frontmatter_lines) - 4} more line(s)" if len(frontmatter_lines) > 4 else ""
        print(
            f"\nPossible front matter (title page / declaration): "
            f"{len(frontmatter_lines)} line(s), {frontmatter_words} words."
            f"\n  {preview}{more}"
        )
        ans = prompt_ynq(
            "Exclude this whole block as title page / declaration / front matter? (y/n/?): ",
            help_text=(
                "QCAA excludes title pages, contents pages, abstracts and declarations. "
                "Everything from the top of the document up to your first real content "
                "heading (Rationale, Introduction, Method, etc.) is front matter if it's "
                "just the title, author, teacher, date, word count, declaration and "
                "signature. Say yes to exclude it. Say no if your response actually "
                "starts before the first heading."
            ),
        )
        if ans == "y":
            title_page_excluded = frontmatter_words
            title_page_words = frontmatter_words
            
            
    # ---- main counting walk ----
    total_words = 0
    numeric_excluded = 0
    struct_excluded = 0
    manual_excluded = 0
    table_candidates = []

    tracker = ExclusionTracker()
    root_node = Node(level=-1, title="(document)")
    node_stack = [root_node]

    title_page_skip_until = first_content_idx if title_page_excluded else None
    if title_page_skip_until is None and title_page_excluded:
        title_page_skip_until = scan_end
    
    for i, block in enumerate(blocks):
        if isinstance(block, Paragraph):
            raw = block.text
        else:
            raw = table_text(block)
        has_start = START_MARK in raw
        has_end = END_MARK in raw
        text = strip_markers(raw)

        if isinstance(block, Paragraph):
            lvl = effective_heading_level(block, confirmed_headings)
            if lvl is not None:
                tracker.visit_heading(lvl, text)
                while node_stack[-1].level >= lvl:
                    node_stack.pop()
                new_node = Node(lvl, text, excluded=tracker.in_struct)
                node_stack[-1].children.append(new_node)
                node_stack.append(new_node)

        block_manual_excluded = tracker.visit_markers(has_start, has_end)
        words = word_count(text)
        current_node = node_stack[-1]

        # Title page blocks: if the user confirmed the cover, drop them.
        # Skip front-matter blocks if the user confirmed exclusion.
        if title_page_excluded and title_page_skip_until is not None and i < title_page_skip_until:
            continue
        if block_manual_excluded:
            manual_excluded += words
            continue
        if tracker.in_struct:
            struct_excluded += words
            continue

        # Visual element exclusion (captions, by-lines confirmed by user)
        block_visual_excluded = False
        for vt in visual_texts:
            if vt and vt in text:
                visual_excluded -= 0  # already accounted below
                block_visual_excluded = True
                break

        if isinstance(block, Table):
            num_in_table = numeric_symbol_count(text)
            preview = text.replace("\n", " ")[:150]
            table_candidates.append((current_node, preview, words, num_in_table))
            total_words += words
            continue

        # Paragraph
        #
        # Accounting rule (must match the table path above):
        #   total_words += words              (raw count, once)
        #   <excluded buckets> += <what's removed>
        #   current_node.own_words += counted (words - all_excluded - numeric)
        #
        # We do NOT reduce `text` before counting; we count the raw text
        # for numeric purposes and compute the section contribution by
        # subtracting every exclusion that applies to this paragraph.
        excluded_here = 0
        
        # Visual elements (captions, by-lines) confirmed by the user
        for vt in visual_texts:
            if vt and vt in text:
                excluded_here += word_count(vt)
                visual_excluded += word_count(vt)

        # Figure inner-text confirmed by the user as data-only
        for it in figure_inner_texts:
            if it and it in text:
                excluded_here += word_count(it)
                figure_inner_excluded += word_count(it)

        # In-text citations confirmed by the user
        for ct in citation_texts:
            if ct and ct in text:
                excluded_here += word_count(ct)
                citation_excluded += word_count(ct)

        # Equations and calculation lines confirmed by the user
        for et in equation_texts:
            if et and et in text:
                excluded_here += word_count(et)
                equation_excluded += word_count(et)

        # Strip confirmed citations and equations from the text before
        # counting numeric tokens, so a year inside "(Smith, 2020)" isn't
        # counted both as part of the citation AND as a number.
        text_for_numeric = text
        for ct in citation_texts:
            if ct and ct in text_for_numeric:
                text_for_numeric = text_for_numeric.replace(ct, " ")
        for et in equation_texts:
            if et and et in text_for_numeric:
                text_for_numeric = text_for_numeric.replace(et, " ")

        num = numeric_symbol_count(text_for_numeric)
        counted = words - excluded_here - num
        if counted < 0:
            counted = 0
        total_words += words
        numeric_excluded += num
        current_node.own_words += counted


    # ---- footnotes and endnotes ----
    footnotes, endnotes = get_footnote_endnote_texts(path)
    fn_words = sum(word_count(t) for t in footnotes)
    en_words = sum(word_count(t) for t in endnotes)
    total_words += fn_words + en_words

    footnote_excluded = 0
    endnote_excluded = 0

    if fn_words > 0:
        ans = prompt_ynmx(
            f"\nFootnotes contain {fn_words} words total.\n"
            "Are they ALL purely bibliographical? (y/n/m for mixed): ",
            help_text=(
                "y = every footnote is just a citation/reference.\n"
                "n = none are purely bibliographic (all count).\n"
                "m = mixed; you'll be asked about each one."
            ),
        )
        if ans == "y":
            footnote_excluded = fn_words
        elif ans == "m":
            print("Confirming footnotes one at a time:")
            for i, t in enumerate(footnotes, 1):
                w = word_count(t)
                preview = t[:100] + ("..." if len(t) > 100 else "")
                a = prompt_ynq(
                    f"\nFootnote {i} ({w} words): \"{preview}\"\n"
                    "Is this purely bibliographical (excluded)? (y/n): "
                )
                if a == "y":
                    footnote_excluded += w
                else:
                    numeric_excluded += numeric_symbol_count(t)
        else:  # n
            for t in footnotes:
                numeric_excluded += numeric_symbol_count(t)

    if en_words > 0:
        ans = prompt_ynmx(
            f"\nEndnotes contain {en_words} words total.\n"
            "Are they ALL purely bibliographical? (y/n/m for mixed): ",
            help_text="Same options as footnotes.",
        )
        if ans == "y":
            endnote_excluded = en_words
        elif ans == "m":
            print("Confirming endnotes one at a time:")
            for i, t in enumerate(endnotes, 1):
                w = word_count(t)
                preview = t[:100] + ("..." if len(t) > 100 else "")
                a = prompt_ynq(
                    f"\nEndnote {i} ({w} words): \"{preview}\"\n"
                    "Is this purely bibliographical (excluded)? (y/n): "
                )
                if a == "y":
                    endnote_excluded += w
                else:
                    numeric_excluded += numeric_symbol_count(t)
        else:
            for t in endnotes:
                numeric_excluded += numeric_symbol_count(t)

    # ---- table questions ----
    table_excluded = 0
    if table_candidates:
        print(f"\nFound {len(table_candidates)} table(s). (Type ? at any prompt for an explanation.)")
        for i, (node, preview, words, num_in_table) in enumerate(table_candidates, 1):
            ans = prompt_dci(
                f"\nTable {i} ({words} words): \"{preview}...\"\n"
                "Is this table:\n"
                "  d = raw or processed data only (excluded)\n"
                "  c = calculation working only (excluded)\n"
                "  i = contains information other than raw/processed data (included)\n"
                "Answer (d/c/i/?): ",
                help_text=RAW_DATA_EXPLANATION,
            )
            if ans == "d":
                # Data-only table: excluded entirely.
                table_excluded += words
            elif ans == "c":
                # Calculation working: excluded entirely under the
                # "numbers, symbols, equations and calculations" rule.
                # Same effect as 'd' but tracked separately so the report
                # can show it.
                table_excluded += words
            else:  # 'i' — include it
                numeric_excluded += num_in_table
                node.own_words += (words - num_in_table)
                
                
    # ---- report ----
    # IMPORTANT: only subtract buckets whose words actually made it into
    # total_words. The main loop's `continue` statements skip title-page,
    # struct-excluded, and manual-marker blocks BEFORE `total_words += words`
    # runs — so those words are already absent from total_words and must
    # not be subtracted again here. They're still displayed in the report
    # so you can see what was skipped.
    excluded = (
        numeric_excluded
        + citation_excluded
        + footnote_excluded
        + endnote_excluded
        + table_excluded
        + visual_excluded
        + figure_inner_excluded
    )
    final_count = max(total_words - excluded, 0)

    print("\n" + "=" * 48)
    print("QCAA WORD COUNT REPORT")
    print("=" * 48)
    print(f"Word total incl. footnotes/endnotes:   {total_words}")
    print("Excluded:")
    print(f"  Numbers/symbols:                       {numeric_excluded}")
    print(f"  Citations (fields + confirmed typed):  {citation_excluded}")
    print(f"  Equations/calculations:                {equation_excluded}")
    print(f"  Bibliographic footnotes:               {footnote_excluded}")
    print(f"  Bibliographic endnotes:                {endnote_excluded}")
    print(f"  Data-only tables:                      {table_excluded}")
    print(f"  Data-only figure inner-text:           {figure_inner_excluded}")
    print(f"  Visual elements (captions/by-lines):   {visual_excluded}")
    print(f"  Title page:                            {title_page_words}")
    print(f"  Contents/Abstract/Bibliography/Appx:   {struct_excluded}")
    print(f"  Manually marked:                       {manual_excluded}")
    print("-" * 48)
    print(f"ESTIMATED QCAA WORD COUNT: {final_count}")
    print("=" * 48)

    # Page estimate
    page_info = estimate_page_count(document)
    if page_info:
        w, h, landscape, wpp = page_info
        if final_count > 0:
            est_pages = max(1, round(final_count / wpp))
            orient = "landscape" if landscape else "portrait"
            print(
                f"\nPAGE ESTIMATE: ~{est_pages} page(s) "
                f"({orient}, ~{wpp} body words/page)."
            )
            print("Word's own pagination will differ based on fonts, tables,")
            print("images, breaks, and heading spacing — check against Word's")
            print("status bar before relying on this figure.")

    # Breakdown
    if root_node.children:
        print("\nBREAKDOWN BY HEADING (counted words only, excluded sections marked):")
        print_tree(root_node)
        if root_node.own_words:
            print(f"\n(+ {root_node.own_words} counted word(s) not under any heading)")
    else:
        print("\nNo headings found in this document, so no section breakdown is available.")

    print(
        "\nStill check by eye: text inside images, unusual citation styles, "
        "SmartArt labels, and any custom reference-manager field codes."
    )


if __name__ == "__main__":
    main()