"""
QCAA Word Count Tool (Python version)
--------------------------------------
Run with:  python qcaa_word_count.py "path/to/your/file.docx"
(or just run it and paste the path when asked)

Requires:  pip install python-docx

WHAT IT AUTOMATES
  - Counts all body text: paragraphs, headings/subheadings, table text,
    quotations (these are just normal text, nothing special needed)
  - Strips out standalone numbers/symbols (42, 3.14, %) while keeping
    abbreviations, units and chemical formulas (kg, KOH, LPG, 37kg)
  - Reads footnotes.xml / endnotes.xml directly (python-docx doesn't
    expose these) and asks once whether they're bibliography-only
  - Detects text inserted via Word's built-in Citation tool (CITATION
    fields), PLUS typed-out citations like (Smith, 2020), Smith (2020),
    or numbered [1]-style citations — asking you to confirm each unique
    one it finds — and excludes them
  - Finds headings that look like Contents/Abstract/Bibliography/
    References/Appendix and excludes everything under them, even if
    they aren't styled with Word's Heading styles (it'll ask you to
    confirm any bold/all-caps line that looks like an unstyled heading)
  - Flags text that looks like an inline equation (e.g. "x = 5cm") and
    asks you to confirm before excluding it — since an equation-shaped
    match can sometimes swallow more surrounding prose than intended,
    this is reviewed rather than stripped automatically
  - Asks you, table by table, whether it's raw/processed data only
    (type ? at that prompt for a fuller explanation of the distinction)
  - Breaks the final count down heading by heading, so you can see
    which sections are eating the most words

WHAT STILL NEEDS YOUR JUDGEMENT
  - Citations in a format the pattern-matcher doesn't recognise
  - Anything else — wrap it in the markers below, typed as plain text
    anywhere in the document:
        [[QCAA_EXCLUDE_START]]  ... content to ignore ...  [[QCAA_EXCLUDE_END]]
    e.g. wrap a title page or blank-page placeholder text in these.
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

START_MARK = "[[QCAA_EXCLUDE_START]]"
END_MARK = "[[QCAA_EXCLUDE_END]]"

STRUCT_KEYWORDS = [
    "table of contents", "contents", "abstract", "bibliography",
    "reference list", "references", "appendix", "appendices",
]

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

Quick test: if you deleted the table and only kept a sentence like
"see Table 1 for full results", would the marker lose anything
beyond raw numbers? If yes (there's discussion/labelling of
substance living only in the table), answer "n" (include it). If
the table is genuinely just numbers/results, answer "y" (exclude).
------------------------------------------------------------------
"""


# ---------------------------------------------------------------- helpers --

def prompt_ynq(message):
    """Ask for y / n / ? and keep looping until one of those is given."""
    while True:
        ans = input(message).strip().lower()
        if ans in ("y", "yes"):
            return "y"
        if ans in ("n", "no"):
            return "n"
        if ans in ("?", "help", "h"):
            return "?"
        print("Please enter y, n, or ?.")


def iter_block_items(document):
    """Yield paragraphs and tables in the order they appear in the body."""
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, document)
        elif child.tag == qn('w:tbl'):
            yield Table(child, document)


def tokenize(text):
    return re.findall(r"\S+", text)


def is_numeric_symbol_token(tok):
    """Digits but no letters -> excluded (42, 3.14, 1,000, %). Letters present
    (kg, KOH, LPG, 37kg) -> left alone, stays counted."""
    has_letter = any(c.isalpha() for c in tok)
    has_digit = any(c.isdigit() for c in tok)
    return has_digit and not has_letter


def word_count(text):
    return len([t for t in tokenize(text) if any(c.isalnum() for c in t)])


def numeric_symbol_count(text):
    return sum(1 for t in tokenize(text) if is_numeric_symbol_token(t))


def heading_level(paragraph):
    # Only Heading 1/2/3... styles form the section tree. The document's
    # "Title" style is treated as ordinary text (usually just the essay
    # title itself, not a container for the sections below it).
    style_name = paragraph.style.name if paragraph.style else ""
    m = re.match(r"Heading (\d+)", style_name or "")
    if m:
        return int(m.group(1))
    return None


def effective_heading_level(paragraph, confirmed_headings):
    """Like heading_level(), but also treats a paragraph as a (level 1)
    heading if its exact text was confirmed by the user as an unstyled
    heading (see find_heading_candidates)."""
    lvl = heading_level(paragraph)
    if lvl is not None:
        return lvl
    text = strip_markers(paragraph.text).strip()
    if text in confirmed_headings:
        return 1
    return None


def find_heading_candidates(document):
    """Find short, distinctly-formatted paragraphs that might be section
    headings without a real Word Heading style applied — common when a
    document started in Google Docs and lost its styles on copy-paste.

    Deliberately conservative: only flags a paragraph if EVERY run in it
    is bold, or the whole thing is in caps. A bold LEAD-IN word followed
    by ordinary sentence text (e.g. "**Caution**: handle acids carefully
    at all times.") won't qualify, since only part of it is bold — that
    kind of partial match is what makes a looser heuristic misfire on
    ordinary emphasised sentences instead of real headings."""
    candidates = {}
    for block in iter_block_items(document):
        if not isinstance(block, Paragraph):
            continue
        if heading_level(block) is not None:
            continue  # already a real Word heading style
        text = strip_markers(block.text).strip()
        if not text or len(text) > 80:
            continue
        runs = [r for r in block.runs if r.text.strip()]
        if not runs:
            continue
        all_bold = all(getattr(r.font, "bold", False) for r in runs)
        letters = [c for c in text if c.isalpha()]
        is_capsy = len(letters) >= 3 and sum(c.isupper() for c in letters) / len(letters) > 0.7
        if (all_bold or is_capsy) and text not in candidates:
            candidates[text] = True
    return candidates


def strip_markers(text):
    return text.replace(START_MARK, "").replace(END_MARK, "")


def table_text(table):
    """Cell text, joined — correctly handling merged cells. python-docx's
    row.cells repeats the SAME cell once per grid column/row it spans (e.g.
    a cell merged across 3 columns shows up 3 times), which would triple-count
    its words if joined naively. We de-duplicate merged cells by identity.

    Important: we compare with `is` against objects we keep alive in a list,
    NOT by stashing id(cell._tc) in a set. Cell wrapper objects are short-lived,
    and once one is garbage-collected Python can reuse its memory address for
    an unrelated cell's wrapper — comparing raw ids across that gap produces
    false "duplicate" matches and silently drops real cells. Keeping the
    actual objects alive in `seen` avoids that entirely."""
    seen = []
    parts = []
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            if any(tc is s for s in seen):
                continue
            seen.append(tc)
            parts.append(cell.text)
    return "\n".join(parts)


def get_footnote_endnote_text(path):
    texts = []
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        for fname, tag in (("word/footnotes.xml", "footnote"), ("word/endnotes.xml", "endnote")):
            if fname in names:
                root = ET.fromstring(z.read(fname))
                for note in root.iter(f"{W_NS}{tag}"):
                    note_type = note.get(f"{W_NS}type")
                    if note_type in ("separator", "continuationSeparator", "continuationNotice"):
                        continue
                    t = "".join(n.text or "" for n in note.iter(f"{W_NS}t"))
                    if t.strip():
                        texts.append(t)
    return texts


def get_citation_texts(document_xml_bytes):
    """Find text produced by Word's CITATION field (Insert Citation / EndNote / Zotero)."""
    root = ET.fromstring(document_xml_bytes)
    texts = []

    # Simple fields: <w:fldSimple w:instr=" CITATION ...">...<w:t>text</w:t>...</w:fldSimple>
    for fs in root.iter(f"{W_NS}fldSimple"):
        instr = fs.get(f"{W_NS}instr", "") or ""
        if "CITATION" in instr.upper():
            t = "".join(n.text or "" for n in fs.iter(f"{W_NS}t"))
            if t.strip():
                texts.append(t)

    # Complex fields: fldChar begin -> instrText -> fldChar separate -> t...t -> fldChar end
    in_field = False
    is_citation = False
    collecting = False
    buffer = ""
    for elem in root.iter():
        tag = elem.tag
        if tag == f"{W_NS}fldChar":
            fct = elem.get(f"{W_NS}fldCharType")
            if fct == "begin":
                in_field, is_citation, collecting, buffer = True, False, False, ""
            elif fct == "separate":
                collecting = True
            elif fct == "end":
                if is_citation and buffer.strip():
                    texts.append(buffer)
                in_field = collecting = is_citation = False
                buffer = ""
        elif tag == f"{W_NS}instrText" and in_field and not collecting:
            if elem.text and "CITATION" in elem.text.upper():
                is_citation = True
        elif tag == f"{W_NS}t" and collecting:
            buffer += elem.text or ""

    return texts


def strip_substrings(text, substrings):
    out = text
    for s in substrings:
        if s and s in out:
            out = out.replace(s, " ")
    return out


# Common author-date citation shapes: (Smith, 2020), (Smith & Jones, 2020),
# (Smith et al., 2020, p. 12), (How Does X Work?, 2025), multiple citations
# separated by semicolons, etc. Requires a comma before the year so it won't
# false-match ordinary parentheticals like "(max-min)" or "(5°C)".
FULL_CITATION_RE = re.compile(
    r"\([^()]{1,120}?,\s*(?:19|20)\d{2}[a-z]?(?:,\s*pp?\.?\s*\d+(?:[-–]\d+)?)?\)"
)
# Bare year in parentheses for narrative-style citations, e.g. "Laidler (2025)".
BARE_YEAR_RE = re.compile(r"\((?:19|20)\d{2}[a-z]?\)")
# Numbered/Vancouver-style citations, e.g. "...as shown previously [3]."
VANCOUVER_RE = re.compile(r"\[\d{1,3}\]")

# A simple inline equation shape: an isolated single letter (a variable),
# an equals-like sign, then up to 80 characters to the next period. This is
# deliberately shown to the user for confirmation rather than stripped
# automatically — the non-greedy match can end up including real
# explanatory prose after the "=" (e.g. "x = 159 seconds for this trial."
# would match in full), so a review step matters here more than usual.
EQUATION_RE = re.compile(
    r"(?<!\w)[A-Za-z\u0394\u03b4\u03b8\u03bb\u03bc\u03c3\u03c0]\s*[=\u2243\u2248]\s*.{1,80}?(?=$|\.)"
)


class ExclusionTracker:
    """Tracks whether we're currently inside a struct-excluded heading
    section (Contents/Bibliography/etc.) or a manually [[QCAA_EXCLUDE_...]]
    marked region, walking blocks in document order. Shared by the citation
    scanner and the main counting loop so they never disagree about what's
    excluded."""

    def __init__(self):
        self.in_struct = False
        self.struct_level = None
        self.in_manual = False

    def visit_heading(self, lvl, text):
        """Call for every heading paragraph (lvl is not None), before
        checking self.in_struct for this block."""
        if self.in_struct and lvl <= self.struct_level:
            self.in_struct = False
            self.struct_level = None
        if not self.in_struct:
            lower = text.lower()
            if any(k in lower for k in STRUCT_KEYWORDS):
                self.in_struct = True
                self.struct_level = lvl

    def visit_markers(self, has_start, has_end):
        """Call for every block. Returns True if THIS block counts as
        manually excluded."""
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
    """Scan paragraph and table text for things that look like typed-out
    in-text citations (as opposed to real Word/EndNote/Zotero citation
    fields, which are handled separately). Returns an ordered dict of
    {matched_text: one_example_context}. Skips anything already inside a
    struct-excluded section (Contents/References/etc.) or a manual
    [[QCAA_EXCLUDE_...]] marker — no point asking about citations that
    are already being thrown out."""
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

        block_manual_excluded = tracker.visit_markers(has_start, has_end)
        if block_manual_excluded or tracker.in_struct:
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
    """Find text shaped like a simple inline equation (e.g. "x = 5cm").
    Every match is returned for the user to confirm rather than stripped
    automatically — see the note on EQUATION_RE for why that matters here."""
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

        block_manual_excluded = tracker.visit_markers(has_start, has_end)
        if block_manual_excluded or tracker.in_struct:
            continue

        for m in EQUATION_RE.finditer(text):
            full = m.group(0)
            if full not in candidates:
                start = max(m.start() - 30, 0)
                end = min(m.end() + 10, len(text))
                candidates[full] = text[start:end]
    return candidates


class Node:
    """One heading in the document's outline tree."""
    __slots__ = ("level", "title", "own_words", "children", "excluded")

    def __init__(self, level, title, excluded=False):
        self.level = level
        self.title = title
        self.own_words = 0          # counted words directly under this heading
        self.children = []
        self.excluded = excluded    # True if this whole section is excluded (Contents/Bibliography/etc.)


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
    """Try to open a native file picker so the user can browse to their file
    instead of typing a path. Returns None if a GUI isn't available (e.g.
    some remote/headless setups) or the user cancels."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:
        return None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select your Word document",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
        )
        root.destroy()
        return path or None
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
        print("Make sure it's a real .docx file — this tool can't read old-style")
        print(".doc files, PDFs, or a Google Docs link directly. If it's a .doc,")
        print("open it in Word and use File > Save As > Word Document (.docx) first.")
        return

    total_words = 0
    numeric_excluded = 0
    struct_excluded = 0
    manual_excluded = 0
    table_candidates = []  # (node, preview, words, numeric_count)

    tracker = ExclusionTracker()

    root_node = Node(level=-1, title="(document)")
    node_stack = [root_node]

    # Unstyled headings: bold/all-caps lines that might be section titles
    # without a real Word Heading style applied (see find_heading_candidates
    # for why this is deliberately conservative).
    confirmed_headings = set()
    heading_candidates = find_heading_candidates(document)
    if heading_candidates:
        print(
            f"\nFound {len(heading_candidates)} line(s) that might be unstyled section "
            "headings (bold or all-caps text, not using Word's Heading style)."
        )
        for text in heading_candidates:
            ans = prompt_ynq(f"\nIs \"{text}\" a section heading (like 'Results' or 'Method')? (y/n/?): ")
            if ans == "?":
                print(
                    "Say yes only if this line is actually the title of the section that "
                    "follows it (like 'Results' or 'Evaluation'). Say no if it's just "
                    "emphasised text within a sentence (like a bolded warning label) — "
                    "wrongly confirming those can cause real content to be excluded if it "
                    "later happens to mention a word like 'references' or 'appendix'."
                )
                ans = prompt_ynq(f"Is \"{text}\" a section heading? (y/n): ")
            if ans == "y":
                confirmed_headings.add(text)

    # Citations: gathered up front so paragraph text can be de-duplicated
    # against them before counting numeric tokens (avoids double-subtracting
    # a number that's part of a citation, e.g. the "2020" in "(Smith, 2020)").
    with zipfile.ZipFile(path) as z:
        doc_xml = z.read("word/document.xml")
    citation_texts = get_citation_texts(doc_xml)

    # Typed-out citations (not real citation fields) won't show up in the
    # above — common if the document started life in Google Docs or was
    # typed by hand. Look for the usual (Author, Year), Author (Year), and
    # [N] shapes and confirm each unique one with the user before treating
    # it as a citation.
    manual_citation_candidates = find_citation_candidates(document, confirmed_headings)
    if manual_citation_candidates:
        print(f"\nFound {len(manual_citation_candidates)} possible typed-out in-text citation(s).")
        for cite, context in manual_citation_candidates.items():
            ans = prompt_ynq(f"\n...{context}...\nIs \"{cite}\" an in-text citation? (y/n/?): ")
            if ans == "?":
                print("Confirm only if this is a source citation, not just a bracketed number or date used for another reason.")
                ans = prompt_ynq(f"Is \"{cite}\" an in-text citation? (y/n): ")
            if ans == "y":
                citation_texts.append(cite)

    citation_excluded = sum(word_count(t) for t in citation_texts)

    # Equations: shown to the user rather than stripped blind, since the
    # match can sometimes include real explanatory prose alongside the
    # equation itself (see EQUATION_RE).
    equation_texts = []
    equation_candidates = find_equation_candidates(document, confirmed_headings)
    if equation_candidates:
        print(f"\nFound {len(equation_candidates)} possible inline equation(s).")
        for eq, context in equation_candidates.items():
            ans = prompt_ynq(f"\n...{context}...\nExclude \"{eq.strip()}\" as an equation? (y/n/?): ")
            if ans == "?":
                print(
                    "Only say yes if this is purely a mathematical expression. If it also "
                    "describes or explains a result in words, say no — otherwise that "
                    "explanatory text gets excluded too, since the match runs to the next "
                    "full stop."
                )
                ans = prompt_ynq(f"Exclude \"{eq.strip()}\" as an equation? (y/n): ")
            if ans == "y":
                equation_texts.append(eq)

    equation_excluded = sum(word_count(t) for t in equation_texts)

    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            raw_text = block.text
        else:
            raw_text = table_text(block)

        has_start = START_MARK in raw_text
        has_end = END_MARK in raw_text
        text = strip_markers(raw_text)

        if isinstance(block, Paragraph):
            lvl = effective_heading_level(block, confirmed_headings)
            if lvl is not None:
                tracker.visit_heading(lvl, text)

                # Update the heading tree
                while node_stack[-1].level >= lvl:
                    node_stack.pop()
                new_node = Node(lvl, text, excluded=tracker.in_struct)
                node_stack[-1].children.append(new_node)
                node_stack.append(new_node)
        else:
            lvl = None

        words = word_count(text)

        # Work out whether *any* part of this block should count as manually excluded.
        # Marker handling is block-granular: if a start/end marker appears in this
        # block, the whole block is treated as excluded (simplest + safest option).
        block_manual_excluded = tracker.visit_markers(has_start, has_end)

        current_node = node_stack[-1]

        if block_manual_excluded:
            manual_excluded += words
        elif tracker.in_struct:
            struct_excluded += words
        else:
            if isinstance(block, Table):
                # Numeric stripping (and section attribution) for tables is
                # deferred until we know whether the whole table gets
                # excluded (see Q&A below) — otherwise a data-only table's
                # numbers get subtracted twice.
                num_in_table = numeric_symbol_count(text)
                preview = text.replace("\n", " ")[:150]
                table_candidates.append((current_node, preview, words, num_in_table))
                total_words += words
            else:
                text_reduced = strip_substrings(text, citation_texts)
                text_reduced = strip_substrings(text_reduced, equation_texts)
                counted = word_count(text_reduced) - numeric_symbol_count(text_reduced)
                total_words += words
                numeric_excluded += numeric_symbol_count(text_reduced)
                current_node.own_words += counted

    # Footnotes / endnotes — added to the total now; numeric stripping is
    # deferred the same way as tables, until we know if they're excluded outright.
    # (Not attributed to a specific heading — footnote anchors aren't reliably
    # traceable back to a section without much more XML plumbing.)
    fn_en_texts = get_footnote_endnote_text(path)
    fn_en_words = sum(word_count(t) for t in fn_en_texts)
    total_words += fn_en_words

    # ---- interactive questions ----
    footnote_excluded = 0
    if fn_en_words > 0:
        ans = input(
            f"\nFootnotes/endnotes contain {fn_en_words} words total.\n"
            "Are ALL of them used only for bibliographical/citation purposes "
            "(none contain content, explanation or argument)? (y/n): "
        ).strip().lower()
        if ans.startswith("y"):
            footnote_excluded = fn_en_words
        else:
            numeric_excluded += sum(numeric_symbol_count(t) for t in fn_en_texts)

    table_excluded = 0
    if table_candidates:
        print(f"\nFound {len(table_candidates)} table(s). (Type ? at any prompt for an explanation.)")
        for i, (node, preview, words, num_in_table) in enumerate(table_candidates, 1):
            while True:
                ans = input(
                    f"\nTable {i} ({words} words): \"{preview}...\"\n"
                    "Does this table contain ONLY raw or processed data? (y/n/?): "
                ).strip().lower()
                if ans in ("?", "help", "h"):
                    print(RAW_DATA_EXPLANATION)
                    continue
                break
            if ans.startswith("y"):
                table_excluded += words
            else:
                numeric_excluded += num_in_table
                node.own_words += (words - num_in_table)

    excluded = numeric_excluded + citation_excluded + footnote_excluded + table_excluded
    final_count = max(total_words - excluded, 0)

    print("\n" + "=" * 44)
    print("QCAA WORD COUNT REPORT")
    print("=" * 44)
    print(f"Word total incl. footnotes/endnotes: {total_words}")
    print("Excluded:")
    print(f"  Numbers/symbols:                          {numeric_excluded}")
    print(f"  Citations (fields + confirmed typed):     {citation_excluded}")
    print(f"  Bibliographic footnotes/endnotes:          {footnote_excluded}")
    print(f"  Data-only tables:                          {table_excluded}")
    print(f"  Contents/Abstract/Bibliography/Appendix:   {struct_excluded}")
    print(f"  Manually marked ([[QCAA_EXCLUDE_...]]):    {manual_excluded}")
    print("-" * 44)
    print(f"ESTIMATED QCAA WORD COUNT: {final_count}")
    print("=" * 44)

    if root_node.children:
        print("\nBREAKDOWN BY HEADING (counted words only, excluded sections marked):")
        print_tree(root_node)
        if root_node.own_words:
            print(f"\n(+ {root_node.own_words} counted word(s) not under any heading)")
    else:
        print("\nNo headings found in this document, so no section breakdown is available.")
        print("Tip: applying Word's Heading 1/2/3 styles to your section titles will")
        print("let this tool show you a per-section breakdown next time.")

    print(
        "\nStill worth checking by eye: equations in running text, and any "
        "citations in a format the pattern-matcher might have missed."
    )


if __name__ == "__main__":
    main()