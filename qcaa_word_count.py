from __future__ import annotations

code = r'''#!/usr/bin/env python3
"""
QCAA Word Count Tool v2
=======================
Run with:  python qcaa_word_count.py "path/to/your/file.docx"
(or just run it - a file picker opens, or paste the path when asked)

Requires:  pip install python-docx

WHAT THIS TOOL DOES
  Estimates the QCAA word count for a response, following the official
  QCAA word-length rules:

  INCLUDED
    all words in the text of the response
    titles, headings and subheadings
    tables/figures/maps/diagrams containing information OTHER than
      raw or processed data (the WHOLE table/figure counts)
    quotations
    footnotes and endnotes (unless purely bibliographical)
    abbreviations, initialisms (LPG), units (kg, m), chemical
      formulas (KOH, HCl) -- anything with letters in it

  EXCLUDED
    title pages, contents pages, abstract, blank pages
    visual elements of written genres (by-lines, banners, captions,
      call-outs in articles/blogs/essays/columns)
    raw or processed data in tables/figures/diagrams
    numbers, symbols, equations and calculations
    bibliography / reference list
    appendixes (supplementary material only)
    page numbers
    in-text citations

QUALITY-OF-LIFE FEATURES (kept from v1)
  - File-picker dialog if you just double-click the script
  - '?' at any prompt shows the relevant rule with examples
  - Plain-English explanations before each category of question
  - Per-heading breakdown at the end
  - Manual override: wrap anything in
        [[QCAA_EXCLUDE_START]] ... [[QCAA_EXCLUDE_END]]
    (typed as plain text) and it is excluded no matter what

NEW IN V2
  - Assessment-type profiles (generic, student-experiment, psmt,
    literary, extended-essay). The QCAA RULES are identical for every
    type -- profiles only tune which candidates are surfaced and the
    suggested default answers, plus guidance text in the prompts.
  - Saved decisions: your answers are stored in
        <file>.qcaa-decisions.json
    Re-run the tool and you are only asked about anything NEW.
  - Non-interactive mode (-n) for batch processing, with a
    "needs review" list instead of prompts.
  - High-confidence auto-exclusions (logged, reviewable with
    --review-auto): TOC/page-number fields, Word "Caption"-styled
    captions next to images, math objects, structural sections.
  - Reads text boxes, SmartArt internal text (word/diagrams/*.xml),
    and flags chart XML the tool cannot judge.
  - Compares against Word's own saved word count as a sanity check.
  - Batch mode: pass several .docx files, get a CSV-style summary.

STILL CHECK BY EYE
  - Text inside images (screenshots, scanned figures)
  - Unusual citation styles the pattern-matcher does not know
  - Tracked changes: the tool counts inserted text and ignores
    deleted text (Word's "final" view) -- accept/reject first if
    your draft is heavily marked up
  - Whether an appendix is genuinely supplementary-only

'''

import argparse
import hashlib
import json
import os
import re
import sys
import zipfile
from collections import defaultdict
from xml.etree import ElementTree as ET

try:
    import docx
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("This tool needs python-docx.\nInstall it with:  pip install python-docx")

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

START_MARK = "[[QCAA_EXCLUDE_START]]"
END_MARK = "[[QCAA_EXCLUDE_END]]"

APP_TITLE = "QCAA Word Count Tool v2"

# ----------------------------------------------------------------------------
# QCAA rule text (shown by --explain and in '?' help)
# ----------------------------------------------------------------------------

QCAA_RULES = """\
QCAA WORD-LENGTH RULES (official summary)
-----------------------------------------
INCLUDED                                    EXCLUDED
all words in the text of the response       title pages
title, headings and subheadings             contents pages
tables/figures/maps/diagrams containing     abstract
  information OTHER than raw or processed   visual elements of written genres*
  data (the WHOLE item counts)              raw or processed data in tables,
quotations                                    figures and diagrams
footnotes/endnotes (unless bibliographic)   numbers, symbols, equations,
abbreviations, initialisms (LPG), units       calculations
  (kg, m), chemical formulas (KOH, HCl)     bibliography / reference list
                                            appendixes (supplementary only)
                                            page numbers
                                            in-text citations
                                            blank pages

* by-lines, banners, captions and call-outs that are the visual elements of
  written genres suitable for print/online publication (literary article,
  blog, essay, column).

Raw data       = results exactly as collected/measured.
Processed data = the same data after basic manipulation (totals, means,
                 percentages, graphs plotted straight from the numbers).
Anything with written interpretation, commentary, qualitative labels beyond
simple column headers, or your own categorisations makes the WHOLE item count.
"""

RAW_DATA_EXPLANATION = """
------------------------------------------------------------------
RAW vs PROCESSED DATA vs "OTHER INFORMATION" (QCAA rule)
------------------------------------------------------------------
QCAA excludes a table/figure/diagram ONLY if it contains nothing
but raw or processed data:

  Raw data       = results exactly as collected/measured, e.g.
            a       each participant's individual score, each
                   trial's reading, unedited survey responses.
  Processed data = that same data after basic manipulation, e.g.
                   totals, means, percentages, a graph plotted
                   straight from the numbers. Still just
                   numbers/results -- no written interpretation.

If the table/figure ALSO contains "information other than raw or
processed data" -- written analysis or commentary inside it,
qualitative/descriptive labels beyond simple column headers,
annotations explaining what the data means, categorisations you
made -- then QCAA counts the WHOLE table/figure, not just the
extra text.

Quick test: if you deleted it and kept only a sentence like "see
Table 1 for full results", would the marker lose anything beyond
raw numbers? If yes, answer "i" (include). If it is genuinely
just numbers/results, answer "d" (exclude data) or "c" if it is
purely calculation working.
------------------------------------------------------------------
"""

VISUAL_ELEMENT_EXPLANATION = """
------------------------------------------------------------------
VISUAL ELEMENTS OF WRITTEN GENRES (QCAA exclusion)
------------------------------------------------------------------
QCAA excludes visual elements that decorate or frame a written
piece rather than carry its argument -- specifically by-lines,
banners, captions and call-outs of the kind you'd see in a
literary article, blog, essay or column.

Examples that SHOULD be excluded:
  - "Figure 3: Reaction rate vs. temperature"  (a label)
  - "By Jane Smith, 12 March 2024"
  - A pull-quote banner restating a sentence
  - A call-out box repeating a key statistic
  - "Source: QCAA (2023)" under a reproduced graphic

Examples that should NOT be excluded:
  - A sentence that happens to start with "Figure 3 shows..."
  - Ordinary body prose that names a figure inline
  - The caption of a figure in a science report, where it
    identifies variables or conditions (judgement call -- when
    unsure, include it)

Say yes only if this line exists purely to label/announce
something else visually, not to carry meaning of its own.
------------------------------------------------------------------
"""

FRONTMATTER_EXPLANATION = """
------------------------------------------------------------------
TITLE PAGE / FRONT MATTER (QCAA exclusion)
------------------------------------------------------------------
QCAA excludes title pages, contents pages, abstracts and
declarations. Everything from the top of the document up to your
first real content heading (Rationale, Introduction, Method...)
is front matter if it is just the title, author, teacher, date,
word count, declaration and signature.

Say yes to exclude the whole block. Say no if your response
actually starts before the first heading.
------------------------------------------------------------------
"""

# ----------------------------------------------------------------------------
# Assessment-type profiles.
# The QCAA counting RULES are the same for every assessment type. Profiles
# only change: (a) guidance text shown in prompts, (b) the suggested default
# answers (always overridable), and (c) which auto-detections are aggressive.
# ----------------------------------------------------------------------------

PROFILES = {
    "generic": {
        "description": "General QCAA response (default)",
        "guidance": "",
        "table_default": "smart",   # smart | d | c | i | ask
        "caption_default": "ask",   # ask | y | n
        "byline_default": "ask",
        "calc_default": "ask",      # ask | y | n  (equation/calc lines)
        "numeric_table_bias": None,
    },
    "student-experiment": {
        "description": "Student experiment / science investigation (IA, EEI, student experiment report)",
        "guidance": (
            "Student experiment tips:\n"
            "  - Tables of repeated measurements / raw readings = raw data (exclude, 'd').\n"
            "  - Tables with means/percentages plotted straight from readings = processed data ('d').\n"
            "  - Risk-assessment, method and results-discussion tables usually contain prose\n"
            "    (hazards, explanations) = information other than data (include, 'i').\n"
            "  - Calculated example lines ('= 4.0 / 40.0') = calculations (exclude, 'c')."
        ),
        "table_default": "smart",
        "caption_default": "ask",
        "byline_default": "n",
        "calc_default": "y",
        "numeric_table_bias": "d",
    },
    "psmt": {
        "description": "Problem Solving and Modelling Task (PSMT) / mathematical modelling",
        "guidance": (
            "PSMT tips:\n"
            "  - Verify/Interpret sections are usually full of calculation working:\n"
            "    chains like '= 3.2 x 4.1' are excluded as calculations.\n"
            "  - A table that is purely algebraic/numeric working = 'c'.\n"
            "  - Tables explaining your assumptions or evaluating the model = 'i'.\n"
            "  - The written Interpret/Evaluate/Verify discussion always counts."
        ),
        "table_default": "smart",
        "caption_default": "ask",
        "byline_default": "n",
        "calc_default": "y",
        "numeric_table_bias": "c",
    },
    "literary": {
        "description": "Written genre for publication (feature article, blog, essay, column, speech)",
        "guidance": (
            "Written-genre tips:\n"
            "  - By-lines, banners, pull-quotes, captions and call-outs are the\n"
            "    classic QCAA visual-element exclusions -- when the tool flags\n"
            "    them, the default answer is yes (exclude).\n"
            "  - The body prose of the piece always counts, including quotations."
        ),
        "table_default": "smart",
        "caption_default": "y",
        "byline_default": "y",
        "calc_default": "n",
        "numeric_table_bias": None,
    },
    "extended-essay": {
        "description": "Extended essay / research task with heavy referencing",
        "guidance": (
            "Research-task tips:\n"
            "  - Footnotes/endnotes are common here: bibliographic ones are\n"
            "    excluded, content notes count.\n"
            "  - The bibliography/reference list and any appendix are excluded\n"
            "    automatically once their headings are recognised."
        ),
        "table_default": "smart",
        "caption_default": "ask",
        "byline_default": "n",
        "calc_default": "ask",
        "numeric_table_bias": None,
    },
}

# ----------------------------------------------------------------------------
# Structural section triggers. Matched against the heading's trimmed,
# lowercased text as an EXACT match or a startswith -- NOT a substring, so
# "References to the Divine in Paradise Lost" does not trigger.
# ----------------------------------------------------------------------------

STRUCT_TRIGGERS = {
    "table of contents": "contents",
    "contents": "contents",
    "list of figures": "contents",
    "list of tables": "contents",
    "abstract": "abstract",
    "bibliography": "bibliography",
    "reference list": "bibliography",
    "references": "bibliography",
    "list of references": "bibliography",
    "works cited": "bibliography",
    "appendix": "appendix",
    "appendices": "appendix",
    "annex": "appendix",
    # Front-matter declarations (QCAA treats these as title-page/front matter)
    "declaration of authenticity": "frontmatter",
    "declaration": "frontmatter",
    "authenticity": "frontmatter",
    "academic integrity": "frontmatter",
    "student declaration": "frontmatter",
    "declaration of originality": "frontmatter",
}

FRONTMATTER_HEADING_TRIGGERS = tuple(
    k for k, v in STRUCT_TRIGGERS.items() if v == "frontmatter"
)

# Field codes treated as citation-producing (Word, Zotero, EndNote, Mendeley).
CITATION_FIELD_CODES = (
    "CITATION",
    "ZOTERO_ITEM",
    "ZOTERO_BIBL",
    "ADDIN EN.CITE",
    "ADDIN ZOTERO_ITEM",
    "MENDELEY",
    "CSL_CITATION",
)

# ----------------------------------------------------------------------------
# Citation / equation patterns (candidate detectors -- matches are confirmed
# by the user or by profile defaults, never silently except field citations)
# ----------------------------------------------------------------------------

FULL_CITATION_RE = re.compile(
    r"\([^()]{1,160}?,\s*(?:19|20)\d{2}[a-z]?"
    r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?"
    r"(?:\s*;\s*[^()]{1,160}?,\s*(?:19|20)\d{2}[a-z]?"
    r"(?:\s*,\s*pp?\.?\s*\d+(?:\s*[-\u2013]\s*\d+)?)?)*\s*\)"
)
NARRATIVE_CITATION_RE = re.compile(
    r"\b[A-Z][A-Za-z'’\-]+(?:\s+(?:and|&)\s+[A-Z][A-Za-z'’\-]+|\s+et\s+al\.?)?"
    r"\s+\((?:19|20)\d{2}[a-z]?\)"
)
BARE_YEAR_RE = re.compile(r"\((?:19|20)\d{2}[a-z]?\)")
VANCOUVER_RE = re.compile(r"\[\d{1,4}(?:\s*[,\-\u2013]\s*\d{1,4})*\]")
IBID_RE = re.compile(r"\(\s*(?:ibid\.?|op\.\s*cit\.?|loc\.\s*cit\.?|l\.c\.)\s*(?:[;,]?\s*[^)]*)?\)", re.IGNORECASE)

EQUATION_RE = re.compile(
    r"""
    (?<![A-Za-z0-9])                    # avoid matching inside words
    [A-Za-zΔα-ωΑ-Ω][A-Za-z0-9_²³⁴⁵⁶⁷⁸⁹]{0,15}\s*
    (?:=|≈|≠|≤|≥)\s*
    [^,.;:!?]{1,60}?
    (?=[,.;:!?]|$)
    """,
    re.VERBOSE,
)
CALC_CONTINUATION_RE = re.compile(
    r"""
    ^\s*
    (?:[+\-×*/÷^=]\s*)?                  # may start with an operator or '='
    \d+(?:\.\d+)?\s*
    (?:[+\-×*/÷^]\s*(?:\d+(?:\.\d+)?|[A-Za-zπΔα-ωΑ-Ω]+))
    (?:\s*[+\-×*/÷^=]\s*(?:\d+(?:\.\d+)?|[A-Za-zπΔα-ωΑ-Ω]+))*
    \s*$
    """,
    re.VERBOSE,
)
CALC_LINE_RE = re.compile(r"^\s*=\s*\S")

CAPTION_TEXT_RE = re.compile(
    r"^\s*(figure|fig\.?|table|chart|graph|diagram|image|photo|map|exhibit)\s*"
    r"[\dIVXivx]+[a-z]?\s*[:.\u2013\u2014\-]",
    re.IGNORECASE,
)
BYLINE_RE = re.compile(
    r"^\s*(by|words by|illustration by|illustrated by|photo by|images? by)\s+\S+",
    re.IGNORECASE,
)
SOURCE_LINE_RE = re.compile(r"^\s*source\s*[:.\u2013\-]\s*\S+", re.IGNORECASE)

# ----------------------------------------------------------------------------
# Small helpers
# ----------------------------------------------------------------------------

USE_COLOR = sys.stdout.isatty() and os.name != "nt" or (os.name == "nt" and os.environ.get("WT_SESSION"))


def c(text, color):
    if not USE_COLOR:
        return text
    codes = {"bold": "\033[1m", "dim": "\033[2m", "red": "\033[31m", "green": "\033[32m", "yellow": "\033[33m", "cyan": "\033[36m"}
    return f"{codes.get(color, '')}{text}\033[0m"


def tokenize(text):
    return re.findall(r"\S+", text)


def is_numeric_symbol_token(tok):
    """Digits but no letters -> excluded (42, 3.14, 1,000, 37%).
    Letters present (kg, KOH, LPG, 37kg) -> left alone."""
    has_letter = any(ch.isalpha() for ch in tok)
    has_digit = any(ch.isdigit() for ch in tok)
    return has_digit and not has_letter


def word_count(text):
    return len([t for t in tokenize(text) if any(ch.isalnum() for ch in t)])


def numeric_symbol_count(text):
    return sum(1 for t in tokenize(text) if is_numeric_symbol_token(t))


def strip_markers(text):
    return text.replace(START_MARK, "").replace(END_MARK, "")


def count_occurrences(text, target):
    return text.count(target) if target else 0


def sha1_file(path, _bufsize=1 << 20):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        while True:
            b = f.read(_bufsize)
            if not b:
                break
            h.update(b)
    return h.hexdigest()


# ----------------------------------------------------------------------------
# Decision store: remembers your answers between runs
# ----------------------------------------------------------------------------

class Decisions:
    """Loads/saves user answers so re-runs only ask about new candidates."""

    def __init__(self, path=None):
        self.path = path
        self.data = {"version": 2, "file_sha1": None, "decisions": {}, "auto": []}
        if path and os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    loaded = json.load(f)
                if loaded.get("version") == 2:
                    self.data.update(loaded)
            except Exception:
                pass

    def saved(self, key):
        return self.data["decisions"].get(key)

    def record_auto(self, key, answer, reason):
        self.data["auto"].append({"key": key, "answer": answer, "reason": reason})

    def resolve(self, key, prompt, choices, default, help_text=None,
                interactive=True, review_auto=False):
        """Returns one of `choices`. Saved answer wins; then interactive
        prompt (with default); then the default itself."""
        saved = self.saved(key)
        if saved in choices:
            default = saved
        if not interactive:
            if saved is None:
                self.data["auto"].append({"key": key, "answer": default, "reason": "default (non-interactive)"})
            return default
        suffix = "/".join(choices)
        while True:
            ans = input(f"{prompt} ({suffix}) [{default}]: ").strip().lower()
            if not ans:
                ans = default
            if ans in choices:
                self.data["decisions"][key] = ans
                return ans
            if ans in ("?", "help", "h") and help_text:
                print(help_text)
                continue
            print(f"Please enter {suffix} or ?.")


# ----------------------------------------------------------------------------
# Document model: walk the body with lxml so text boxes, fields and math are
# all visible to us (python-docx's paragraph.text misses text-box content).
# ----------------------------------------------------------------------------

class Doc:
    def __init__(self, path):
        self.path = path
        self.document = docx.Document(path)
        self.body = self.document.element.body
        self.style_names = {s.style_id: s.name for s in self.document.styles}
        # blocks: list of ("p"|"tbl", element)
        self.blocks = [
            (kind, el)
            for el in self.body
            for kind in ("p", "tbl")
            if el.tag == qn(f"w:{kind}")
        ]
        self.fields = self._analyse_fields()
        self.app_props = self._read_app_props()
        self.diagram_texts = self._read_diagram_texts()
        self.chart_parts = self._list_chart_parts()

    # -- text extraction ------------------------------------------------------

    @staticmethod
    def visible_text(el):
        """All w:t descendants -- includes text-box text and field results,
        excludes field instructions and tracked deletions."""
        return "".join(t.text or "" for t in el.iter(qn("w:t")))

    @staticmethod
    def has_math(el):
        return el.find(f".//{qn('m:oMath')}") is not None or \
               el.find(f".//{qn('m:oMathPara')}") is not None

    @staticmethod
    def has_drawing(el):
        for tag in ("w:drawing", "w:pict", "w:object"):
            if el.find(f".//{qn(tag)}") is not None:
                return True
        return False

    def style_name(self, el):
        ps = el.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
        if ps is None:
            return ""
        return self.style_names.get(ps.get(qn("w:val")), "")

    def heading_level(self, el):
        m = re.match(r"Heading (\d+)", self.style_name(el) or "")
        return int(m.group(1)) if m else None

    def is_caption_styled(self, el):
        name = (self.style_name(el) or "").lower()
        return name.startswith("caption")

    def para_bold_ratio(self, el):
        runs = [r for r in el.iter(qn("w:r"))
                if (r.text or "").strip() or any(t.text for t in r.iter(qn("w:t")))]
        if not runs:
            return 0.0
        bold = 0
        for r in runs:
            rpr = r.find(qn("w:rPr"))
            if rpr is not None and (rpr.find(qn("w:b")) is not None or rpr.find(qn("w:bCs")) is not None):
                bold += 1
        return bold / len(runs)

    # -- field analysis -------------------------------------------------------

    def _analyse_fields(self):
        """Walk the body once, tracking complex and simple fields.
        Returns dict with citation_texts, toc_paras, page_paras, other_fields."""
        citation_texts, other_fields = [], []
        toc_paras, page_paras = set(), set()
        stack = []
        current_p = None
        for el in self.body.iter():
            tag = el.tag
            if tag == qn("w:p"):
                current_p = el
            elif tag == qn("w:fldChar"):
                fct = el.get(qn("w:fldCharType"))
                if fct == "begin":
                    stack.append({"instr": "", "collecting": False, "buf": "",
                                  "paras": set([current_p]) if current_p is not None else set()})
                elif fct == "separate" and stack:
                    stack[-1]["collecting"] = True
                elif fct == "end" and stack:
                    cur = stack.pop()
                    self._classify_field(cur["instr"], cur["buf"], cur["paras"],
                                         citation_texts, toc_paras, page_paras, other_fields)
            elif tag == qn("w:instrText"):
                if stack and not stack[-1]["collecting"]:
                    stack[-1]["instr"] += el.text or ""
            elif tag == qn("w:t"):
                if stack:
                    if stack[-1]["collecting"]:
                        stack[-1]["buf"] += el.text or ""
                    if current_p is not None:
                        stack[-1]["paras"].add(current_p)
            elif tag == qn("w:fldSimple"):
                instr = el.get(qn("w:instr"), "") or ""
                buf = "".join(t.text or "" for t in el.iter(qn("w:t")))
                paras = {current_p} if current_p is not None else set()
                self._classify_field(instr, buf, paras,
                                     citation_texts, toc_paras, page_paras, other_fields)
        return {"citation_texts": citation_texts, "toc_paras": toc_paras,
                "page_paras": page_paras, "other_fields": other_fields}

    @staticmethod
    def _classify_field(instr, buf, paras, citation_texts, toc_paras, page_paras, other_fields):
        up = (instr or "").upper().strip()
        if up.startswith("TOC") or up.startswith("TOA"):
            toc_paras.update(paras)
        elif up.startswith("PAGEREF"):
            toc_paras.update(paras)   # TOC entries & "see page X" cross-refs
        elif up.startswith("PAGE") or up.startswith("NUMPAGES"):
            page_paras.update(paras)
        elif any(code in up for code in CITATION_FIELD_CODES):
            if buf.strip():
                citation_texts.append(buf)
        elif up.startswith(("REF ", "STYLEREF", "SEQ", "HYPERLINK")):
            pass  # keep visible text; not an exclusion category
        else:
            if buf.strip():
                other_fields.append((up[:40], buf[:60]))

    # -- package parts --------------------------------------------------------

    def _read_app_props(self):
        try:
            with zipfile.ZipFile(self.path) as z:
                if "docProps/app.xml" not in z.namelist():
                    return {}
                root = ET.fromstring(z.read("docProps/app.xml"))
            props = {}
            for child in root:
                tag = child.tag.split("}")[-1]
                if tag in ("Pages", "Words", "Characters", "Paragraphs", "Lines") and child.text:
                    try:
                        props[tag] = int(child.text)
                    except ValueError:
                        pass
            return props
        except Exception:
            return {}

    def _read_diagram_texts(self):
        """SmartArt text lives in word/diagrams/data*.xml -- readable after all."""
        out = {}
        try:
            with zipfile.ZipFile(self.path) as z:
                for name in z.namelist():
                    if name.startswith("word/diagrams/") and name.endswith(".xml"):
                        try:
                            root = ET.fromstring(z.read(name))
                        except ET.ParseError:
                            continue
                        ts = [(t.text or "").strip() for t in root.iter(f"{A_NS}t")]
                        ts = [t for t in ts if t]
                        if ts:
                            out[name] = " ".join(ts)
        except Exception:
            pass
        return out

    def _list_chart_parts(self):
        try:
            with zipfile.ZipFile(self.path) as z:
                return [n for n in z.namelist()
                        if n.startswith("word/charts/") and n.endswith(".xml")]
        except Exception:
            return []


# ----------------------------------------------------------------------------
# Candidate finders. Each returns candidates keyed by a stable decision key,
# carrying the block index where found (so exclusions apply only there --
# never to identical text elsewhere in the body).
# ----------------------------------------------------------------------------

class ExclusionTracker:
    """Tracks structural exclusions (contents/abstract/bibliography/appendix)
    and manual [[QCAA_EXCLUDE]] marker regions across the block walk."""

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
                # Only treat an exact heading as a structural section.
                # This prevents legitimate headings such as
                # "References to the Divine in Paradise Lost"
                # from being mistaken for a reference list.
                if lower == trigger:
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


def find_heading_candidates(doc, confirmed_headings):
    """Short, bold-majority or mostly-caps paragraphs that might be unstyled
    section headings. More forgiving than v1: a heading with one non-bold
    run (e.g. the number in '1. Introduction') still qualifies."""
    seen, ordered = set(), []
    for i, (kind, el) in enumerate(doc.blocks):
        if kind != "p" or doc.heading_level(el) is not None:
            continue
        text = strip_markers(doc.visible_text(el)).strip()
        if not text or len(text) > 90 or text in seen:
            continue
        if text.endswith((".", "?", "!")) and len(text) > 40:
            continue  # almost certainly a sentence
        bold_ratio = doc.para_bold_ratio(el)
        letters = [ch for ch in text if ch.isalpha()]
        is_capsy = len(letters) >= 3 and sum(ch.isupper() for ch in letters) / len(letters) > 0.7
        looks_numbered = bool(re.match(r"^\d+(\.\d+)*[.)]?\s+\S", text)) and bold_ratio >= 0.5
        if bold_ratio >= 0.8 or is_capsy or looks_numbered:
            seen.add(text)
            ordered.append((i, text))
    return ordered


def find_citation_candidates(doc, confirmed_headings):
    """Typed-out in-text citations (field citations are handled separately).
    Returns {key: (match, context)}."""
    candidates = {}
    tracker = ExclusionTracker()
    for i, (kind, el) in enumerate(doc.blocks):
        raw = doc.visible_text(el)
        text = strip_markers(raw).replace("\n", " ")
        has_start, has_end = START_MARK in raw, END_MARK in raw
        if kind == "p":
            lvl = effective_heading_level(doc, el, confirmed_headings)
            if lvl is not None:
                tracker.visit_heading(lvl, text)
        if tracker.visit_markers(has_start, has_end) or tracker.in_struct:
            continue
        if el in doc.fields["toc_paras"] or el in doc.fields["page_paras"]:
            continue
        for pattern in (FULL_CITATION_RE, NARRATIVE_CITATION_RE, VANCOUVER_RE, IBID_RE, BARE_YEAR_RE):
            for m in pattern.finditer(text):
                full = m.group(0)
                key = "cite:" + full
                if key not in candidates:
                    start = max(m.start() - 40, 0)
                    end = min(m.end() + 20, len(text))
                    candidates[key] = (full, text[start:end], i)
    return candidates


def find_equation_candidates(doc, confirmed_headings):
    """Inline equations and calculation lines. Math objects (OMML) are
    reported separately -- their text is never in w:t so they are already
    excluded automatically."""
    candidates = {}
    math_paras = []
    tracker = ExclusionTracker()
    for i, (kind, el) in enumerate(doc.blocks):
        if kind != "p":
            continue
        raw = doc.visible_text(el)
        text = strip_markers(raw)
        has_start, has_end = START_MARK in raw, END_MARK in raw
        lvl = effective_heading_level(doc, el, confirmed_headings)
        if lvl is not None:
            tracker.visit_heading(lvl, text)
        if tracker.visit_markers(has_start, has_end) or tracker.in_struct:
            continue
        if el in doc.fields["toc_paras"] or el in doc.fields["page_paras"]:
            continue
        if doc.has_math(el):
            math_paras.append(i)
        for m in EQUATION_RE.finditer(text):
            full = m.group(0).strip()
            if full and len(full) <= 70:
                key = "eq:" + full
                if key not in candidates:
                    start = max(m.start() - 30, 0)
                    end = min(m.end() + 10, len(text))
                    candidates[key] = (full, text[start:end], i)
        stripped = text.strip()
        if (CALC_CONTINUATION_RE.match(stripped) or CALC_LINE_RE.match(stripped)) and stripped:
            key = "eq:" + stripped
            if key not in candidates:
                candidates[key] = (stripped, stripped, i)
    return candidates, math_paras


def find_visual_candidates(doc, confirmed_headings):
    """Captions (styled or text-shaped), by-lines and Source: lines.
    Only paragraphs adjacent to a drawing/table, or styled as captions, so
    ordinary prose starting with 'Figure 3 shows...' is not flagged."""
    candidates = {}
    tracker = ExclusionTracker()
    n = len(doc.blocks)
    for i, (kind, el) in enumerate(doc.blocks):
        if kind != "p":
            continue
        raw = doc.visible_text(el)
        text = strip_markers(raw).strip()
        has_start, has_end = START_MARK in raw, END_MARK in raw
        lvl = effective_heading_level(doc, el, confirmed_headings)
        if lvl is not None:
            tracker.visit_heading(lvl, text)
        if tracker.visit_markers(has_start, has_end) or tracker.in_struct:
            continue
        if el in doc.fields["toc_paras"] or el in doc.fields["page_paras"]:
            continue
        if not text:
            continue
        # adjacency: drawing paragraph before/after, or table before/after
        adj_drawing = False
        for j in (i - 1, i + 1):
            if 0 <= j < n:
                jkind, jel = doc.blocks[j]
                if jkind == "p" and doc.has_drawing(jel):
                    adj_drawing = True
                if jkind == "tbl":
                    adj_drawing = True
        short = len(text) <= 160
        is_caption_style = doc.is_caption_styled(el)
        looks_caption = bool(CAPTION_TEXT_RE.match(text))
        looks_byline = bool(BYLINE_RE.match(text))
        looks_source = bool(SOURCE_LINE_RE.match(text)) and short
        if short:
            if (is_caption_style and (looks_caption or adj_drawing)) or \
               (looks_caption and (adj_drawing or is_caption_style)):
                candidates.setdefault("vis:" + text, (text, "caption", i, text))
            elif looks_caption and adj_drawing:
                candidates.setdefault("vis:" + text, (text, "caption", i, text))
            elif looks_byline:
                match = BYLINE_RE.match(text)
                candidates.setdefault("vis:" + match.group(0), (match.group(0), "by-line", i, text))
            elif looks_source and adj_drawing:
                candidates.setdefault("vis:" + text, (text, "source line", i, text))
    return candidates


def find_figure_candidates(doc, confirmed_headings):
    """Paragraphs containing drawings/picts/objects. Text directly inside the
    paragraph (e.g. a text box) is captured; SmartArt text is read from the
    diagrams part and attached to the same candidate where possible."""
    candidates = []
    tracker = ExclusionTracker()
    for i, (kind, el) in enumerate(doc.blocks):
        if kind != "p":
            continue
        raw = doc.visible_text(el)
        text = strip_markers(raw)
        has_start, has_end = START_MARK in raw, END_MARK in raw
        lvl = effective_heading_level(doc, el, confirmed_headings)
        if lvl is not None:
            tracker.visit_heading(lvl, text)
        if tracker.visit_markers(has_start, has_end) or tracker.in_struct:
            continue
        if el in doc.fields["toc_paras"] or el in doc.fields["page_paras"]:
            continue
        if doc.has_drawing(el):
            inner = text.strip()
            candidates.append({"idx": len(candidates) + 1, "block": i, "inner": inner})
    # SmartArt cannot safely be matched to drawing candidates by
    # list position. Keep it as a separate review candidate.
    for part, txt in doc.diagram_texts.items():
        candidates.append({
            "idx": len(candidates) + 1,
            "block": None,
            "inner": "",
            "smartart": (part, txt),
        })
    return candidates
        


def effective_heading_level(doc, el, confirmed_headings):
    lvl = doc.heading_level(el)
    if lvl is not None:
        return lvl
    text = strip_markers(doc.visible_text(el)).strip()
    if text in confirmed_headings:
        return 1
    return None


# ----------------------------------------------------------------------------
# Heading tree for the per-section breakdown
# ----------------------------------------------------------------------------

class Node:
    __slots__ = ("level", "title", "own_words", "children", "excluded")

    def __init__(self, level, title, excluded=False):
        self.level = level
        self.title = title
        self.own_words = 0
        self.children = []
        self.excluded = excluded


def subtotal(node):
    return node.own_words + sum(subtotal(ch) for ch in node.children)


def print_tree(node, indent=0, out=print):
    for child in node.children:
        total = subtotal(child)
        tag = "  [excluded from count]" if child.excluded else ""
        label = child.title.strip() or "(untitled heading)"
        out(f"{'    ' * indent}- {label}: {total} word{'s' if total != 1 else ''}{tag}")
        print_tree(child, indent + 1, out)


# ----------------------------------------------------------------------------
# Front matter detection
# ----------------------------------------------------------------------------

def is_frontmatter_heading_text(text):
    lower = text.strip().lower().rstrip(":.-")

    for trigger in FRONTMATTER_HEADING_TRIGGERS:
        if lower == trigger:
            return True

    return False


def find_front_matter(doc, confirmed_headings):
    """Everything before the first real content heading (a heading that is
    not itself a front-matter trigger like 'Declaration'). Returns
    (block_indices, lines, words) or None."""
    first_content = None
    for i, (kind, el) in enumerate(doc.blocks):
        if kind != "p":
            continue
        lvl = effective_heading_level(doc, el, confirmed_headings)
        if lvl is None:
            continue
        text = strip_markers(doc.visible_text(el)).strip()
        if is_frontmatter_heading_text(text):
            continue
        first_content = i
        break
    indices, lines = [], []
    scan_end = first_content if first_content is not None else len(doc.blocks)
    for i, (kind, el) in enumerate(doc.blocks):
        if i >= scan_end:
            break
        t = strip_markers(doc.visible_text(el)).strip()
        if t:
            lines.append(t)
            indices.append(i)
    words = sum(word_count(t) for t in lines)
    has_declaration = any(is_frontmatter_heading_text(t) for t in lines)
    if not lines:
        return None
    if first_content is None and not has_declaration:
        return None  # no content heading and no declaration: too risky to offer
    if has_declaration or (words <= 300 and len(lines) <= 15):
        return indices, lines, words
    return None


# ----------------------------------------------------------------------------
# Notes (footnotes / endnotes)
# ----------------------------------------------------------------------------

def get_note_texts(path):
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
                ntype = note.get(f"{W_NS}type")
                if ntype in ("separator", "continuationSeparator", "continuationNotice"):
                    continue
                t = "".join(n.text or "" for n in note.iter(f"{W_NS}t"))
                if t.strip():
                    bucket.append(t)
    return footnotes, endnotes


def looks_bibliographic(note_text):
    t = note_text.lower()
    score = 0
    if re.search(r"(?:19|20)\d{2}", t):
        score += 1
    if re.search(r"\b(pp?\.|vol\.|no\.|ed\.|doi|http|www\.|isbn|issn)\b", t):
        score += 1
    if re.search(r"\b(press|journal|university|phd|thesis|edition|trans\.|vols?)\b", t):
        score += 1
    if re.search(r"[\"'“”‘’]", t) and score >= 1:
        score += 1
    return score >= 2


# ----------------------------------------------------------------------------
# Table smart default
# ----------------------------------------------------------------------------

def table_smart_default(text, profile):
    """Suggest d/c/i for a table based on its content, biased by profile."""
    toks = tokenize(text)
    if not toks:
        return "d"
    numeric = sum(1 for t in toks if is_numeric_symbol_token(t))
    alpha = sum(1 for t in toks if any(ch.isalpha() for ch in t))
    operators = sum(1 for t in toks if t in "+-×*/÷^=" or re.fullmatch(r"[+\-×*/÷^=]+", t))
    has_sentence = bool(re.search(r"[a-z]{3,}\s+[a-z]{3,}\s+[a-z]{3,}", text.lower()))
    ratio_num = numeric / len(toks)
    bias = profile.get("numeric_table_bias")
    if has_sentence or (alpha > 8 and ratio_num < 0.6):
        return "i"
    if operators >= 3 and ratio_num > 0.3:
        return "c" if bias in (None, "c") else bias
    if ratio_num > 0.7:
        return bias if bias in ("d", "c") else "d"
    return "i"


# ----------------------------------------------------------------------------
# File picker (kept from v1)
# ----------------------------------------------------------------------------

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


# ----------------------------------------------------------------------------
# Main per-file analysis
# ----------------------------------------------------------------------------

def analyse_file(path, args, profile, decisions, out=print):
    document = docx.Document(path)
    doc = Doc(path)

    current_sha1 = sha1_file(path)

    if decisions.data.get("file_sha1") and decisions.data["file_sha1"] != current_sha1:
        out(c(
            "Warning: the file has changed since these saved decisions "
            "were made. Old position-based decisions will be cleared.",
            "yellow"
        ))

        decisions.data["decisions"] = {}
        decisions.data["auto"] = []

    decisions.data["file_sha1"] = current_sha1

    interactive = not args.non_interactive

    if profile["guidance"]:
        out(c("\n" + profile["guidance"], "cyan"))

    # ---- 1. unstyled heading confirmation ----------------------------------
    confirmed_headings = set()
    heading_candidates = find_heading_candidates(doc, confirmed_headings)
    if heading_candidates:
        out(f"\nFound {len(heading_candidates)} line(s) that might be unstyled section headings.")
        for i, text in heading_candidates:
            ans = decisions.resolve(
                f"heading:{text}",
                f'\nIs "{text}" a section heading (like "Results" or "Method")?',
                ("y", "n"), "n",
                help_text="Say yes only if this line titles the section that follows it. "
                          "Say no if it is just emphasised text inside a sentence.",
                interactive=interactive,
            )
            if ans == "y":
                confirmed_headings.add(text)

    # ---- 2. front matter ----------------------------------------------------
    frontmatter = find_front_matter(doc, confirmed_headings)
    fm_indices, fm_words = set(), 0
    if frontmatter:
        indices, lines, words = frontmatter
        preview = " | ".join(lines[:4])
        more = f" ... and {len(lines) - 4} more line(s)" if len(lines) > 4 else ""
        out(f"\nPossible front matter (title page / declaration): {len(lines)} line(s), {words} words.")
        out(f"  {preview}{more}")
        ans = decisions.resolve(
            "frontmatter", "Exclude this whole block as title page / declaration / front matter?",
            ("y", "n"), "y" if words <= 300 else "n",
            help_text=FRONTMATTER_EXPLANATION, interactive=interactive,
        )
        if ans == "y":
            fm_indices, fm_words = set(indices), words

    # ---- 3. citations --------------------------------------------------------
    citation_texts = list(doc.fields["citation_texts"])
    if citation_texts:
        out(f"\nAuto-detected {len(citation_texts)} citation field(s) (Word/Zotero/EndNote/Mendeley) -- excluded.")
        for ct in citation_texts:
            decisions.record_auto("cite:" + ct, "field", "citation field")
    manual_cites = find_citation_candidates(doc, confirmed_headings)
    if manual_cites:
        out(f"\nFound {len(manual_cites)} possible typed-out in-text citation(s).")
        for key, (match, context, _bi) in manual_cites.items():
            is_bare_year = bool(BARE_YEAR_RE.fullmatch(match))
            default = "n" if is_bare_year else "y"
            ans = decisions.resolve(
                key, f'\n...{context}...\nIs "{match}" an in-text citation?',
                ("y", "n"), default,
                help_text="Confirm only if this is a source citation. Bare years like "
                          "(2020) default to no because they often appear in ordinary text.",
                interactive=interactive,
            )
            if ans == "y":
                citation_texts.append(match)

    # ---- 4. equations / calculations ----------------------------------------
    equation_items = {}   # key -> (text, block_index)
    psmt_unrecognised_math = []
    eq_candidates, math_paras = find_equation_candidates(doc, confirmed_headings)
    if math_paras:
        out(f"\nAuto-noted {len(math_paras)} math object paragraph(s) (Word equation editor) "
            "-- their content is excluded automatically as equations.")
    if eq_candidates:
        out(f"\nFound {len(eq_candidates)} possible equation/calculation item(s).")
        for key, (match, context, bi) in eq_candidates.items():
            calcish = bool(CALC_CONTINUATION_RE.match(match.strip()) or CALC_LINE_RE.match(match.strip()))
            default = profile["calc_default"] if profile["calc_default"] in ("y", "n") else ("y" if calcish else "n")
            ans = decisions.resolve(
                key, f'\n...{context}...\nExclude "{match.strip()}" as an equation/calculation?',
                ("y", "n"), default,
                help_text="Yes if this is a mathematical expression or a worked-calculation "
                          "line (e.g. '= 4.0 / 40.0'). No if it is ordinary prose that happens "
                          "to contain an equals sign.",
                interactive=interactive,
            )
            if ans == "y":
                equation_items[key] = (match, bi)

    # ---- 5. visual elements --------------------------------------------------
    visual_items = {}     # key -> (substring, block_index)
    vis_candidates = find_visual_candidates(doc, confirmed_headings)
    if vis_candidates:
        out(f"\nFound {len(vis_candidates)} possible visual element(s) (captions/by-lines/source lines).")
        for key, (substr, kind, bi, full) in vis_candidates.items():
            if kind == "caption":
                default = profile["caption_default"] if profile["caption_default"] in ("y", "n") else "n"
            elif kind == "by-line":
                default = profile["byline_default"] if profile["byline_default"] in ("y", "n") else "n"
            else:
                default = "y" if kind == "source line" else "n"
            ans = decisions.resolve(
                key, f'\n({kind}) "{full}"\nExclude this as a visual element?',
                ("y", "n"), default,
                help_text=VISUAL_ELEMENT_EXPLANATION, interactive=interactive,
            )
            if ans == "y":
                visual_items[key] = (substr, bi)

    # ---- 6. figures -----------------------------------------------------------
    figure_data_only = {}  # block_index -> inner text to exclude
    figure_add = []        # (words, label) to ADD (SmartArt prose the tool can now read)
    fig_candidates = find_figure_candidates(doc, confirmed_headings)
    if fig_candidates:
        out(f"\nFound {len(fig_candidates)} figure/diagram candidate(s).")
        for fig in fig_candidates:
            label = f"Figure candidate {fig['idx']}"
            parts = []
            if fig.get("inner"):
                parts.append(f'in-paragraph text: "{fig["inner"][:120]}"')
            if fig.get("smartart"):
                part_name, txt = fig["smartart"]
                parts.append(f"SmartArt text ({part_name}): \"{txt[:120]}\"")
            shown = "\n  ".join(parts) if parts else "(no readable internal text)"
            default = "n"
            ans = decisions.resolve(
                f'figure:{fig["idx"]}',
                f"\n{label} has:\n  {shown}\n"
                "Does this figure contain information other than raw/processed data?",
                ("y", "n"), default,
                help_text=RAW_DATA_EXPLANATION, interactive=interactive,
            )
            if ans == "n":
                if fig.get("inner") and fig["block"] is not None:
                    figure_data_only[fig["block"]] = fig["inner"]
                if fig.get("smartart"):
                    figure_add.append((0, label))  # nothing to add: text not in body gross
            else:
                if fig.get("smartart"):
                    figure_add.append((word_count(fig["smartart"][1]), label + " (SmartArt text)"))

    # ---- 7. footnotes / endnotes ----------------------------------------------
    footnotes, endnotes = get_note_texts(path)
    fn_words = sum(word_count(t) for t in footnotes)
    en_words = sum(word_count(t) for t in endnotes)
    fn_exclude, en_exclude = 0, 0

    def ask_notes(kind, notes, total_words):
        if total_words == 0:
            return 0, []
        bib = [looks_bibliographic(t) for t in notes]
        default = "y" if (bib and all(bib)) else ("m" if any(bib) else "n")
        ans = decisions.resolve(
            f"{kind}:all",
            f"\n{kind.capitalize()} contain {total_words} words total.\n"
            "Are they ALL purely bibliographical?",
            ("y", "n", "m"), default,
            help_text="y = every note is just a citation/reference.\n"
                      "n = none are purely bibliographic (all count).\n"
                      "m = mixed; you will be asked about each one.",
            interactive=interactive,
        )
        excluded = 0
        if ans == "y":
            return total_words, []
        if ans == "m":
            out(f"Confirming {kind} one at a time:")
            kept_numeric = []
            for i, t in enumerate(notes, 1):
                w = word_count(t)
                preview = t[:100] + ("..." if len(t) > 100 else "")
                a = decisions.resolve(
                    f"{kind}:{i}",
                    f'\n{kind.capitalize()} {i} ({w} words): "{preview}"\n'
                    "Is this purely bibliographical (excluded)?",
                    ("y", "n"), "y" if looks_bibliographic(t) else "n",
                    interactive=interactive,
                )
                if a == "y":
                    excluded += w
                else:
                    kept_numeric.append(t)
            return excluded, kept_numeric
        return 0, notes

    fn_excluded, fn_kept = ask_notes("footnotes", footnotes, fn_words)
    en_excluded, en_kept = ask_notes("endnotes", endnotes, en_words)

    # ---- 8. counting walk -------------------------------------------------------
    gross = 0
    buckets = defaultdict(int)
    tracker = ExclusionTracker()
    root = Node(level=-1, title="(document)")
    stack = [root]
    front_node = Node(level=0, title="(front matter / title page)", excluded=True)
    if fm_indices:
        root.children.append(front_node)

    table_answers = {}    # block index -> d/c/i
    table_info = {}       # block index -> (words, preview)

        # pass A: collect tables for decisions
    for i, (kind, el) in enumerate(doc.blocks):
        if kind == "tbl":
            text = strip_markers(doc.visible_text(el))
            table_info[i] = (
                word_count(text),
                text.replace("\n", " ")[:150]
            )

    if table_info:
        out(f"\nFound {len(table_info)} table(s).")
        for i, (words, preview) in table_info.items():
            td = profile["table_default"]
            default = (
                table_smart_default(
                    doc.visible_text(doc.blocks[i][1]),
                    profile
                )
                if td == "smart"
                else td
            )

            ans = decisions.resolve(
                f"table:{i}",
                f'\nTable at block {i} ({words} words): "{preview}..."\n'
                "Is this table:\n"
                "  d = raw or processed data only (excluded)\n"
                "  c = calculation working only (excluded)\n"
                "  i = contains information other than raw/processed data (included)",
                ("d", "c", "i"),
                default,
                help_text=RAW_DATA_EXPLANATION,
                interactive=interactive,
            )
            table_answers[i] = ans
    for i, (kind, el) in enumerate(doc.blocks):
        raw = doc.visible_text(el)
        has_start, has_end = START_MARK in raw, END_MARK in raw
        text = strip_markers(raw)
        words = word_count(text)

        if i in fm_indices:
            gross += words
            buckets["title page"] += words
            front_node.own_words += words
            continue

        if kind == "p":
            lvl = effective_heading_level(doc, el, confirmed_headings)
            if lvl is not None:
                tracker.visit_heading(lvl, text)
                while stack[-1].level >= lvl:
                    stack.pop()
                node = Node(lvl, text, excluded=tracker.in_struct)
                stack[-1].children.append(node)
                stack.append(node)

        manual = tracker.visit_markers(has_start, has_end)
        if manual:
            gross += words
            buckets["manually marked"] += words
            continue
        if tracker.in_struct:
            gross += words
            buckets[tracker.struct_kind or "structural"] += words
            continue
        if kind == "p" and el in doc.fields["toc_paras"]:
            gross += words
            buckets["contents fields (TOC/page refs)"] += words
            continue
        if kind == "p" and el in doc.fields["page_paras"]:
            gross += words
            buckets["page-number fields"] += words
            continue

        gross += words
        node = stack[-1]

        if kind == "tbl":
            ans = table_answers.get(i, "i")

            if ans in ("d", "c"):
                # The entire table is excluded.
                buckets["tables (data/calculation only)"] += words
            else:
                # An included table counts as a whole.
                # Do NOT remove numbers/symbols from it.
                node.own_words += words

            continue

        # paragraph: subtract per-block exclusions, then numbers
        excluded_here = 0
        for key, (substr, bi) in visual_items.items():
            if bi == i:
                n_occ = count_occurrences(text, substr)
                if n_occ:
                    w = word_count(substr) * n_occ
                    excluded_here += w
                    buckets["visual elements"] += w
        if i in figure_data_only:
            inner = figure_data_only[i]
            if inner and inner in text:
                w = word_count(inner)
                excluded_here += w
                buckets["figure inner text (data only)"] += w
        for key, (eq, bi) in equation_items.items():
            if bi == i:
                n_occ = count_occurrences(text, eq)
                if n_occ:
                    w = word_count(eq) * n_occ
                    excluded_here += w
                    buckets["equations/calculations"] += w
        # citations: apply at every occurrence, whole response
        text_for_numeric = text
        for ct in citation_texts:
            n_occ = count_occurrences(text, ct)
            if n_occ:
                w = word_count(ct) * n_occ
                excluded_here += w
                buckets["in-text citations"] += w
                text_for_numeric = text_for_numeric.replace(ct, " ")
        for key, (eq, bi) in equation_items.items():
            if bi == i:
                text_for_numeric = text_for_numeric.replace(eq, " ")
        num = numeric_symbol_count(text)
        buckets["numbers/symbols"] += num

        counted_words = max(words - excluded_here - num, 0)

        if args.profile == "psmt":
            # Flag paragraphs that look mathematical but were not
            # completely identified as equations/calculations.
            math_chars = len(re.findall(
                r"[0-9+\-×*/÷^=<>≤≥≈√²³⁴⁵⁶⁷⁸⁹]",
                text
            ))

            if math_chars >= 3 and counted_words > 0:
                psmt_unrecognised_math.append(
                    (i, text.strip(), counted_words)
                )

        node.own_words += counted_words
    # notes accounting
    if args.profile == "psmt" and psmt_unrecognised_math:
        emit_debug = out
        emit_debug(
            "\nPSMT diagnostic: mathematical-looking paragraphs that still "
            "contributed counted words:"
        )

        for bi, txt, wc in psmt_unrecognised_math[:20]:
            emit_debug(
                f"  Block {bi}: +{wc} counted word(s): {txt[:180]}"
            )

        if len(psmt_unrecognised_math) > 20:
            emit_debug(
                f"  ... and {len(psmt_unrecognised_math) - 20} more."
            )
    gross += fn_words + en_words
    buckets["footnotes (bibliographic)"] += fn_excluded
    buckets["endnotes (bibliographic)"] += en_excluded
    for t in fn_kept:
        buckets["numbers/symbols"] += numeric_symbol_count(t)
    for t in en_kept:
        buckets["numbers/symbols"] += numeric_symbol_count(t)

    # SmartArt prose the user said counts (text not in body gross)
    added_words = 0
    for w_add, label in figure_add:
        if w_add:
            gross += w_add
            root.own_words += w_add
            added_words += w_add

    total_excluded = sum(buckets.values())
    final_count = max(gross - total_excluded, 0)

    # ---- report -------------------------------------------------------------
    lines_out = []
    def emit(s=""):
        out(s)
        lines_out.append(s)

    emit("\n" + "=" * 52)
    emit(c("QCAA WORD COUNT REPORT", "bold"))
    emit(f"File: {os.path.basename(path)}   Profile: {args.profile}   Mode: {'interactive' if interactive else 'non-interactive'}")
    emit("=" * 52)
    emit(f"Gross countable words (incl. footnotes/endnotes): {gross}")
    emit("Excluded:")
    label_map = [
        ("numbers/symbols", "Numbers/symbols"),
        ("in-text citations", "In-text citations"),
        ("equations/calculations", "Equations/calculations"),
        ("visual elements", "Visual elements (captions/by-lines)"),
        ("figure inner text (data only)", "Figure inner text (data-only figures)"),
        ("tables (data/calculation only)", "Tables (data/calculation only)"),
        ("footnotes (bibliographic)", "Bibliographic footnotes"),
        ("endnotes (bibliographic)", "Bibliographic endnotes"),
        ("title page", "Title page / front matter"),
        ("contents", "Contents pages"),
        ("abstract", "Abstract"),
        ("bibliography", "Bibliography / reference list"),
        ("appendix", "Appendixes"),
        ("frontmatter", "Declarations (front matter)"),
        ("contents fields (TOC/page refs)", "Contents fields (TOC/page refs)"),
        ("page-number fields", "Page-number fields"),
        ("manually marked", "Manually marked regions"),
        ("structural", "Other structural sections"),
    ]
    for key, label in label_map:
        if buckets.get(key):
            emit(f"  {label + ':':<42}{buckets[key]}")
    emit("-" * 52)
    emit(c(f"ESTIMATED QCAA WORD COUNT: {final_count}", "green"))
    emit("=" * 52)

    if doc.app_props.get("Words"):
        w = doc.app_props["Words"]
        emit(f"\nReference: Word's own last-saved count = {w} words "
             f"({doc.app_props.get('Pages', '?')} pages).")
        naive = gross - buckets.get("title page", 0) - buckets.get("contents", 0) \
            - buckets.get("bibliography", 0) - buckets.get("appendix", 0) \
            - buckets.get("abstract", 0) - buckets.get("frontmatter", 0)
        if abs(naive - w) > max(0.05 * w, 20):
            emit(c(f"Note: this tool's gross figure ({naive}) differs from Word's ({w}) by "
                   f"more than 5%. Common causes: tracked changes, text boxes Word includes, "
                   f"or content in headers/footers.", "yellow"))
    if doc.chart_parts:
        emit(f"\nNote: {len(doc.chart_parts)} embedded chart part(s) detected "
             f"({', '.join(os.path.basename(p) for p in doc.chart_parts)}). "
             "Chart cached data is not words; judge the chart itself by eye.")

    if root.children:
        emit("\nBREAKDOWN BY HEADING (counted words only; excluded sections marked):")
        print_tree(root, out=emit)
        if root.own_words:
            emit(f"\n(+ {root.own_words} counted word(s) not under any heading)")

    auto_items = [a for a in decisions.data["auto"] if a.get("answer")]
    if auto_items:
        emit("\nAuto/default decisions this run (review with --review-auto if unsure):")
        for a in auto_items[-12:]:
            emit(f"  [{a['answer']}] {a['key'][:70]}  ({a['reason']})")
        if len(auto_items) > 12:
            emit(f"  ... and {len(auto_items) - 12} more")

    emit("\nStill check by eye: text inside images, unusual citation styles, "
         "whether appendix content is genuinely supplementary, and heavily "
         "tracked-changed drafts (accept/reject first).")
    emit(c(f"ESTIMATED QCAA WORD COUNT: {final_count}", "green"))

    return {
        "file": os.path.basename(path),
        "final_count": final_count,
        "gross": gross,
        "excluded": dict(buckets),
        "word_app_words": doc.app_props.get("Words"),
        "word_app_pages": doc.app_props.get("Pages"),
        "breakdown_tree": lines_out,
    }


# ----------------------------------------------------------------------------
# CLI
# ----------------------------------------------------------------------------

def build_parser():
    p = argparse.ArgumentParser(
        prog="qcaa_word_count.py",
        description="Estimate the QCAA word count of a .docx response.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("files", nargs="*", help=".docx file(s); omit to use the file picker")
    p.add_argument("-p", "--profile", default="generic",
                   choices=list(PROFILES.keys()),
                   help="assessment-type profile (rules are identical; only defaults/guidance differ)")
    p.add_argument("--list-profiles", action="store_true", help="show available profiles and exit")
    p.add_argument("-n", "--non-interactive", action="store_true",
                   help="never prompt; use defaults and list them for review")
    p.add_argument(
        "--review-auto",action="store_true",help="reserved for reviewing automatic exclusions (currently informational)")
    p.add_argument("-d", "--decisions", help="decisions file to load/save (default: <file>.qcaa-decisions.json)")
    p.add_argument("--no-save", action="store_true", help="do not write the decisions file")
    p.add_argument("-r", "--report", help="also write the full report to this file")
    p.add_argument("--json", action="store_true", help="print a machine-readable JSON summary")
    p.add_argument("--explain", action="store_true", help="print the QCAA rules and exit")
    p.add_argument("--no-color", action="store_true", help="disable coloured output")
    return p


def main(argv=None):
    global USE_COLOR
    args = build_parser().parse_args(argv)
    if args.no_color:
        USE_COLOR = False

    if args.explain:
        print(QCAA_RULES)
        return 0
    if args.list_profiles:
        print("Available profiles (QCAA rules are the same for all):")
        for name, prof in PROFILES.items():
            print(f"  {name:<20} {prof['description']}")
        return 0

    files = list(args.files)
    if not files:
        picked = get_path_via_dialog()
        if picked:
            files.append(picked)
        else:
            files.append(input("Path to your .docx file: ").strip().strip('"'))

    profile = PROFILES[args.profile]
    results = []
    exit_code = 0

    for path in files:
        if not os.path.isfile(path):
            print(f"\nCouldn't find a file at: {path}")
            print("Check the path is correct and try again.")
            exit_code = 2
            continue
        if not path.lower().endswith(".docx"):
            print(f"\n'{path}' doesn't look like a .docx file.")
            print("If it's a .doc, open it in Word and use File > Save As > Word Document (.docx) first.")
            exit_code = 2
            continue

        dec_path = args.decisions or (path + ".qcaa-decisions.json")
        decisions = Decisions(dec_path)

        try:
            result = analyse_file(path, args, profile, decisions)
        except Exception as e:
            print(f"\nCouldn't analyse '{path}': {e}")
            exit_code = 1
            continue

        if not args.no_save and not args.non_interactive:
            try:
                decisions.data["decisions"] = decisions.data["decisions"]
                with open(dec_path, "w", encoding="utf-8") as f:
                    json.dump(decisions.data, f, indent=2, ensure_ascii=False)
                print(f"\nDecisions saved to {dec_path} (re-run to skip answered questions).")
            except Exception as e:
                print(f"(Couldn't save decisions: {e})")

        if args.report:
            try:
                with open(args.report, "a", encoding="utf-8") as f:
                    f.write("\n".join(result["breakdown_tree"]) + "\n\n")
            except Exception as e:
                print(f"(Couldn't write report: {e})")

        results.append(result)

    if len(results) > 1 or args.json:
        print("\nSUMMARY")
        print("-" * 60)
        print(f"{'File':<40}{'QCAA words':>12}")
        for r in results:
            print(f"{r['file']:<40}{r['final_count']:>12}")
        if args.json:
            print("\nJSON:")
            print(json.dumps(results, indent=2, ensure_ascii=False))

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
'''
with open('/mnt/agents/output/qcaa_word_count.py', 'w', encoding='utf-8') as f:
    f.write(code)
print("written", len(code), "chars")
'''
