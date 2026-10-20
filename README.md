# QCAA Word Count Tool

A Python tool that estimates a **QCAA-compliant word count** for a `.docx`
file. It applies the QCAA rules on what counts and what doesn't (headings,
quotations, tables, footnotes, citations, appendixes, numbers, symbols,
equations), rather than reporting Word's raw count.

Word's own counter treats every visible word the same. QCAA's rules don't —
raw data tables, bibliographies, citations, and standalone numbers are
excluded; headings, quotations, and abbreviations are included. This tool
automates the parts a computer *can* reliably decide, and asks you a quick
question about the parts that genuinely need a human judgement call.

The result is a colour-coded HTML preview you can open in your browser:
green text counted, red text excluded, magenta for citations, grey for
excluded regions, cyan for headings. Click any word to see which rule
was applied to it. Print to PDF for a clean artefact.

---

## Built with AI

This tool was built by **Edward Carroll**, with heavy use of **Claude
(Anthropic)** and **DeepSeek** as coding assistants. The rules, decisions
about what to count, testing, and QCAA interpretation are mine; the
code was written in collaboration with those AI tools. If you'd like to
credit AI assistance in your own repo, the standard is a short note like
this one — there's no formal convention, so any clear statement is fine.

---

## What it does

**Automatic (reliable):**

- Counts body text: paragraphs, headings/subheadings, table text
- Excludes standalone numbers and symbols (`42`, `3.14`, `%`)
- Keeps abbreviations, units and chemical formulas (`kg`, `KOH`, `LPG`, `37kg`)
- Reads footnotes and endnotes directly from the file, classifies them
  as bibliographic or commentary, and prompts for confirmation
- Detects APA 7 in-text citations — both parenthetical `(Smith, 2020)`
  and narrative `Smith (2020)` — and offers to catch ones it missed
- Finds headings matching Contents / Abstract / Bibliography / References /
  Appendix and excludes everything under them
- Handles merged table cells (a common source of over- or under-counting
  in naive scripts)
- Detects captions and figure/table labels, prompts for each one
- Detects and reports text inside floating text boxes, drawings and
  SmartArt (it can't extract the text, but tells you what's there)

**Interactive (your call, prompted by the tool):**

- **Tables** — for each one, choose: count the whole table, exclude the
  whole table (raw/processed data), or count just the header row
- **Captions and figure labels** — exclude (QCAA default) or count
  (if they carry information)
- **Footnotes and endnotes** — exclude (purely bibliographic) or count
  (contains commentary)
- **Appendixes** — exclude if supplementary-only, or count
- **Citations the pattern matcher didn't catch** — confirm or reject each

**Reports:**

- Final word count, with a per-rule breakdown of what was excluded and why
- Per-section word counts (which sections are eating the most words)
- Vocabulary richness, longest paragraph
- Full decision log exported as JSON or CSV for auditing

---

## QCAA rules it follows

<img width="1226" height="885" alt="QCAA rules" src="https://github.com/user-attachments/assets/163b1e06-057b-4608-8e43-4e066ed6ed44" />

Source: [QCAA — School assessment policies (8.2)](https://www.qcaa.qld.edu.au/senior/certificates-and-qualifications/qce-qcia-handbook/8-school-assessment-policies/8.2-integrating-learning-assessment)

The tool applies the same rules regardless of assessment type (student
experiment, PSMT, extended response, essay). It doesn't need to be told
what kind of document it's looking at.

---

## Installation

Requires **Python 3.10 or newer**.

### Quick install (Windows)

1. Install Python from [python.org](https://www.python.org/downloads/) —
   **tick "Add Python to PATH"** on the first screen.
2. Download this repo (green **Code** button → **Download ZIP**) and
   extract it somewhere sensible.
3. Double-click `install-windows.bat` in the extracted folder. That's it
   — it sets up the virtual environment, installs dependencies, and runs
   the tool.

### Quick install (macOS / Linux)

```bash
git clone https://github.com/Prototype0S/qcaa-word-count.git
cd qcaa-word-count
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python qcaa_word_count.py
```

### Manual install (any platform)

```bash
pip install python-docx
```

That's the only hard dependency. Or:

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
python qcaa_word_count.py
```

With no argument, the tool opens a file picker (on Windows, macOS and most
Linux desktops with Tkinter installed) or asks you to paste a path.

You can also pass a file directly:

```bash
python qcaa_word_count.py "path/to/your/file.docx"
```

### The startup menu

Running the tool (with or without a file) shows a menu:

```
[1] Include everything  (fastest)
[2] Review each decision  (most accurate)
[3] What are the QCAA rules?
[4] Exit
```

- **Include everything** — auto-accepts every suggestion. Numbers,
  symbols, equations and citations are still excluded (they always are);
  tables, captions, footnotes and appendixes are all counted. Fast,
  gives you an upper-bound estimate.
- **Review each decision** — prompts you about every ambiguous case.
  Slower, most accurate.
- **What are the QCAA rules?** — prints the rules and returns to the menu.

Both modes generate the HTML preview and open it in your browser.

### Command-line flags

```
--explain              Print the QCAA rules and exit
--include-all          Skip the menu (mode 1)
--review               Skip the menu (mode 2)
--open                 Open the HTML preview (default)
--no-open              Don't open the HTML preview
--html-path PATH       Where to write the HTML preview
--report {json,csv}    Export a full decision log
--report-path PATH     Where to write the decision log
--help                 Show all options
```

### The HTML preview

The preview is the main output. It shows:

- The final word count, with a per-rule exclusion breakdown
- A statistics panel — words per section, vocabulary richness,
  longest paragraph
- The full document, colour-coded: green = counted, red = excluded,
  magenta = citations, grey = excluded regions, cyan = headings
- **Click any coloured word** to see the rule that was applied
- **Dark mode toggle** (remembers your choice)
- **Copy summary** button
- **Print / Save PDF** button

---

## Example results

<img width="1207" height="640" alt="Example word count report" src="https://github.com/user-attachments/assets/744134f9-e625-4f7f-a836-7414ad16b647" />

---

## What still needs your own judgement

- **Citation styles other than APA 7.** The pattern matcher handles
  `(Author, Year)` and `Author (Year)` shapes. Numbered styles like
  `[1]` and other formats (MLA, Chicago, Harvard) are not currently
  detected, and I don't have plans to add them — use `--report json`
  to audit what was missed.
- **Text inside floating text boxes, drawings and SmartArt.** The tool
  detects these and reports how many exist, but can't extract their
  text. If your document has prose in text boxes, it won't be counted.
- **Struct-based exclusion requires real Heading styles.** Contents /
  Abstract / References / Appendix detection fires only when those
  headings use Word's `Heading 1/2/3...` paragraph styles. If a document
  was drafted in Google Docs and pasted into Word without reapplying
  styles, none of the structural detection will trigger. **Apply real
  heading styles first.**
- **Tables need a human call.** The classifier gets most tables right,
  but "information" vs "raw data" is genuinely subjective. Review the
  tool's suggestions.

---

## Disclaimer

This is an estimation tool, not an official QCAA word-count certification.
Some exclusion/inclusion calls in the QCAA rules are subjective — whether
a table counts as "raw/processed data" or "other information", whether a
caption is a visual element or an identifying label, and so on. Always
sanity-check the final number, especially around tables, footnotes, and
citation styles the pattern matcher might not recognise.

Use of AI in the construction of this tool is disclosed at the top of
this README. If you're using the tool for assessment, check with your
teacher whether AI-assisted tools need to be disclosed in your own
acknowledgement.

---

## Contributing

Issues and pull requests welcome. Areas that would benefit most:

- Detecting additional citation styles
- Text box / nested table extraction
- Tests for additional document structures

For questions or comments, please use the **Discussions** section — I'll
do my best to answer them.
