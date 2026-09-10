# QCAA Word Count Tool
**Note that this was made with Claude, I don't know how to state that because I'm new with using GitHub.**


A script that estimates a QCAA-compliant word count for a `.docx` file, following
the QCAA rules on what counts and what doesn't (headings, quotations, tables,
footnotes, citations, appendixes, etc.) rather than just reporting Word's raw
word count.

Word's own count treats every visible word the same way. QCAA's rules don't —
raw data tables, bibliographies, citations, and standalone numbers are meant
to be excluded, while headings, quotations, and abbreviations are meant to be
included. This script automates the parts of that distinction that a computer
*can* reliably make, and asks you a quick question for the parts that
genuinely need a human judgement call.

## What it automates

- Counts all body text: paragraphs, headings/subheadings, table text and
  quotations (these are just normal text, nothing special needed)
- Strips out standalone numbers/symbols (`42`, `3.14`, `%`) while keeping
  abbreviations, units and chemical formulas (`kg`, `KOH`, `LPG`, `37kg`)
- Reads footnotes/endnotes directly from the file and asks once whether
  they're bibliography-only
- Detects text inserted via Word's built-in Citation tool (or
  EndNote/Zotero), **and** typed-out citations like `(Smith, 2020)` or
  `Smith (2020)` — asking you to confirm each unique one it finds
- Finds headings that look like Contents/Abstract/Bibliography/
  References/Appendix and excludes everything under them
- Asks you, table by table, whether it's raw/processed data only (type `?`
  at that prompt for a fuller explanation of the distinction)
- Breaks the final count down heading by heading, so you can see which
  sections are eating the most words
- Correctly handles merged table cells (a common source of accidental
  over- or under-counting in naive word-count scripts)

## What still needs your own judgement

- Citations in a format the pattern-matcher doesn't recognise (it looks for
  `(Author, Year)` / `Author (Year)` shapes — numbered styles like `[1]`
  aren't currently covered)
- Equations/formulas mixed into running text
- Text inside floating text boxes, or a table nested inside another table
  (both are edge cases python-docx doesn't expose directly)
- Anything else — wrap it in plain-text markers anywhere in the document
  and the script will exclude it:

  ```
  [[QCAA_EXCLUDE_START]]  ... content to ignore ...  [[QCAA_EXCLUDE_END]]
  ```

  (e.g. wrap a title page or a blank-page placeholder in these)

Struct-based exclusion (Contents/References/etc.) relies on the section
titles actually using Word's **Heading 1/2/3...** paragraph styles. If a
document was drafted in Google Docs and pasted into Word without reapplying
heading styles, none of that structural detection will fire — apply real
heading styles first.

## Installation

Requires Python 3.8+.

```bash
pip install -r requirements.txt
```

(or just `pip install python-docx`)

## Usage

```bash
python qcaa_word_count.py "path/to/your/file.docx"
```

Or run it with no argument and it will either open a file picker (if a
graphical environment is available) or ask you to paste a path.

The script will then interactively ask you about:
- any typed-out citations it finds (confirm/reject each unique one)
- any tables (whether each one is raw/processed data only)
- footnotes/endnotes (whether they're all bibliography-only)

...and finish with a full report: the estimated word count, a breakdown of
what was excluded and why, and a per-heading breakdown of the final count.

## Disclaimer

This is an estimation tool, not an official QCAA word-count certification.
Some exclusion/inclusion calls in the QCAA rules are genuinely subjective
(e.g. whether a table counts as "raw/processed data" or "other information").
Always sanity-check the final number, especially around tables, footnotes,
and any citation style the pattern-matcher might not recognise.

## Contributing

Issues and pull requests welcome — in particular around detecting
additional citation styles, text box / nested table support, and page-count
exclusions.
