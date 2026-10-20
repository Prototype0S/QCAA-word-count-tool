# Advanced Usage

You've installed the tool, you've run it once, and it works. This guide
is for everything beyond the double-click workflow: running the tool
from a terminal, passing flags, and understanding what each of the extra
modes is actually good for.

You don't need any of this to use the tool. The `.bat` (Windows) or
`.command` (macOS) launch does everything the average user needs. This
document is for the cases where the default workflow isn't quite what
you want.

---

## Why use a terminal at all?

The double-click launcher does exactly one thing: it opens a file picker,
runs the tool interactively, and shows the report. That's the right
behaviour 90% of the time.

The other 10% is where the terminal comes in:

- **You're iterating on a draft.** You edit one sentence, re-run the tool,
  and want to know whether the count went up or down. Typing a command
  is faster than clicking through a file picker.
- **You want just the number, not the full report.** `--summary` prints
  one line. That's it.
- **You want to inspect *what* was excluded, not just how many.** The
  `--report json` flag writes a full audit log you can open in any text
  editor.
- **You want to see the HTML preview without committing to decisions.**
  `--preview-only` shows the colour-coded document with the *default*
  decisions applied, and doesn't ask you anything.
- **You want to compare against a previous run.** `--check` shows
  what changed since last time.
- **You're processing several documents.** Script the loop.

None of these are required. But if you find yourself re-running the tool
often (editing a draft, tweaking tables, checking word counts), the
terminal workflow saves a lot of clicking.

---

## Opening a terminal with the tool ready

You need three things in your terminal for the commands in this guide
to work:

1. You're in the folder that contains `qcaa_word_count.py`
2. The virtual environment is activated (the prompt starts with `(.venv)`)
3. `python-docx` is installed

There are two ways to get there.

### Windows — the easy way

There should be an `open-terminal.bat` file in the same folder as
`qcaa_word_count.py`. **Double-click it.**

A Command Prompt window opens with `(.venv)` at the start of the prompt,
already inside the tool's folder. Every command in this guide works
directly from here. Type `exit` when you're done.

If `open-terminal.bat` isn't there, use the manual method below.

### Windows — the manual way

1. Open **File Explorer** and navigate to the folder containing
   `qcaa_word_count.py`
2. Click the address bar at the top
3. Type `cmd` and press Enter

A Command Prompt opens in that folder. Then:

```
.venv\Scripts\activate.bat
```

You should see `(.venv)` appear at the start of the prompt. Confirm
Python is working:

```
python --version
```

You should see `Python 3.15.x` (or whichever version you installed).
If you get "Python is not recognised", see [Troubleshooting](#troubleshooting)
at the bottom.

### macOS — the easy way

Double-click `open-terminal.command`. A Terminal window opens with
the venv active. Type `exit` when done.

If it's missing, use the manual method.

### macOS — the manual way

Open Terminal, then:

```bash
cd /path/to/your/qcaa-word-count
source .venv/bin/activate
```

Replace `/path/to/your/qcaa-word-count` with wherever you extracted the
tool. You should see `(.venv)` appear at the prompt.

### Linux

Same as macOS — `cd` into the tool's folder, then
`source .venv/bin/activate`.

---

## The basics: running the tool with a file argument

Once the terminal is open and the venv is active, you don't need to use
the file picker. Just pass the `.docx` path:

```
python qcaa_word_count.py "your-file.docx"
```

The path can be absolute or relative to where you are. If the `.docx`
is in the same folder as the tool, `"your-file.docx"` is enough. If it's
elsewhere, use the full path:

```
python qcaa_word_count.py "C:\Users\you\Documents\your-file.docx"
```

...or on macOS/Linux:

```
python qcaa_word_count.py "/Users/you/Documents/your-file.docx"
```

**Always wrap paths in quotes** if they contain spaces or parentheses.
`"10.3 Science(7).docx"` works. `10.3 Science(7).docx` doesn't — the
shell splits it into three arguments.

**Drag-and-drop shortcut (Windows and macOS)**: you can drag a `.docx`
file from File Explorer/Finder directly into the terminal window. The
shell pastes the full quoted path for you. This is faster than typing.

Type the start of the command:

```
python qcaa_word_count.py 
```

...then drag the file into the terminal. The shell appends the quoted
path automatically:

```
python qcaa_word_count.py "C:\Users\you\Downloads\your-file.docx"
```

Now you have a template. Add flags after the path.

---

## The modes, and what each is good for

### `--summary` — one line, done

```
python qcaa_word_count.py "your-file.docx" --summary
```

**What it does**: auto-accepts every suggestion (same as picking option
`[1]` from the menu), skips the HTML preview entirely, and prints only
the final count. No prompts, no report, no browser window.

**Output**:
```
your-file.docx    1083
```

**When to use it**:
- You're editing a draft and want a quick check: *"did I go over the limit?"*
- You're scripting. `--summary` is designed to be pipeable.
- You want the count without committing to any decisions — because
  `--summary` uses the default (count-everything) behaviour.

**What you lose**: the HTML preview, the report, and the chance to
override decisions. It's a "just tell me the number" mode.

**Does it save decisions?** Yes, by default. If you don't want it to,
add `--no-save`.

### `--preview-only` — see the colour-coded document without answering prompts

```
python qcaa_word_count.py "your-file.docx" --preview-only
```

**What it does**: runs the whole pipeline, auto-accepts every suggestion
(like `--summary`), generates the HTML preview, and opens it in the
browser. No terminal report, no prompts.

**When to use it**:
- You want to see what the tool *would* do before committing to specific
  decisions.
- You're showing the tool to someone else and don't want to click
  through prompts.
- You want to look at the colour-coded document and decide if any
  decisions are wrong, *then* run `--review` to fix them.

**The workflow**: `--preview-only` → look at HTML → spot a wrong decision →
run `--review` → fix it → run again to confirm.

### `--check` — compare against the last run

```
python qcaa_word_count.py "your-file.docx" --check
```

**What it does**: loads the saved decisions from the previous run,
compares this run's count and per-section breakdown against the previous
run, and prints the difference.

**Output** (when you've edited since the last run):
```
Comparison with last run:
  Last run:  1083 words
  This run:  1102 words
  Change:    +19

  Per section:
    Rationale                                367   +12
    Modifications                            74    =
    Risk Management                         143    =
    Evaluation                              325    +7
    Conclusion                               82    =
```

**When to use it**:
- After editing, to see whether the edit made the count better or worse.
- To track progress: *"I was 40 words over, now I'm 12 over."*
- To confirm your changes actually landed (sometimes you edit the wrong
  paragraph and don't notice).

**It needs a previous run** — the first `--check` on a fresh file prints
nothing. Run once without `--check` to establish a baseline, then use
`--check` after every edit.

**Note**: `--check` still prompts interactively. If you've already
decided everything in a previous run, most prompts will be skipped
automatically (because saved decisions are restored). So `--check` on a
second run is usually fast.

### `--review` — the full interactive experience

```
python qcaa_word_count.py "your-file.docx" --review
```

**What it does**: skips the startup menu and goes straight to interactive
review. Same as picking `[2]` from the menu.

**When to use it**: when you're scripting or writing a `.bat` and want
to force interactive mode without showing the menu.

### `--include-all` — accept everything, but still show the report

```
python qcaa_word_count.py "your-file.docx" --include-all
```

**What it does**: same as `--summary` for decisions (auto-accepts),
but still shows the full report and HTML preview.

**Difference from `--summary`**: `--summary` prints one line and exits.
`--include-all` runs the full output pipeline. Use `--include-all` when
you want the count *and* the colour-coded document but don't want to
answer prompts.

---

## The decision-file flags

Every run saves your decisions to `<filename>_qcaa_decisions.json` next
to your `.docx`. The next run restores them automatically.

### `--no-save` — don't write the file this run

```
python qcaa_word_count.py "your-file.docx" --no-save
```

Useful when you're just exploring (`--summary`, `--preview-only`) and
don't want to leave a decisions file behind.

### `--no-load` — ignore any saved file

```
python qcaa_word_count.py "your-file.docx" --no-load
```

Runs fresh even if a decisions file exists, without deleting it. Useful
when you want to see what the auto-classifier would do, without wiping
your previous choices.

### `--reset` — delete the saved file, then run normally

```
python qcaa_word_count.py "your-file.docx" --reset
```

Deletes the decisions file, then runs as if it were the first run. Use
this when:
- You want to redo every prompt from scratch.
- The saved decisions are giving a weird result and you want to
  start over.
- You've edited the document significantly and the saved decisions are
  mostly stale anyway.

### `--report json` / `--report csv` — export a full audit log

```
python qcaa_word_count.py "your-file.docx" --report json
```

**What it does**: in addition to everything else, writes a `<filename>_qcaa_decisions.json`
file containing a complete record of every span, its decision, the rule
that was applied, and the reason. This is the "here's exactly what I
did and why" audit trail.

**When to use it**:
- You want to justify the count to a teacher or marker.
- You're debugging a suspicious number and need to see per-word decisions.
- You want a portable record that survives even if you delete the tool.

The JSON is large (every word is a separate entry) but it's the
complete picture. Open it in VS Code, Notepad++, or any text editor.

`--report csv` writes the same data as comma-separated values, which
opens cleanly in Excel or Google Sheets. The CSV is per-word rows with
`block_index, span_text, span_type, decision, rule, reason` columns.
Useful for filtering: *"show me every span where rule is `I-EQUATION`"*.

### `--report-path` — write the log somewhere else

```
python qcaa_word_count.py "your-file.docx" --report json --report-path "path/to/audit.json"
```

Useful when you want the audit log somewhere specific, not next to the
`.docx`.

---

## Other flags

### `--html-path` — put the preview somewhere else

```
python qcaa_word_count.py "your-file.docx" --html-path "path/to/preview.html"
```

Default is next to the `.docx`. Useful for keeping previews away from
the source file.

### `--no-open` — write the HTML but don't launch the browser

```
python qcaa_word_count.py "your-file.docx" --no-open
```

Useful when:
- You're running many documents in a loop and don't want 20 browser tabs.
- You're on a remote connection and can't open a browser.
- You want to control when the preview opens.

### `--open` — force browser opening (default)

Explicit force. Useful only if you've set `--no-open` as a habit and want
to override once.

### `--tutorial` — see the first-run tutorial again

```
python qcaa_word_count.py --tutorial
```

The tutorial normally appears only on the first run. This flag replays
it. Useful if you read it once, forgot what it said, and want a reminder.

### `--explain` — print the QCAA rules and exit

```
python qcaa_word_count.py --explain
```

Prints the full word-length rules summary. Doesn't need a file. Use it
when you want to check exactly what counts before running anything.

### `--help` — see every flag

```
python qcaa_word_count.py --help
```

Prints all flags with short descriptions. When in doubt, this is the
authoritative list.

---

## Workflows: three real scenarios

### Scenario A — iterating on a draft

You've written an essay, it's 1,200 words, and the limit is 1,500.
You're going to keep editing it for another two hours.

**First run** (establishes the baseline):
```
python qcaa_word_count.py "essay.docx" --review
```
Answer the prompts once. Note the count.

**Then, after every edit**:
```
python qcaa_word_count.py "essay.docx" --check
```
Most prompts are skipped automatically (decisions are saved). You see
the count and the per-section diff. If the diff shows "Rationale +12",
you know your edit to the rationale added 12 words.

**When you're done**:
```
python qcaa_word_count.py "essay.docx" --summary
```
Final count, one line, done.

### Scenario B — checking what was excluded

The tool says 1,083 but you thought it was closer to 1,180. You want to
know where the difference is.

**Step 1** — get the per-rule breakdown in the HTML preview:
```
python qcaa_word_count.py "essay.docx" --preview-only
```
Open the HTML, scroll to the panel below the word count. It shows
`excluded:I-CITATION: 29 word(s)`, `excluded:I-NUMBER: 32 word(s)`, etc.

**Step 2** — if that's not enough, get the full per-word audit:
```
python qcaa_word_count.py "essay.docx" --report json --no-open
```
Open the JSON in VS Code, search for the rule you're curious about
(e.g. `"rule": "I-CITATION"`). Every occurrence is listed.

**Step 3** — if a specific exclusion is wrong (e.g. a caption was excluded
that you wanted counted), run:
```
python qcaa_word_count.py "essay.docx" --review
```
Change the specific decision and re-check.

### Scenario C — batch processing

You have a folder of ten drafts and you want a table of counts.

**Windows** (Command Prompt):
```
cd "path\to\your\drafts"
for %f in (*.docx) do python "path\to\qcaa-word-count\qcaa_word_count.py" "%f" --summary --no-save
```

**macOS / Linux**:
```
cd /path/to/your/drafts
for f in *.docx; do python /path/to/qcaa-word-count/qcaa_word_count.py "$f" --summary --no-save; done
```

The output is a list:
```
draft1.docx    834
draft2.docx    1201
draft3.docx    967
...
```

`--no-save` stops the tool writing decisions files for every draft (if
you're just counting, you don't need them).

---

## What each mode *doesn't* do

- **`--summary` doesn't save the HTML preview.** If you want the
  colour-coded document, use `--preview-only` or the default mode.
- **`--preview-only` doesn't prompt you.** If you want to override a
  decision, use `--review` after looking at the preview.
- **`--check` needs a previous run.** The first `--check` on a new file
  prints nothing because there's nothing to compare against.
- **`--reset` deletes the saved decisions.** If you have a good set of
  decisions and want to keep them, don't use `--reset`. Use `--no-load`
  to see what a fresh run would produce, or `--no-save` to make changes
  without overwriting.

---

## Troubleshooting

### "Python is not recognised"

You haven't activated the venv, or Python isn't on PATH. From the
tool's folder:

```
.venv\Scripts\activate.bat        (Windows)
source .venv/bin/activate         (macOS/Linux)
```

The prompt should show `(.venv)`. If it doesn't, the activation failed.

### "ModuleNotFoundError: No module named 'docx'"

The venv is active, but `python-docx` isn't installed in it. Run:

```
python -m pip install python-docx
```

### "I can't find the saved decisions file"

It's next to your `.docx`, not next to the tool. So if your essay is at
`Documents/essay.docx`, the decisions file is at
`Documents/essay_qcaa_decisions.json`.

To open it in a text editor:

- **Windows**: right-click the file → Open With → Notepad
- **macOS**: right-click → Open With → TextEdit
- **Or**: `notepad "path/to/your-file_qcaa_decisions.json"` from a
  terminal (Windows) / `open -t "path/to/your-file_qcaa_decisions.json"`
  (macOS)

Don't edit it directly — just look at it.

### "The tool is asking me the same question every time"

The saved decisions file might not be matching the current document
(e.g. you renamed a heading, moved a table, or edited a caption). The
tool restores decisions that still match; anything that changed gets
re-asked.

If this happens a lot, the workflow is:
1. Fix all the decisions once with `--review`
2. Run again — most should be restored
3. Fix the few that are asked again

Or delete the JSON file and start over with `--reset`.

### "I want to start completely fresh"

```
python qcaa_word_count.py "your-file.docx" --reset
```

Deletes the saved decisions and re-runs from scratch.

---

## A quick reference card

| Task | Command |
|---|---|
| Run interactively (from terminal) | `python qcaa_word_count.py "file.docx" --review` |
| Just the count, one line | `python qcaa_word_count.py "file.docx" --summary` |
| Preview without committing | `python qcaa_word_count.py "file.docx" --preview-only` |
| Compare with last run | `python qcaa_word_count.py "file.docx" --check` |
| Full report, no prompts | `python qcaa_word_count.py "file.docx" --include-all` |
| Export full audit log | `python qcaa_word_count.py "file.docx" --report json` |
| Start over, forget decisions | `python qcaa_word_count.py "file.docx" --reset` |
| See all flags | `python qcaa_word_count.py --help` |
| See the QCAA rules | `python qcaa_word_count.py --explain` |
| See the tutorial again | `python qcaa_word_count.py --tutorial` |

---

## One more thing

If you find yourself typing the same command repeatedly, make a shortcut:

1. Write a small `.bat` (Windows) or `.sh` (macOS/Linux) file with the
   command, e.g. `check-essay.bat`:
   ```bat
   @echo off
   cd /d "%~dp0"
   .venv\Scripts\python.exe qcaa_word_count.py "essay.docx" --check
   pause
   ```
2. Save it in the tool's folder.
3. Double-click it whenever you want to run that specific command.

That's the shortcut between "full terminal workflow" and "double-click
launcher". Neither is wrong; pick whichever fits how you work.