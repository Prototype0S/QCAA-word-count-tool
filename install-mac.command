#!/bin/bash
# ============================================================
#  QCAA Word Count Tool — macOS launcher
#
#  Double-click in Finder to set up and run.
#  First run creates .venv and installs python-docx; later runs
#  go straight to the tool.
# ============================================================

cd "$(dirname "$0")" || {
    echo "ERROR: could not change to the script's folder."
    exit 1
}

pause() {
    if [ -t 0 ]; then
        if [ -n "${BASH_VERSION:-}" ]; then
            read -n 1 -s -r -p "Press any key to close..." _ || true
        else
            read -r -p "Press Enter to close..." _ || true
        fi
        echo
    fi
}

fail() {
    echo
    echo "ERROR: $*"
    echo
    pause
    exit 1
}

# --- macOS Gatekeeper quarantine check ------------------------------
if xattr -l "$0" 2>/dev/null | grep -q "com.apple.quarantine"; then
    echo
    echo "============================================================"
    echo "  macOS SECURITY NOTICE"
    echo "============================================================"
    echo
    echo "This launcher was downloaded from the internet and macOS"
    echo "has quarantined it. The old 'right-click → Open' method"
    echo "NO LONGER WORKS on macOS 15 (Sequoia) and later."
    echo
    echo "To run it:"
    echo "  1. Double-click this file and click 'Done' in the warning"
    echo "     (do NOT click 'Move to Trash')."
    echo "  2. Open System Settings → Privacy & Security."
    echo "  3. Scroll to the Security section."
    echo "  4. Click 'Open Anyway' next to this file."
    echo "  5. Enter your Mac password when asked."
    echo
    echo "Alternatively, open Terminal and run:"
    echo "  xattr -cr \"$0\""
    echo "Then double-click this file again."
    echo
    pause
    exit 1
fi

# --- Executable bit check -------------------------------------------
if [ ! -x "$0" ]; then
    echo
    echo "============================================================"
    echo "  PERMISSION FIX NEEDED"
    echo "============================================================"
    echo
    echo "This launcher is not marked as executable, so macOS"
    echo "won't run it. Fix it with one command in Terminal:"
    echo
    echo "  chmod +x \"$0\""
    echo
    echo "Then double-click this file again."
    echo
    pause
    exit 1
fi

echo
echo "============================================================"
echo "  QCAA WORD COUNT TOOL"
echo "============================================================"
echo

[ -f "qcaa_word_count.py" ] || fail "qcaa_word_count.py is missing. Keep this launcher in the same folder as the tool."
[ -d "qcaa_word_count" ] || fail "The qcaa_word_count/ package folder is missing. Keep the whole folder together."

case "$(pwd)" in
    */Downloads|*/Downloads/*|/tmp/*|*/Temp/*)
        echo "NOTE: This folder is in Downloads or a temporary location."
        echo "      Move it to Documents so the virtual environment is kept."
        echo
        ;;
esac

# --- Find Python 3.10+ ----------------------------------------------
PY=""
for candidate in \
    python3.15 python3.14 python3.13 python3.12 python3.11 python3.10 \
    python3 python \
    /usr/local/bin/python3 \
    /opt/homebrew/bin/python3 \
    /Library/Frameworks/Python.framework/Versions/Current/bin/python3
do
    if command -v "$candidate" >/dev/null 2>&1; then
        if "$candidate" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)' >/dev/null 2>&1; then
            PY="$(command -v "$candidate")"
            break
        fi
    fi
done

if [ -z "$PY" ]; then
    echo "Python 3.10 or newer was not found."
    echo
    echo "  Install Python 3.15 (current stable) from:"
    echo "      https://www.python.org/downloads/"
    echo
    pause
    exit 1
fi

PYVER="$("$PY" -c 'import sys; print("%d.%d.%d" % sys.version_info[:3])')"
echo "Found Python $PYVER ($PY)"

# --- Create / repair the virtual environment ------------------------
VENV_PY=".venv/bin/python"
venv_ok() {
    [ -x "$VENV_PY" ] && "$VENV_PY" -c 'import sys' >/dev/null 2>&1
}

if ! venv_ok; then
    echo "Setting up virtual environment..."
    rm -rf .venv
    if ! "$PY" -m venv .venv; then
        fail "Could not create the virtual environment.
  Homebrew users: run 'brew install python' (includes venv/pip).
  Otherwise reinstall Python from https://www.python.org/downloads/"
    fi
    venv_ok || fail "The virtual environment was created but is not usable.
  Delete the .venv folder and run this launcher again."
fi

# --- Install dependencies -------------------------------------------
if ! "$VENV_PY" -c 'import docx, docx.opc' >/dev/null 2>&1; then
    echo "Installing dependencies (python-docx)..."
    "$VENV_PY" -m pip install --quiet --upgrade pip >/dev/null 2>&1 || true
    if ! "$VENV_PY" -m pip install --quiet python-docx; then
        fail "Could not install python-docx.
  Check your internet connection. On a school network, PyPI may be
  blocked; try a phone hotspot or ask IT for the proxy address."
    fi
fi

# --- Smoke test -----------------------------------------------------
# Quiet on success, loud on failure.
if [ -f "selftest.py" ]; then
    if ! "$VENV_PY" selftest.py >/dev/null 2>&1; then
        echo
        echo "WARNING: The tool's own sanity check failed."
        echo "         The installer worked but the tool may not run."
        echo "         Details:"
        echo
        "$VENV_PY" selftest.py
        echo
        pause
        exit 1
    fi
fi

# --- Fallback file picker if tkinter is missing ---------------------
if [ $# -eq 0 ] && ! "$VENV_PY" -c 'import tkinter' >/dev/null 2>&1; then
    echo "This Python has no tkinter, so the built-in picker is unavailable."
    echo "Choose your .docx in the dialog instead:"
    picked="$(osascript -e 'POSIX path of (choose file with prompt "Select your Word document")' 2>/dev/null)" || picked=""
    if [ -n "$picked" ]; then
        set -- "$picked"
    else
        echo "No file selected — the tool will open its menu without a file."
    fi
    echo
fi

# --- Run ------------------------------------------------------------
echo "Starting..."
echo
"$VENV_PY" qcaa_word_count.py "$@"
STATUS=$?

if [ "$STATUS" -ne 0 ]; then
    echo
    echo "The tool exited with an error (code $STATUS). See the message above."
fi

echo
pause
exit "$STATUS"