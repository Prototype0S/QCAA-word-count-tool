#!/bin/bash
# ============================================================
#  QCAA Word Count Tool — macOS launcher
# ============================================================

cd "$(dirname "$0")" || exit 1

pause() {
    if [ -t 0 ]; then
        read -n 1 -s -r -p "Press any key to close..." _ || true
        echo
    fi
}

# --- Permission check ---------------------------------------
if [ ! -x "$0" ]; then
    echo
    echo "This launcher is not executable. Run this in Terminal:"
    echo "  chmod +x \"$0\""
    echo
    pause
    exit 1
fi

# --- macOS Gatekeeper quarantine check ----------------------
if xattr -l "$0" 2>/dev/null | grep -q "com.apple.quarantine"; then
    echo
    echo "============================================================"
    echo "  macOS SECURITY NOTICE"
    echo "============================================================"
    echo
    echo "This launcher was downloaded and macOS has quarantined it."
    echo
    echo "To allow it:"
    echo "  1. Double-click, click 'Done' (NOT 'Move to Trash')."
    echo "  2. Open System Settings -> Privacy & Security."
    echo "  3. Click 'Open Anyway' next to this file."
    echo "  4. Enter your Mac password when asked."
    echo
    echo "Or in Terminal:  xattr -cr \"$0\""
    echo
    pause
    exit 1
fi

# --- Banner -------------------------------------------------
echo
echo "============================================================"
echo "  QCAA WORD COUNT TOOL"
echo "============================================================"
echo

# --- Find Python 3.10+ --------------------------------------
PY=""
for c in python3.15 python3.14 python3.13 python3.12 python3.11 python3.10 python3 python; do
    if command -v "$c" >/dev/null 2>&1; then
        if "$c" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3,10) else 1)' 2>/dev/null; then
            PY="$c"
            break
        fi
    fi
done

if [ -z "$PY" ]; then
    echo "ERROR: Python 3.10 or newer was not found."
    echo
    echo "Install Python 3.15 (current stable) from:"
    echo "  https://www.python.org/downloads/"
    echo
    echo "Or with Homebrew:  brew install python@3.14"
    echo
    pause
    exit 1
fi

# --- Run setup + tool ---------------------------------------
"$PY" setup.py "$@"
STATUS=$?

if [ "$STATUS" -ne 0 ]; then
    echo
    echo "The tool exited with an error (code $STATUS)."
fi

pause
exit "$STATUS"