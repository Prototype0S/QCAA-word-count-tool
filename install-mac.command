#!/bin/bash
cd "$(dirname "$0")" || exit 1

if [ ! -x "$0" ]; then
    echo "This launcher is not executable. Run:"
    echo "  chmod +x \"$0\""
    read -n 1 -s -r -p "Press any key to close..."
    exit 1
fi

if xattr -l "$0" 2>/dev/null | grep -q "com.apple.quarantine"; then
    echo
    echo "macOS has quarantined this file. To allow it:"
    echo "  1. Double-click, click 'Done' (NOT 'Move to Trash')"
    echo "  2. System Settings → Privacy & Security → Open Anyway"
    echo
    echo "Or in Terminal:  xattr -cr \"$0\""
    read -n 1 -s -r -p "Press any key to close..."
    exit 1
fi

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
    echo "Python 3.10+ was not found."
    echo "Install from: https://www.python.org/downloads/"
    echo "Or: brew install python@3.14"
    read -n 1 -s -r -p "Press any key to close..."
    exit 1
fi

"$PY" setup.py "$@"
STATUS=$?

if [ -t 0 ]; then
    read -n 1 -s -r -p "Press any key to close..."
    echo
fi
exit "$STATUS"