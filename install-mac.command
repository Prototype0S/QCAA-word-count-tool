#!/usr/bin/env bash
# QCAA Word Count Tool — macOS launcher
# Double-click this file in Finder to set up and run the tool.
# The first run installs Python dependencies; subsequent runs just launch.

set -e
cd "$(dirname "$0")"

# ---------------------------------------------------------------------------
# Pretty banner
# ---------------------------------------------------------------------------
echo
echo "============================================================"
echo "  QCAA WORD COUNT TOOL"
echo "============================================================"
echo

# ---------------------------------------------------------------------------
# Warn if running from Downloads or a temporary folder
# ---------------------------------------------------------------------------
case "$(pwd)" in
    */Downloads/*|*/Downloads|/tmp/*|*/Temp/*)
        echo "NOTE: You're running this from Downloads or a temporary folder."
        echo "      Move the whole folder to Documents first (e.g."
        echo "      ~/Documents/qcaa-word-count) and run again."
        echo
        ;;
esac

# ---------------------------------------------------------------------------
# Find a suitable Python (3.10+, prefer 3.15)
# ---------------------------------------------------------------------------
# Try the newest first so anyone with 3.15 installed gets it automatically.
PYTHON_BIN=""
for candidate in python3.15 python3.14 python3.13 python3.12 python3.11 python3.10 python3 python; do
    if command -v "$candidate" >/dev/null 2>&1; then
        if "$candidate" -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)" >/dev/null 2>&1; then
            PYTHON_BIN="$candidate"
            break
        fi
    fi
done

if [ -z "$PYTHON_BIN" ]; then
    echo "ERROR: Python 3.10 or newer was not found."
    echo
    echo "  RECOMMENDED: Python 3.15 (much faster for this tool)"
    echo "  Download:    https://www.python.org/downloads/release/python-3150/"
    echo "               (pick the macOS installer, 10.15+)"
    echo
    echo "  1. Run the installer."
    echo "  2. Double-click this file again."
    echo
    echo "Alternative (if you use Homebrew): brew install python@3.15"
    echo
    read -n 1 -s -r -p "Press any key to close..."
    echo
    exit 1
fi

PYVER=$("$PYTHON_BIN" -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')")
echo "Found Python $PYVER ($PYTHON_BIN)"

# ---------------------------------------------------------------------------
# Soft-recommend 3.15 (works on 3.10+, faster on 3.15)
# ---------------------------------------------------------------------------
if ! "$PYTHON_BIN" -c "import sys; sys.exit(0 if sys.version_info >= (3,15) else 1)" >/dev/null 2>&1; then
    echo
    echo "NOTE: You're on Python $PYVER. The tool works, but 3.15 is"
    echo "      noticeably faster. If you want the speedup, upgrade from:"
    echo "      https://www.python.org/downloads/release/python-3150/"
    echo
fi

# ---------------------------------------------------------------------------
# Create the virtual environment if needed
# ---------------------------------------------------------------------------
if [ ! -x ".venv/bin/python" ]; then
    echo "Setting up virtual environment..."
    "$PYTHON_BIN" -m venv .venv
    if [ $? -ne 0 ]; then
        echo "ERROR: Could not create the virtual environment."
        read -n 1 -s -r -p "Press any key to close..."
        echo
        exit 1
    fi
fi

# ---------------------------------------------------------------------------
# Install dependencies if needed
# ---------------------------------------------------------------------------
if ! .venv/bin/python -c "import docx" >/dev/null 2>&1; then
    echo "Installing dependencies..."
    .venv/bin/python -m pip install --quiet --upgrade pip
    if ! .venv/bin/python -m pip install --quiet python-docx; then
        echo
        echo "ERROR: Could not install python-docx."
        echo
        echo "If you're on a school or work network, PyPI might be blocked."
        echo "Try from a phone hotspot, or ask IT for the proxy address."
        echo
        read -n 1 -s -r -p "Press any key to close..."
        echo
        exit 1
    fi
fi

# ---------------------------------------------------------------------------
# Run the tool
# ---------------------------------------------------------------------------
echo "Starting..."
echo
.venv/bin/python qcaa_word_count.py
STATUS=$?

if [ $STATUS -ne 0 ]; then
    echo
    echo "The tool exited with an error. See the message above."
fi

echo
read -n 1 -s -r -p "Press any key to close..."
echo