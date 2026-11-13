"""One-time setup, then launch the tool."""
from __future__ import annotations

import json
import shutil
import subprocess
import sys
import urllib.request
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOOL = HERE / "qcaa_word_count.py"
PKG = HERE / "qcaa_word_count"
OLD_VENV = HERE / ".venv"

PYTHON_DOWNLOADS_URL = "https://www.python.org/downloads/"
EOL_API_URL = "https://endoflife.date/api/python.json"


def fail(msg: str) -> None:
    print()
    print("ERROR:", msg)
    print()
    if sys.stdin.isatty():
        try:
            input("Press Enter to close...")
        except (EOFError, KeyboardInterrupt):
            pass
    sys.exit(1)


def has_docx(python: str) -> bool:
    return subprocess.call(
        [python, "-c", "import docx, docx.opc"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    ) == 0


def install_docx(python: str) -> bool:
    print("First run: installing python-docx (this only happens once)...")
    subprocess.call(
        [python, "-m", "pip", "install", "--quiet", "--upgrade", "pip"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    for extra in (
        ["--user"],
        [],
        ["--user", "--no-cache-dir", "--index-url", "https://pypi.org/simple/"],
    ):
        subprocess.call(
            [python, "-m", "pip", "install", "--quiet", *extra, "python-docx"]
        )
        if has_docx(python):
            return True
    return False


# ---------------------------------------------------------------------------
# Python version check
# ---------------------------------------------------------------------------

def _parse_version(v: str) -> tuple[int, ...]:
    """Turn '3.15.2' into (3, 15, 2). Tolerates suffixes like '3.15.2rc1'."""
    parts = []
    for chunk in v.split(".")[:3]:
        digits = "".join(c for c in chunk if c.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts)


def _latest_python_version() -> str | None:
    """Ask endoflife.date for the newest stable Python release.

    Returns None on any failure (offline, DNS, timeout, malformed JSON).
    Uses only the stdlib — no extra dependency.
    """
    try:
        req = urllib.request.Request(
            EOL_API_URL,
            headers={"User-Agent": "qcaa-word-count-setup"},
        )
        with urllib.request.urlopen(req, timeout=3) as r:
            data = json.load(r)
        # The API returns a list, newest first. Each entry has "latest".
        for entry in data:
            cycle = entry.get("cycle", "")
            # Skip pre-releases / non-standard cycles (e.g. "3.16" is fine,
            # "3.15-dev" is not).
            if not cycle or not all(p.isdigit() for p in cycle.split(".")):
                continue
            latest = entry.get("latest")
            if latest:
                return latest
    except Exception:
        pass
    return None


def check_python_version() -> None:
    """Report the current Python version and nudge if a newer one exists.

    Never fails the install. Never blocks the tool. Best-effort only.
    """
    current = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    print(f"Python {current}")

    latest = _latest_python_version()
    if latest is None:
        # Offline or API unreachable — silent. Don't nag the user.
        return

    try:
        if _parse_version(current) >= _parse_version(latest):
            return
    except (TypeError, ValueError):
        return

    print()
    print(f"  A newer Python is available: {latest}")
    print(f"  You have {current}. Updating is optional but recommended.")
    print(f"  Download: {PYTHON_DOWNLOADS_URL}")
    print()
    print(f"  On Windows: run the installer from {PYTHON_DOWNLOADS_URL}")
    print(f"              and tick \"Add Python to PATH\".")
    print(f"  On Mac:     run the universal2 installer, then re-open this launcher.")
    print()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(argv: list[str]) -> int:
    if not TOOL.is_file():
        fail("qcaa_word_count.py is missing. Keep the whole project together.")
    if not (PKG / "__init__.py").is_file():
        fail("The qcaa_word_count/ package is missing or incomplete.")

    if OLD_VENV.exists():
        shutil.rmtree(OLD_VENV, ignore_errors=True)

    python = sys.executable
    if not has_docx(python):
        if not install_docx(python):
            fail(
                "Could not install python-docx.\n\n"
                "  Check your internet connection.\n"
                "  If the error mentions 'lxml' or 'Microsoft Visual C++',\n"
                "  Python 3.15 may be too new — try Python 3.14:\n"
                f"      {PYTHON_DOWNLOADS_URL}\n"
                "  School networks may block pypi.org."
            )

    # Report version + nudge if newer exists. Runs once per setup.
    check_python_version()

    print()
    return subprocess.call([python, str(TOOL), *argv])


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        sys.exit(130)