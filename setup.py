"""One-time setup, then launch the tool.

Called by install-windows.bat / install-mac.command.
Or directly: python setup.py [tool args...]
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
TOOL = HERE / "qcaa_word_count.py"
PKG = HERE / "qcaa_word_count"
OLD_VENV = HERE / ".venv"


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
    """True if python can import the real python-docx.

    `docx.opc` only exists in the real package, not the abandoned `docx`.
    Output is suppressed so a failed check doesn't spam tracebacks.
    """
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


def main(argv: list[str]) -> int:
    if not TOOL.is_file():
        fail("qcaa_word_count.py is missing. Keep the whole project together.")
    if not (PKG / "__init__.py").is_file():
        fail("The qcaa_word_count/ package is missing or incomplete.")

    # Silent cleanup of any leftover venv from old installer versions.
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
                "      https://www.python.org/downloads/\n"
                "  School networks may block pypi.org."
            )

    print()
    return subprocess.call([python, str(TOOL), *argv])


if __name__ == "__main__":
    try:
        sys.exit(main(sys.argv[1:]))
    except KeyboardInterrupt:
        sys.exit(130)
