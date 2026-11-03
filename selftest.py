"""Smoke test — run once after install, before launching the tool.

Checks that the pieces the tool depends on are all present and
working. Prints one line per check. Exit code 0 = all good.

Run manually with:
    .venv/bin/python selftest.py        # macOS
    .venv\\Scripts\\python.exe selftest.py    # Windows
"""
from __future__ import annotations

import sys
from pathlib import Path


def check(name: str, fn) -> bool:
    try:
        fn()
        print(f"  [OK]   {name}")
        return True
    except Exception as e:
        print(f"  [FAIL] {name}: {e}")
        return False


def main() -> int:
    print("Running smoke test...")

    def _docx():
        import docx  # noqa: F401
        import docx.opc  # noqa: F401

    def _files():
        if not Path("qcaa_word_count.py").is_file():
            raise FileNotFoundError("qcaa_word_count.py is missing")
        if not Path("qcaa_word_count").is_dir():
            raise FileNotFoundError("qcaa_word_count/ folder is missing")

    def _package():
        sys.path.insert(0, ".")
        import qcaa_word_count  # noqa: F401

    def _classifier():
        from qcaa_word_count.helpers import classify_span
        from qcaa_word_count.model import Decision

        cases = [
            ("word", Decision.COUNT),
            ("37°C", Decision.COUNT),
            ("100", Decision.EXCLUDE),
            ("42", Decision.EXCLUDE),
            ("A = πr²".split()[0], Decision.EXCLUDE),  # "A"
        ]
        for tok, expected in cases:
            _, rec = classify_span(tok)
            if rec.decision != expected:
                raise AssertionError(
                    f"classify_span({tok!r}) gave {rec.decision.value}, "
                    f"expected {expected.value}"
                )

    ok = True
    ok &= check("python-docx importable", _docx)
    ok &= check("tool files present", _files)
    ok &= check("qcaa_word_count package importable", _package)
    ok &= check("QCAA span classifier sanity", _classifier)

    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())