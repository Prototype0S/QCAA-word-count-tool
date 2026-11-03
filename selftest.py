"""Quick check that the tool's dependencies are usable.

Run manually with: python selftest.py
Exit code 0 = all good.
"""
import sys
from pathlib import Path


def check(name, fn):
    try:
        fn()
        print(f"  [OK]   {name}")
        return True
    except Exception as e:
        print(f"  [FAIL] {name}: {e}")
        return False


def main():
    print("Running smoke test...")

    def _docx():
        import docx
        import docx.opc  # only real python-docx has this

    def _files():
        if not Path("qcaa_word_count.py").is_file():
            raise FileNotFoundError("qcaa_word_count.py")
        if not Path("qcaa_word_count/__init__.py").is_file():
            raise FileNotFoundError("qcaa_word_count/__init__.py")

    def _package():
        sys.path.insert(0, ".")
        import qcaa_word_count.cli  # noqa: F401

    def _classifier():
        from qcaa_word_count.helpers import classify_span
        from qcaa_word_count.model import Decision
        for tok, expected in [
            ("word", Decision.COUNT),
            ("37°C", Decision.COUNT),
            ("5kg", Decision.COUNT),
            ("42", Decision.EXCLUDE),
            ("100", Decision.EXCLUDE),
            ("πr²", Decision.EXCLUDE),
            ("$", Decision.EXCLUDE),
        ]:
            _, rec = classify_span(tok)
            if rec.decision != expected:
                raise AssertionError(
                    f"{tok!r} gave {rec.decision.value}, "
                    f"expected {expected.value}"
                )

    ok = True
    ok &= check("python-docx importable", _docx)
    ok &= check("tool files present", _files)
    ok &= check("package imports", _package)
    ok &= check("classifier sanity", _classifier)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
