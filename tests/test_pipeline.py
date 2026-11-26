"""End-to-end tests for the QCAA word count pipeline.

Each test builds a .docx on the fly, runs the production pipeline with
auto_accept=True (no interactive prompts), and asserts a known count.

These are deliberately hand-checkable: each fixture has a comment
explaining the expected count, including the heading (headings count
under QCAA rules).
"""
from __future__ import annotations

import contextlib
import io
import json
import os
import tempfile
import unittest

from docx import Document as DocxDocument
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


def _make_docx(paragraphs: list[str], path: str,
               heading: str = "Introduction") -> None:
    """Build a .docx with an optional body heading.

    The heading matters: the tool relies on Heading styles to tell a
    title page from body text, and a headed document is the normal case.
    Headings themselves count as words, so callers must add 1 for it.
    """
    doc = DocxDocument()
    if heading:
        doc.add_heading(heading, level=1)
    for p in paragraphs:
        doc.add_paragraph(p)
    doc.save(path)


def _add_text_box(paragraph, text: str) -> None:
    """Insert a Word text box (VML shape) holding ``text`` into a paragraph."""
    run = paragraph.add_run()
    xml = (
        f'<w:pict {nsdecls("w")}>'
        '<v:shape xmlns:v="urn:schemas-microsoft-com:vml">'
        f'<v:textbox><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p>'
        '</w:txbxContent></v:textbox></v:shape></w:pict>'
    )
    run._r.append(parse_xml(xml))


def _prepare_document(path: str):
    """Run the production pipeline up to (and including) auto-resolution."""
    from qcaa_word_count.interactive import resolve_all
    from qcaa_word_count.pipeline import _prepare

    document = _prepare(path)
    with contextlib.redirect_stdout(io.StringIO()):
        resolve_all(document, auto_accept=True)
    return document


def _run(path: str) -> dict:
    """Run the real pipeline end-to-end with auto-accept and return the result.

    Uses the production ``_prepare`` (rather than re-assembling the
    pipeline steps here) so the tests exercise exactly the order the
    tool uses.
    """
    from qcaa_word_count.count import count_words

    return count_words(_prepare_document(path))


class TestBasicCounting(unittest.TestCase):
    def test_plain_paragraph(self):
        # heading "Introduction" (1) + "The quick brown fox jumps over
        # the lazy dog" (9) = 10
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(["The quick brown fox jumps over the lazy dog"], path)
            self.assertEqual(_run(path)["total"], 10)

    def test_numbers_excluded(self):
        # heading (1) + "There were 5 cats" -> "There were cats" (3) = 4
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(["There were 5 cats"], path)
            self.assertEqual(_run(path)["total"], 4)

    def test_number_with_unit_counts_the_unit(self):
        # heading (1) + "The mass was 5kg" (4).
        # "5kg" is one measurement token and a unit counts under the
        # QCAA rules, so this is 4 words, not 3.
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(["The mass was 5kg"], path)
            self.assertEqual(_run(path)["total"], 5)

    def test_unit_alone_counts(self):
        # heading (1) + "measured in kg" (3) = 4
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(["measured in kg"], path)
            self.assertEqual(_run(path)["total"], 4)

    def test_symbol_excluded(self):
        # heading (1) + "the + sign" -> "the sign" (2) = 3
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(["the + sign"], path)
            self.assertEqual(_run(path)["total"], 3)

    def test_equation_excluded(self):
        # heading (1) + "x = 5" -> 0 (variable, symbol, number) = 1
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(["x = 5"], path)
            self.assertEqual(_run(path)["total"], 1)

    def test_abbreviations_units_and_formulas_count(self):
        # heading (1) + LPG, KOH, HCl, H2O, CO2 all count (5) plus the
        # connecting prose (3: "and", "and", "are") = 9
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(["LPG KOH HCl H2O CO2 are and and"], path)
            self.assertEqual(_run(path)["total"], 9)


class TestRegions(unittest.TestCase):
    def test_references_excluded(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            doc = DocxDocument()
            doc.add_heading("Introduction", level=1)
            doc.add_paragraph("This is body text with five words.")
            doc.add_heading("References", level=1)
            doc.add_paragraph("Smith, J. (2020). A book. Publisher.")
            doc.save(path)
            # heading "Introduction" (1) + body (7) = 8.
            # The "References" heading and its entry are excluded.
            self.assertEqual(_run(path)["total"], 8)

    def test_short_opening_without_heading_is_not_lost(self):
        # Regression: a document that opens with short paragraphs and no
        # Heading styles used to become 100% "title page" and report 0.
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            doc = DocxDocument()
            doc.add_paragraph("The experiment failed.")
            doc.add_paragraph("This is a normal body paragraph.")
            doc.save(path)
            # "The experiment failed." (3) + "This is a normal body
            # paragraph." (6) = 9
            self.assertEqual(_run(path)["total"], 9)

    def test_cover_page_with_stated_count_stays_excluded(self):
        # A real cover page announces its word count, so it must NOT be
        # rescued into the body.
        from qcaa_word_count.stated import extract_stated_count

        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(
                ["Hour Maker", "A Business Pitch", "Edward Carroll",
                 "Word count: 797"],
                path, heading="",
            )
            document = _prepare_document(path)
            self.assertEqual(extract_stated_count(document), 797)
            self.assertEqual(_run(path)["total"], 0)


class TestCitations(unittest.TestCase):
    def test_apa_citation_excluded(self):
        # heading (1) + "Research shows this is true" (5) = 6;
        # "(Smith, 2020)." is excluded.
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(
                ["Research shows this is true (Smith, 2020)."], path,
            )
            self.assertEqual(_run(path)["total"], 6)


class TestTextBoxes(unittest.TestCase):
    def test_text_box_counted_exactly_once(self):
        # Regression: box text was counted both in the host paragraph and
        # as its own TEXT_BOX block.
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            doc = DocxDocument()
            doc.add_heading("Introduction", level=1)
            para = doc.add_paragraph("Body paragraph here.")
            _add_text_box(para, "Boxed caption text")
            doc.save(path)

            document = _prepare_document(path)
            # heading (1) + body (3) + box (3) = 7
            from qcaa_word_count.count import count_words
            self.assertEqual(count_words(document)["total"], 7)

            # The host paragraph must not contain the box text.
            paragraphs = [b for b in document.blocks
                          if b.block_type.value == "paragraph"]
            self.assertTrue(paragraphs)
            for block in paragraphs:
                self.assertNotIn("Boxed caption text", block.text)


class TestTables(unittest.TestCase):
    def test_data_table_excluded(self):
        # heading (1) + "Body text goes here." (4) = 5; numbers-only
        # table is raw data and is excluded.
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            doc = DocxDocument()
            doc.add_heading("Introduction", level=1)
            doc.add_paragraph("Body text goes here.")
            table = doc.add_table(rows=2, cols=2)
            table.cell(0, 0).text = "1"
            table.cell(0, 1).text = "2"
            table.cell(1, 0).text = "3"
            table.cell(1, 1).text = "4"
            doc.save(path)
            self.assertEqual(_run(path)["total"], 5)


class TestCaptions(unittest.TestCase):
    def _document_with(self, path: str, text: str):
        doc = DocxDocument()
        doc.add_heading("Introduction", level=1)
        doc.add_paragraph("Body text here.")
        doc.add_paragraph(text)
        doc.save(path)
        return _prepare_document(path)

    def test_prose_figure_is_not_a_caption(self):
        # "Figure 3 shows the trend." is prose, not a caption.
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            document = self._document_with(path, "Figure 3 shows the trend.")
            captions = [b for b in document.blocks if b.is_visual]
            self.assertEqual(captions, [])

    def test_labeled_caption_is_flagged(self):
        # "Figure 3: Trend over time" is a caption and is flagged for
        # review (the user decides whether it counts).
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            document = self._document_with(path, "Figure 3: Trend over time")
            captions = [b for b in document.blocks if b.is_visual]
            self.assertEqual(len(captions), 1)


class TestUserOverrides(unittest.TestCase):
    def _document(self, path: str):
        _make_docx(["One two three."], path)
        return _prepare_document(path)

    def _payload_for_first_counted(self, document):
        from qcaa_word_count.model import Decision

        target = next(
            s for b in document.blocks for s in b.spans
            if s.decision_record.decision == Decision.COUNT
        )
        return f"{target.source.block_index}:{target.source.char_start}"

    def test_new_payload_shape_applies(self):
        from qcaa_word_count.count import count_words
        from qcaa_word_count.pipeline import _apply_user_overrides

        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            document = self._document(path)
            before = count_words(document)["total"]
            sid = self._payload_for_first_counted(document)
            with contextlib.redirect_stdout(io.StringIO()):
                _apply_user_overrides(
                    document, {"spans": {sid: "exclude"}, "tables": {},
                               "total": before},
                )
            self.assertEqual(count_words(document)["total"], before - 1)

    def test_legacy_envelope_still_applies(self):
        from qcaa_word_count.count import count_words
        from qcaa_word_count.pipeline import _apply_user_overrides

        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            document = self._document(path)
            before = count_words(document)["total"]
            sid = self._payload_for_first_counted(document)
            with contextlib.redirect_stdout(io.StringIO()):
                _apply_user_overrides(
                    document, {"overrides": {"spans": {sid: "exclude"}}},
                )
            self.assertEqual(count_words(document)["total"], before - 1)

    def test_html_export_shape_is_read_by_cli(self):
        # Regression: the HTML preview exports {spans, tables, total},
        # but the CLI used to look only for a top-level "overrides" key
        # and silently applied nothing.
        from qcaa_word_count.cli import _load_overrides

        payload = {
            "spans": {"0:0": "exclude", "4:12": "count"},
            "tables": {"0": "exclude_all"},
            "total": 42,
        }
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "qcaa_overrides.json")
            with open(path, "w", encoding="utf-8") as f:
                json.dump(payload, f)
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(_load_overrides(path), payload)

    def test_missing_overrides_file_returns_none(self):
        from qcaa_word_count.cli import _load_overrides

        with contextlib.redirect_stdout(io.StringIO()):
            self.assertIsNone(_load_overrides("does-not-exist.json"))


class TestCliFlags(unittest.TestCase):
    def test_html_preview_opens_by_default(self):
        from qcaa_word_count.cli import _build_parser

        self.assertTrue(_build_parser().parse_args([]).open_browser)

    def test_no_open_suppresses_browser(self):
        from qcaa_word_count.cli import _build_parser

        args = _build_parser().parse_args(["--no-open"])
        self.assertFalse(args.open_browser)

    def test_open_flag_is_explicit_and_wins(self):
        from qcaa_word_count.cli import _build_parser

        args = _build_parser().parse_args(["--no-open", "--open"])
        self.assertTrue(args.open_browser)


class TestDecisionLog(unittest.TestCase):
    def test_json_export_contains_all_blocks(self):
        from qcaa_word_count.ingest import ingest
        from qcaa_word_count.output import export_decisions

        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "t.docx")
            _make_docx(["One two three."], path)
            result = _run(path)
            document = ingest(path)
            json_out = export_decisions(document, result, format="json")
            self.assertIn("total", json_out)
            self.assertIn("blocks", json_out)


if __name__ == "__main__":
    unittest.main()
