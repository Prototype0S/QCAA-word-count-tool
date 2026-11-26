"""Regression tests for the span/region/citation accuracy rules.

Converted from a pytest-only file (pytest is not a dependency of this
tool, so the tests never ran). Run with:

    python -m unittest tests.test_accuracy -v
"""
from __future__ import annotations

import tempfile
import unittest

from qcaa_word_count.citations import (
    _author_is_plausible,
    find_apa_citations,
)
from qcaa_word_count.helpers import (
    classify_span,
    is_equation_token,
    is_number_unit_token,
)
from qcaa_word_count.model import (
    Block, BlockType, Citation, Decision, Document, Region, SourceLocation, Span,
)
from qcaa_word_count.regions import (
    _looks_like_body_prose,
    _looks_like_reference,
)


class TestOrdinals(unittest.TestCase):
    def test_ordinals_are_not_number_units(self):
        for tok in ("1st", "2nd", "3rd", "4th", "10th", "21st"):
            self.assertFalse(is_number_unit_token(tok), tok)

    def test_tech_adjectives_are_not_number_units(self):
        for tok in ("3D", "4K", "5G", "2FA"):
            self.assertFalse(is_number_unit_token(tok), tok)

    def test_real_units_are_still_number_units(self):
        for tok in ("5kg", "100km", "2mL", "37°C"):
            self.assertTrue(is_number_unit_token(tok), tok)


class TestCompoundNumbers(unittest.TestCase):
    def test_compound_numbers_are_not_equations(self):
        for tok in ("1-to-1", "2-in-1", "3-on-3", "1-by-1"):
            self.assertFalse(is_equation_token(tok), tok)


class TestCitationPunctuation(unittest.TestCase):
    def _make_doc(self, text: str) -> Document:
        from qcaa_word_count.helpers import make_spans_for_text

        doc = Document(filename="test.docx")
        block = Block(block_type=BlockType.PARAGRAPH, block_index=0)
        block.raw_text = text
        block.spans = make_spans_for_text(text, block_index=0)
        block.region = Region.BODY
        doc.blocks.append(block)
        return doc

    def test_sentence_final_citation_fully_excluded(self):
        from qcaa_word_count.citations import apply_citation_exclusions

        doc = self._make_doc("Some text (Smith, 2020).")
        doc.citations = [Citation(text="(Smith, 2020)", source="auto")]
        apply_citation_exclusions(doc)
        counted = [s.text for s in doc.blocks[0].spans
                   if s.decision_record.decision == Decision.COUNT]
        self.assertNotIn("(Smith,", counted)
        self.assertNotIn("2020).", counted)

    def test_citation_mid_sentence_fully_excluded(self):
        from qcaa_word_count.citations import apply_citation_exclusions

        doc = self._make_doc("As (Smith, 2020) argued.")
        doc.citations = [Citation(text="(Smith, 2020)", source="auto")]
        apply_citation_exclusions(doc)
        counted = [s.text for s in doc.blocks[0].spans
                   if s.decision_record.decision == Decision.COUNT]
        self.assertNotIn("(Smith,", counted)
        self.assertNotIn("2020)", counted)


class TestNonAuthorCitation(unittest.TestCase):
    def test_figure_year_not_a_citation(self):
        self.assertEqual(find_apa_citations("(Figure, 2020)"), [])

    def test_table_year_not_a_citation(self):
        self.assertEqual(find_apa_citations("(Table, 2019)"), [])

    def test_smith_year_is_a_citation(self):
        result = find_apa_citations("(Smith, 2020)")
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0][0], "(Smith, 2020)")

    def test_corporate_multiword_is_a_citation(self):
        result = find_apa_citations("(Australian Bureau of Statistics, 2026)")
        self.assertEqual(len(result), 1)

    def test_malformed_period_comma_is_a_citation(self):
        result = find_apa_citations("(Australian Bureau of Statistics. , 2026)")
        self.assertEqual(len(result), 1)


class TestEquationHeavyBrackets(unittest.TestCase):
    def _make_span(self, text: str) -> Span:
        return Span(text=text, source=SourceLocation(0, 0, len(text)))

    def test_prose_bracket_in_equation_heavy_is_word(self):
        _, rec = classify_span("(labelled", equation_heavy=True)
        self.assertEqual(rec.decision, Decision.COUNT)

    def test_numeric_bracket_in_equation_heavy_is_equation(self):
        _, rec = classify_span("(3.5)", equation_heavy=True)
        self.assertEqual(rec.decision, Decision.EXCLUDE)

    def test_single_letter_bracket_in_equation_heavy_is_equation(self):
        _, rec = classify_span("(m)", equation_heavy=True)
        self.assertEqual(rec.decision, Decision.EXCLUDE)


class TestReferenceDetection(unittest.TestCase):
    def test_apa_entry(self):
        self.assertTrue(_looks_like_reference(
            "Smith, J. (2020). The title. CUP."))

    def test_author_year_entry(self):
        self.assertTrue(_looks_like_reference("Smith (2020). A report."))

    def test_author_and_year_entry(self):
        self.assertTrue(_looks_like_reference(
            "Smith and Jones (2020). A report."))

    def test_organisation_year_entry(self):
        self.assertTrue(_looks_like_reference(
            "Australian Bureau of Statistics (2026). National population estimate."
        ))

    def test_url_entry(self):
        self.assertTrue(_looks_like_reference(
            "Retrieved 20 October 2026, from https://example.com/report.pdf"
        ))

    def test_entry_with_stray_period(self):
        self.assertTrue(_looks_like_reference(
            "Australian Bureau of Statistics. . (2026). Quarterly estimates "
            "[Data set]. . https://dataexplorer.abs.gov.au/vis?tm=Quarterly"
        ))

    def test_entry_with_data_set_period(self):
        self.assertTrue(_looks_like_reference(
            "Australian Bureau of Statistics. (2026). Quarterly estimates "
            "[Data set]. https://dataexplorer.abs.gov.au/"
        ))

    def test_long_prose_with_url_is_not_reference(self):
        self.assertFalse(_looks_like_reference(
            "In 2020 the Reserve Bank published a report on the state of the "
            "economy. It found that growth had slowed significantly. This was "
            "confirmed by subsequent analysis. The pattern continued into the "
            "following year. Economists remain divided on the causes. "
            "https://example.com/paper"
        ))

    def test_prose_with_year_and_journal_hint_is_not_reference(self):
        self.assertFalse(_looks_like_reference(
            "In 2020 the Journal of Economics published a study. The findings "
            "were significant. Further work is needed. This was confirmed. "
            "The pattern continued."
        ))


class TestTitlePageProse(unittest.TestCase):
    def test_short_title_is_not_prose(self):
        self.assertFalse(_looks_like_body_prose("Hour Maker"))

    def test_byline_is_not_prose(self):
        self.assertFalse(_looks_like_body_prose(
            "A Business Pitch By Edward Carroll"))

    def test_stated_count_line_is_not_prose(self):
        self.assertFalse(_looks_like_body_prose("Word Count: 1500"))

    def test_two_sentences_is_prose(self):
        self.assertTrue(_looks_like_body_prose(
            "What am I trying to make? An app to help young students."))

    def test_one_long_sentence_is_prose(self):
        self.assertTrue(_looks_like_body_prose(
            "I chose the name Hour Maker as it captures the core idea of the "
            "business to help young people."
        ))


class TestCitationPersistence(unittest.TestCase):
    def test_manual_citations_survive_round_trip(self):
        from qcaa_word_count.state import load_decisions, save_decisions

        with tempfile.TemporaryDirectory() as tmp:
            save_path = f"{tmp}/state.json"
            doc = Document(filename="test.docx")
            doc.citations = [Citation(text="(ABS, 2026)", source="manual")]
            save_decisions(doc, "fake.docx", save_path=save_path, total=100)

            doc2 = Document(filename="test.docx")
            load_decisions(doc2, save_path)
            self.assertTrue(any(
                c.text == "(ABS, 2026)" and c.source == "manual"
                for c in doc2.citations
            ))


class TestTokeniser(unittest.TestCase):
    def test_tokenize_splits_on_whitespace_only(self):
        from qcaa_word_count.helpers import tokenize

        self.assertEqual(tokenize("Hello, world (Smith, 2020)."),
                         ["Hello,", "world", "(Smith,", "2020)."])


if __name__ == "__main__":
    unittest.main()
