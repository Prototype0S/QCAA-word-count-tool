"""Generate qcaa_stress_test.docx — a document that exercises every rule
the QCAA word count tool tries to handle, plus a few cases it
deliberately doesn't.

Run:  python make_stress_test.py

Delete the generated file after each test — the tool caches decisions
against block_index + text hash, so you want a fresh render each time.
"""
from docx import Document
from docx.shared import Pt


def add_para(doc, text, style=None):
    p = doc.add_paragraph(text, style=style) if style else doc.add_paragraph(text)
    return p


def main():
    doc = Document()

    # ------------------------------------------------------------------
    # Title page — should be excluded entirely
    # ------------------------------------------------------------------
    doc.add_paragraph("Stress Test Document")
    doc.add_paragraph("Prepared by QA Test")
    doc.add_paragraph("Word Count: 900")
    doc.add_paragraph("Date: 1 November 2026")
    add_para(doc, "A subtitle with no full stop")

    # ------------------------------------------------------------------
    # Opening prose (this is the edge case that caught us once)
    # ------------------------------------------------------------------
    add_para(doc, "This document exists to test every branch of the "
                  "QCAA word count tool. It contains citations, numbers, "
                  "units, equations, tables, footnotes, and a few things "
                  "the tool will not catch.")

    # ------------------------------------------------------------------
    # Heading: units and measurements (tests the new measurement rule)
    # ------------------------------------------------------------------
    doc.add_heading("Units and measurements", level=1)
    add_para(doc, "Distances of 30m, 5 km, 100 kilometres, and 2.5 mL all "
                  "appear in this sentence. Temperatures hit 37°C and 100 °F. "
                  "Angles reach 45 degrees. A single letter m appears, and "
                  "so does the number 42 on its own.")
    add_para(doc, "Compare 30m with 30 m and 30 meters. The tool should "
                  "count all three the same way now.")

    # ------------------------------------------------------------------
    # Heading: abbreviations and criteria labels
    # ------------------------------------------------------------------
    doc.add_heading("Abbreviations and criteria labels", level=1)
    add_para(doc, "This paragraph references C6, C8, and A1 as if they were "
                  "criteria. It also mentions USB, HDMI, and the PNG format. "
                  "Acronyms like NASA and ABS should count as words.")

    # ------------------------------------------------------------------
    # Heading: equations (should be excluded when equation-heavy)
    # ------------------------------------------------------------------
    doc.add_heading("Equations and calculations", level=1)
    add_para(doc, "The area of a circle is A = πr², where r is the radius. "
                  "So 5 × 3 = 15, and x² + y² = z² holds in Euclidean "
                  "geometry. A simple calculation is 12 + 8 = 20.")
    add_para(doc, "In prose, the variable x is often used generically, but "
                  "in this sentence it appears mostly as a letter.")

    # ------------------------------------------------------------------
    # Heading: citations in many forms
    # ------------------------------------------------------------------
    doc.add_heading("Citations in every format", level=1)
    add_para(doc, "A simple citation looks like (Smith, 2020). Multi-author "
                  "is (Smith & Jones, 2020) or (Smith, Jones, & Brown, 2020). "
                  "Et-al is (Smith et al., 2020) and with a page is "
                  "(Smith, 2020, p. 170).")
    add_para(doc, "Same author, two works: (Smith, 2020a, 2020b). Multiple "
                  "citations in one paren: (Smith, 2020; Jones, 2021). "
                  "Secondary source: (Smith, 2020, as cited in Jones, 2021).")
    add_para(doc, "Corporate author: (Australian Bureau of Statistics, 2021). "
                  "Long title: (A Guide to All Schools in Brisbane, Queensland "
                  "- Australian Schools Directory, 2021). No date: "
                  "(Smith, n.d.). Narrative form: Smith (2020) argues that "
                  "this is true. Another: According to the AKG First Job to "
                  "Future Career Report Australia (n.d.), 78% of young people "
                  "lack knowledge.")

    # ------------------------------------------------------------------
    # Heading: false-positive traps (should NOT be citations)
    # ------------------------------------------------------------------
    doc.add_heading("Things that look like citations but aren't", level=1)
    add_para(doc, "The figure (Figure 3, 2020) is not a citation. Neither is "
                  "(Note the following, 2020). A parenthetical like "
                  "(In 2020, this happened) is also prose. So is "
                  "(The results were clear, 2020) — the tool should skip it.")

    # ------------------------------------------------------------------
    # Heading: a table that should be included whole
    # ------------------------------------------------------------------
    doc.add_heading("Information table", level=1)
    table = doc.add_table(rows=3, cols=3)
    table.style = "Table Grid"
    table.cell(0, 0).text = "Design"
    table.cell(0, 1).text = "Advantage"
    table.cell(0, 2).text = "Disadvantage"
    table.cell(1, 0).text = "Flying wing"
    table.cell(1, 1).text = "Very efficient and low drag in cruise"
    table.cell(1, 2).text = "Often unstable and harder to fly"
    table.cell(2, 0).text = "Delta wing"
    table.cell(2, 1).text = "Good balance of speed and stability"
    table.cell(2, 2).text = "Higher low-speed drag than a flying wing"

    # ------------------------------------------------------------------
    # Heading: a table with raw data (user must decide)
    # ------------------------------------------------------------------
    doc.add_heading("Raw data table", level=1)
    data = doc.add_table(rows=5, cols=3)
    data.style = "Table Grid"
    data.cell(0, 0).text = "Trial"
    data.cell(0, 1).text = "Value"
    data.cell(0, 2).text = "Note"
    for i in range(1, 5):
        data.cell(i, 0).text = str(i)
        data.cell(i, 1).text = f"{i * 1.5:.1f}"
        data.cell(i, 2).text = ""

    # ------------------------------------------------------------------
    # Heading: a table that contains a citation (tests table citation fix)
    # ------------------------------------------------------------------
    doc.add_heading("Table with citations inside", level=1)
    cite_table = doc.add_table(rows=2, cols=2)
    cite_table.style = "Table Grid"
    cite_table.cell(0, 0).text = "Claim"
    cite_table.cell(0, 1).text = "Source"
    cite_table.cell(1, 0).text = "Population of Brisbane is large"
    cite_table.cell(1, 1).text = "(Australian Bureau of Statistics, 2021)"

    # ------------------------------------------------------------------
    # Heading: text box simulation (plain paragraph, since python-docx
    # can't easily insert real text boxes — but the tool treats any
    # paragraph with a specific style as a candidate)
    # ------------------------------------------------------------------
    doc.add_heading("Quotations and lists", level=1)
    add_para(doc, "Below is a block quote. QCAA counts quotations, "
                  "so this should remain green.")
    add_para(doc, '"The quick brown fox jumps over the lazy dog," said '
                  "the tester, \"and it counts as words.\"")
    add_para(doc, "A bulleted list:", style="List Bullet")
    add_para(doc, "First item in the list", style="List Bullet")
    add_para(doc, "Second item with a number 42", style="List Bullet")
    add_para(doc, "Third item with a unit 5kg", style="List Bullet")

    # ------------------------------------------------------------------
    # Heading: ambiguous things
    # ------------------------------------------------------------------
    doc.add_heading("Ambiguous cases", level=1)
    add_para(doc, "A date like 3/4/2026 is ambiguous by region. A symbol "
                  "like & is definitely excluded. An email address such as "
                  "test@example.com contains mixed tokens. A URL like "
                  "https://example.com/path?query=1 should be a symbol "
                  "in a normal document but a reference entry at the end.")
    add_para(doc, "A single lowercase letter v is ambiguous — in prose it "
                  "counts, but in an equation-heavy block it does not. The "
                  "same goes for the capital V as a unit or a variable.")

    # ------------------------------------------------------------------
    # References (this whole section should be excluded)
    # ------------------------------------------------------------------
    doc.add_heading("References", level=1)
    add_para(doc, "Smith, J. (2020). The citation manual for students. Wiley.")
    add_para(doc, "Jones, A., & Brown, B. (2019). Effects of sleep on memory. "
                  "Journal of Cognitive Research, 12(3), 45–60.")
    add_para(doc, "Australian Bureau of Statistics. (2021). 2021 Brisbane "
                  "Census QuickStats. https://www.abs.gov.au/census/")
    add_para(doc, "AKG. (n.d.). First Job to Future Career Report. "
                  "https://akg.com.au/report.pdf")

    # ------------------------------------------------------------------
    # Appendix (should be flagged, not auto-decided)
    # ------------------------------------------------------------------
    doc.add_heading("Appendix A", level=1)
    add_para(doc, "Supplementary material that should not be counted unless "
                  "the user decides to include it. This paragraph contains "
                  "the number 999 and the formula y = 2x + 1.")

    doc.save("qcqa_stress_test.docx")
    print("Created qcqa_stress_test.docx")


if __name__ == "__main__":
    main()