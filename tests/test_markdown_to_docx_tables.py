import sys
import os
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from markdown_to_docx import MarkdownToDocx

from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


def _convert(md: str):
    converter = MarkdownToDocx()
    return converter.convert(md)


def _tables(doc):
    return doc.tables


class TestMarkdownTableDetection(unittest.TestCase):
    """Tests for _is_table_line and _is_separator_line helpers."""

    def test_pipe_row_is_table_line(self):
        self.assertTrue(MarkdownToDocx._is_table_line("| A | B | C |"))

    def test_separator_is_table_line(self):
        self.assertTrue(MarkdownToDocx._is_table_line("| --- | --- | --- |"))

    def test_no_pipe_is_not_table_line(self):
        self.assertFalse(MarkdownToDocx._is_table_line("Normal text"))

    def test_row_without_outer_pipes_is_table_line(self):
        # Two or more pipes → table row
        self.assertTrue(MarkdownToDocx._is_table_line("A | B | C"))

    def test_separator_detected(self):
        self.assertTrue(MarkdownToDocx._is_separator_line("|---|---|"))
        self.assertTrue(MarkdownToDocx._is_separator_line("|:---|---:|:---:|"))
        self.assertFalse(MarkdownToDocx._is_separator_line("| A | B |"))


class TestSplitTableCells(unittest.TestCase):
    def test_basic_split(self):
        cells = MarkdownToDocx._split_table_cells("| A | B | C |")
        self.assertEqual(cells, ["A", "B", "C"])

    def test_no_outer_pipes(self):
        cells = MarkdownToDocx._split_table_cells("A | B | C")
        self.assertEqual(cells, ["A", "B", "C"])

    def test_escaped_pipe(self):
        cells = MarkdownToDocx._split_table_cells(r"| A\|B | C |")
        self.assertEqual(cells, ["A|B", "C"])

    def test_whitespace_trimmed(self):
        cells = MarkdownToDocx._split_table_cells("|  hello  |  world  |")
        self.assertEqual(cells, ["hello", "world"])


class TestParseAlignments(unittest.TestCase):
    def test_default_left(self):
        aligns = MarkdownToDocx._parse_alignments("|---|---|")
        self.assertEqual(aligns, ["left", "left"])

    def test_right_alignment(self):
        aligns = MarkdownToDocx._parse_alignments("|---|---:|")
        self.assertEqual(aligns, ["left", "right"])

    def test_center_alignment(self):
        aligns = MarkdownToDocx._parse_alignments("|:---:|:---:|")
        self.assertEqual(aligns, ["center", "center"])

    def test_mixed_alignments(self):
        aligns = MarkdownToDocx._parse_alignments("|:---|---:|:---:|")
        self.assertEqual(aligns, ["left", "right", "center"])


class TestBasicTableConversion(unittest.TestCase):
    SIMPLE_TABLE = "| Name | Age |\n|------|-----|\n| Alice | 30 |\n| Bob | 25 |"

    def test_table_is_created(self):
        doc = _convert(self.SIMPLE_TABLE)
        self.assertEqual(len(_tables(doc)), 1, "Expected exactly one DOCX table")

    def test_table_dimensions(self):
        doc = _convert(self.SIMPLE_TABLE)
        table = _tables(doc)[0]
        self.assertEqual(len(table.rows), 3)  # header + 2 data rows
        self.assertEqual(len(table.columns), 2)

    def test_header_row_is_bold(self):
        doc = _convert(self.SIMPLE_TABLE)
        table = _tables(doc)[0]
        for cell in table.rows[0].cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    self.assertTrue(run.bold, f"Header cell '{run.text}' should be bold")

    def test_data_rows_not_bold(self):
        doc = _convert(self.SIMPLE_TABLE)
        table = _tables(doc)[0]
        for row in table.rows[1:]:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        self.assertFalse(run.bold, f"Data cell '{run.text}' should not be bold")

    def test_cell_text_content(self):
        doc = _convert(self.SIMPLE_TABLE)
        table = _tables(doc)[0]
        self.assertEqual(table.cell(0, 0).text, "Name")
        self.assertEqual(table.cell(0, 1).text, "Age")
        self.assertEqual(table.cell(1, 0).text, "Alice")
        self.assertEqual(table.cell(1, 1).text, "30")
        self.assertEqual(table.cell(2, 0).text, "Bob")
        self.assertEqual(table.cell(2, 1).text, "25")

    def test_no_extra_paragraphs_for_table_lines(self):
        """Table lines must NOT produce standalone paragraphs."""
        doc = _convert(self.SIMPLE_TABLE)
        # All paragraphs come from inside table cells, not from the body
        body_paragraphs = [p for p in doc.paragraphs if p.text.strip()]
        # python-docx includes an empty paragraph after a table; ignore empty ones
        self.assertEqual(len(body_paragraphs), 0,
                         f"Unexpected body paragraphs: {[p.text for p in body_paragraphs]}")


class TestTableWithoutOuterPipes(unittest.TestCase):
    TABLE_NO_OUTER = "Name | Age\n---|---\nAlice | 30"

    def test_table_created(self):
        doc = _convert(self.TABLE_NO_OUTER)
        self.assertEqual(len(_tables(doc)), 1)

    def test_cell_text(self):
        doc = _convert(self.TABLE_NO_OUTER)
        table = _tables(doc)[0]
        self.assertEqual(table.cell(0, 0).text, "Name")
        self.assertEqual(table.cell(1, 0).text, "Alice")


class TestTableAlignment(unittest.TestCase):
    TABLE = "| L | R | C |\n|:---|---:|:---:|\n| a | b | c |"

    def test_alignments(self):
        doc = _convert(self.TABLE)
        table = _tables(doc)[0]
        # Header row
        self.assertEqual(table.cell(0, 0).paragraphs[0].alignment, WD_PARAGRAPH_ALIGNMENT.LEFT)
        self.assertEqual(table.cell(0, 1).paragraphs[0].alignment, WD_PARAGRAPH_ALIGNMENT.RIGHT)
        self.assertEqual(table.cell(0, 2).paragraphs[0].alignment, WD_PARAGRAPH_ALIGNMENT.CENTER)
        # Data row
        self.assertEqual(table.cell(1, 0).paragraphs[0].alignment, WD_PARAGRAPH_ALIGNMENT.LEFT)
        self.assertEqual(table.cell(1, 1).paragraphs[0].alignment, WD_PARAGRAPH_ALIGNMENT.RIGHT)
        self.assertEqual(table.cell(1, 2).paragraphs[0].alignment, WD_PARAGRAPH_ALIGNMENT.CENTER)


class TestTableEscapedPipes(unittest.TestCase):
    def test_escaped_pipe_in_cell(self):
        md = r"| A\|B | C |" + "\n|---|---|\n" + r"| x\|y | z |"
        doc = _convert(md)
        table = _tables(doc)[0]
        self.assertEqual(table.cell(0, 0).text, "A|B")
        self.assertEqual(table.cell(1, 0).text, "x|y")


class TestTablePaddingShortRows(unittest.TestCase):
    def test_short_row_padded(self):
        md = "| A | B | C |\n|---|---|---|\n| only_one |"
        doc = _convert(md)
        table = _tables(doc)[0]
        self.assertEqual(len(table.columns), 3)
        self.assertEqual(table.cell(1, 0).text, "only_one")
        self.assertEqual(table.cell(1, 1).text, "")
        self.assertEqual(table.cell(1, 2).text, "")


class TestTableDoesNotBreakOtherElements(unittest.TestCase):
    def test_heading_before_table(self):
        md = "# Title\n\n| A | B |\n|---|---|\n| 1 | 2 |"
        doc = _convert(md)
        headings = [p for p in doc.paragraphs if p.style.name.startswith("Heading")]
        self.assertEqual(len(headings), 1)
        self.assertEqual(headings[0].text, "Title")
        self.assertEqual(len(_tables(doc)), 1)

    def test_paragraph_after_table(self):
        md = "| A | B |\n|---|---|\n| 1 | 2 |\n\nFinal paragraph."
        doc = _convert(md)
        self.assertEqual(len(_tables(doc)), 1)
        body_paras = [p for p in doc.paragraphs if p.text.strip()]
        self.assertTrue(any("Final paragraph" in p.text for p in body_paras))

    def test_code_block_not_treated_as_table(self):
        md = "```\n| fake | table |\n|---|---|\n```"
        doc = _convert(md)
        self.assertEqual(len(_tables(doc)), 0, "Code block content must not become a table")

    def test_table_at_eof_no_trailing_newline(self):
        md = "| X | Y |\n|---|---|\n| 1 | 2 |"
        doc = _convert(md)
        self.assertEqual(len(_tables(doc)), 1)


if __name__ == "__main__":
    unittest.main()
