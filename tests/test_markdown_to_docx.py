import unittest
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from markdown_to_docx import MarkdownToDocx


class TestMarkdownToDocx(unittest.TestCase):

    def _get_runs(self, markdown_input):
        """Helper: convert markdown and return all paragraph runs."""
        converter = MarkdownToDocx()
        doc = converter.convert(markdown_input)
        runs = []
        for para in doc.paragraphs:
            for run in para.runs:
                runs.append(run)
        return runs

    def test_bold(self):
        runs = self._get_runs("**Bold Text**")
        bold_runs = [r for r in runs if r.bold]
        self.assertTrue(len(bold_runs) > 0, "Expected at least one bold run")
        self.assertEqual(bold_runs[0].text, "Bold Text")

    def test_italic(self):
        runs = self._get_runs("*Italic Text*")
        italic_runs = [r for r in runs if r.italic]
        self.assertTrue(len(italic_runs) > 0, "Expected at least one italic run")
        self.assertEqual(italic_runs[0].text, "Italic Text")

    def test_strikethrough(self):
        runs = self._get_runs("~~Strikethrough Text~~")
        strike_runs = [r for r in runs if r.font.strike]
        self.assertTrue(len(strike_runs) > 0, "Expected at least one strikethrough run")
        self.assertEqual(strike_runs[0].text, "Strikethrough Text")

    def test_code_block(self):
        markdown_input = "```\nCode Block\n```"
        converter = MarkdownToDocx()
        doc = converter.convert(markdown_input)
        paragraphs = doc.paragraphs
        self.assertTrue(len(paragraphs) > 0, "Expected at least one paragraph")
        self.assertIn("Code Block", paragraphs[0].text)

    def test_headings(self):
        converter = MarkdownToDocx()
        doc = converter.convert("# Heading 1")
        headings = [p for p in doc.paragraphs if p.style.name.startswith('Heading')]
        self.assertTrue(len(headings) > 0, "Expected at least one heading paragraph")
        self.assertEqual(headings[0].text, "Heading 1")

    def test_lists(self):
        converter = MarkdownToDocx()
        doc = converter.convert("- Item 1\n- Item 2")
        list_paras = [p for p in doc.paragraphs if 'List' in p.style.name]
        self.assertEqual(len(list_paras), 2, "Expected two list items")

    def test_bold_at_end_of_paragraph(self):
        """Test that bold text at the end of a paragraph (starting and ending with **)
        is correctly converted. Regression test for the bug where bold was not applied."""
        text = ("The parties' rights and obligations under this Agreement will bind and "
                "inure to the benefit of their respective successors, heirs, executors "
                "and administrators and permitted assigns. "
                "**Neither party may assign this Agreement without the prior written "
                "consent of the other party, except in connection with a merger, "
                "acquisition or sale of substantially all assets.**")
        runs = self._get_runs(text)
        bold_runs = [r for r in runs if r.bold]
        self.assertTrue(len(bold_runs) > 0,
                        "Bold text starting and ending with ** should be converted to bold")
        self.assertIn("Neither party", bold_runs[0].text)

    def test_bold_entire_line(self):
        """Test that a line where the entire content is wrapped in ** is bold."""
        runs = self._get_runs("**This entire sentence is bold.**")
        bold_runs = [r for r in runs if r.bold]
        self.assertTrue(len(bold_runs) > 0, "Entire bold line should have bold run")
        self.assertEqual(bold_runs[0].text, "This entire sentence is bold.")

    def test_bold_with_period_before_closing(self):
        """Test bold text ending with a period before the closing **."""
        runs = self._get_runs("**text.**")
        bold_runs = [r for r in runs if r.bold]
        self.assertTrue(len(bold_runs) > 0)
        self.assertEqual(bold_runs[0].text, "text.")

    def test_normal_text_not_bold(self):
        """Test that normal text without ** markers is not bold."""
        runs = self._get_runs("Normal text without bold.")
        for run in runs:
            self.assertFalse(run.bold, "Normal text should not be bold")

    def test_mixed_bold_and_normal(self):
        """Test paragraph with both normal and bold text."""
        runs = self._get_runs("Normal start. **Bold end.**")
        normal_runs = [r for r in runs if not r.bold]
        bold_runs = [r for r in runs if r.bold]
        self.assertTrue(len(normal_runs) > 0, "Should have normal runs")
        self.assertTrue(len(bold_runs) > 0, "Should have bold runs")
        self.assertIn("Normal start.", normal_runs[0].text)
        self.assertIn("Bold end.", bold_runs[0].text)


if __name__ == '__main__':
    unittest.main()