"""Tests for pdf_to_markdown.py – layout-aware PDF → Markdown conversion."""

import io
import os
import re
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import fitz  # PyMuPDF

from pdf_to_markdown import (
    _apply_list_detection,
    _clean_text,
    _detect_columns,
    _extract_embedded_images,
    _font_size_stats,
    _format_span,
    _heading_level,
    _list_prefix,
    _should_do_ocr,
    _sort_blocks_reading_order,
    _split_image,
    _table_to_html,
    _table_to_markdown,
    pdf_to_markdown,
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_text_pdf(pages_content: list) -> bytes:
    """Create an in-memory PDF with text pages.

    *pages_content* is a list of dicts with keys:
      - ``texts``: list of (x, y, text, fontsize) tuples
    """
    doc = fitz.open()
    for page_spec in pages_content:
        page = doc.new_page(width=595, height=842)
        for x, y, text, size in page_spec.get("texts", []):
            page.insert_text((x, y), text, fontsize=size, color=(0, 0, 0))
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _save_pdf(tmp_dir: Path, name: str, pages_content: list) -> Path:
    path = tmp_dir / name
    path.write_bytes(_make_text_pdf(pages_content))
    return path


# ---------------------------------------------------------------------------
# _clean_text
# ---------------------------------------------------------------------------

class TestCleanText(unittest.TestCase):
    def test_strips_lines(self):
        result = _clean_text("  hello  \n  world  ")
        self.assertEqual(result, "hello\nworld")

    def test_collapses_blank_lines(self):
        result = _clean_text("a\n\n\n\nb")
        self.assertNotIn("\n\n\n", result)

    def test_trims_trailing_blank(self):
        result = _clean_text("hello\n\n")
        self.assertFalse(result.endswith("\n"))


# ---------------------------------------------------------------------------
# _should_do_ocr
# ---------------------------------------------------------------------------

class TestShouldDoOcr(unittest.TestCase):
    def test_auto_no_text_triggers_ocr(self):
        self.assertTrue(_should_do_ocr("auto", "", False))

    def test_auto_with_text_no_images(self):
        self.assertFalse(_should_do_ocr("auto", "some text", False))

    def test_no_text_mode_with_text(self):
        self.assertFalse(_should_do_ocr("no_text", "text", True))

    def test_images_mode_with_image(self):
        self.assertTrue(_should_do_ocr("images", "text", True))

    def test_images_mode_without_image(self):
        self.assertFalse(_should_do_ocr("images", "text", False))


# ---------------------------------------------------------------------------
# _split_image
# ---------------------------------------------------------------------------

class TestSplitImage(unittest.TestCase):
    def _blank_img(self, w: int = 100, h: int = 100):
        from PIL import Image
        return Image.new("RGB", (w, h), color=(255, 255, 255))

    def test_split_none(self):
        img = self._blank_img()
        result = _split_image(img, "none")
        self.assertEqual(len(result), 1)

    def test_split_2(self):
        img = self._blank_img(200, 100)
        parts = _split_image(img, "2")
        self.assertEqual(len(parts), 2)
        self.assertEqual(parts[0].size, (100, 100))
        self.assertEqual(parts[1].size, (100, 100))

    def test_split_4(self):
        img = self._blank_img(200, 200)
        parts = _split_image(img, "4")
        self.assertEqual(len(parts), 4)
        for p in parts:
            self.assertEqual(p.size, (100, 100))

    def test_split_invalid(self):
        img = self._blank_img()
        with self.assertRaises(ValueError):
            _split_image(img, "3")  # type: ignore[arg-type]


# ---------------------------------------------------------------------------
# _font_size_stats
# ---------------------------------------------------------------------------

class TestFontSizeStats(unittest.TestCase):
    def _make_blocks(self, sizes):
        blocks = []
        for s in sizes:
            blocks.append(
                {
                    "type": 0,
                    "bbox": (0, 0, 100, 20),
                    "lines": [
                        {
                            "spans": [
                                {"text": "text", "flags": 0, "font": "Helvetica",
                                 "size": s, "bbox": (0, 0, 50, 20)}
                            ]
                        }
                    ],
                }
            )
        return blocks

    def test_median_and_max(self):
        blocks = self._make_blocks([10, 12, 14, 16])
        med, mx = _font_size_stats(blocks)
        self.assertEqual(mx, 16)
        self.assertAlmostEqual(med, 13.0, places=1)

    def test_empty_blocks(self):
        med, mx = _font_size_stats([])
        self.assertEqual(med, 10.0)
        self.assertEqual(mx, 10.0)

    def test_non_text_blocks_ignored(self):
        blocks = [{"type": 1, "bbox": (0, 0, 100, 100)}]  # image block
        med, mx = _font_size_stats(blocks)
        self.assertEqual(med, 10.0)


# ---------------------------------------------------------------------------
# _heading_level
# ---------------------------------------------------------------------------

class TestHeadingLevel(unittest.TestCase):
    def test_large_size_h1(self):
        level = _heading_level(22.0, 0, "Helvetica", 11.0)
        self.assertEqual(level, 1)

    def test_medium_size_h2(self):
        level = _heading_level(18.0, 0, "Helvetica", 11.0)
        self.assertEqual(level, 2)

    def test_body_size_no_heading(self):
        level = _heading_level(11.0, 0, "Helvetica", 11.0)
        self.assertIsNone(level)

    def test_bold_flag_slightly_larger(self):
        level = _heading_level(11.5, 16, "Helvetica", 11.0)
        self.assertIsNotNone(level)

    def test_bold_in_font_name(self):
        level = _heading_level(11.5, 0, "Arial-Bold", 11.0)
        self.assertIsNotNone(level)

    def test_zero_median_returns_none(self):
        level = _heading_level(14.0, 0, "Helvetica", 0.0)
        self.assertIsNone(level)


# ---------------------------------------------------------------------------
# _list_prefix
# ---------------------------------------------------------------------------

class TestListPrefix(unittest.TestCase):
    def test_bullet_circle(self):
        result = _list_prefix("• Item text")
        self.assertIsNotNone(result)
        marker, content = result
        self.assertEqual(marker, "-")
        self.assertEqual(content, "Item text")

    def test_bullet_middle_dot(self):
        result = _list_prefix("· Item text")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "-")

    def test_bullet_dash(self):
        result = _list_prefix("- Item text")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "-")

    def test_ordered_number(self):
        result = _list_prefix("1. First item")
        self.assertIsNotNone(result)
        marker, content = result
        self.assertEqual(marker, "1.")
        self.assertEqual(content, "First item")

    def test_ordered_paren(self):
        result = _list_prefix("2) Second item")
        self.assertIsNotNone(result)
        self.assertEqual(result[0], "2.")

    def test_no_prefix(self):
        self.assertIsNone(_list_prefix("Regular text"))

    def test_empty_string(self):
        self.assertIsNone(_list_prefix(""))


# ---------------------------------------------------------------------------
# _format_span
# ---------------------------------------------------------------------------

class TestFormatSpan(unittest.TestCase):
    BOLD_FLAG = 16
    ITALIC_FLAG = 2

    def test_plain_text_unchanged(self):
        result = _format_span("hello", 0, "Helvetica")
        self.assertEqual(result.strip(), "hello")

    def test_bold_flag(self):
        result = _format_span("bold", self.BOLD_FLAG, "Helvetica")
        self.assertIn("**", result)

    def test_italic_flag(self):
        result = _format_span("italic", self.ITALIC_FLAG, "Helvetica")
        self.assertIn("*", result)
        self.assertNotIn("**", result)

    def test_bold_and_italic_flags(self):
        result = _format_span("bolditalic", self.BOLD_FLAG | self.ITALIC_FLAG, "Helvetica")
        self.assertIn("***", result)

    def test_bold_font_name(self):
        result = _format_span("text", 0, "Arial-Bold")
        self.assertIn("**", result)

    def test_italic_font_name(self):
        result = _format_span("text", 0, "Times-Italic")
        self.assertIn("*", result)

    def test_whitespace_only_unchanged(self):
        result = _format_span("   ", self.BOLD_FLAG, "Helvetica")
        self.assertEqual(result, "   ")


# ---------------------------------------------------------------------------
# _table_to_html
# ---------------------------------------------------------------------------

class TestTableToHtml(unittest.TestCase):
    def test_basic_table(self):
        data = [["A", "B"], ["1", "2"], ["3", "4"]]
        html = _table_to_html(data)
        self.assertIn("<table>", html)
        self.assertIn("<th>A</th>", html)
        self.assertIn("<td>1</td>", html)
        self.assertIn("</table>", html)

    def test_empty_data(self):
        self.assertEqual(_table_to_html([]), "")

    def test_single_row(self):
        data = [["X", "Y"]]
        html = _table_to_html(data)
        self.assertIn("<th>X</th>", html)


# ---------------------------------------------------------------------------
# _table_to_markdown
# ---------------------------------------------------------------------------

class TestTableToMarkdown(unittest.TestCase):
    def test_basic_table(self):
        data = [["Name", "Age"], ["Alice", "30"], ["Bob", "25"]]
        md = _table_to_markdown(data)
        self.assertIn("| Name |", md)
        self.assertIn("| --- |", md)
        self.assertIn("| Alice |", md)

    def test_empty_data(self):
        self.assertEqual(_table_to_markdown([]), "")

    def test_pipe_escaping(self):
        data = [["A|B", "C"]]
        md = _table_to_markdown(data)
        self.assertIn("\\|", md)


# ---------------------------------------------------------------------------
# _detect_columns
# ---------------------------------------------------------------------------

class TestDetectColumns(unittest.TestCase):
    def _make_block(self, x0, x1, y0=0, y1=20):
        return {"bbox": (x0, y0, x1, y1), "type": 0}

    def test_single_column(self):
        blocks = [self._make_block(50, 300, y0=i * 25) for i in range(5)]
        self.assertEqual(_detect_columns(blocks, 595), 1)

    def test_two_columns(self):
        # Left column blocks
        left = [self._make_block(50, 250, y0=i * 25) for i in range(4)]
        # Right column blocks
        right = [self._make_block(320, 520, y0=i * 25) for i in range(4)]
        result = _detect_columns(left + right, 595)
        self.assertEqual(result, 2)

    def test_empty_blocks(self):
        self.assertEqual(_detect_columns([], 595), 1)

    def test_full_width_blocks_single_column(self):
        # All blocks span full width → single column
        blocks = [self._make_block(50, 545, y0=i * 30) for i in range(5)]
        self.assertEqual(_detect_columns(blocks, 595), 1)


# ---------------------------------------------------------------------------
# _sort_blocks_reading_order
# ---------------------------------------------------------------------------

class TestSortBlocksReadingOrder(unittest.TestCase):
    def _block(self, x0, y0, x1=None, y1=None):
        x1 = x1 or x0 + 100
        y1 = y1 or y0 + 20
        return {"bbox": (x0, y0, x1, y1), "type": 0}

    def test_single_column_sorted_top_to_bottom(self):
        blocks = [self._block(50, 200), self._block(50, 100), self._block(50, 300)]
        result = _sort_blocks_reading_order(blocks, 1, 595)
        ys = [b["bbox"][1] for b in result]
        self.assertEqual(ys, sorted(ys))

    def test_two_column_left_before_right(self):
        left = [self._block(50, i * 30, x1=200) for i in range(3)]
        right = [self._block(350, i * 30, x1=500) for i in range(3)]
        result = _sort_blocks_reading_order(left + right, 2, 595)
        # Left column blocks should come before right column blocks
        idxs_left = [i for i, b in enumerate(result) if b["bbox"][0] < 297]
        idxs_right = [i for i, b in enumerate(result) if b["bbox"][0] >= 297]
        self.assertTrue(max(idxs_left) < min(idxs_right))


# ---------------------------------------------------------------------------
# _apply_list_detection
# ---------------------------------------------------------------------------

class TestApplyListDetection(unittest.TestCase):
    def test_bullet_converted(self):
        lines = ["· Item one", "· Item two"]
        result = _apply_list_detection(lines)
        self.assertTrue(all(ln.startswith("- ") for ln in result))

    def test_ordered_converted(self):
        lines = ["1. First", "2. Second"]
        result = _apply_list_detection(lines)
        self.assertTrue(all(re.match(r"\d+\.", ln) for ln in result))

    def test_plain_text_unchanged(self):
        lines = ["Normal paragraph text."]
        result = _apply_list_detection(lines)
        self.assertEqual(result, lines)


# ---------------------------------------------------------------------------
# _extract_embedded_images
# ---------------------------------------------------------------------------

class TestExtractEmbeddedImages(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def test_no_images_returns_empty(self):
        doc = fitz.open()
        page = doc.new_page()
        result = _extract_embedded_images(page, doc, self.tmp / "assets", 0)
        self.assertEqual(result, [])

    def test_assets_dir_created_only_when_needed(self):
        doc = fitz.open()
        page = doc.new_page()
        assets = self.tmp / "assets"
        _extract_embedded_images(page, doc, assets, 0)
        # No images → directory should NOT be created
        self.assertFalse(assets.exists())


# ---------------------------------------------------------------------------
# pdf_to_markdown – integration tests
# ---------------------------------------------------------------------------

class TestPdfToMarkdown(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def test_missing_file_raises(self):
        with self.assertRaises(FileNotFoundError):
            pdf_to_markdown(self.tmp / "nonexistent.pdf", extract_images=False)

    def test_single_page_text_extraction(self):
        pdf_path = _save_pdf(
            self.tmp,
            "simple.pdf",
            [{"texts": [(72, 100, "Hello World", 12)]}],
        )
        md = pdf_to_markdown(pdf_path, extract_images=False)
        self.assertIn("Hello World", md)

    def test_heading_detection(self):
        """Larger font text should be emitted as a Markdown heading."""
        pdf_path = _save_pdf(
            self.tmp,
            "headings.pdf",
            [
                {
                    "texts": [
                        (72, 80, "Big Heading", 24),
                        (72, 140, "Body text is smaller.", 12),
                    ]
                }
            ],
        )
        md = pdf_to_markdown(pdf_path, extract_images=False)
        self.assertRegex(md, r"#+ Big Heading")

    def test_list_detection(self):
        """Bullet-prefixed lines should be converted to Markdown list items."""
        pdf_path = _save_pdf(
            self.tmp,
            "list.pdf",
            [
                {
                    "texts": [
                        (72, 100, "- First item", 12),
                        (72, 120, "- Second item", 12),
                    ]
                }
            ],
        )
        md = pdf_to_markdown(pdf_path, extract_images=False)
        self.assertIn("- First item", md)

    def test_ordered_list_detection(self):
        pdf_path = _save_pdf(
            self.tmp,
            "ordered_list.pdf",
            [{"texts": [(72, 100, "1. First", 12), (72, 120, "2. Second", 12)]}],
        )
        md = pdf_to_markdown(pdf_path, extract_images=False)
        self.assertRegex(md, r"1\.")

    def test_multi_page_separator(self):
        pdf_path = _save_pdf(
            self.tmp,
            "twopage.pdf",
            [
                {"texts": [(72, 100, "Page one text", 12)]},
                {"texts": [(72, 100, "Page two text", 12)]},
            ],
        )
        md = pdf_to_markdown(pdf_path, extract_images=False)
        self.assertIn("---", md)
        self.assertIn("Page one text", md)
        self.assertIn("Page two text", md)

    def test_extract_images_false_no_assets(self):
        """When extract_images=False, no assets directory should be created."""
        pdf_path = _save_pdf(
            self.tmp,
            "no_images.pdf",
            [{"texts": [(72, 100, "text", 12)]}],
        )
        pdf_to_markdown(pdf_path, extract_images=False)
        assets = self.tmp / "no_images_assets"
        self.assertFalse(assets.exists())

    def test_custom_assets_dir(self):
        """Custom assets_dir should be used for image export."""
        pdf_path = _save_pdf(
            self.tmp,
            "custom_assets.pdf",
            [{"texts": [(72, 100, "text", 12)]}],
        )
        custom_dir = self.tmp / "my_assets"
        pdf_to_markdown(pdf_path, extract_images=True, assets_dir=custom_dir)
        # No embedded images → directory should not be created
        self.assertFalse(custom_dir.exists())

    def test_layout_mode_ocr_requires_tesseract(self):
        """layout_mode='ocr' should attempt OCR (or raise if unavailable)."""
        import pdf_to_markdown as m
        orig = m.pytesseract
        m.pytesseract = None  # type: ignore[assignment]
        try:
            pdf_path = _save_pdf(
                self.tmp,
                "ocr_mode.pdf",
                [{"texts": [(72, 100, "text", 12)]}],
            )
            with self.assertRaises(RuntimeError):
                pdf_to_markdown(pdf_path, layout_mode="ocr", extract_images=False)
        finally:
            m.pytesseract = orig

    def test_table_mode_off_skips_tables(self):
        """table_mode='off' should not raise and should still return text."""
        pdf_path = _save_pdf(
            self.tmp,
            "no_table.pdf",
            [{"texts": [(72, 100, "Some text here", 12)]}],
        )
        md = pdf_to_markdown(pdf_path, table_mode="off", extract_images=False)
        self.assertIn("Some text", md)

    def test_table_mode_markdown(self):
        """table_mode='markdown' should not crash (no tables in this PDF)."""
        pdf_path = _save_pdf(
            self.tmp,
            "md_table.pdf",
            [{"texts": [(72, 100, "Some text", 12)]}],
        )
        md = pdf_to_markdown(pdf_path, table_mode="markdown", extract_images=False)
        self.assertIn("Some text", md)

    def test_text_mode_raw(self):
        pdf_path = _save_pdf(
            self.tmp,
            "raw.pdf",
            [{"texts": [(72, 100, "raw text", 12)]}],
        )
        md = pdf_to_markdown(pdf_path, text_mode="raw", extract_images=False)
        self.assertIn("raw text", md)

    def test_invalid_text_mode_raises(self):
        pdf_path = _save_pdf(
            self.tmp,
            "inv.pdf",
            [{"texts": [(72, 100, "text", 12)]}],
        )
        with self.assertRaises(ValueError):
            pdf_to_markdown(
                pdf_path,
                text_mode="invalid",  # type: ignore[arg-type]
                extract_images=False,
            )


# ---------------------------------------------------------------------------
# CLI integration tests
# ---------------------------------------------------------------------------

class TestMain(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def test_basic_cli(self):
        from pdf_to_markdown import main

        pdf_path = _save_pdf(
            self.tmp,
            "cli.pdf",
            [{"texts": [(72, 100, "CLI test", 12)]}],
        )
        out_path = self.tmp / "cli.md"
        rc = main([str(pdf_path), str(out_path), "--no-extract-images"])
        self.assertEqual(rc, 0)
        self.assertTrue(out_path.exists())
        self.assertIn("CLI test", out_path.read_text(encoding="utf-8"))

    def test_default_output_path(self):
        from pdf_to_markdown import main

        pdf_path = _save_pdf(
            self.tmp,
            "default_out.pdf",
            [{"texts": [(72, 100, "text", 12)]}],
        )
        rc = main([str(pdf_path), "--no-extract-images"])
        self.assertEqual(rc, 0)
        expected_md = pdf_path.with_suffix(".md")
        self.assertTrue(expected_md.exists())

    def test_cli_layout_mode(self):
        from pdf_to_markdown import main

        pdf_path = _save_pdf(
            self.tmp,
            "layout.pdf",
            [{"texts": [(72, 100, "layout text", 12)]}],
        )
        out_path = self.tmp / "layout.md"
        rc = main([
            str(pdf_path), str(out_path),
            "--layout-mode", "pymupdf",
            "--no-extract-images",
        ])
        self.assertEqual(rc, 0)

    def test_cli_table_mode(self):
        from pdf_to_markdown import main

        pdf_path = _save_pdf(
            self.tmp,
            "tbl.pdf",
            [{"texts": [(72, 100, "table text", 12)]}],
        )
        out_path = self.tmp / "tbl.md"
        rc = main([
            str(pdf_path), str(out_path),
            "--table-mode", "html",
            "--no-extract-images",
        ])
        self.assertEqual(rc, 0)

    def test_cli_assets_dir(self):
        from pdf_to_markdown import main

        pdf_path = _save_pdf(
            self.tmp,
            "assets_cli.pdf",
            [{"texts": [(72, 100, "img text", 12)]}],
        )
        out_path = self.tmp / "assets_cli.md"
        assets_dir = self.tmp / "my_assets"
        rc = main([
            str(pdf_path), str(out_path),
            "--assets-dir", str(assets_dir),
        ])
        self.assertEqual(rc, 0)

    def test_cli_missing_file(self):
        from pdf_to_markdown import main

        rc = main([str(self.tmp / "missing.pdf"), "--no-extract-images"])
        self.assertEqual(rc, 1)


if __name__ == "__main__":
    unittest.main()
