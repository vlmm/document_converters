"""Тестове за pdf_utils.py – модул за комбиниране и разделяне на PDF файлове."""

import io
import os
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pypdf import PdfReader, PdfWriter

from pdf_utils import _parse_page_ranges, merge_pdfs, split_pdf


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_pdf(num_pages: int) -> bytes:
    """Create a minimal in-memory PDF with *num_pages* blank pages."""
    writer = PdfWriter()
    for _ in range(num_pages):
        writer.add_blank_page(width=72, height=72)
    buf = io.BytesIO()
    writer.write(buf)
    return buf.getvalue()


def _save_pdf(tmp_dir: Path, name: str, num_pages: int) -> Path:
    """Save a minimal PDF with *num_pages* blank pages to *tmp_dir/name*."""
    path = tmp_dir / name
    path.write_bytes(_make_pdf(num_pages))
    return path


# ---------------------------------------------------------------------------
# _parse_page_ranges
# ---------------------------------------------------------------------------

class TestParsePageRanges(unittest.TestCase):
    """Unit tests for the _parse_page_ranges helper."""

    def test_single_page(self):
        self.assertEqual(_parse_page_ranges("1", 5), [0])

    def test_range(self):
        self.assertEqual(_parse_page_ranges("2-4", 5), [1, 2, 3])

    def test_mixed(self):
        self.assertEqual(_parse_page_ranges("1,3-5", 5), [0, 2, 3, 4])

    def test_last_page(self):
        self.assertEqual(_parse_page_ranges("5", 5), [4])

    def test_spaces_ignored(self):
        self.assertEqual(_parse_page_ranges(" 1 , 3 - 4 ", 5), [0, 2, 3])

    def test_duplicates_removed(self):
        result = _parse_page_ranges("1-3,2-4", 5)
        self.assertEqual(result, [0, 1, 2, 3])

    def test_full_document(self):
        self.assertEqual(_parse_page_ranges("1-3", 3), [0, 1, 2])

    def test_invalid_page_zero(self):
        with self.assertRaises(ValueError):
            _parse_page_ranges("0", 5)

    def test_invalid_page_out_of_range(self):
        with self.assertRaises(ValueError):
            _parse_page_ranges("6", 5)

    def test_invalid_range_reversed(self):
        with self.assertRaises(ValueError):
            _parse_page_ranges("5-3", 5)

    def test_invalid_text(self):
        with self.assertRaises(ValueError):
            _parse_page_ranges("abc", 5)

    def test_range_end_out_of_range(self):
        with self.assertRaises(ValueError):
            _parse_page_ranges("1-10", 5)


# ---------------------------------------------------------------------------
# merge_pdfs
# ---------------------------------------------------------------------------

class TestMergePdfs(unittest.TestCase):
    """Integration tests for merge_pdfs."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def test_merge_two_files(self):
        a = _save_pdf(self.tmp, "a.pdf", 2)
        b = _save_pdf(self.tmp, "b.pdf", 3)
        out = self.tmp / "merged.pdf"
        result = merge_pdfs([a, b], out)
        self.assertEqual(result, out)
        self.assertTrue(out.exists())
        reader = PdfReader(str(out))
        self.assertEqual(len(reader.pages), 5)

    def test_merge_single_file(self):
        a = _save_pdf(self.tmp, "single.pdf", 4)
        out = self.tmp / "out.pdf"
        merge_pdfs([a], out)
        reader = PdfReader(str(out))
        self.assertEqual(len(reader.pages), 4)

    def test_merge_preserves_order(self):
        """Pages from first file appear before pages from second file."""
        a = _save_pdf(self.tmp, "first.pdf", 1)
        b = _save_pdf(self.tmp, "second.pdf", 1)
        out = self.tmp / "order.pdf"
        merge_pdfs([a, b], out)
        reader = PdfReader(str(out))
        self.assertEqual(len(reader.pages), 2)

    def test_merge_empty_list_raises(self):
        out = self.tmp / "empty.pdf"
        with self.assertRaises(ValueError):
            merge_pdfs([], out)

    def test_merge_missing_file_raises(self):
        out = self.tmp / "out.pdf"
        with self.assertRaises(FileNotFoundError):
            merge_pdfs([self.tmp / "nonexistent.pdf"], out)

    def test_merge_creates_output_directory(self):
        a = _save_pdf(self.tmp, "a.pdf", 1)
        out = self.tmp / "subdir" / "out.pdf"
        merge_pdfs([a], out)
        self.assertTrue(out.exists())

    def test_merge_three_files(self):
        files = [_save_pdf(self.tmp, f"p{i}.pdf", i + 1) for i in range(3)]
        out = self.tmp / "three.pdf"
        merge_pdfs(files, out)
        reader = PdfReader(str(out))
        # 1+2+3 = 6 pages
        self.assertEqual(len(reader.pages), 6)


# ---------------------------------------------------------------------------
# split_pdf
# ---------------------------------------------------------------------------

class TestSplitPdf(unittest.TestCase):
    """Integration tests for split_pdf."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.src = _save_pdf(self.tmp, "source.pdf", 5)
        self.out_dir = self.tmp / "output"

    def test_split_all_pages(self):
        results = split_pdf(self.src, self.out_dir)
        self.assertEqual(len(results), 5)
        for path in results:
            self.assertTrue(path.exists())
            reader = PdfReader(str(path))
            self.assertEqual(len(reader.pages), 1)

    def test_split_naming_convention(self):
        results = split_pdf(self.src, self.out_dir)
        names = [p.name for p in results]
        self.assertIn("source_page_1.pdf", names)
        self.assertIn("source_page_5.pdf", names)

    def test_split_with_page_range(self):
        results = split_pdf(self.src, self.out_dir, pages="2-4")
        self.assertEqual(len(results), 1)
        reader = PdfReader(str(results[0]))
        self.assertEqual(len(reader.pages), 3)

    def test_split_with_single_page(self):
        results = split_pdf(self.src, self.out_dir, pages="3")
        self.assertEqual(len(results), 1)
        reader = PdfReader(str(results[0]))
        self.assertEqual(len(reader.pages), 1)

    def test_split_with_mixed_ranges(self):
        results = split_pdf(self.src, self.out_dir, pages="1,3-5")
        self.assertEqual(len(results), 1)
        reader = PdfReader(str(results[0]))
        self.assertEqual(len(reader.pages), 4)

    def test_split_missing_file_raises(self):
        with self.assertRaises(FileNotFoundError):
            split_pdf(self.tmp / "missing.pdf", self.out_dir)

    def test_split_creates_output_directory(self):
        nested = self.tmp / "deep" / "nested"
        split_pdf(self.src, nested)
        self.assertTrue(nested.exists())

    def test_split_invalid_pages_raises(self):
        with self.assertRaises(ValueError):
            split_pdf(self.src, self.out_dir, pages="0")

    def test_split_page_out_of_range_raises(self):
        with self.assertRaises(ValueError):
            split_pdf(self.src, self.out_dir, pages="10")


# ---------------------------------------------------------------------------
# CLI (main)
# ---------------------------------------------------------------------------

class TestMain(unittest.TestCase):
    """Tests for the main() CLI entry point."""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())

    def test_merge_cli(self):
        from pdf_utils import main
        a = _save_pdf(self.tmp, "a.pdf", 2)
        b = _save_pdf(self.tmp, "b.pdf", 1)
        out = str(self.tmp / "cli_merged.pdf")
        rc = main(["merge", "-o", out, str(a), str(b)])
        self.assertEqual(rc, 0)
        self.assertTrue(Path(out).exists())
        reader = PdfReader(out)
        self.assertEqual(len(reader.pages), 3)

    def test_split_cli_all_pages(self):
        from pdf_utils import main
        src = _save_pdf(self.tmp, "src.pdf", 3)
        out_dir = str(self.tmp / "split_out")
        rc = main(["split", str(src), "-o", out_dir])
        self.assertEqual(rc, 0)
        parts = list(Path(out_dir).glob("*.pdf"))
        self.assertEqual(len(parts), 3)

    def test_split_cli_with_pages(self):
        from pdf_utils import main
        src = _save_pdf(self.tmp, "src2.pdf", 5)
        out_dir = str(self.tmp / "split_pages_out")
        rc = main(["split", str(src), "-o", out_dir, "--pages", "1-2"])
        self.assertEqual(rc, 0)
        parts = list(Path(out_dir).glob("*.pdf"))
        self.assertEqual(len(parts), 1)
        reader = PdfReader(str(parts[0]))
        self.assertEqual(len(reader.pages), 2)


if __name__ == "__main__":
    unittest.main()
