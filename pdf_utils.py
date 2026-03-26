"""Utilities for combining, splitting, and exporting PDF files as images.

Supports:
- Merging multiple PDF files into one
- Splitting a PDF file by individual pages or page ranges
- Extracting PDF pages as PNG images
- Command-line usage (merge / split / extract-images subcommands)
"""

import argparse
import re
import sys
from pathlib import Path
from typing import List, Optional

try:
    from pypdf import PdfReader, PdfWriter  # type: ignore
except ImportError:  # pragma: no cover - clear message if dependency missing
    PdfReader = None  # type: ignore
    PdfWriter = None  # type: ignore

try:
    import fitz  # PyMuPDF  # type: ignore
except ImportError:  # pragma: no cover
    fitz = None  # type: ignore


def _check_dependency() -> None:
    """Check that pypdf is installed; raise RuntimeError if it is not."""
    if PdfReader is None:
        raise RuntimeError(
            "Missing dependency: pypdf.\n"
            "Install it with: pip install pypdf"
        )


def _check_fitz_dependency() -> None:
    """Check that PyMuPDF (fitz) is installed; raise RuntimeError if it is not."""
    if fitz is None:  # pragma: no cover
        raise RuntimeError(
            "Missing dependency: PyMuPDF.\n"
            "Install it with: pip install pymupdf"
        )


def merge_pdfs(input_paths: List[Path], output_path: Path) -> Path:
    """Merge multiple PDF files into a single output file.

    Args:
        input_paths: Ordered sequence of paths to the input PDF files.
        output_path: Path for the output PDF file.

    Returns:
        The path to the created output file.

    Raises:
        RuntimeError: If pypdf is not installed.
        FileNotFoundError: If any of the input files does not exist.
        ValueError: If the list of input files is empty.
    """
    _check_dependency()

    if not input_paths:
        raise ValueError("At least one input PDF file is required.")

    for path in input_paths:
        if not path.exists():
            raise FileNotFoundError(f"Input file not found: {path}")

    writer = PdfWriter()

    for path in input_paths:
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "wb") as fout:
        writer.write(fout)

    return output_path


def _parse_page_ranges(spec: str, total_pages: int) -> List[int]:
    """Parse a page-range string and return a list of 0-based page indices.

    The format is comma-separated ranges or individual page numbers (1-based).
    Examples: ``"1"`` → [0], ``"1-3"`` → [0,1,2], ``"1,3-5"`` → [0,2,3,4].

    Args:
        spec: Range string (e.g. ``"1-3,5,7-9"``).
        total_pages: Total number of pages in the document.

    Returns:
        List of unique 0-based page indices, preserving order.

    Raises:
        ValueError: On invalid format or page numbers outside the valid range.
    """
    indices: List[int] = []
    seen = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            bounds = part.split("-", 1)
            try:
                start = int(bounds[0].strip())
                end = int(bounds[1].strip())
            except ValueError:
                raise ValueError(f"Invalid page range: '{part}'")
            if start < 1 or end < 1:
                raise ValueError(
                    f"Page numbers must be positive: '{part}'"
                )
            if start > end:
                raise ValueError(
                    f"Start page must be <= end page: '{part}'"
                )
            if end > total_pages:
                raise ValueError(
                    f"Page {end} is beyond the document ({total_pages} pages)"
                )
            for i in range(start - 1, end):
                if i not in seen:
                    indices.append(i)
                    seen.add(i)
        else:
            try:
                page_num = int(part)
            except ValueError:
                raise ValueError(f"Invalid page number: '{part}'")
            if page_num < 1:
                raise ValueError(
                    f"Page numbers must be positive: '{part}'"
                )
            if page_num > total_pages:
                raise ValueError(
                    f"Page {page_num} is beyond the document ({total_pages} pages)"
                )
            if (page_num - 1) not in seen:
                indices.append(page_num - 1)
                seen.add(page_num - 1)
    return indices


def split_pdf(
    input_path: Path,
    output_dir: Path,
    pages: Optional[str] = None,
) -> List[Path]:
    """Split a PDF file into individual pages or specified page ranges.

    When ``pages`` is not provided, each page is written to a separate file
    (``<stem>_page_1.pdf``, ``<stem>_page_2.pdf``, …).

    When ``pages`` is provided, only those pages are extracted into *one* file.
    The format is comma-separated ranges (1-based), e.g. ``"1-3,5"``.

    Args:
        input_path: Path to the input PDF file.
        output_dir: Directory where the output files will be written.
        pages: Optional page-range string for selective extraction (e.g. ``"2-4"``).

    Returns:
        List of paths to all created output files.

    Raises:
        RuntimeError: If pypdf is not installed.
        FileNotFoundError: If the input file does not exist.
        ValueError: On invalid page-range format.
    """
    _check_dependency()

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    reader = PdfReader(str(input_path))
    total_pages = len(reader.pages)
    stem = input_path.stem
    output_dir.mkdir(parents=True, exist_ok=True)

    output_paths: List[Path] = []

    if pages is not None:
        # Extract the specified pages into a single output file
        indices = _parse_page_ranges(pages, total_pages)
        writer = PdfWriter()
        for idx in indices:
            writer.add_page(reader.pages[idx])
        # Build a safe filename from the sanitised page spec (keep only digits, commas and hyphens)
        safe_spec = re.sub(r"[^\d,\-]", "", pages.replace(" ", ""))
        out_file = output_dir / f"{stem}_pages_{safe_spec}.pdf"
        with open(out_file, "wb") as fout:
            writer.write(fout)
        output_paths.append(out_file)
    else:
        # One file per page
        for idx in range(total_pages):
            writer = PdfWriter()
            writer.add_page(reader.pages[idx])
            out_file = output_dir / f"{stem}_page_{idx + 1}.pdf"
            with open(out_file, "wb") as fout:
                writer.write(fout)
            output_paths.append(out_file)

    return output_paths


def extract_images_from_pdf(
    input_path: Path,
    output_dir: Path,
    pages: Optional[str] = None,
    dpi: int = 200,
) -> List[Path]:
    """Render PDF pages as PNG images.

    Each selected page is saved as a separate PNG file named
    ``<stem>_page_<N>.png`` where *N* is the 1-based page number within the
    original document.

    Args:
        input_path: Path to the input PDF file.
        output_dir: Directory where the PNG images will be written.
        pages: Optional page-range string (e.g. ``"1-3,5"``).  When omitted,
            all pages are exported.
        dpi: Resolution for rendering (default: 200).

    Returns:
        List of paths to all created PNG files, in page order.

    Raises:
        RuntimeError: If PyMuPDF is not installed.
        FileNotFoundError: If the input file does not exist.
        ValueError: On invalid page-range format.
    """
    _check_fitz_dependency()

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    stem = input_path.stem

    doc = fitz.open(str(input_path))
    total_pages = len(doc)

    if pages is not None:
        indices = _parse_page_ranges(pages, total_pages)
    else:
        indices = list(range(total_pages))

    # Compute zoom factor from DPI (PDF default is 72 dpi)
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)

    output_paths: List[Path] = []
    for idx in indices:
        page = doc[idx]
        pixmap = page.get_pixmap(matrix=matrix)
        out_file = output_dir / f"{stem}_page_{idx + 1}.png"
        pixmap.save(str(out_file))
        output_paths.append(out_file)

    doc.close()
    return output_paths


def main(argv: Optional[List[str]] = None) -> int:
    """Main entry point for the command-line interface.

    Supports three subcommands:

    * ``merge`` – merge PDF files::

        pdf_utils.py merge -o combined.pdf file1.pdf file2.pdf

    * ``split`` – split a PDF file::

        pdf_utils.py split input.pdf -o output_dir/
        pdf_utils.py split input.pdf -o output_dir/ --pages 1-3,5

    * ``extract-images`` – export PDF pages as PNG images::

        pdf_utils.py extract-images input.pdf -o output_dir/
        pdf_utils.py extract-images input.pdf -o output_dir/ --pages 1-3 --dpi 300
    """
    parser = argparse.ArgumentParser(
        description="Combine, split, and export PDF files."
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    # --- merge subcommand ---
    merge_parser = subparsers.add_parser(
        "merge", help="Merge multiple PDF files into one."
    )
    merge_parser.add_argument(
        "inputs",
        nargs="+",
        metavar="INPUT",
        help="Input PDF files (in order).",
    )
    merge_parser.add_argument(
        "-o", "--output",
        required=True,
        metavar="OUTPUT",
        help="Output PDF file.",
    )

    # --- split subcommand ---
    split_parser = subparsers.add_parser(
        "split", help="Split a PDF file into separate files."
    )
    split_parser.add_argument("input", help="Input PDF file.")
    split_parser.add_argument(
        "-o", "--output-dir",
        default=".",
        metavar="DIR",
        help="Output directory (default: current directory).",
    )
    split_parser.add_argument(
        "--pages",
        default=None,
        metavar="RANGES",
        help=(
            "Page ranges to extract (e.g. '1-3,5'). "
            "If not specified, every page is written to a separate file."
        ),
    )

    # --- extract-images subcommand ---
    img_parser = subparsers.add_parser(
        "extract-images", help="Export PDF pages as PNG images."
    )
    img_parser.add_argument("input", help="Input PDF file.")
    img_parser.add_argument(
        "-o", "--output-dir",
        default=".",
        metavar="DIR",
        help="Output directory for PNG files (default: current directory).",
    )
    img_parser.add_argument(
        "--pages",
        default=None,
        metavar="RANGES",
        help=(
            "Page ranges to export (e.g. '1-3,5'). "
            "If not specified, all pages are exported."
        ),
    )
    img_parser.add_argument(
        "--dpi",
        type=int,
        default=200,
        metavar="DPI",
        help="Rendering resolution in DPI (default: 200).",
    )

    args = parser.parse_args(argv)

    try:
        if args.command == "merge":
            output = merge_pdfs(
                [Path(p).expanduser().resolve() for p in args.inputs],
                Path(args.output).expanduser().resolve(),
            )
            print(f"Saved merged PDF: {output}")

        elif args.command == "split":
            outputs = split_pdf(
                Path(args.input).expanduser().resolve(),
                Path(args.output_dir).expanduser().resolve(),
                pages=args.pages,
            )
            for out in outputs:
                print(f"Saved: {out}")

        else:  # extract-images
            outputs = extract_images_from_pdf(
                Path(args.input).expanduser().resolve(),
                Path(args.output_dir).expanduser().resolve(),
                pages=args.pages,
                dpi=args.dpi,
            )
            for out in outputs:
                print(f"Saved: {out}")

    except Exception as exc:  # pragma: no cover - CLI error path
        sys.stderr.write(f"Error: {exc}\n")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
