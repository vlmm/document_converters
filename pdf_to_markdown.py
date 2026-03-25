import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Literal, Optional, Sequence, Tuple

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover
    fitz = None  # type: ignore

try:
    from pypdf import PdfReader  # type: ignore
except ImportError:  # pragma: no cover
    PdfReader = None  # type: ignore

try:
    import pytesseract  # type: ignore
except ImportError:  # pragma: no cover
    pytesseract = None  # type: ignore

try:
    from PIL import Image  # type: ignore
except ImportError:  # pragma: no cover
    Image = None  # type: ignore


OcrMode = Literal["auto", "no_text", "images", "images_or_no_text"]
SplitMode = Literal["none", "2", "4"]
TextMode = Literal["clean", "raw"]

def _clean_text(text: str) -> str:
    """Clean OCR or extracted text.

    - Normalizes line endings
    - Trims each line
    - Collapses multiple blank lines
    """
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.strip() for ln in text.split("\n")]

    out: List[str] = []
    blank = 0
    for ln in lines:
        if not ln:
            blank += 1
            if blank <= 1:
                out.append("")
            continue
        blank = 0
        out.append(ln)

    while out and out[-1] == "":
        out.pop()

    return "\n".join(out)

def _page_has_images_pymupdf(page) -> bool:
    """Return True if the PyMuPDF page contains any image objects."""
    try:
        return bool(page.get_images(full=True))
    except Exception:
        return False

def _extract_text_pypdf(reader: "PdfReader", page_index: int) -> str:
    """Extract text from a PDF page using pypdf."""
    page = reader.pages[page_index]
    txt = page.extract_text() or ""
    return txt

def _render_page_to_pil(doc, page_index: int, dpi: int) -> "Image.Image":
    """Render a PDF page to a PIL image using PyMuPDF."""
    if fitz is None or Image is None:
        raise RuntimeError(
            "Missing dependencies for rasterization: PyMuPDF and Pillow.\n"
            "Install with: pip install PyMuPDF Pillow"
        )

    page = doc.load_page(page_index)
    zoom = dpi / 72.0
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    return img

def _split_image(img: "Image.Image", mode: SplitMode) -> List["Image.Image"]:
    """Split image into parts according to split mode.

    Ordering is fixed:
    - mode "2": left, right
    - mode "4": top-left, top-right, bottom-left, bottom-right
    """
    if mode == "none":
        return [img]

    w, h = img.size

    if mode == "2":
        mid = w // 2
        left = img.crop((0, 0, mid, h))
        right = img.crop((mid, 0, w, h))
        return [left, right]

    if mode == "4":
        mid_x = w // 2
        mid_y = h // 2
        tl = img.crop((0, 0, mid_x, mid_y))
        tr = img.crop((mid_x, 0, w, mid_y))
        bl = img.crop((0, mid_y, mid_x, h))
        br = img.crop((mid_x, mid_y, w, h))
        return [tl, tr, bl, br]

    raise ValueError(f"Unsupported split mode: {mode}")

def _ocr_image(img: "Image.Image", lang: str) -> str:
    if pytesseract is None:
        raise RuntimeError(
            "Missing dependency: pytesseract.\nInstall it with: pip install pytesseract\n"
            "Also install the system Tesseract OCR binary and language data."
        )

    # Note: this requires the OS-level tesseract binary.
    return pytesseract.image_to_string(img, lang=lang) or ""

def _should_do_ocr(mode: OcrMode, extracted_text: str, page_has_images: bool) -> bool:
    extracted_text_stripped = (extracted_text or "").strip()
    no_text = len(extracted_text_stripped) == 0

    if mode == "no_text":
        return no_text
    if mode == "images":
        return page_has_images
    if mode in ("images_or_no_text", "auto"):
        return page_has_images or no_text

    raise ValueError(f"Unsupported ocr mode: {mode}")

def pdf_to_markdown(
    input_path: Path,
    *,
    ocr_mode: OcrMode = "auto",
    ocr_lang: str = "bul",
    split: SplitMode = "none",
    text_mode: TextMode = "clean",
    dpi: int = 300,
) -> str:
    """Convert a PDF file to Markdown.

    The converter extracts PDF text and optionally runs OCR on pages.

    OCR can be triggered based on:
    - missing extracted text
    - presence of embedded images

    Split modes are intended for scans that contain multiple pages per PDF page.
    """
    if fitz is None:
        raise RuntimeError(
            "Missing dependency: PyMuPDF.\nInstall it with: pip install PyMuPDF"
        )

    if PdfReader is None:
        raise RuntimeError(
            "Missing dependency: pypdf.\nInstall it with: pip install pypdf"
        )

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    doc = fitz.open(str(input_path))
    reader = PdfReader(str(input_path))

    md_lines: List[str] = []

    page_count = doc.page_count
    for page_index in range(page_count):
        page = doc.load_page(page_index)

        extracted_text = _extract_text_pypdf(reader, page_index)
        page_has_images = _page_has_images_pymupdf(page)

        do_ocr = _should_do_ocr(ocr_mode, extracted_text, page_has_images)

        parts_text: List[str] = []
        if do_ocr:
            img = _render_page_to_pil(doc, page_index, dpi=dpi)
            for part_img in _split_image(img, split):
                part_txt = _ocr_image(part_img, lang=ocr_lang)
                parts_text.append(part_txt)
        else:
            parts_text.append(extracted_text)

        page_text = "\n\n".join([t for t in parts_text if t is not None])

        if text_mode == "clean":
            page_text = _clean_text(page_text)
        elif text_mode == "raw":
            page_text = page_text.replace("\r\n", "\n").replace("\r", "\n")
        else:
            raise ValueError(f"Unsupported text mode: {text_mode}")

        # Separate pages with a horizontal rule.
        if page_index > 0:
            md_lines.append("---")
            md_lines.append("")

        if page_text.strip():
            md_lines.append(page_text)

    while md_lines and not md_lines[-1].strip():
        md_lines.pop()

    return "\n".join(md_lines)

def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Convert a PDF file to Markdown with optional OCR.")
    parser.add_argument("input", help="Path to input PDF file")
    parser.add_argument(
        "output",
        nargs="?",
        help="Path to output .md file (default: same name with .md extension)",
    )

    parser.add_argument(
        "--ocr-mode",
        choices=["auto", "no_text", "images", "images_or_no_text"],
        default="auto",
        help="When to run OCR (default: auto).",
    )
    parser.add_argument(
        "--ocr-lang",
        default="bul",
        help="Tesseract language code (default: bul).",
    )
    parser.add_argument(
        "--split",
        choices=["none", "2", "4"],
        default="none",
        help="Split scanned page image before OCR: none, 2 (left/right), 4 (quadrants).",
    )
    parser.add_argument(
        "--text-mode",
        choices=["clean", "raw"],
        default="clean",
        help="Clean or raw output text (default: clean).",
    )
    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="Rasterization DPI for OCR (default: 300).",
    )

    args = parser.parse_args(list(argv) if argv is not None else None)

    input_path = Path(args.input).expanduser().resolve()
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = input_path.with_suffix(".md")

    try:
        md = pdf_to_markdown(
            input_path,
            ocr_mode=args.ocr_mode,
            ocr_lang=args.ocr_lang,
            split=args.split,
            text_mode=args.text_mode,
            dpi=args.dpi,
        )
    except Exception as exc:  # pragma: no cover
        sys.stderr.write(f"Error: {exc}\n")
        return 1

    output_path.write_text(md, encoding="utf-8")
    print(f"Wrote Markdown to: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
