"""PDF to Markdown converter.

Converts PDF files to Markdown with support for:
- Layout-aware extraction (headings, lists, tables, multi-column text)
- Inline formatting: bold, italic, underline, strikethrough
- Embedded image extraction
- OCR fallback for scanned pages (requires Tesseract)

Dependencies
------------
Required:
  pip install PyMuPDF pypdf

Optional (OCR support):
  pip install pytesseract Pillow
  # Also install the Tesseract binary and desired language data, e.g.:
  # sudo apt install tesseract-ocr tesseract-ocr-bul
"""

import argparse
import re
import statistics
import sys
from pathlib import Path
from typing import Dict, List, Literal, Optional, Sequence, Tuple

try:
    import fitz  # PyMuPDF
    _FITZ_BOLD = fitz.TEXT_FONT_BOLD      # 16
    _FITZ_ITALIC = fitz.TEXT_FONT_ITALIC  # 2
except ImportError:  # pragma: no cover
    fitz = None  # type: ignore
    _FITZ_BOLD = 16
    _FITZ_ITALIC = 2

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
LayoutMode = Literal["auto", "pymupdf", "ocr"]
TableMode = Literal["auto", "html", "markdown", "off"]

# ---------------------------------------------------------------------------
# Heading heuristics: font-size ratio relative to median body text
# ---------------------------------------------------------------------------
_HEADING_SCALES: List[Tuple[float, int]] = [
    (1.9, 1),
    (1.55, 2),
    (1.30, 3),
    (1.15, 4),
    (1.05, 5),
]

# ---------------------------------------------------------------------------
# List detection patterns
# ---------------------------------------------------------------------------
_UNORDERED_RE = re.compile(
    r"^(?P<bullet>[•·◦▪▸►▶●○◆◇\u2022\u00B7\u25E6\u25AA]|\-|\*|\+)\s+(?P<text>.+)",
    re.DOTALL,
)
_ORDERED_RE = re.compile(
    r"^(?P<num>\d+)[.)]\s+(?P<text>.+)",
    re.DOTALL,
)


# ===========================================================================
# Legacy helpers (preserved for backward compatibility)
# ===========================================================================

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
    """Extract plain text from a PDF page using pypdf."""
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
        return [img.crop((0, 0, mid, h)), img.crop((mid, 0, w, h))]

    if mode == "4":
        mid_x, mid_y = w // 2, h // 2
        return [
            img.crop((0, 0, mid_x, mid_y)),
            img.crop((mid_x, 0, w, mid_y)),
            img.crop((0, mid_y, mid_x, h)),
            img.crop((mid_x, mid_y, w, h)),
        ]

    raise ValueError(f"Unsupported split mode: {mode}")


def _ocr_image(img: "Image.Image", lang: str) -> str:
    """Run Tesseract OCR on a PIL image and return plain text."""
    if pytesseract is None:
        raise RuntimeError(
            "Missing dependency: pytesseract.\nInstall it with: pip install pytesseract\n"
            "Also install the system Tesseract OCR binary and language data."
        )
    return pytesseract.image_to_string(img, lang=lang) or ""


def _should_do_ocr(mode: OcrMode, extracted_text: str, page_has_images: bool) -> bool:
    """Decide whether to run OCR based on the selected mode."""
    extracted_text_stripped = (extracted_text or "").strip()
    no_text = len(extracted_text_stripped) == 0

    if mode == "no_text":
        return no_text
    if mode == "images":
        return page_has_images
    if mode in ("images_or_no_text", "auto"):
        return page_has_images or no_text

    raise ValueError(f"Unsupported ocr mode: {mode}")


# ===========================================================================
# Layout-aware helpers (PyMuPDF path)
# ===========================================================================

def _font_size_stats(blocks: list) -> Tuple[float, float]:
    """Return (median_size, max_size) of all text spans in *blocks*.

    Falls back to ``(10.0, 10.0)`` when no size data is available.
    """
    sizes: List[float] = []
    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                size = span.get("size", 0)
                if size > 0:
                    sizes.append(size)
    if not sizes:
        return 10.0, 10.0
    return statistics.median(sizes), max(sizes)


def _heading_level(
    size: float, flags: int, font_name: str, median: float
) -> Optional[int]:
    """Return a Markdown heading level (1–6) or ``None`` for body text.

    Uses font-size ratio relative to *median* body size.  A bold span that
    is slightly larger than the median is treated as at least H5.
    """
    if median <= 0:
        return None
    is_bold = bool(flags & _FITZ_BOLD) or "bold" in font_name.lower()
    ratio = size / median

    for scale, level in _HEADING_SCALES:
        if ratio >= scale:
            return level

    # Bold + slightly larger than median → treat as heading level 5
    if is_bold and ratio >= 1.02 and size > median:
        return 5

    return None


def _list_prefix(text: str) -> Optional[Tuple[str, str]]:
    """Detect a list bullet/number prefix.

    Returns ``(marker, content)`` where *marker* is ``"-"`` for unordered or
    ``"N."`` for ordered lists, or ``None`` if no prefix is detected.
    """
    stripped = text.lstrip()
    m = _UNORDERED_RE.match(stripped)
    if m:
        return "-", m.group("text")
    m = _ORDERED_RE.match(stripped)
    if m:
        return f"{m.group('num')}.", m.group("text")
    return None


def _format_span(text: str, flags: int, font_name: str) -> str:
    """Wrap *text* with Markdown bold/italic markers based on font attributes."""
    if not text.strip():
        return text

    fn_lower = font_name.lower()
    is_bold = bool(flags & _FITZ_BOLD) or ("bold" in fn_lower and "notbold" not in fn_lower)
    is_italic = bool(flags & _FITZ_ITALIC) or "italic" in fn_lower or "oblique" in fn_lower

    core = text.strip()
    leading = text[: len(text) - len(text.lstrip())]
    trailing = text[len(text.rstrip()) :]

    if is_bold and is_italic:
        core = f"***{core}***"
    elif is_bold:
        core = f"**{core}**"
    elif is_italic:
        core = f"*{core}*"

    return leading + core + trailing


def _detect_marks(
    span_bbox: Tuple[float, float, float, float], drawings: list
) -> Tuple[bool, bool]:
    """Detect underline and strikethrough for a span from page drawings.

    Inspects *drawings* (from ``page.get_drawings()``) for horizontal lines
    that overlap the span's horizontal extent and cross either the bottom
    (underline) or the mid-line (strikethrough) of the span.

    Returns ``(is_underline, is_strikethrough)``.
    """
    x0, y0, x1, y1 = span_bbox
    span_height = max(y1 - y0, 1.0)
    span_width = max(x1 - x0, 1.0)

    is_underline = False
    is_strikethrough = False

    for path in drawings:
        for item in path.get("items", []):
            if not item or item[0] != "l":
                continue
            try:
                p1, p2 = item[1], item[2]
                # Skip non-horizontal lines (allow up to 2 pt vertical drift)
                if abs(p2.y - p1.y) > 2:
                    continue
                lx0 = min(p1.x, p2.x)
                lx1 = max(p1.x, p2.x)
                ly = (p1.y + p2.y) / 2
                # Require at least 30 % horizontal overlap with the span
                x_overlap = min(x1, lx1) - max(x0, lx0)
                if x_overlap < span_width * 0.30:
                    continue
                rel_y = (ly - y0) / span_height
                if rel_y > 0.80:
                    is_underline = True
                elif 0.35 < rel_y < 0.65:
                    is_strikethrough = True
            except (AttributeError, IndexError, ZeroDivisionError):
                continue

    return is_underline, is_strikethrough


def _table_to_html(data: List[List]) -> str:
    """Convert table data (list-of-rows of strings) to an HTML ``<table>``."""
    if not data:
        return ""

    def cell_text(cell) -> str:
        return str(cell or "").strip().replace("\n", " ")

    rows: List[str] = ["<table>"]
    for row_idx, row in enumerate(data):
        if row_idx == 0:
            rows.append("<thead><tr>")
            for cell in row:
                rows.append(f"  <th>{cell_text(cell)}</th>")
            rows.append("</tr></thead>")
            rows.append("<tbody>")
        else:
            rows.append("<tr>")
            for cell in row:
                rows.append(f"  <td>{cell_text(cell)}</td>")
            rows.append("</tr>")
    rows.append("</tbody></table>")
    return "\n".join(rows)


def _table_to_markdown(data: List[List]) -> str:
    """Convert table data to a GitHub-Flavoured Markdown table."""
    if not data:
        return ""

    col_count = max((len(row) for row in data), default=0)

    def cell(row: list, i: int) -> str:
        return str(row[i] if i < len(row) else "").strip().replace("|", "\\|")

    lines: List[str] = []
    header = [cell(data[0], i) for i in range(col_count)]
    lines.append("| " + " | ".join(header) + " |")
    lines.append("|" + "|".join(" --- " for _ in range(col_count)) + "|")
    for row in data[1:]:
        lines.append("| " + " | ".join(cell(row, i) for i in range(col_count)) + " |")
    return "\n".join(lines)


def _extract_embedded_images(
    page,
    doc,
    assets_dir: Path,
    page_index: int,
) -> List[Tuple[Path, str]]:
    """Extract embedded raster/vector images from *page*.

    Images are saved to *assets_dir* and the function returns a list of
    ``(saved_path, markdown_reference)`` pairs.  The markdown reference uses
    a path relative to *assets_dir*'s parent so it resolves correctly when
    the ``.md`` file sits next to *assets_dir*.
    """
    results: List[Tuple[Path, str]] = []
    try:
        image_list = page.get_images(full=True)
    except Exception:
        return results

    if not image_list:
        return results

    assets_dir.mkdir(parents=True, exist_ok=True)
    seen_xrefs: set = set()

    for img_index, img_info in enumerate(image_list):
        xref = img_info[0]
        if xref in seen_xrefs:
            continue
        seen_xrefs.add(xref)
        try:
            img_data = doc.extract_image(xref)
            if not img_data:
                continue
            ext = img_data.get("ext", "png")
            img_bytes = img_data.get("image", b"")
            if not img_bytes:
                continue
            filename = f"page_{page_index + 1:03d}_img_{img_index + 1:02d}.{ext}"
            img_path = assets_dir / filename
            img_path.write_bytes(img_bytes)
            # Reference uses <assets_dir_name>/<filename> for portability
            md_ref = f"![image]({assets_dir.name}/{filename})"
            results.append((img_path, md_ref))
        except Exception:
            continue

    return results


def _detect_columns(text_blocks: list, page_width: float) -> int:
    """Heuristically detect the number of text columns (1 or 2).

    Examines the horizontal positions of *text_blocks* and returns 2 when
    the blocks clearly separate into two groups that each cover roughly one
    half of the page width.
    """
    if not text_blocks or page_width <= 0:
        return 1

    mid = page_width / 2
    left_count = 0
    right_count = 0
    spanning_count = 0

    for block in text_blocks:
        x0, _y0, x1, _y1 = block["bbox"]
        block_width = x1 - x0
        if block_width > page_width * 0.60:
            spanning_count += 1
            continue
        center_x = (x0 + x1) / 2
        if center_x < mid:
            left_count += 1
        else:
            right_count += 1

    total = left_count + right_count
    if total == 0:
        return 1
    # Both columns must have content; spanning blocks should be a minority
    if (
        left_count >= 2
        and right_count >= 2
        and spanning_count <= max(1, total // 4)
    ):
        return 2
    return 1


def _sort_blocks_reading_order(
    blocks: list, num_columns: int, page_width: float
) -> list:
    """Return *blocks* sorted in natural reading order.

    For single-column documents blocks are ordered top-to-bottom, left-to-right.
    For two-column layouts the left column is read fully before the right column;
    full-width blocks (headers, captions, footers) are inserted at their natural
    vertical position relative to the columns.
    """
    if num_columns == 1 or page_width <= 0:
        return sorted(blocks, key=lambda b: (b["bbox"][1], b["bbox"][0]))

    mid = page_width / 2
    left_col: List[dict] = []
    right_col: List[dict] = []
    full_width: List[dict] = []

    for block in blocks:
        x0, _y0, x1, _y1 = block["bbox"]
        bw = x1 - x0
        if bw > page_width * 0.60:
            full_width.append(block)
        elif (x0 + x1) / 2 < mid:
            left_col.append(block)
        else:
            right_col.append(block)

    left_col.sort(key=lambda b: b["bbox"][1])
    right_col.sort(key=lambda b: b["bbox"][1])
    full_width.sort(key=lambda b: b["bbox"][1])

    if not (left_col or right_col):
        return full_width

    min_col_y = min(b["bbox"][1] for b in left_col + right_col)
    max_col_y = max(b["bbox"][3] for b in left_col + right_col)

    result: List[dict] = []
    result.extend(b for b in full_width if b["bbox"][1] < min_col_y)
    result.extend(left_col)
    result.extend(right_col)
    result.extend(b for b in full_width if b["bbox"][1] >= max_col_y)
    return result


def _block_to_markdown_lines(
    block: dict,
    drawings: list,
    median_size: float,
) -> List[str]:
    """Convert a single PyMuPDF text block to Markdown-formatted lines.

    Each line in the block becomes one output line.  Span-level formatting
    (bold, italic, underline, strikethrough) is applied per span; heading
    detection is based on the dominant font size of the first span.
    """
    lines_out: List[str] = []

    for line in block.get("lines", []):
        spans = line.get("spans", [])
        if not spans:
            continue

        line_parts: List[str] = []
        for span in spans:
            raw_text = span.get("text", "")
            if not raw_text:
                continue

            flags = span.get("flags", 0)
            font_name = span.get("font", "")
            formatted = _format_span(raw_text, flags, font_name)

            # Underline / strikethrough from vector drawings
            is_ul, is_st = _detect_marks(span["bbox"], drawings)
            if is_st:
                core = formatted.strip()
                formatted = formatted.replace(core, f"~~{core}~~")
            if is_ul:
                core = formatted.strip()
                formatted = formatted.replace(core, f"<u>{core}</u>")

            line_parts.append(formatted)

        if not line_parts:
            continue

        line_text = "".join(line_parts).strip()
        if not line_text:
            continue

        # Heading detection from first span
        first_span = spans[0]
        size = first_span.get("size", median_size)
        flags = first_span.get("flags", 0)
        font_name = first_span.get("font", "")
        level = _heading_level(size, flags, font_name, median_size)
        if level:
            lines_out.append(f"{'#' * level} {line_text}")
        else:
            lines_out.append(line_text)

    return lines_out


def _apply_list_detection(lines: List[str]) -> List[str]:
    """Post-process a list of lines to emit Markdown list items."""
    result: List[str] = []
    for line in lines:
        info = _list_prefix(line)
        if info:
            marker, content = info
            result.append(f"{marker} {content}")
        else:
            result.append(line)
    return result


def _extract_page_layout_pymupdf(
    page,
    doc,
    page_index: int,
    assets_dir: Optional[Path],
    table_mode: TableMode,
    extract_images: bool,
) -> str:
    """Layout-aware extraction for a single digital PDF page.

    Uses ``page.get_text("dict")`` for rich layout information (blocks, lines,
    spans with font metrics) and ``page.find_tables()`` for table detection.

    Parameters
    ----------
    page:
        PyMuPDF page object.
    doc:
        PyMuPDF document object (needed for image extraction).
    page_index:
        Zero-based page index (used to name exported image files).
    assets_dir:
        Directory for exported images.  Pass ``None`` to skip image export.
    table_mode:
        Controls how detected tables are serialised; ``"off"`` disables table
        detection entirely.
    extract_images:
        When ``True`` (and *assets_dir* is not ``None``) embedded images are
        exported to *assets_dir*.
    """
    page_dict = page.get_text(
        "dict",
        flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE,
    )
    page_width: float = page_dict.get("width", page.rect.width)
    blocks: list = page_dict.get("blocks", [])

    # Vector drawings for underline / strikethrough detection
    try:
        drawings: list = page.get_drawings()
    except Exception:
        drawings = []

    # Font-size statistics for heading classification
    text_blocks = [b for b in blocks if b.get("type") == 0]
    median_size, _max_size = _font_size_stats(text_blocks)

    # -----------------------------------------------------------------------
    # Table extraction via PyMuPDF (requires PyMuPDF ≥ 1.23)
    # -----------------------------------------------------------------------
    table_bboxes: List[Tuple[float, float, float, float]] = []
    # Map: top-y → rendered table string
    pending_tables: List[Tuple[float, str]] = []

    if table_mode != "off":
        try:
            tabs = page.find_tables()
            for tab in tabs:
                try:
                    data = tab.extract()
                    if not data:
                        continue
                    bbox = tab.bbox  # (x0, y0, x1, y1)
                    table_bboxes.append(bbox)
                    if table_mode in ("html", "auto"):
                        tbl_str = _table_to_html(data)
                    else:
                        tbl_str = _table_to_markdown(data)
                    pending_tables.append((bbox[1], tbl_str))
                except Exception:
                    continue
        except AttributeError:
            pass  # find_tables() not available in this PyMuPDF version

    pending_tables.sort(key=lambda t: t[0])

    # -----------------------------------------------------------------------
    # Embedded image extraction
    # -----------------------------------------------------------------------
    image_refs: List[str] = []
    if extract_images and assets_dir is not None:
        for _path, md_ref in _extract_embedded_images(page, doc, assets_dir, page_index):
            image_refs.append(md_ref)

    # -----------------------------------------------------------------------
    # Assemble output in reading order
    # -----------------------------------------------------------------------
    num_cols = _detect_columns(text_blocks, page_width)
    ordered_blocks = _sort_blocks_reading_order(text_blocks, num_cols, page_width)

    all_lines: List[str] = []
    table_idx = 0

    for block in ordered_blocks:
        bx0, by0, bx1, by1 = block["bbox"]

        # Skip blocks that are covered by a detected table
        in_table = False
        for tx0, ty0, tx1, ty1 in table_bboxes:
            x_overlap = max(0.0, min(tx1, bx1) - max(tx0, bx0))
            y_overlap = max(0.0, min(ty1, by1) - max(ty0, by0))
            if x_overlap > 0 and y_overlap > 0:
                in_table = True
                break
        if in_table:
            continue

        # Insert tables whose top-y precedes this block
        while table_idx < len(pending_tables):
            t_y, t_str = pending_tables[table_idx]
            if t_y <= by0:
                all_lines.extend(["", t_str, ""])
                table_idx += 1
            else:
                break

        block_lines = _block_to_markdown_lines(block, drawings, median_size)
        block_lines = _apply_list_detection(block_lines)
        all_lines.extend(block_lines)

    # Flush any tables that come after all text blocks
    while table_idx < len(pending_tables):
        _t_y, t_str = pending_tables[table_idx]
        all_lines.extend(["", t_str, ""])
        table_idx += 1

    # Append image references at the end of the page
    if image_refs:
        all_lines.append("")
        all_lines.extend(image_refs)

    return "\n".join(all_lines)


# ===========================================================================
# Improved OCR path
# ===========================================================================

def _ocr_page_with_layout(img: "Image.Image", lang: str) -> str:
    """Run Tesseract with word-box output and reconstruct a basic layout.

    Uses ``pytesseract.image_to_data`` to obtain per-word bounding boxes and
    confidence scores, then groups words into lines by their OCR block/paragraph/line
    keys.  The result is similar to ``image_to_string`` but preserves the
    natural reading order more reliably and allows basic list detection.

    Falls back to plain ``image_to_string`` if ``image_to_data`` is unavailable.
    """
    if pytesseract is None:
        raise RuntimeError(
            "Missing dependency: pytesseract.\nInstall it with: pip install pytesseract\n"
            "Also install the system Tesseract OCR binary and language data."
        )

    try:
        data: Dict = pytesseract.image_to_data(
            img,
            lang=lang,
            output_type=pytesseract.Output.DICT,
        )
    except Exception:
        return pytesseract.image_to_string(img, lang=lang) or ""

    # Group words by (block_num, par_num, line_num)
    words_by_line: Dict[Tuple[int, int, int], List[dict]] = {}
    n = len(data.get("text", []))
    for i in range(n):
        try:
            conf = int(data["conf"][i])
        except (ValueError, TypeError):
            conf = 0
        if conf < 0:
            continue
        text = data["text"][i]
        if not text or not text.strip():
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        words_by_line.setdefault(key, []).append(
            {"text": text, "x": data["left"][i]}
        )

    if not words_by_line:
        return ""

    lines_text: List[str] = []
    for key in sorted(words_by_line):
        words = sorted(words_by_line[key], key=lambda w: w["x"])
        lines_text.append(" ".join(w["text"] for w in words))

    return "\n".join(_apply_list_detection(lines_text))


# ===========================================================================
# Main conversion function
# ===========================================================================

def pdf_to_markdown(
    input_path: Path,
    *,
    ocr_mode: OcrMode = "auto",
    ocr_lang: str = "bul",
    split: SplitMode = "none",
    text_mode: TextMode = "clean",
    dpi: int = 300,
    layout_mode: LayoutMode = "auto",
    table_mode: TableMode = "auto",
    extract_images: bool = True,
    assets_dir: Optional[Path] = None,
) -> str:
    """Convert a PDF file to Markdown.

    Layout-aware extraction (headings, lists, tables, multi-column text, inline
    formatting, embedded images) is used for digital PDFs when PyMuPDF can read
    real text objects from the file.  Scanned pages fall back to Tesseract OCR.

    Parameters
    ----------
    input_path:
        Path to the input PDF file.
    ocr_mode:
        Controls when to run OCR.
        ``"auto"`` / ``"images_or_no_text"``: OCR if the page has images *or*
        extracted text is empty.
        ``"no_text"``: OCR only when extracted text is empty.
        ``"images"``: OCR only when images are detected on the page.
    ocr_lang:
        Tesseract language code (default ``"bul"``).
    split:
        Split mode for scanned landscape pages that contain two pages
        side-by-side.  ``"none"`` (default), ``"2"`` (left/right),
        ``"4"`` (quadrants).
    text_mode:
        ``"clean"`` (default): trims lines and collapses blank lines.
        ``"raw"``: keeps text closer to the raw extraction/OCR output.
        Applies only to the OCR path and the legacy plain-text path.
    dpi:
        Rasterisation DPI for OCR (default 300).
    layout_mode:
        ``"auto"`` / ``"pymupdf"``: use PyMuPDF layout extraction for pages
        that contain extractable text; fall back to OCR for scanned pages.
        ``"ocr"``: always use the OCR path (overrides layout detection).
    table_mode:
        Controls table output format.
        ``"auto"`` / ``"html"``: emit HTML ``<table>`` (best fidelity).
        ``"markdown"``: emit a GFM Markdown table.
        ``"off"``: skip table detection.
    extract_images:
        When ``True`` (default) embedded images are exported to *assets_dir*
        and linked from the Markdown output.
    assets_dir:
        Directory where extracted images are saved.  Defaults to
        ``<input_stem>_assets/`` next to the input file when ``None``.
        Set to ``None`` explicitly and ``extract_images=False`` to disable.
    """
    if fitz is None:
        raise RuntimeError(
            "Missing dependency: PyMuPDF.\nInstall it with: pip install PyMuPDF"
        )

    if text_mode not in ("clean", "raw"):
        raise ValueError(f"Unsupported text mode: {text_mode!r}")

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    # Resolve default assets directory
    effective_assets_dir: Optional[Path] = None
    if extract_images:
        effective_assets_dir = assets_dir or (
            input_path.parent / f"{input_path.stem}_assets"
        )

    doc = fitz.open(str(input_path))

    # pypdf is still used for the legacy plain-text path (ocr fallback check)
    reader = None
    if PdfReader is not None:
        try:
            reader = PdfReader(str(input_path))
        except Exception:
            reader = None

    md_lines: List[str] = []
    page_count = doc.page_count

    for page_index in range(page_count):
        page = doc.load_page(page_index)

        # Decide which extraction path to use for this page
        use_ocr = layout_mode == "ocr"

        if not use_ocr:
            # Check whether the page has usable text
            page_text_raw = page.get_text("text").strip()
            page_has_imgs = _page_has_images_pymupdf(page)

            if layout_mode == "auto":
                # Fall back to OCR when there is no extractable text
                if reader is not None:
                    extracted_flat = _extract_text_pypdf(reader, page_index)
                else:
                    extracted_flat = page_text_raw
                use_ocr = _should_do_ocr(ocr_mode, extracted_flat, page_has_imgs)

        if use_ocr:
            # ---------------------------------------------------------------
            # OCR path (scanned pages)
            # ---------------------------------------------------------------
            pil_img = _render_page_to_pil(doc, page_index, dpi=dpi)
            parts_text: List[str] = []
            for part_img in _split_image(pil_img, split):
                try:
                    part_txt = _ocr_page_with_layout(part_img, lang=ocr_lang)
                except Exception:
                    part_txt = _ocr_image(part_img, lang=ocr_lang)
                parts_text.append(part_txt)

            page_text = "\n\n".join(t for t in parts_text if t)
            if text_mode == "clean":
                page_text = _clean_text(page_text)
            elif text_mode == "raw":
                page_text = page_text.replace("\r\n", "\n").replace("\r", "\n")
            else:
                raise ValueError(f"Unsupported text mode: {text_mode}")

        else:
            # ---------------------------------------------------------------
            # Layout-aware PyMuPDF path (digital PDFs)
            # ---------------------------------------------------------------
            page_text = _extract_page_layout_pymupdf(
                page,
                doc,
                page_index,
                effective_assets_dir,
                table_mode,
                extract_images,
            )
            if text_mode == "clean":
                page_text = _clean_text(page_text)

        # Separate pages with a horizontal rule
        if page_index > 0:
            md_lines.append("---")
            md_lines.append("")

        if page_text.strip():
            md_lines.append(page_text)

    while md_lines and not md_lines[-1].strip():
        md_lines.pop()

    return "\n".join(md_lines)


# ===========================================================================
# CLI
# ===========================================================================

def main(argv: Optional[Sequence[str]] = None) -> int:
    """Command-line entry point."""
    parser = argparse.ArgumentParser(
        description=(
            "Convert a PDF file to Markdown.\n\n"
            "For digital PDFs the converter uses PyMuPDF's layout-aware extraction "
            "to preserve headings, lists, tables, inline formatting, and images.  "
            "For scanned pages it falls back to Tesseract OCR."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="Path to input PDF file.")
    parser.add_argument(
        "output",
        nargs="?",
        help="Path to output .md file (default: same name with .md extension).",
    )
    parser.add_argument(
        "--ocr-mode",
        choices=["auto", "no_text", "images", "images_or_no_text"],
        default="auto",
        help=(
            "When to run OCR (default: auto).  "
            "'auto'/'images_or_no_text': OCR if the page has images or no text.  "
            "'no_text': OCR only when extracted text is empty.  "
            "'images': OCR only when images are detected."
        ),
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
        help=(
            "Split scanned page image before OCR: "
            "none (default), 2 (left/right halves), 4 (quadrants)."
        ),
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
        help="Rasterisation DPI for OCR (default: 300).",
    )
    parser.add_argument(
        "--layout-mode",
        choices=["auto", "pymupdf", "ocr"],
        default="auto",
        help=(
            "Extraction strategy (default: auto).  "
            "'auto'/'pymupdf': use PyMuPDF layout for digital PDFs, OCR for scanned.  "
            "'ocr': always use OCR."
        ),
    )
    parser.add_argument(
        "--table-mode",
        choices=["auto", "html", "markdown", "off"],
        default="auto",
        help=(
            "Table output format (default: auto → HTML).  "
            "'html': emit <table> HTML.  "
            "'markdown': emit GFM Markdown tables.  "
            "'off': skip table detection."
        ),
    )
    parser.add_argument(
        "--extract-images",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Extract and link embedded images (default: on).",
    )
    parser.add_argument(
        "--assets-dir",
        default=None,
        help=(
            "Directory for exported images "
            "(default: <output_stem>_assets/ next to the output file)."
        ),
    )

    args = parser.parse_args(list(argv) if argv is not None else None)

    input_path = Path(args.input).expanduser().resolve()
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = input_path.with_suffix(".md")

    # Resolve assets directory relative to the output file
    assets_dir: Optional[Path] = None
    if args.extract_images:
        if args.assets_dir:
            assets_dir = Path(args.assets_dir).expanduser().resolve()
        else:
            assets_dir = output_path.parent / f"{output_path.stem}_assets"

    try:
        md = pdf_to_markdown(
            input_path,
            ocr_mode=args.ocr_mode,
            ocr_lang=args.ocr_lang,
            split=args.split,
            text_mode=args.text_mode,
            dpi=args.dpi,
            layout_mode=args.layout_mode,
            table_mode=args.table_mode,
            extract_images=args.extract_images,
            assets_dir=assets_dir,
        )
    except Exception as exc:  # pragma: no cover
        sys.stderr.write(f"Error: {exc}\n")
        return 1

    output_path.write_text(md, encoding="utf-8")
    print(f"Wrote Markdown to: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
