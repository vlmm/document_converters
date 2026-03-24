import argparse
import re
import sys
from pathlib import Path

try:
    from docx import Document  # type: ignore
except ImportError:  # pragma: no cover - clear message if dependency missing
    Document = None  # type: ignore

# Marker open/close tags in consistent application order:
# bold  ̷ italic ̷ underline ̷ strikethrough
_MARKER_ORDER = ["**", "_", "<u>", "~~"]
_MARKER_OPEN = {"**": "**", "_": "_", "<u>": "<u>", "~~": "~~"}
_MARKER_CLOSE = {"**": "**", "_": "_", "<u>": "</u>", "~~": "~~"}

# One indent step: 0.5 inch in EMU (English Metric Units).
# 914 400 EMU = 1 inch, so 0.5 inch ̷ 457 200 EMU.
_INDENT_STEP_EMU = 457_200

# Page-break marker emitted when we detect a Word page break.
_PAGE_BREAK_MARKER = '<div style="page-break-after: always;"></div>'

def _heading_level(style_name: str | None) -> int:
    """
    Infer Markdown heading level from a paragraph style name.
    Common styles in Word: 'Heading 1', 'Heading 2', etc.
    """
    if not style_name:
        return 0
    style_name = style_name.lower()
    if "heading" in style_name:
        # Try to extract the first digit after 'heading'
        for ch in style_name:
            if ch.isdigit():
                return max(1, min(6, int(ch)))
        return 1
    return 0

def _paragraph_contains_page_break(p) -> bool:
    """Return True if paragraph *p* contains a hard page break."""
    try:
        # <w:br w:type="page"/> may appear in runs.
        return bool(p._p.xpath('.//w:br[@w:type="page"]'))
    except Exception:
        return False

def _runs_to_markdown_text(runs) -> str:
    """
    Convert a sequence of docx runs to Markdown inline formatting.

    Supports bold (**), italic (_), underline (<u>...</u>), and
    strikethrough (~~).

    Uses an ordered active-marker stack so that opening/closing order is
    always consistent and nesting is valid when multiple styles toggle within
    a paragraph.  When transitioning between run styles the function:
      1. Closes markers (innermost first) that differ from the target set.
      2. Opens the new markers that were not already open.
    This avoids mis-nesting such as **_text**_ and ensures that combined
    styles like bold+italic produce properly nested output (**_text_**).

    Also normalizes embedded line breaks inside a single run: Word may store
    them as '\n' or '\v'. In Markdown output, we convert them to literal
    '<br>' (especially important inside table cells).
    """
    result: list[str] = []
    active: list[str] = []  # currently open markers, in opening order

    for r in runs:
        text = r.text or ""
        if not text:
            continue

        # Normalize in-run line breaks to HTML breaks.
        # (python-docx sometimes returns '\n' or '\v' for manual line breaks.)
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = text.replace("\v", "\n")
        if "\n" in text:
            parts = text.split("\n")
            text = "<br>".join(parts)

        target_bold = bool(getattr(r, "bold", False))
        target_italic = bool(getattr(r, "italic", False))
        target_underline = bool(getattr(r, "underline", False))
        target_strike = bool(getattr(getattr(r, "font", None), "strike", False))

        # Build the desired marker list in a fixed, consistent order.
        target: list[str] = []
        if target_bold:
            target.append("**")
        if target_italic:
            target.append("_")
        if target_underline:
            target.append("<u>")
        if target_strike:
            target.append("~~")

        if active != target:
            # Find how many leading markers are identical in both lists.
            common_prefix = 0
            for i in range(min(len(active), len(target))):
                if active[i] == target[i]:
                    common_prefix += 1
                else:
                    break

            # Close markers from innermost (end of list) down to the split point.
            for m in reversed(active[common_prefix:]):
                result.append(_MARKER_CLOSE[m])

            # Open new markers from the split point onwards.
            for m in target[common_prefix:]:
                result.append(_MARKER_OPEN[m])

            active = target[:]

        result.append(text)

    # Close all remaining open markers from innermost to outermost.
    for m in reversed(active):
        result.append(_MARKER_CLOSE[m])

    return "".join(result)

def _get_list_info(p) -> tuple[bool, int, bool]:
    """
    Return ``(is_list, nesting_level, is_numbered)`` for paragraph *p*.

    Detection strategy (in priority order):
    1. XML ``<w:numPr>`` element ̷ most reliable; provides both the list flag
       and the zero-based ``ilvl`` nesting level directly.
    2. Style name heuristic (``List Bullet`` / ``List Number``) ̷ fallback when
       the paragraph has the right style but no explicit numbering XML (rare).
    """
    style_name = getattr(getattr(p, "style", None), "name", "") or ""
    style_lower = style_name.lower()
    is_numbered_style = "list number" in style_lower
    is_bullet_style = "list bullet" in style_lower

    # --- primary: inspect numbering XML ---
    try:
        pPr = p._p.pPr  # paragraph properties element
        if pPr is not None:
            numPr = pPr.numPr  # numbering properties
            if numPr is not None:
                ilvl_elem = numPr.ilvl
                ilvl_val = int(ilvl_elem.val) if ilvl_elem is not None else 0
                numId_elem = numPr.numId
                num_id_val = (
                    int(numId_elem.val)
                    if numId_elem is not None and numId_elem.val is not None
                    else 0
                )
                if num_id_val > 0:
                    # Full numbering definition present: use ilvl for nesting.
                    return True, ilvl_val, is_numbered_style
                if is_bullet_style or is_numbered_style:
                    # numId absent but ilvl is set (e.g. via style-only list);
                    # still use ilvl for the nesting level.
                    return True, ilvl_val, is_numbered_style
    except (AttributeError, Exception):
        pass

    # --- fallback: style name only, no numPr ---
    if is_bullet_style or is_numbered_style:
        level = _indent_level_from_paragraph(p)
        return True, level, is_numbered_style

    return False, 0, False

def _indent_level_from_paragraph(p) -> int:
    """Estimate nesting level from the paragraph's left indent (0.5 inch per level)."""
    try:
        left_indent = p.paragraph_format.left_indent
        if left_indent is not None and left_indent > 0:
            return max(0, int(left_indent) // _INDENT_STEP_EMU)
    except (AttributeError, Exception):
        pass
    return 0

def _get_paragraph_indent_level(p) -> int:
    """Return the indentation level for a *non-list* paragraph.
    Both left_indent and first_line_indent are considered; the larger value wins.
    Level 0 ̷ no leading spaces; level N ̷ 2*N leading spaces in the output."""
    try:
        pf = p.paragraph_format
        left_indent = pf.left_indent
        first_line_indent = pf.first_line_indent

        indent_emu = 0
        if left_indent is not None and left_indent > 0:
            indent_emu = max(indent_emu, int(left_indent))
        if first_line_indent is not None and first_line_indent > 0:
            indent_emu = max(indent_emu, int(first_line_indent))

        if indent_emu > 0:
            return max(0, indent_emu // _INDENT_STEP_EMU)
    except (AttributeError, Exception):
        pass
    return 0

def _normalize_marker_whitespace(text: str, marker: str) -> str:
    """Normalise whitespace around a Markdown marker (e.g. ** or ~~):
    - strip space/tab immediately after an opening marker
    - strip space/tab immediately before a closing marker
    - insert a space after a closing marker when it is directly followed by a
      non-space character"""
    m = re.escape(marker)
    # **  text** -> **text**
    text = re.sub(rf"{m}[ \t]+(\S)", rf"{marker}\1", text)
    # **text  ** -> **text**
    text = re.sub(rf"(\S)[ \t]+{m}", rf"\1{marker}", text)
    # **text**X -> **text** X
    # The negated character class uses the *first* character of the marker
    # (e.g. '*' for '**', '~' for '~~') to avoid matching a span that itself
    # contains the same marker characters.  This prevents the regex from
    # accidentally consuming nested markers of the same type.
    text = re.sub(
        rf"{m}([^{re.escape(marker[0])}]+){m}(\S)",
        rf"{marker}\1{marker} \2",
        text,
    )
    return text

def _paragraph_to_md_line(p, in_table: bool = False) -> str:
    """Convert a single paragraph object to its Markdown line representation.

    Handles:
    - Inline formatting (bold, italic, underline, strikethrough).
    - Bullet and numbered lists with nesting (``  -``, ``    -``, ...).
    - Headings.
    - Indented non-list paragraphs (leading spaces outside tables, or
      ``&nbsp;`` sequences inside tables).
    - Letter-clause paragraphs like ``(a)``, ``(b)`` are never converted to
      list items.

    Returns an empty string for empty/whitespace-only paragraphs.

    Special cases:
    - If the paragraph contains a hard page break, we emit a page-break marker.
      Note: if the page break is embedded within text, we currently emit the
      marker as a standalone line (keeps behavior simple and predictable)."""
    if _paragraph_contains_page_break(p):
        return _PAGE_BREAK_MARKER

    text = _runs_to_markdown_text(p.runs)
    text = _normalize_marker_whitespace(text, "**")
    text = _normalize_marker_whitespace(text, "~~")

    if not text or not text.strip():
        return ""

    # Letter clauses like (a), (b) ̷ do not convert to list items.
    # The \*{0,2} groups account for optional bold markers (**) that may wrap
    # the clause when the run is bold-formatted.
    is_letter_clause = bool(
        re.match(r"^\s*\*{0,2}\(\s*[a-zA-Z]\s*\)\*{0,2}", text)
    )

    is_list, level, is_numbered = _get_list_info(p)

    if is_list and not is_letter_clause:
        indent = "  " * level
        marker = "1." if is_numbered else "-"
        return f"{indent}{marker} {text}"

    # Heading paragraph.
    head_level = _heading_level(getattr(getattr(p, "style", None), "name", None))
    if head_level > 0:
        return f"{'#' * head_level} {text}"

    # Indented non-list paragraph.
    indent_level = _get_paragraph_indent_level(p)
    if indent_level > 0:
        if in_table:
            return "&nbsp;" * (2 * indent_level) + text
        return "  " * indent_level + text

    return text

def _convert_cell_to_md(cell) -> str:
    """Convert a table cell to a single Markdown string.

    - Each paragraph inside the cell is converted independently (including
      list items and indentation).
    - Multiple paragraphs are joined with ``<br>``.
    - Empty paragraphs (blank lines) are preserved as empty strings, which
      become additional ``<br>`` separators.
    - Leading spaces in each resulting line are replaced by ``&nbsp;``
      sequences so that Markdown renderers preserve the indentation inside
      the table cell.

    Important: Markdown tables must not contain literal newlines inside a cell.
    This function ensures we always return a single-line string for the cell."""
    lines: list[str] = []
    for p in cell.paragraphs:
        line = _paragraph_to_md_line(p, in_table=True)
        # Preserve blank lines as empty entries so they become <br><br> etc.
        if line == "":
            lines.append("")
        else:
            lines.append(line)

    if not lines:
        return ""

    # Replace leading ASCII spaces with &nbsp; so Markdown tables render them.
    processed: list[str] = []
    for line in lines:
        if line == "":
            processed.append("")
            continue
        stripped = line.lstrip(" ")
        n_spaces = len(line) - len(stripped)
        if n_spaces:
            processed.append("&nbsp;" * n_spaces + stripped)
        else:
            processed.append(line)

    # Ensure no literal newlines slip into table cells.
    processed = [s.replace("\n", "<br>") for s in processed]

    return "<br>".join(processed)

def _table_row_cells(row) -> list:
    """Return the list of cells for a row, without duplicates."""
    # python-docx exposes cells via a grid that can cause duplicate cell objects
    # for vertically merged cells (vMerge). Use the underlying XML tc elements
    # to get the actual cell count.
    try:
        tcs = row._tr.xpath("./w:tc")
        if tcs is not None:
            return [row.table._cell(tc, row._tr) for tc in tcs]
    except Exception:
        pass
    return list(row.cells)

def docx_to_markdown(input_path: Path) -> str:
    """Convert a .docx file to a Markdown string.

    Notes:
    - Handles headings, bullet/numbered lists (with nesting), and normal
      paragraphs.
    - Tables are converted to Markdown pipe-tables; multi-paragraph cells are
      joined with ``<br>`` and indentation is preserved via ``&nbsp;``.
    - Blank paragraphs are preserved as blank lines outside tables, and as
      additional ``<br>`` inside table cells.
    - Hard page breaks are emitted as:
      <div style="page-break-after: always;"></div>
    """
    if Document is None:
        raise RuntimeError(
            "Missing dependency: python-docx.\n"
            "Install it with: pip install python-docx"
        )

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    doc = Document(str(input_path))

    md_lines: list[str] = []

    # We iterate over top-level block items in order: paragraphs and tables.
    # python-docx does not expose a high-level API for this, so we inspect
    # the document element tree.
    from docx.table import Table  # type: ignore
    from docx.text.paragraph import Paragraph  # type: ignore

    def iter_block_items(parent):
        """Yield each paragraph and table child within *parent*, in document order."""
        parent_elm = parent.element.body
        for child in parent_elm.iterchildren():
            if child.tag.endswith("}p"):
                yield Paragraph(child, parent)
            elif child.tag.endswith("}tbl"):
                yield Table(child, parent)

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            line = _paragraph_to_md_line(block)
            if not line:
                md_lines.append("")
                continue
            md_lines.append(line)

        elif isinstance(block, Table):
            rows = list(block.rows)
            if not rows:
                continue

            header_row_cells = _table_row_cells(rows[0])
            header_cells = [_convert_cell_to_md(cell) for cell in header_row_cells]
            if any(header_cells):
                md_lines.append("| " + " | ".join(header_cells) + " |")
                md_lines.append("| " + " | ".join("---" for _ in header_cells) + " |")
                for row in rows[1:]:
                    row_cells = _table_row_cells(row)
                    cells = [_convert_cell_to_md(cell) for cell in row_cells]
                    # Ensure consistent column count.
                    if len(cells) < len(header_cells):
                        cells.extend([""] * (len(header_cells) - len(cells)))
                    elif len(cells) > len(header_cells):
                        cells = cells[: len(header_cells)]
                    md_lines.append("| " + " | ".join(cells) + " |")
            else:
                # Fallback: emit as plain text rows
                for row in rows:
                    row_cells = _table_row_cells(row)
                    cells = [_convert_cell_to_md(cell) for cell in row_cells]
                    md_lines.append(" | ".join(cells))

        else:
            # Unknown block type placeholder
            md_lines.append("<!-- Unsupported block element -->")

        # Add a blank line after each block for spacing
        if md_lines and md_lines[-1] != "":
            md_lines.append("")

    # Remove trailing blank lines
    while md_lines and not md_lines[-1].strip():
        md_lines.pop()

    return "\n".join(md_lines)

def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Convert a .docx file to Markdown.")
    parser.add_argument("input", help="Path to input .docx file")
    parser.add_argument(
        "output",
        nargs="?",
        help="Path to output .md file (default: same name with .md extension)",
    )

    args = parser.parse_args(argv)

    input_path = Path(args.input).expanduser().resolve()
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = input_path.with_suffix(".md")

    try:
        md = docx_to_markdown(input_path)
    except Exception as exc:  # pragma: no cover - CLI error path
        sys.stderr.write(f"Error: {exc}\n")
        return 1

    output_path.write_text(md, encoding="utf-8")
    print(f"Wrote Markdown to: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())