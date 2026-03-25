import argparse
import re
import sys
from pathlib import Path

try:
    from docx import Document  # type: ignore
except ImportError:  # pragma: no cover - clear message if dependency missing
    Document = None  # type: ignore

from utilities.docx_wml import (
    ListFormat,
    build_list_formats,
    extract_numpr,
    paragraph_contains_page_break,
    xpath as _xpath,
)

_MARKER_ORDER = ["**", "_", "<u>", "~~"]
_MARKER_OPEN = {"**": "**", "_": "_", "<u>": "<u>", "~~": "~~"}
_MARKER_CLOSE = {"**": "**", "_": "_", "<u>": "</u>", "~~": "~~"}

_INDENT_STEP_EMU = 457_200

_PAGE_BREAK_MARKER = '<div style="page-break-after: always;"></div>'

def _heading_level(style_name: str | None) -> int:
    if not style_name:
        return 0
    style_name = style_name.lower()
    if "heading" in style_name:
        for ch in style_name:
            if ch.isdigit():
                return max(1, min(6, int(ch)))
        return 1
    return 0

def _is_non_black_color(r) -> bool:
    """Return True if the run has an explicit non-black color.

    Auto/default color (None rgb) and explicit black (#000000) are treated as
    black.  Any other explicit RGB color is considered non-black.
    """
    try:
        color = getattr(getattr(r, "font", None), "color", None)
        if color is None:
            return False
        rgb = color.rgb
        if rgb is None:
            return False
        # python-docx RGBColor is an int subclass; 0 == #000000.
        # Fall back to string comparison for other representations.
        try:
            return int(rgb) != 0
        except (TypeError, ValueError):
            return str(rgb).upper() != "000000"
    except Exception:
        return False


def _runs_to_markdown_text(runs, italic_non_black: bool = False) -> str:
    result: list[str] = []
    active: list[str] = []  # currently open markers, in opening order

    for r in runs:
        text = r.text or ""
        if not text:
            continue

        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = text.replace("\v", "\n")
        if "\n" in text:
            parts = text.split("\n")
            text = "<br>".join(parts)

        target_bold = bool(getattr(r, "bold", False))
        target_italic = bool(getattr(r, "italic", False)) or (
            italic_non_black and _is_non_black_color(r)
        )
        target_underline = bool(getattr(r, "underline", False))
        target_strike = bool(getattr(getattr(r, "font", None), "strike", False))

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
            common_prefix = 0
            for i in range(min(len(active), len(target))):
                if active[i] == target[i]:
                    common_prefix += 1
                else:
                    break

            for m in reversed(active[common_prefix:]):
                result.append(_MARKER_CLOSE[m])

            for m in target[common_prefix:]:
                result.append(_MARKER_OPEN[m])

            active = target[:]

        result.append(text)

    for m in reversed(active):
        result.append(_MARKER_CLOSE[m])

    return "".join(result)

def _indent_level_from_paragraph(p) -> int:
    try:
        left_indent = p.paragraph_format.left_indent
        if left_indent is not None and left_indent > 0:
            return max(0, int(left_indent) // _INDENT_STEP_EMU)
    except (AttributeError, Exception):
        pass
    return 0

def _get_paragraph_indent_level(p) -> int:
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
    m = re.escape(marker)
    text = re.sub(rf"{m}[ \t]+(\S)", rf"{marker}\1", text)
    text = re.sub(rf"(\S)[ \t]+{m}", rf"\1{marker}", text)
    text = re.sub(
        rf"{m}([^{re.escape(marker[0])}]+){m}(\S)",
        rf"{marker}\1{marker} \2",
        text,
    )
    return text

def _int_to_roman(n: int, upper: bool = True) -> str:
    if n <= 0:
        return str(n)
    vals = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ]
    out: list[str] = []
    x = n
    for v, s in vals:
        while x >= v:
            out.append(s)
            x -= v
    res = "".join(out)
    return res if upper else res.lower()

def _int_to_alpha(n: int, upper: bool = False) -> str:
    # 1 -> a, 26 -> z, 27 -> aa ...
    if n <= 0:
        return str(n)
    x = n
    chars: list[str] = []
    while x > 0:
        x -= 1
        chars.append(chr(ord("A" if upper else "a") + (x % 26)))
        x //= 26
    return "".join(reversed(chars))

class _ListState:
    def __init__(self, formats: dict[tuple[int, int], ListFormat]):
        self.formats = formats
        self.counters: dict[tuple[int, int], int] = {}

    def next_marker(self, num_id: int, ilvl: int) -> str:
        fmt = self.formats.get((num_id, ilvl))
        if fmt is None:
            # fallback
            n = self.counters.get((num_id, ilvl), 0) + 1
            self.counters[(num_id, ilvl)] = n
            return f"{n}."
        if fmt.kind == "bullet":
            return "-"
        current = self.counters.get((num_id, ilvl))
        if current is None:
            current = fmt.start
        else:
            current += 1
        self.counters[(num_id, ilvl)] = current

        # Build displayed token based on kind; then inject into lvl_text
        if fmt.kind == "decimal":
            token = str(current)
        elif fmt.kind == "upperRoman":
            token = _int_to_roman(current, upper=True)
        elif fmt.kind == "lowerRoman":
            token = _int_to_roman(current, upper=False)
        elif fmt.kind == "upperLetter":
            token = _int_to_alpha(current, upper=True)
        elif fmt.kind == "lowerLetter":
            token = _int_to_alpha(current, upper=False)
        else:
            token = str(current)

        # Word uses %1, %2... for levels; we only render current level (%1)
        # As a pragmatic approximation: replace any %<digit> with token
        marker = re.sub(r"%\d+", token, fmt.lvl_text or ("%1."))
        marker = marker.strip()
        return marker

def _get_list_info(p) -> tuple[bool, int, bool]:
    """
    Return (is_list, nesting_level, is_numbered_style)
    """
    style_name = getattr(getattr(p, "style", None), "name", "") or ""
    style_lower = style_name.lower()
    is_numbered_style = "list number" in style_lower
    is_bullet_style = "list bullet" in style_lower

    has_numpr, _num_id, ilvl = extract_numpr(p)
    if has_numpr:
        return True, ilvl, is_numbered_style

    if is_bullet_style or is_numbered_style:
        # Try to read ilvl directly from XML (numId may be absent or 0)
        try:
            pPr = p._p.pPr
            if pPr is not None and pPr.numPr is not None:
                ilvl_elem = pPr.numPr.ilvl
                if ilvl_elem is not None and ilvl_elem.val is not None:
                    return True, int(ilvl_elem.val), is_numbered_style
        except Exception:
            pass
        level = _indent_level_from_paragraph(p)
        return True, level, is_numbered_style

    return False, 0, False

def _paragraph_to_md_line(
    p,
    list_state: _ListState | None = None,
    in_table: bool = False,
    italic_non_black: bool = False,
) -> str:
    if paragraph_contains_page_break(p):
        return _PAGE_BREAK_MARKER

    text = _runs_to_markdown_text(p.runs, italic_non_black=italic_non_black)
    text = _normalize_marker_whitespace(text, "**")
    text = _normalize_marker_whitespace(text, "~~")

    if not text or not text.strip():
        return ""

    is_letter_clause = bool(re.match(r"^\s*\*{0,2}\(\s*[a-zA-Z]\s*\)\*{0,2}", text))

    is_list, level, is_numbered_style = _get_list_info(p)
    _has_numpr, num_id, ilvl = extract_numpr(p)

    if is_list and not is_letter_clause:
        indent = "  " * level

        # Bullet list (or unknown without numId)
        if num_id <= 0:
            marker = "-" if not is_numbered_style else "1."
            return f"{indent}{marker} {text}"

        # Numbered/bulleted list from numbering definitions
        marker = "-" if not is_numbered_style else "1."
        if list_state is not None:
            marker = list_state.next_marker(num_id, ilvl)

        # If it's a bullet-like marker, emit '-' (Markdown bullet)
        if marker in ("-", "•", "–", "—", "o"):
            return f"{indent}- {text}"

        # If marker is standard decimal like "3." prefer Markdown ordered list syntax:
        if re.fullmatch(r"\d+\.", marker):
            return f"{indent}{marker} {text}"

        # Otherwise (roman/alpha/custom patterns) simulate in plain Markdown:
        # use bullet + literal marker text to keep visual closeness and allow restarts.
        return f"{indent}- {marker} {text}"

    head_level = _heading_level(getattr(getattr(p, "style", None), "name", None))
    if head_level > 0:
        return f"{'#' * head_level} {text}"

    indent_level = _get_paragraph_indent_level(p)
    if indent_level > 0:
        if in_table:
            return "&nbsp;" * (2 * indent_level) + text
        return "  " * indent_level + text

    return text

def _convert_cell_to_md(
    cell, list_state: _ListState | None = None, italic_non_black: bool = False
) -> str:
    lines: list[str] = []
    for p in cell.paragraphs:
        line = _paragraph_to_md_line(
            p, list_state=list_state, in_table=True, italic_non_black=italic_non_black
        )
        if line:
            lines.append(line)

    if not lines:
        return ""

    processed: list[str] = []
    for line in lines:
        stripped = line.lstrip(" ")
        n_spaces = len(line) - len(stripped)
        if n_spaces:
            processed.append("&nbsp;" * n_spaces + stripped)
        else:
            processed.append(line)

    processed = [s.replace("\n", "<br>") for s in processed]

    return "<br>".join(processed)

def _table_row_cells(row) -> list:
    """Return unique cells for a table row, skipping merged-cell duplicates.

    python-docx's ``row.cells`` returns the same ``_Cell`` object for every
    grid column that a merged cell occupies (both horizontal and vertical
    merges).  By comparing the underlying ``w:tc`` XML elements we can filter
    duplicates and return each logical cell exactly once.
    """
    seen: set[int] = set()
    result = []
    for cell in row.cells:
        tc_id = id(cell._tc)
        if tc_id not in seen:
            seen.add(tc_id)
            result.append(cell)
    return result

def docx_to_markdown(input_path: Path, italic_non_black: bool = False) -> str:
    if Document is None:
        raise RuntimeError(
            "Missing dependency: python-docx.\n" "Install it with: pip install python-docx"
        )

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    doc = Document(str(input_path))

    list_formats = build_list_formats(doc)
    list_state = _ListState(list_formats)

    md_lines: list[str] = []

    from docx.table import Table  # type: ignore
    from docx.text.paragraph import Paragraph  # type: ignore

    def iter_block_items(parent):
        parent_elm = parent.element.body
        for child in parent_elm.iterchildren():
            if child.tag.endswith("}p"):
                yield Paragraph(child, parent)
            elif child.tag.endswith("}tbl"):
                yield Table(child, parent)

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            line = _paragraph_to_md_line(
                block, list_state=list_state, italic_non_black=italic_non_black
            )
            if not line:
                md_lines.append("")
                continue
            md_lines.append(line)

        elif isinstance(block, Table):
            rows = list(block.rows)
            if not rows:
                continue

            header_row_cells = _table_row_cells(rows[0])
            header_cells = [
                _convert_cell_to_md(
                    cell, list_state=list_state, italic_non_black=italic_non_black
                )
                for cell in header_row_cells
            ]
            if any(header_cells):
                md_lines.append("| " + " | ".join(header_cells) + " |")
                md_lines.append("| " + " | ".join("---" for _ in header_cells) + " |")
                for row in rows[1:]:
                    row_cells = _table_row_cells(row)
                    cells = [
                        _convert_cell_to_md(
                            cell,
                            list_state=list_state,
                            italic_non_black=italic_non_black,
                        )
                        for cell in row_cells
                    ]
                    if len(cells) < len(header_cells):
                        cells.extend([""] * (len(header_cells) - len(cells)))
                    elif len(cells) > len(header_cells):
                        cells = cells[: len(header_cells)]
                    md_lines.append("| " + " | ".join(cells) + " |")
            else:
                for row in rows:
                    row_cells = _table_row_cells(row)
                    cells = [
                        _convert_cell_to_md(
                            cell,
                            list_state=list_state,
                            italic_non_black=italic_non_black,
                        )
                        for cell in row_cells
                    ]
                    md_lines.append(" | ".join(cells))

        else:
            md_lines.append("<!-- Unsupported block element -->")

        if md_lines and md_lines[-1] != "":
            md_lines.append("")

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
    parser.add_argument(
        "--italic-non-black",
        action="store_true",
        default=False,
        help=(
            "Wrap any run whose text color is not black (and not the default/auto color) "
            "in Markdown italic markers (_..._).  Black is defined as explicit #000000 or "
            "the document default (auto) color."
        ),
    )

    args = parser.parse_args(argv)

    input_path = Path(args.input).expanduser().resolve()
    if args.output:
        output_path = Path(args.output).expanduser().resolve()
    else:
        output_path = input_path.with_suffix(".md")

    try:
        md = docx_to_markdown(input_path, italic_non_black=args.italic_non_black)
    except Exception as exc:  # pragma: no cover - CLI error path
        sys.stderr.write(f"Error: {exc}\n")
        return 1

    output_path.write_text(md, encoding="utf-8")
    print(f"Wrote Markdown to: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())