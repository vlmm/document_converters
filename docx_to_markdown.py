import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path

try:
    from docx import Document  # type: ignore
except ImportError:  # pragma: no cover - clear message if dependency missing
    Document = None  # type: ignore

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


def _paragraph_contains_page_break(p) -> bool:
    try:
        return bool(p._p.xpath('.//w:br[@w:type="page"]'))
    except Exception:
        return False


def _runs_to_markdown_text(runs) -> str:
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
        target_italic = bool(getattr(r, "italic", False))
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


@dataclass
class _ListFormat:
    kind: str  # "decimal" | "upperRoman" | "lowerRoman" | "upperLetter" | "lowerLetter" | "bullet" | "unknown"
    lvl_text: str  # e.g. "%1." or "%1)"
    start: int = 1


class _ListState:
    def __init__(self, formats: dict[tuple[int, int], _ListFormat]):
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


def _extract_numpr(p) -> tuple[bool, int, int]:
    """Return (has_numpr, numId, ilvl)."""
    try:
        pPr = p._p.pPr
        if pPr is None or pPr.numPr is None:
            return False, 0, 0
        numPr = pPr.numPr
        ilvl_elem = numPr.ilvl
        ilvl_val = int(ilvl_elem.val) if ilvl_elem is not None and ilvl_elem.val is not None else 0
        numId_elem = numPr.numId
        num_id_val = int(numId_elem.val) if numId_elem is not None and numId_elem.val is not None else 0
        if num_id_val <= 0:
            return False, 0, 0
        return True, num_id_val, ilvl_val
    except Exception:
        return False, 0, 0


def _get_list_info(p) -> tuple[bool, int, bool, int, int]:
    """
    Return (is_list, nesting_level, is_numbered_style, numId, ilvl)
    """
    style_name = getattr(getattr(p, "style", None), "name", "") or ""
    style_lower = style_name.lower()
    is_numbered_style = "list number" in style_lower
    is_bullet_style = "list bullet" in style_lower

    has_numpr, num_id, ilvl = _extract_numpr(p)
    if has_numpr:
        return True, ilvl, is_numbered_style, num_id, ilvl

    if is_bullet_style or is_numbered_style:
        level = _indent_level_from_paragraph(p)
        return True, level, is_numbered_style, 0, level

    return False, 0, False, 0, 0


def _build_list_formats(doc) -> dict[tuple[int, int], _ListFormat]:
    """
    Best-effort parse of numbering definitions for list marker formats.
    """
    formats: dict[tuple[int, int], _ListFormat] = {}
    try:
        numbering = doc.part.numbering_part.element  # CT_Numbering
    except Exception:
        return formats

    # Map abstractNumId -> {ilvl -> lvl element}
    abstract_lvls: dict[int, dict[int, object]] = {}
    for abs_el in numbering.xpath("./w:abstractNum"):
        try:
            abs_id = int(abs_el.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId"))
        except Exception:
            continue
        lvl_map: dict[int, object] = {}
        for lvl in abs_el.xpath("./w:lvl"):
            try:
                ilvl = int(lvl.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl"))
            except Exception:
                continue
            lvl_map[ilvl] = lvl
        abstract_lvls[abs_id] = lvl_map

    # numId -> abstractNumId
    num_to_abs: dict[int, int] = {}
    # numId overrides: (numId, ilvl) -> startOverride
    overrides_start: dict[tuple[int, int], int] = {}

    for num in numbering.xpath("./w:num"):
        try:
            num_id = int(num.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId"))
        except Exception:
            continue
        abs_ref = num.xpath("./w:abstractNumId")
        if abs_ref:
            try:
                abs_id = int(abs_ref[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"))
                num_to_abs[num_id] = abs_id
            except Exception:
                pass

        for lvl_override in num.xpath("./w:lvlOverride"):
            try:
                ilvl = int(lvl_override.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ilvl"))
            except Exception:
                continue
            start_ov = lvl_override.xpath("./w:startOverride")
            if start_ov:
                try:
                    start_val = int(start_ov[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"))
                    overrides_start[(num_id, ilvl)] = start_val
                except Exception:
                    pass

    def _kind_from_numfmt(numfmt: str) -> str:
        m = (numfmt or "").lower()
        if m in ("decimal",):
            return "decimal"
        if m in ("upperroman",):
            return "upperRoman"
        if m in ("lowerroman",):
            return "lowerRoman"
        if m in ("upperletter", "upperalpha"):
            return "upperLetter"
        if m in ("lowerletter", "loweralpha"):
            return "lowerLetter"
        if m in ("bullet",):
            return "bullet"
        return "unknown"

    # Build (numId, ilvl) -> ListFormat
    for num_id, abs_id in num_to_abs.items():
        lvl_map = abstract_lvls.get(abs_id, {})
        for ilvl, lvl in lvl_map.items():
            numfmt_el = lvl.xpath("./w:numFmt")
            lvltext_el = lvl.xpath("./w:lvlText")
            start_el = lvl.xpath("./w:start")

            numfmt = numfmt_el[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if numfmt_el else ""
            lvltext = lvltext_el[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val") if lvltext_el else "%1."
            start = 1
            if start_el:
                try:
                    start = int(start_el[0].get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"))
                except Exception:
                    start = 1

            if (num_id, ilvl) in overrides_start:
                start = overrides_start[(num_id, ilvl)]

            formats[(num_id, ilvl)] = _ListFormat(kind=_kind_from_numfmt(numfmt), lvl_text=lvltext, start=start)

    return formats


def _paragraph_to_md_line(p, list_state: _ListState | None = None, in_table: bool = False) -> str:
    if _paragraph_contains_page_break(p):
        return _PAGE_BREAK_MARKER

    text = _runs_to_markdown_text(p.runs)
    text = _normalize_marker_whitespace(text, "**")
    text = _normalize_marker_whitespace(text, "~~")

    if not text or not text.strip():
        return ""

    is_letter_clause = bool(re.match(r"^\s*\*{0,2}\(\s*[a-zA-Z]\s*\)\*{0,2}", text))

    is_list, level, is_numbered_style, num_id, ilvl = _get_list_info(p)

    if is_list and not is_letter_clause:
        indent = "  " * level

        # Bullet list (or unknown without numId)
        if num_id <= 0:
            marker = "-" if not is_numbered_style else "1."
            return f"{indent}{marker} {text}"

        # Numbered/bulleted list from numbering definitions
        marker = "1."
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


def _convert_cell_to_md(cell, list_state: _ListState | None = None) -> str:
    lines: list[str] = []
    for p in cell.paragraphs:
        line = _paragraph_to_md_line(p, list_state=list_state, in_table=True)
        if line == "":
            lines.append("")
        else:
            lines.append(line)

    if not lines:
        return ""

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

    processed = [s.replace("\n", "<br>") for s in processed]

    return "<br>".join(processed)


def _table_row_cells(row) -> list:
    try:
        tcs = row._tr.xpath("./w:tc")
        if tcs is not None:
            return [row.table._cell(tc, row._tr) for tc in tcs]
    except Exception:
        pass
    return list(row.cells)


def docx_to_markdown(input_path: Path) -> str:
    if Document is None:
        raise RuntimeError(
            "Missing dependency: python-docx.\n"
            "Install it with: pip install python-docx"
        )

    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    doc = Document(str(input_path))

    list_formats = _build_list_formats(doc)
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
            line = _paragraph_to_md_line(block, list_state=list_state)
            if not line:
                md_lines.append("")
                continue
            md_lines.append(line)

        elif isinstance(block, Table):
            rows = list(block.rows)
            if not rows:
                continue

            header_row_cells = _table_row_cells(rows[0])
            header_cells = [_convert_cell_to_md(cell, list_state=list_state) for cell in header_row_cells]
            if any(header_cells):
                md_lines.append("| " + " | ".join(header_cells) + " |")
                md_lines.append("| " + " | ".join("---" for _ in header_cells) + " |")
                for row in rows[1:]:
                    row_cells = _table_row_cells(row)
                    cells = [_convert_cell_to_md(cell, list_state=list_state) for cell in row_cells]
                    if len(cells) < len(header_cells):
                        cells.extend([""] * (len(header_cells) - len(cells)))
                    elif len(cells) > len(header_cells):
                        cells = cells[: len(header_cells)]
                    md_lines.append("| " + " | ".join(cells) + " |")
            else:
                for row in rows:
                    row_cells = _table_row_cells(row)
                    cells = [_convert_cell_to_md(cell, list_state=list_state) for cell in row_cells]
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
