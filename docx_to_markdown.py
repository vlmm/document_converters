import argparse
import re
import sys
from pathlib import Path

try:
    from docx import Document  # type: ignore
except ImportError:  # pragma: no cover - clear message if dependency missing
    Document = None  # type: ignore


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


def _runs_to_markdown_text(runs) -> str:
    """
    Convert a sequence of docx runs to Markdown inline formatting.
    Supports bold, italic, underline и зачертаване (strike-through).

    Използваме state machine, за да:
    - не слагаме отделни `**` за всеки run
    - не „чупим“ bold/italic върху интервали и табове
    """
    result: list[str] = []
    cur_bold = False
    cur_italic = False
    cur_underline = False
    cur_strike = False

    for r in runs:
        text = r.text or ""
        if not text:
            continue

        target_bold = bool(getattr(r, "bold", False))
        target_italic = bool(getattr(r, "italic", False))
        target_underline = bool(getattr(r, "underline", False))
        target_strike = bool(getattr(getattr(r, "font", None), "strike", False))

        # Затваряме маркери, които вече не са активни
        if cur_underline and not target_underline:
            result.append("</u>")
            cur_underline = False
        if cur_italic and not target_italic:
            result.append("_")
            cur_italic = False
        if cur_bold and not target_bold:
            result.append("**")
            cur_bold = False
        if cur_strike and not target_strike:
            result.append("~~")
            cur_strike = False

        # Отваряме нови маркери, които стават активни
        if target_bold and not cur_bold:
            result.append("**")
            cur_bold = True
        if target_italic and not cur_italic:
            result.append("_")
            cur_italic = True
        if target_underline and not cur_underline:
            result.append("<u>")
            cur_underline = True
        if target_strike and not cur_strike:
            result.append("~~")
            cur_strike = True

        result.append(text)

    # Затваряме всичко останало в края на параграфа
    if cur_underline:
        result.append("</u>")
    if cur_italic:
        result.append("_")
    if cur_bold:
        result.append("**")
    if cur_strike:
        result.append("~~")

    return "".join(result)


def _is_list_paragraph(p) -> bool:
    """
    Heuristic: treat paragraphs with 'List Bullet' or 'List Number' styles as lists.
    This does not inspect the underlying numbering XML, but works for common cases.
    """
    style_name = getattr(getattr(p, "style", None), "name", "") or ""
    style_name = style_name.lower()
    return "list bullet" in style_name or "list number" in style_name


def _normalize_marker_whitespace(text: str, marker: str) -> str:
    """
    Нормализира whitespace около даден Markdown маркер (напр. ** или ~~):
    - маха space/таб веднага след отварящия маркер
    - маха space/таб точно преди затварящия маркер
    - ако след затварящия маркер има символ, добавя интервал
    """
    m = re.escape(marker)
    # marker   textmarker -> markertextmarker
    text = re.sub(rf"{m}[ \t]+(\S)", rf"{marker}\1", text)
    # text   marker -> textmarker
    text = re.sub(rf"(\S)[ \t]+{m}", rf"\1{marker}", text)
    # markertextmarkerX -> markertextmarker X
    text = re.sub(rf"{m}([^{m}]+){m}(\S)", rf"{marker}\1{marker} \2", text)
    return text


def _list_marker(p) -> str:
    style_name = getattr(getattr(p, "style", None), "name", "") or ""
    style_name = style_name.lower()
    if "number" in style_name:
        return "1."
    return "-"  # default bullet


def docx_to_markdown(input_path: Path) -> str:
    """
    Convert a .docx file to a Markdown string.

    Notes:
    - Handles headings, bullet/numbered lists, and normal paragraphs.
    - Tables and images are emitted as simple placeholders.
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
    from docx.document import Document as _DocType  # type: ignore
    from docx.table import Table  # type: ignore
    from docx.text.paragraph import Paragraph  # type: ignore

    def iter_block_items(parent):
        """
        Yield each paragraph and table child within *parent*, in document order.
        Source: adapted from python-docx documentation examples.
        """
        parent_elm = parent.element.body
        for child in parent_elm.iterchildren():
            if child.tag.endswith("}p"):
                yield Paragraph(child, parent)
            elif child.tag.endswith("}tbl"):
                yield Table(child, parent)

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = _runs_to_markdown_text(block.runs)
            # Нормализираме whitespace около основните inline маркери
            text = _normalize_marker_whitespace(text, "**")
            text = _normalize_marker_whitespace(text, "~~")
            if not text:
                # Preserve blank line
                md_lines.append("")
                continue

            # Клауза от вида (a), (b) ... – да не я превръщаме в списък
            is_letter_clause = bool(
                re.match(r"^\s*\*{0,2}\(\s*[a-zA-Z]\s*\)\*{0,2}", text)
            )

            if _is_list_paragraph(block) and not is_letter_clause:
                marker = _list_marker(block)
                md_lines.append(f"{marker} {text}")
            else:
                level = _heading_level(getattr(getattr(block, "style", None), "name", None))
                if level > 0:
                    md_lines.append(f"{'#' * level} {text}")
                else:
                    md_lines.append(text)

        else:  # Table or other block
            if isinstance(block, Table):
                # Simple Markdown table conversion: first row as header if possible
                rows = list(block.rows)
                if not rows:
                    continue
                header_cells = [cell.text.strip() for cell in rows[0].cells]
                if any(header_cells):
                    md_lines.append("| " + " | ".join(header_cells) + " |")
                    md_lines.append(
                        "| " + " | ".join("---" for _ in header_cells) + " |"
                    )
                    for row in rows[1:]:
                        cells = [cell.text.strip() for cell in row.cells]
                        md_lines.append("| " + " | ".join(cells) + " |")
                else:
                    # Fallback: emit as plain text rows
                    for row in rows:
                        cells = [cell.text.strip() for cell in row.cells]
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
    parser = argparse.ArgumentParser(
        description="Convert a .docx file to Markdown."
    )
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

