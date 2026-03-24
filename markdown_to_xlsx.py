"""
Markdown -> XLSX (real formatting + tables).

Rules:
- width-cols controls ONLY merged text blocks (headings + paragraphs).
- Tables are NOT limited by width-cols: every markdown table row becomes one excel row with adjacent cells.

Inline formatting supports combinations/nesting:
  **bold**, *italic* / _italic_, ~~strike~~, <u>underline</u>, `code`
Examples:
  ## *test*
  **bold _italic_**
  <u>**bold under**</u>

Dependency: XlsxWriter
Install: pip install XlsxWriter
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

try:
    import xlsxwriter  # type: ignore
except ImportError:  # pragma: no cover
    xlsxwriter = None  # type: ignore


_TABLE_SEPARATOR_RE = re.compile(
    r"^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$"
)

def is_table_separator_line(line: str) -> bool:
    return bool(_TABLE_SEPARATOR_RE.match(line))

def looks_like_table_row(line: str) -> bool:
    s = line.strip()
    if "|" not in s:
        return False
    if is_table_separator_line(s):
        return True
    if s.startswith("|") or s.endswith("|"):
        return True
    return s.count("|") >= 2

def split_table_row(line: str) -> list[str]:
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    return [c.strip() for c in s.split("|")]


_HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")

def parse_heading(line: str) -> tuple[int, str] | None:
    m = _HEADING_RE.match(line.rstrip())
    if not m:
        return None
    return len(m.group(1)), m.group(2).strip()


def parse_inline_runs(text: str) -> list[tuple[str, frozenset[str]]]:
    """
    Toggle-based inline parser supporting combined/nested styles.
    Tokens:
      ** bold, * italic, _ italic, ~~ strike, ` code, <u> underline
    """
    runs: list[tuple[str, frozenset[str]]] = []
    buf: list[str] = []
    style: set[str] = set()

    def flush():
        if not buf:
            return
        s = "".join(buf)
        if s:
            runs.append((s, frozenset(style)))
        buf.clear()

    i = 0
    while i < len(text):
        if text.startswith("<u>", i):
            flush(); style.add("underline"); i += 3; continue
        if text.startswith("</u>", i):
            flush(); style.discard("underline"); i += 4; continue

        if text.startswith("**", i):
            flush()
            if "bold" in style: style.remove("bold")
            else: style.add("bold")
            i += 2
            continue

        if text.startswith("~~", i):
            flush()
            if "strike" in style: style.remove("strike")
            else: style.add("strike")
            i += 2
            continue

        if text.startswith("`", i):
            flush()
            if "code" in style: style.remove("code")
            else: style.add("code")
            i += 1
            continue

        if text.startswith("*", i):
            flush()
            if "italic" in style: style.remove("italic")
            else: style.add("italic")
            i += 1
            continue

        if text.startswith("_", i):
            flush()
            if "italic" in style: style.remove("italic")
            else: style.add("italic")
            i += 1
            continue

        buf.append(text[i])
        i += 1

    flush()

    merged: list[tuple[str, frozenset[str]]] = []
    for t, st in runs:
        if not t:
            continue
        if merged and merged[-1][1] == st:
            merged[-1] = (merged[-1][0] + t, st)
        else:
            merged.append((t, st))
    return merged


@dataclass
class Block:
    kind: str  # heading | paragraph | table
    lines: list[str]
    heading_level: int = 0


def parse_blocks(md: str) -> list[Block]:
    lines = md.splitlines()
    blocks: list[Block] = []
    cur_para: list[str] = []
    cur_table: list[str] = []
    in_table = False

    def flush_para():
        nonlocal cur_para
        if not cur_para:
            return
        joined = " ".join(l.strip() for l in cur_para if l.strip()).strip()
        if joined:
            blocks.append(Block(kind="paragraph", lines=[joined]))
        cur_para = []

    def flush_table():
        nonlocal cur_table, in_table
        if cur_table:
            blocks.append(Block(kind="table", lines=cur_table[:]))
        cur_table = []
        in_table = False

    for line in lines:
        h = parse_heading(line)
        if h is not None and not in_table:
            flush_para()
            lvl, text = h
            blocks.append(Block(kind="heading", lines=[text], heading_level=lvl))
            continue

        if looks_like_table_row(line):
            flush_para()
            in_table = True
            cur_table.append(line)
            continue

        if in_table:
            flush_table()

        if not line.strip():
            flush_para()
        else:
            cur_para.append(line)

    if in_table:
        flush_table()
    flush_para()
    return blocks


class MarkdownToXlsx:
    def __init__(self, width_cols: int = 4):
        if xlsxwriter is None:
            raise RuntimeError("Missing dependency: XlsxWriter. Install with: pip install XlsxWriter")
        self.width_cols = max(1, int(width_cols))
        self._format_cache: dict[tuple, object] = {}

    def convert_file(self, input_file: str, output_file: Optional[str] = None) -> Path:
        in_path = Path(input_file).expanduser().resolve()
        if not in_path.exists():
            raise FileNotFoundError(f"Input file not found: {in_path}")

        out_path = Path(output_file).expanduser().resolve() if output_file else in_path.with_suffix(".xlsx")
        md = in_path.read_text(encoding="utf-8")
        self.convert(md, out_path)
        return out_path

    def _get_inline_format(self, wb, style: frozenset[str], *, base_font_size: int = 11, base_bold: bool = False):
        key = (tuple(sorted(style)), base_font_size, base_bold)
        if key in self._format_cache:
            return self._format_cache[key]

        is_code = "code" in style
        props = {
            "font_name": "Courier New" if is_code else "Calibri",
            "font_size": 10 if is_code else base_font_size,
        }
        if is_code:
            props["font_color"] = "#C00000"

        props["bold"] = base_bold or ("bold" in style)
        if "italic" in style:
            props["italic"] = True
        if "underline" in style:
            props["underline"] = True
        if "strike" in style:
            props["font_strikeout"] = True

        fmt = wb.add_format(props)
        self._format_cache[key] = fmt
        return fmt

    def _runs_to_plain_or_rich(self, wb, runs: list[tuple[str, frozenset[str]]], *, base_font_size: int = 11, base_bold: bool = False):
        if not runs:
            return "", None
        has_style = any(bool(st) for _, st in runs)
        plain = "".join(t for t, _ in runs)
        if not has_style:
            return plain, None

        parts: list = []
        for t, st in runs:
            if not t:
                continue
            if st:
                parts.extend([self._get_inline_format(wb, st, base_font_size=base_font_size, base_bold=base_bold), t])
            else:
                parts.append(t)

        if not parts:
            return plain, None
        if not isinstance(parts[-1], str):
            parts.append("")
        if len(parts) < 3:
            return plain, None
        return None, parts

    def _write_cell_inline(self, wb, ws, row: int, col: int, text: str, base_fmt, *, base_font_size: int = 11, base_bold: bool = False):
        runs = parse_inline_runs(text)
        plain, rich = self._runs_to_plain_or_rich(wb, runs, base_font_size=base_font_size, base_bold=base_bold)
        if rich is None:
            ws.write(row, col, plain, base_fmt)
        else:
            ws.write_rich_string(row, col, *rich, base_fmt)

    def convert(self, md: str, out_path: Path) -> None:
        wb = xlsxwriter.Workbook(str(out_path))
        ws = wb.add_worksheet("Sheet1")

        fmt_para = wb.add_format({"text_wrap": True, "valign": "top", "font_name": "Calibri", "font_size": 11})
        fmt_table = wb.add_format({"text_wrap": True, "valign": "top", "font_name": "Calibri", "font_size": 11, "border": 1})
        fmt_table_header = wb.add_format({"text_wrap": True, "valign": "top", "font_name": "Calibri", "font_size": 11, "bold": True, "border": 1, "bg_color": "#F2F2F2"})

        # set widths for text area columns
        for c in range(self.width_cols):
            ws.set_column(c, c, 12)

        heading_sizes = {1: 18, 2: 16, 3: 14, 4: 13, 5: 12, 6: 11}
        heading_base = {
            lvl: wb.add_format({"bold": True, "font_name": "Calibri", "font_size": heading_sizes[lvl], "text_wrap": True, "valign": "vcenter"})
            for lvl in range(1, 7)
        }

        row = 0
        max_col_used = self.width_cols - 1

        for block in parse_blocks(md):
            if block.kind == "heading":
                lvl = max(1, min(6, block.heading_level))
                text = block.lines[0].strip()

                ws.merge_range(row, 0, row, self.width_cols - 1, "", heading_base[lvl])
                self._write_cell_inline(wb, ws, row, 0, text, heading_base[lvl], base_font_size=heading_sizes[lvl], base_bold=True)
                ws.set_row(row, 24)
                row += 2
                continue

            if block.kind == "paragraph":
                text = block.lines[0].strip()
                ws.merge_range(row, 0, row, self.width_cols - 1, "", fmt_para)
                self._write_cell_inline(wb, ws, row, 0, text, fmt_para)
                ws.set_row(row, 40)
                row += 2
                continue

            if block.kind == "table":
                table_rows: list[list[str]] = []
                for ln in block.lines:
                    if is_table_separator_line(ln):
                        continue
                    table_rows.append(split_table_row(ln))

                if not table_rows:
                    continue

                for r_idx, cells in enumerate(table_rows):
                    base_fmt = fmt_table_header if r_idx == 0 else fmt_table
                    for c_idx, cell_text in enumerate(cells):
                        self._write_cell_inline(wb, ws, row, c_idx, cell_text, base_fmt, base_bold=(r_idx == 0))
                        if c_idx > max_col_used:
                            max_col_used = c_idx
                    ws.set_row(row, 22)
                    row += 1

                row += 1
                continue

        # make extra table columns readable
        for c in range(self.width_cols, max_col_used + 1):
            ws.set_column(c, c, 14)

        wb.close()


def main(argv: Optional[list[str]] = None) -> int:
    p = argparse.ArgumentParser(description="Convert Markdown (.md) to XLSX.")
    p.add_argument("input", help="Input markdown file")
    p.add_argument("-o", "--output", help="Output .xlsx file (default: input name with .xlsx)")
    p.add_argument("--width-cols", type=int, default=4, help="Merged text width in columns (headings/paragraphs only). Default: 4")
    args = p.parse_args(argv)

    try:
        conv = MarkdownToXlsx(width_cols=args.width_cols)
        out = conv.convert_file(args.input, args.output)
    except Exception as e:
        sys.stderr.write(f"Error: {e}\n")
        return 1

    print(f"Wrote XLSX to: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
