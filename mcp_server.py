"""
MCP Server for document_converters (FastMCP)

Metadata (JSON) — машинно-четим, AI може да обновява:
{
  "app_name": "document-converters-mcp",
  "description": "MCP server exposing all document converter functions: DOCX/PPTX/PDF/Markdown/JSON/XLSX conversions.",
  "version": "1.0",
  "author": "auto-generated",
  "tool_contracts": [
    {
      "name": "ping",
      "description": "Health check.",
      "inputs": {},
      "output": {"status": "string", "time": "string"}
    },
    {
      "name": "docx_to_markdown",
      "description": "Convert a .docx file to Markdown text.",
      "inputs": {"input_path": "string", "italic_non_black": "bool"},
      "output": {"status": "string", "markdown": "string"}
    },
    {
      "name": "pptx_to_markdown",
      "description": "Convert a .pptx file to Markdown text.",
      "inputs": {"input_path": "string"},
      "output": {"status": "string", "markdown": "string"}
    },
    {
      "name": "pdf_to_markdown",
      "description": "Convert a .pdf file to Markdown text with optional OCR.",
      "inputs": {
        "input_path": "string",
        "ocr_mode": "string",
        "ocr_lang": "string",
        "split": "string",
        "text_mode": "string",
        "dpi": "int",
        "layout_mode": "string",
        "table_mode": "string",
        "extract_images": "bool",
        "assets_dir": "string"
      },
      "output": {"status": "string", "markdown": "string"}
    },
    {
      "name": "markdown_to_docx",
      "description": "Convert Markdown text to a .docx file.",
      "inputs": {
        "markdown_text": "string",
        "output_path": "string",
        "table_borders": "bool",
        "table_border_color": "string"
      },
      "output": {"status": "string", "output_path": "string"}
    },
    {
      "name": "json_to_docx",
      "description": "Convert a JSON document structure to a .docx file.",
      "inputs": {"data": "object", "output_path": "string"},
      "output": {"status": "string", "output_path": "string"}
    },
    {
      "name": "markdown_to_xlsx",
      "description": "Convert a Markdown file to an .xlsx file.",
      "inputs": {
        "input_path": "string",
        "output_path": "string",
        "width_cols": "int"
      },
      "output": {"status": "string", "output_path": "string"}
    }
  ]
}
"""

import logging
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Optional

from fastmcp import FastMCP

# -------------------------
# App metadata (AI can edit)
# -------------------------
APP_META = {
    "app_name": os.getenv("MCP_APP_NAME", "document-converters-mcp"),
    "description": "MCP server exposing all document converter functions",
    "version": "1.0",
    "author": os.getenv("MCP_APP_AUTHOR", "auto-generated"),
    "tool_contracts": [
        {
            "name": "ping",
            "description": "Health check.",
            "inputs": {},
            "output": {"status": "string", "time": "string"},
        },
        {
            "name": "docx_to_markdown",
            "description": "Convert a .docx file to Markdown text.",
            "inputs": {"input_path": "string", "italic_non_black": "bool"},
            "output": {"status": "string", "markdown": "string"},
        },
        {
            "name": "pptx_to_markdown",
            "description": "Convert a .pptx file to Markdown text.",
            "inputs": {"input_path": "string"},
            "output": {"status": "string", "markdown": "string"},
        },
        {
            "name": "pdf_to_markdown",
            "description": "Convert a .pdf file to Markdown text with optional OCR.",
            "inputs": {
                "input_path": "string",
                "ocr_mode": "string",
                "ocr_lang": "string",
                "split": "string",
                "text_mode": "string",
                "dpi": "int",
                "layout_mode": "string",
                "table_mode": "string",
                "extract_images": "bool",
                "assets_dir": "string",
            },
            "output": {"status": "string", "markdown": "string"},
        },
        {
            "name": "markdown_to_docx",
            "description": "Convert Markdown text to a .docx file.",
            "inputs": {
                "markdown_text": "string",
                "output_path": "string",
                "table_borders": "bool",
                "table_border_color": "string",
            },
            "output": {"status": "string", "output_path": "string"},
        },
        {
            "name": "json_to_docx",
            "description": "Convert a JSON document structure to a .docx file.",
            "inputs": {"data": "object", "output_path": "string"},
            "output": {"status": "string", "output_path": "string"},
        },
        {
            "name": "markdown_to_xlsx",
            "description": "Convert a Markdown file to an .xlsx file.",
            "inputs": {
                "input_path": "string",
                "output_path": "string",
                "width_cols": "int",
            },
            "output": {"status": "string", "output_path": "string"},
        },
    ],
}

# -------------------------
# Init
# -------------------------
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(APP_META["app_name"])
mcp = FastMCP(APP_META["app_name"])


# -------------------------
# Utilities
# -------------------------
def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


# -------------------------
# Tools
# -------------------------
@mcp.tool
def ping() -> Dict[str, Any]:
    """Health check; returns status and current UTC time."""
    return {"status": "ok", "time": now_iso()}


@mcp.tool
def docx_to_markdown(
    input_path: str,
    italic_non_black: bool = False,
) -> Dict[str, Any]:
    """Convert a .docx file to Markdown text.

    Parameters
    ----------
    input_path:
        Absolute or relative path to the input .docx file.
    italic_non_black:
        When True, wrap any run whose text color is not black in Markdown
        italic markers (_..._).
    """
    try:
        from docx_to_markdown import docx_to_markdown as _convert  # type: ignore
    except ImportError as exc:
        return {"status": "error", "message": f"Import error: {exc}"}

    path = Path(input_path).expanduser().resolve()
    try:
        md = _convert(path, italic_non_black=italic_non_black)
        return {"status": "ok", "markdown": md}
    except FileNotFoundError as exc:
        return {"status": "error", "message": str(exc)}
    except Exception as exc:
        logger.exception("docx_to_markdown failed for %s", input_path)
        return {"status": "error", "message": str(exc)}


@mcp.tool
def pptx_to_markdown(input_path: str) -> Dict[str, Any]:
    """Convert a .pptx file to Markdown text.

    Parameters
    ----------
    input_path:
        Absolute or relative path to the input .pptx file.
    """
    try:
        from pptx_to_markdown import pptx_to_markdown as _convert  # type: ignore
    except ImportError as exc:
        return {"status": "error", "message": f"Import error: {exc}"}

    path = Path(input_path).expanduser().resolve()
    try:
        md = _convert(path)
        return {"status": "ok", "markdown": md}
    except FileNotFoundError as exc:
        return {"status": "error", "message": str(exc)}
    except Exception as exc:
        logger.exception("pptx_to_markdown failed for %s", input_path)
        return {"status": "error", "message": str(exc)}


@mcp.tool
def pdf_to_markdown(
    input_path: str,
    ocr_mode: str = "auto",
    ocr_lang: str = "bul",
    split: str = "none",
    text_mode: str = "clean",
    dpi: int = 300,
    layout_mode: str = "auto",
    table_mode: str = "auto",
    extract_images: bool = True,
    assets_dir: Optional[str] = None,
) -> Dict[str, Any]:
    """Convert a .pdf file to Markdown text.

    Parameters
    ----------
    input_path:
        Absolute or relative path to the input .pdf file.
    ocr_mode:
        Controls when to run OCR. One of: ``"auto"`` (default),
        ``"no_text"``, ``"images"``, ``"images_or_no_text"``.
    ocr_lang:
        Tesseract language code (default ``"bul"``).
    split:
        Page split mode. One of: ``"none"`` (default), ``"2"``, ``"4"``.
    text_mode:
        Text cleanup mode. One of: ``"clean"`` (default), ``"raw"``.
    dpi:
        Rasterisation DPI for OCR (default 300).
    layout_mode:
        Layout extraction mode. One of: ``"auto"`` (default), ``"pymupdf"``,
        ``"ocr"``.
    table_mode:
        Table output format. One of: ``"auto"`` (default), ``"html"``,
        ``"markdown"``, ``"off"``.
    extract_images:
        When True (default) embedded images are exported to *assets_dir*.
    assets_dir:
        Directory for extracted images. Defaults to ``<input_stem>_assets/``
        next to the input file when omitted.
    """
    try:
        from pdf_to_markdown import pdf_to_markdown as _convert  # type: ignore
    except ImportError as exc:
        return {"status": "error", "message": f"Import error: {exc}"}

    path = Path(input_path).expanduser().resolve()
    assets = Path(assets_dir).expanduser().resolve() if assets_dir else None
    try:
        md = _convert(
            path,
            ocr_mode=ocr_mode,
            ocr_lang=ocr_lang,
            split=split,
            text_mode=text_mode,
            dpi=dpi,
            layout_mode=layout_mode,
            table_mode=table_mode,
            extract_images=extract_images,
            assets_dir=assets,
        )
        return {"status": "ok", "markdown": md}
    except FileNotFoundError as exc:
        return {"status": "error", "message": str(exc)}
    except Exception as exc:
        logger.exception("pdf_to_markdown failed for %s", input_path)
        return {"status": "error", "message": str(exc)}


@mcp.tool
def markdown_to_docx(
    markdown_text: str,
    output_path: str,
    table_borders: bool = True,
    table_border_color: str = "#000000",
) -> Dict[str, Any]:
    """Convert Markdown text to a .docx file.

    Parameters
    ----------
    markdown_text:
        Markdown-formatted string to convert.
    output_path:
        Absolute or relative path for the output .docx file.
    table_borders:
        When True (default) tables are rendered with visible borders.
    table_border_color:
        Hex color string for table borders (default ``"#000000"``).
    """
    try:
        from markdown_to_docx import markdown_to_docx as _convert  # type: ignore
    except ImportError as exc:
        return {"status": "error", "message": f"Import error: {exc}"}

    out = Path(output_path).expanduser().resolve()
    try:
        _convert(
            markdown_text,
            str(out),
            table_borders=table_borders,
            table_border_color=table_border_color,
        )
        return {"status": "ok", "output_path": str(out)}
    except Exception as exc:
        logger.exception("markdown_to_docx failed, output=%s", output_path)
        return {"status": "error", "message": str(exc)}


@mcp.tool
def json_to_docx(data: Dict[str, Any], output_path: str) -> Dict[str, Any]:
    """Convert a JSON document structure to a .docx file.

    Parameters
    ----------
    data:
        JSON document object.  Expected shape::

            {
              "document": {
                "sections": [
                  {"type": "paragraph", "text": "...", "style": {...}},
                  {"type": "table", "rows_content": [[...], ...], ...},
                  {"type": "image", "description": "..."}
                ]
              }
            }

    output_path:
        Absolute or relative path for the output .docx file.
    """
    try:
        from json_to_docx import create_docx_from_json  # type: ignore
    except ImportError as exc:
        return {"status": "error", "message": f"Import error: {exc}"}

    out = Path(output_path).expanduser().resolve()
    try:
        doc = create_docx_from_json(data)
        doc.save(str(out))
        return {"status": "ok", "output_path": str(out)}
    except Exception as exc:
        logger.exception("json_to_docx failed, output=%s", output_path)
        return {"status": "error", "message": str(exc)}


@mcp.tool
def markdown_to_xlsx(
    input_path: str,
    output_path: Optional[str] = None,
    width_cols: int = 4,
) -> Dict[str, Any]:
    """Convert a Markdown file to an .xlsx file.

    Parameters
    ----------
    input_path:
        Absolute or relative path to the input .md file.
    output_path:
        Absolute or relative path for the output .xlsx file.  Defaults to
        the same name as *input_path* with a ``.xlsx`` extension.
    width_cols:
        Number of merged columns for headings and paragraph blocks
        (default 4).  Table columns are not affected.
    """
    try:
        from markdown_to_xlsx import MarkdownToXlsx  # type: ignore
    except ImportError as exc:
        return {"status": "error", "message": f"Import error: {exc}"}

    try:
        conv = MarkdownToXlsx(width_cols=width_cols)
        out = conv.convert_file(input_path, output_path)
        return {"status": "ok", "output_path": str(out)}
    except FileNotFoundError as exc:
        return {"status": "error", "message": str(exc)}
    except Exception as exc:
        logger.exception("markdown_to_xlsx failed for %s", input_path)
        return {"status": "error", "message": str(exc)}


# -------------------------
# Main
# -------------------------
if __name__ == "__main__":
    logger.info(
        "%s starting (v%s) at %s",
        APP_META["app_name"],
        APP_META["version"],
        now_iso(),
    )
    mcp.run()
