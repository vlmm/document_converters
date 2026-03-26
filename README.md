# document_converters

A small collection of Python scripts for converting between common office/document formats.

The repository currently includes standalone modules (scripts) that can be used either as:
- importable Python functions, or
- command-line tools by running the module directly (e.g. `python docx_to_markdown.py input.docx`).

## Modules / Tools

### `docx_to_markdown.py` (DOCX → Markdown)

Converts a Word `.docx` file to Markdown.

**Features**
- Headings: Word paragraph styles like *Heading 1..6* → `#..######`
- Lists: common *List Bullet* / *List Number* styles → `-` / `1.`
- Inline formatting: **bold**, _italic_, <u>underline</u>, and ~~strikethrough~~ (based on run formatting)
- Tables: converted into Markdown tables (first row used as header when present)
- Other/unsupported blocks: emitted as placeholders

**CLI usage**
```bash
python docx_to_markdown.py input.docx [output.md]
```

### `pptx_to_markdown.py` (PPTX → Markdown)

Converts a PowerPoint `.pptx` presentation to Markdown.

**Features**
- Slide titles → `# Heading`
- Slide subtitles → `## Heading`
- Body text → bullet lists (`- ...`) with indentation based on paragraph level
- Tables → Markdown tables
- Images → placeholders like `![name](name)`
- Speaker notes → blockquote section (`> **Notes:** ...`)
- Slides are separated by a horizontal rule (`---`)

**CLI usage**
```bash
python pptx_to_markdown.py input.pptx [output.md]
```

### `markdown_to_docx.py` (Markdown → DOCX)

Converts Markdown text (or a Markdown file) into a Word `.docx` document.

**Features**
- Headings: `#`..`####`
- Lists: `-` / `*` and ordered lists like `1.`
- Inline formatting: `**bold**`, `*italic*`/`_italic_`, `~~strikethrough~~`, and `` `inline code` ``
- Code blocks: fenced blocks (``` ... ```) are added as monospace paragraphs

**CLI usage**
```bash
python markdown_to_docx.py input.md -o output.docx
```

### `pdf_to_markdown.py` (PDF → Markdown with optional OCR)

Converts a `.pdf` to Markdown. For digital PDFs the converter uses
PyMuPDF's layout-aware extraction to preserve structure and formatting.
For scanned pages it falls back to Tesseract OCR.

**Features**
- **Headings** – detected from font-size heuristics; emitted as `#`–`#####`
- **Lists** – unordered (`-`) and ordered (`1.`) lists detected via bullet/number prefixes
- **Inline formatting** – **bold** and *italic* from PDF font flags/names; underline `<u>` and ~~strikethrough~~ from vector line detection
- **Tables** – detected via PyMuPDF's `find_tables()` (PyMuPDF ≥ 1.23); emitted as HTML `<table>` by default (most faithful) or as GFM Markdown tables
- **Multi-column layouts** – columns are detected by clustering block X-positions; left column is read before right column
- **Embedded images** – exported to an `<stem>_assets/` directory and linked from the Markdown output
- **OCR fallback** – pages without extractable text are rasterised and processed by Tesseract with word-box reconstruction (better list/structure detection than plain `image_to_string`)

**Dependencies**
```
pip install PyMuPDF pypdf            # required
pip install pytesseract Pillow       # required for OCR fallback
# also install the Tesseract binary and language data, e.g.:
# sudo apt install tesseract-ocr tesseract-ocr-bul
```

**CLI usage**
```bash
python pdf_to_markdown.py input.pdf [output.md] \
  [--layout-mode auto|pymupdf|ocr] \
  [--table-mode auto|html|markdown|off] \
  [--extract-images | --no-extract-images] \
  [--assets-dir PATH] \
  [--ocr-mode auto|no_text|images|images_or_no_text] \
  [--ocr-lang bul] \
  [--split none|2|4] \
  [--text-mode clean|raw] \
  [--dpi 300]
```

**Arguments**
- `--layout-mode` (default: `auto`)
  - `auto` / `pymupdf`: use PyMuPDF layout extraction for digital PDFs; OCR for scanned pages
  - `ocr`: always use the OCR path
- `--table-mode` (default: `auto`)
  - `auto` / `html`: emit HTML `<table>` (best fidelity for complex tables)
  - `markdown`: emit GFM Markdown tables
  - `off`: skip table detection
- `--extract-images` / `--no-extract-images` (default: on)
  - When on, embedded images are exported and linked from the Markdown
- `--assets-dir PATH`
  - Directory for exported images (default: `<output_stem>_assets/` next to the output file)
- `--ocr-mode`
  - `auto` / `images_or_no_text`: OCR if the page has images *or* extracted text is empty
  - `no_text`: OCR only when extracted text is empty
  - `images`: OCR only when images are detected on the page
- `--split` (useful for scanned landscape pages containing two pages side-by-side)
  - `none`: do not split
  - `2`: split into left / right halves
  - `4`: split into quadrants (top-left, top-right, bottom-left, bottom-right)
- `--text-mode`
  - `clean`: trims lines and collapses blank lines
  - `raw`: keeps text closer to the original extraction/OCR output

## Development

Tests are located under `tests/`.

```bash
python -m unittest
```
