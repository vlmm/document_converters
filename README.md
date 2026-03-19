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

Converts a `.pdf` to Markdown. It first tries to extract text from the PDF, and can optionally run OCR on pages depending on the selected mode.

> Note: `pdf_to_markdown.py` exists in the `pdf-to-markdown-ocr` branch and may not yet be merged into `main`.

**OCR support**
- Uses Tesseract via `pytesseract`
- Can work with Cyrillic as long as the appropriate Tesseract language data is installed (default `--ocr-lang bul`)

**CLI usage**
```bash
python pdf_to_markdown.py input.pdf [output.md] \n  --ocr-mode auto|no_text|images|images_or_no_text \n  --ocr-lang bul \n  --split none|2|4 \n  --text-mode clean|raw \n  --dpi 300
```

**Arguments**
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
