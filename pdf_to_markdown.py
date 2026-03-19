import fitz  # PyMuPDF
import pypdf
import pytesseract
from PIL import Image
import argparse
import os


def clean_text(text):
    """Normalize whitespace and remove blank lines."""
    return '\n'.join([line.strip() for line in text.split('\n') if line.strip()])


def pdf_to_markdown(input_file, output_file=None, ocr_mode='auto', ocr_lang='bul', split='none', text_mode='clean', dpi=300):
    """Convert PDF to Markdown with OCR."""
    # Implement PDF processing and OCR detection here
    # Read PDF and process
    doc = fitz.open(input_file)
    markdown = ''
    # Process each page
    for page in doc:
        # OCR and image detection logic goes here
        
    # Handle output
    if output_file:
        with open(output_file, 'w') as f:
            f.write(markdown)
    else:
        print(markdown)


def main():
    parser = argparse.ArgumentParser(description='Convert PDF to Markdown')
    parser.add_argument('input', type=str, help='Input PDF file')
    parser.add_argument('--output', type=str, help='Output Markdown file')
    parser.add_argument('--ocr-mode', choices=['auto', 'no_text', 'images', 'images_or_no_text'], default='auto', help='OCR mode')
    parser.add_argument('--ocr-lang', default='bul', help='OCR language')
    parser.add_argument('--split', choices=['none', '2', '4'], default='none', help='Split mode')
    parser.add_argument('--text-mode', choices=['clean', 'raw'], default='clean', help='Text mode')
    parser.add_argument('--dpi', type=int, default=300, help='DPI for image conversion')
    args = parser.parse_args()
    
    pdf_to_markdown(args.input, args.output, args.ocr_mode, args.ocr_lang, args.split, args.text_mode, args.dpi)


if __name__ == '__main__':
    main()