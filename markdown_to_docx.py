# markdown_to_docx.py

"""
This module provides functionalities to convert Markdown files to DOCX format.
"""

import markdown
from docx import Document


def convert_markdown_to_docx(markdown_text, docx_filename):
    """
    Convert Markdown text to DOCX and save it to a file.
    
    :param markdown_text: Markdown formatted string
    :param docx_filename: Filename for the output DOCX file
    """
    # Convert Markdown to HTML
    html = markdown.markdown(markdown_text)
    
    # Create a new Document
    doc = Document()
    doc.add_heading('Converted Document', level=1)
    
    # Add HTML content to the DOCX file
    doc.add_paragraph(html)
    
    # Save the DOCX file
    doc.save(docx_filename)
    

if __name__ == '__main__':
    sample_markdown = "# Sample Markdown\nThis is a sample markdown content that will be converted to DOCX."
    convert_markdown_to_docx(sample_markdown, 'output.docx')
