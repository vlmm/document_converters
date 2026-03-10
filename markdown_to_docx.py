"""
Python модул за конвертиране на Markdown в DOCX формат
Поддържа използване от командния ред
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import argparse
import os
from pathlib import Path
from typing import List, Tuple


class MarkdownToDocx:
    def __init__(self):
        self.doc = Document()
        self.in_code_block = False
        self.code_block_lines = []

    def convert(self, markdown_text: str) -> Document:
        """Конвертира markdown текст в DOCX документ"""
        lines = markdown_text.split('\n')
        
        for line in lines:
            self._process_line(line)
        
        # Добавяне на остатъчни редове от code блок
        if self.in_code_block:
            self._add_code_block()
        
        return self.doc

    def _process_line(self, line: str):
        """Обработва един ред от markdown"""
        
        # Code блокове
        if line.strip().startswith('```'):
            if not self.in_code_block:
                self.in_code_block = True
                self.code_block_lines = []
            else:
                self._add_code_block()
                self.in_code_block = False
            return
        
        if self.in_code_block:
            self.code_block_lines.append(line)
            return
        
        # Пропускане на празни редове
        if not line.strip():
            return
        
        # Заглавия
        if line.startswith('# '):
            self._add_heading(line[2:].strip(), level=1)
        elif line.startswith('## '):
            self._add_heading(line[3:].strip(), level=2)
        elif line.startswith('### '):
            self._add_heading(line[4:].strip(), level=3)
        elif line.startswith('#### '):
            self._add_heading(line[5:].strip(), level=4)
        
        # Списъци
        elif line.strip().startswith('- ') or line.strip().startswith('* '):
            self._add_list_item(line.strip()[2:].strip())
        elif re.match(r'^\d+\.\s', line.strip()):
            match = re.match(r'^\d+\.\s(.+)', line.strip())
            if match:
                self._add_list_item(match.group(1), ordered=True)
        
        # Обикновен текст с форматиране
        else:
            self._add_paragraph_with_formatting(line)

    def _add_heading(self, text: str, level: int):
        """Добавя заглавие"""
        heading = self.doc.add_heading(text, level=level)
        heading.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    def _add_list_item(self, text: str, ordered: bool = False):
        """Добавя елемент на списък"""
        paragraph = self.doc.add_paragraph(text, style='List Number' if ordered else 'List Bullet')

    def _add_code_block(self):
        """Добавя code блок"""
        code_text = '\n'.join(self.code_block_lines)
        paragraph = self.doc.add_paragraph(code_text)
        
        # Форматиране на код
        for run in paragraph.runs:
            run.font.name = 'Courier New'
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Светлосиво фоново оцветяване
        shading_elm = self._shade_paragraph(paragraph)

    def _add_paragraph_with_formatting(self, text: str):
        """Добавя параграф с bold, italic, strikethrough и inline код форматиране"""
        paragraph = self.doc.add_paragraph()
        
        # Обработка на форматирането в правилния ред (от най-специфично към по-общо)
        # Паттерни: inline код, bold (**), зачеркнат (~~), italic (*)
        tokens = self._tokenize_formatting(text)
        
        for token_type, token_text in tokens:
            run = paragraph.add_run(token_text)
            
            if token_type == 'bold':
                run.bold = True
            elif token_type == 'italic':
                run.italic = True
            elif token_type == 'strikethrough':
                run.font.strike = True
            elif token_type == 'code':
                run.font.name = 'Courier New'
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(192, 0, 0)

    def _tokenize_formatting(self, text: str):
        """Разделя текста на токени с информация за форматирането"""
        tokens = []
        i = 0
        
        while i < len(text):
            # Inline код - приоритет 1 (най-висок)
            if text[i:i+1] == '`' and i + 1 < len(text):
                end = text.find('`', i + 1)
                if end != -1:
                    tokens.append(('code', text[i+1:end]))
                    i = end + 1
                    continue
            
            # Bold - приоритет 2
            if text[i:i+2] == '**':
                end = text.find('**', i + 2)
                if end != -1:
                    tokens.append(('bold', text[i+2:end]))
                    i = end + 2
                    continue
            
            # Зачеркнат текст - приоритет 3
            if text[i:i+2] == '~~':
                end = text.find('~~', i + 2)
                if end != -1:
                    tokens.append(('strikethrough', text[i+2:end]))
                    i = end + 2
                    continue
            
            # Italic - приоритет 4
            if text[i] == '*' or text[i] == '_':
                marker = text[i]
                end = text.find(marker, i + 1)
                if end != -1:
                    tokens.append(('italic', text[i+1:end]))
                    i = end + 1
                    continue
            
            # Обикновен текст
            next_special = len(text)
            for pattern in ['`', '**', '~~', '*', '_']:
                pos = text.find(pattern, i)
                if pos != -1 and pos < next_special:
                    next_special = pos
            
            if next_special > i:
                tokens.append(('normal', text[i:next_special]))
                i = next_special
            else:
                tokens.append(('normal', text[i]))
                i += 1
        
        return tokens

    def _shade_paragraph(self, paragraph):
        """Добавя сива фонова оцветяване на параграф"""
        from docx.oxml import parse_xml
        shading_elm = parse_xml(r'<w:shd {} w:fill="E7E6E6"/>'.format(
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        ))
        paragraph._element.get_or_add_pPr().append(shading_elm)
        return shading_elm

    def save(self, filename: str):
        """Запазва документа като DOCX файл"""
        self.doc.save(filename)


def markdown_to_docx(markdown_text: str, output_filename: str):
    """Простна функция за конвертиране на markdown в docx"""
    converter = MarkdownToDocx()
    doc = converter.convert(markdown_text)
    converter.save(output_filename)


def convert_file(input_file: str, output_file: str = None):
    """Конвертира markdown файл в docx файл"""
    
    # Проверка дали входния файл съществува
    if not os.path.exists(input_file):
        print(f"❌ Грешка: Файлът '{input_file}' не съществува!")
        return False
    
    # Проверка дали файлът е markdown
    if not input_file.lower().endswith(('.md', '.markdown')):
        print(f"⚠️  Внимание: Файлът не е markdown формат")
    
    # Генериране на изходния файл ако не е посочен
    if output_file is None:
        base_name = Path(input_file).stem
        output_file = f"{base_name}.docx"
    
    try:
        # Прочитане на markdown файла
        with open(input_file, 'r', encoding='utf-8') as f:
            markdown_content = f.read()
        
        print(f"📖 Прочитане на файл: {input_file}")
        
        # Конвертиране
        print(f"⏳ Конвертиране на markdown в docx...")
        markdown_to_docx(markdown_content, output_file)
        
        print(f"✅ Успешно! Документ създаден: {output_file}")
        return True
        
    except Exception as e:
        print(f"❌ Грешка при конвертиране: {str(e)}")
        return False


def main():
    """Главна функция за командния ред"""
    parser = argparse.ArgumentParser(
        description='Конвертира Markdown файлове в DOCX формат',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примери за използване:
  python markdown_to_docx.py README.md
  python markdown_to_docx.py README.md -o output.docx
  python markdown_to_docx.py document.md --output report.docx

Поддържани форматирания:
  **bold текст** - удебелен текст
  ~~strikethrough~~ - зачеркнат текст
  *italic текст* - наклонен текст
  `inline код` - вътрешен код
        """
    )
    
    parser.add_argument(
        'input',
        help='Пътя до markdown файла за конвертиране'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='Пътя до изходния DOCX файл (по подразбиране се използва същото име с .docx разширение)',
        metavar='OUTPUT_FILE'
    )
    
    parser.add_argument(
        '-v', '--version',
        action='version',
        version='%(prog)s 1.0'
    )
    
    args = parser.parse_args()
    
    # Конвертиране на файла
    success = convert_file(args.input, args.output)
    
    # Връщане на статус код
    return 0 if success else 1


if __name__ == "__main__":
    exit(main())
