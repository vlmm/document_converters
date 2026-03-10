import unittest

class TestMarkdownToDocx(unittest.TestCase):

    def test_bold(self):
        markdown_input = "**Bold Text**"
        expected_output = "<w:b>Bold Text</w:b>"
        self.assertEqual(markdown_to_docx(markdown_input), expected_output)

    def test_italic(self):
        markdown_input = "*Italic Text*"
        expected_output = "<w:i>Italic Text</w:i>"
        self.assertEqual(markdown_to_docx(markdown_input), expected_output)

    def test_strikethrough(self):
        markdown_input = "~~Strikethrough Text~~"
        expected_output = "<w:strike>Strikethrough Text</w:strike>"
        self.assertEqual(markdown_to_docx(markdown_input), expected_output)

    def test_code_block(self):
        markdown_input = "```
        Code Block
        ```"
        expected_output = "<w:code>Code Block</w:code>"
        self.assertEqual(markdown_to_docx(markdown_input), expected_output)

    def test_headings(self):
        markdown_input = "# Heading 1"
        expected_output = "<w:h1>Heading 1</w:h1>"
        self.assertEqual(markdown_to_docx(markdown_input), expected_output)

    def test_lists(self):
        markdown_input = "* Item 1\n* Item 2"
        expected_output = "<w:list><w:item>Item 1</w:item><w:item>Item 2</w:item></w:list>"
        self.assertEqual(markdown_to_docx(markdown_input), expected_output)

if __name__ == '__main__':
    unittest.main()