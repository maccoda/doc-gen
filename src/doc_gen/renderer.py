"""
Module containing the custom renderer to be used for the markdown parsing.
"""

import mistune

from doc_elements import *

class IdentityRenderer(mistune.Renderer):
    """
    Renderer which just returns a typed object of the arguments for later
    processing to construct the DOCX document
    """

    def __init__(self, **kwargs):
        super().__init__()
        self.dir_name = ""

    def in_file_directory(self, dir_name):
        """Sets the directory for the current markdown file"""
        self.dir_name = dir_name

    def placeholder(self):
        return []

    def header(self, text, level, raw=None):
        return [Heading(level, text)]

    def paragraph(self, text):
        return [Paragraph(text)]

    def double_emphasis(self, text):
        return [BoldText(text)]

    def emphasis(self, text):
        return [ItalicText(text)]

    def text(self, text):
        return [Text(text)]

    def image(self, src, title, text):
        return [Image(self.dir_name + src)]

    def link(self, link, title, text):
        return [Link(link, text)]

    def list(self, body, ordered=True):
        return [WrittenList(body, ordered)]

    def list_item(self, text):
        return [ListElement(text)]

    def table(self, header, body):
        print('table')
        print(header)
        print(body)
        return [Table(header, body)]

    def table_row(self, content):
        print("row")
        print(content)
        return [TableRow(content)]

    def table_cell(self, content, **flags):
        return [TableCell(content)]
