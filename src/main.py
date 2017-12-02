"""
Doc-Gen
Documentation generator from markdown files.
"""

import mistune
from doc_elements import Heading, Paragraph, BoldText, ItalicText, Text, DocumentElement
from docx import Document


class IdentityRenderer(mistune.Renderer):
    """
    Renderer which just returns a typed object of the arguments for later
    processing to construct the DOCX document
    """

    def __init__(self, **kwargs):
        super().__init__()

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


class DocumentBuilder:
    """Builder of the DOCX document from DocumentElements"""

    def __init__(self, name):
        self.document = Document()
        self.name = name

    def create(self, elements):
        """Creates and saves a DOCX file from the elements provided"""
        # The top level items are just the block level items captured
        for item in elements:
            if isinstance(item, DocumentElement):
                item.append_to_document(self.document)
            else:
                raise Exception

        self.document.save(self.name)


def main():
    """Main method taking the markdown file and creating DOCX file"""
    with open('/home/dmaccora/workspace/Python/doc-gen/tests/test1.md') as in_file:
        doc_render = IdentityRenderer()
        md = mistune.Markdown(renderer=doc_render)
        result_list = md.output(in_file.read())

        DocumentBuilder('output.docx').create(result_list)


if __name__ == '__main__':
    main()
