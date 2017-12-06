"""
Module containing all of parsed elements from the Markdown file and the relevant
mappings to the DOCX types
"""


class DocumentElement:
    """Single DOCX document element. Capable of appending itself to a document"""

    def append_to_document(self, document, parent_level):
        """
        Append this element to the document with all heading levels being
        under the one provided.
        """
        pass


class SingleElementTextWrapper(DocumentElement):
    """Element expecting only a single Text element to be within its content"""

    def __init__(self, text):
        # Expect only a single Text element
        if len(text) > 1:
            raise Exception
        self.text = text[0]


class StyledText(SingleElementTextWrapper):
    """Text element with specialized styling. e.g. Bold & Italic"""

    def append_to_document(self, document, parent_level):
        run = document.paragraphs[len(document.paragraphs) -
                                  1].add_run(self.text.as_str())
        self.add_style_to_run(run)

    def add_style_to_run(self, run):
        """Add styling to the provided Run element"""
        pass


class Heading(SingleElementTextWrapper):
    """Heading element in document. Contains text and the level of heading"""

    def __init__(self, level, text):
        self.level = level
        super().__init__(text)

    def append_to_document(self, document, parent_level):
        document.add_heading(self.text.as_str(), level=(self.level + parent_level))


class Paragraph(DocumentElement):
    """
    Single paragraph element. This is a glob of text that is expected to be
    separated by a new line, mapping directly to a Paragraph element in the
    Document
    """

    def __init__(self, text_elements):
        self.text_elements = text_elements

    def append_to_document(self, document, parent_level):
        # Add a new paragraph
        document.add_paragraph("")
        for item in self.text_elements:
            item.append_to_document(document, parent_level)


class BoldText(StyledText):
    """Bold text element"""

    def add_style_to_run(self, run):
        run.bold = True


class ItalicText(StyledText):
    """Italic text element"""

    def add_style_to_run(self, run):
        run.italic = True


class Text(DocumentElement):
    """Basic text element, containing only plain string"""

    def __init__(self, text):
        self.text = text

    def append_to_document(self, document, parent_level):
        document.paragraphs[len(document.paragraphs) -
                            1].add_run(self.text)

    def as_str(self):
        """Returns the string representation of the this Text element"""
        return self.text
