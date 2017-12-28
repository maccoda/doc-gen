"""
Module containing all of parsed elements from the Markdown file and the relevant
mappings to the DOCX types
"""
import os
import docx

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

    def heading_text(self):
        """Return the heading text"""
        return self.text

    def append_to_document(self, document, parent_level):
        document.add_heading(self.text.as_str(),
                             level=(self.level + parent_level))


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

class Image(DocumentElement):
    """Image element"""

    def __init__(self, image_path):
        self.path = os.path.abspath(image_path)

    def append_to_document(self, document, parent_level):
        document.add_picture(self.path)

class Link(SingleElementTextWrapper):
    """Hyperlink element"""

    def __init__(self, url, text):
        self.url = url
        super().__init__(text)

    def append_to_document(self, document, parent_level):
        paragraph = document.paragraphs[len(document.paragraphs) - 1]

        part = paragraph.part
        r_id = part.relate_to(
            self.url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        print(r_id)
        hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
        hyperlink.set(docx.oxml.shared.qn('r:id'), r_id,)

        new_run = docx.oxml.shared.OxmlElement('w:r')
        rPr = docx.oxml.shared.OxmlElement('w:rPr')

        new_run.append(rPr)
        new_run.text = self.text.as_str()
        hyperlink.append(new_run)

        paragraph._p.append(hyperlink)
