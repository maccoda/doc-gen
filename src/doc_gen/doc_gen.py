"""
Doc-Gen
Documentation generator from markdown files.
"""
import sys
import getopt

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

    def __init__(self, name, templ_name=None):
        self.document = Document()
        self.name = name
        if templ_name:
            self.template_doc = Document(templ_name)
        else:
            self.template_doc = None

    def create(self, elements):
        """Creates and saves a DOCX file from the elements provided"""
        if self.template_doc:
            self._populate_with_template()
        for item in elements:
            if isinstance(item, DocumentElement):
                item.append_to_document(self.document)
            else:
                raise Exception

        self.document.save(self.name)

    def _populate_with_template(self):
        """Populates the document with the template content"""
        for elem in self.template_doc.paragraphs:
            style_name = elem.style.name
            style = self.document.styles[style_name]
            para = self.document.add_paragraph(elem.text)
            para.style = style


def print_help():
    """Print help message"""
    print("doc_gen.py -i <input> -o <output>")
    print("-i,--input\t name of Markdown input file (with file extension)")
    print("-o,--output\t name to write output DOCX file to (with file extension)")


def parse_args(args):
    """Parse the provided arguments and return a dictionary of received arguments"""
    try:
        opts, args = getopt.getopt(
            args, "hi:o:t:", ["input", "output", "help", "template"])
    except getopt.GetoptError:
        print_help()
        sys.exit(2)

    returned = {}
    for opt, arg in opts:
        if opt in ('-h', '--help'):
            print_help()
            sys.exit()
        elif opt in ('-i', '--input'):
            returned['input_file'] = arg
        elif opt in ('-o', '--output'):
            returned['output_file'] = arg
        elif opt in ('-t', '--template'):
            returned['template_file'] = arg

    return returned


def get_args(args):
    """Parse and check correct arguments present"""
    parsed = parse_args(args)
    if 'input_file' not in parsed or 'output_file' not in parsed:
        print("Error: Not all arguments provided")
        print_help()
        sys.exit(2)

    return (parsed['input_file'], parsed['output_file'], parsed.get('template_file'))


def main(input_file_name, output_file_name, template_name):
    """Main method taking the markdown file and creating DOCX file"""

    with open(input_file_name) as in_file:
        doc_render = IdentityRenderer()
        md = mistune.Markdown(renderer=doc_render)
        result_list = md.output(in_file.read())

        DocumentBuilder(output_file_name, template_name).create(result_list)


if __name__ == '__main__':
    input_file_name, output_file_name, template_name = get_args(sys.argv[1:])
    main(input_file_name, output_file_name, template_name)
