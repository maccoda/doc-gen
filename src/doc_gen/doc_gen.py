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


def print_help():
    """Print help message"""
    print("doc_gen.py -i <input> -o <output>")
    print("-i,--input\t name of Markdown input file (with file extension)")
    print("-o,--output\t name to write output DOCX file to (with file extension)")

def parse_args(args):
    """Parse the provided arguments and return a dictionary of received arguments"""
    try:
        opts, args = getopt.getopt(args, "hi:o:", ["input", "output", "help"])
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

    return returned

def get_args(args):
    """Parse and check correct arguments present"""
    parsed = parse_args(args)
    if 'input_file' not in parsed or 'output_file' not in parsed:
        print("Error: Not all arguments provided")
        print_help()
        sys.exit(2)

    return (parsed['input_file'], parsed['output_file'])

def main(args):
    """Main method taking the markdown file and creating DOCX file"""

    input_file_name, output_file_name = get_args(args)

    with open(input_file_name) as in_file:
        doc_render = IdentityRenderer()
        md = mistune.Markdown(renderer=doc_render)
        result_list = md.output(in_file.read())

        DocumentBuilder(output_file_name).create(result_list)


if __name__ == '__main__':
    main(sys.argv[1:])
