"""Integration tests"""

import sys
import os
from docx import Document
from docx.text.parfmt import ParagraphFormat

def rel_to_abs_path(relative):
    """Returns the absolute path of the file given as a relative path"""
    return os.path.abspath(os.path.join(os.path.dirname(__file__), relative))


PATH = rel_to_abs_path('../doc_gen')
sys.path.insert(0, PATH)
import doc_gen

DEF_STYLE = 'Default Paragraph Font'

def test_all():
    """High level integration test"""
    in_path = rel_to_abs_path('./test1.md')
    out_path = rel_to_abs_path('./output.docx')
    doc_gen.main([in_path, out_path])
    # Now to test
    document = Document(out_path)
    paragraphs = document.paragraphs
    assert len(paragraphs) == 6
    # Check the headings
    assert_heading(paragraphs[0], 'Heading 1', 'Heading 1')
    assert_heading(paragraphs[1], 'Heading 2', 'Heading 2')
    assert_heading(paragraphs[4], 'Another Heading 2', 'Heading 2')

    # Check the basic paragraphs/text
    assert_text_runs(paragraphs[3], 'Now a second paragraph')
    assert_text_runs(paragraphs[5], 'Let us add some more text.')
    assert_text_runs(
        paragraphs[2], 'Some paragraph text with bold and italics.',
        [DEF_STYLE, 'bold', DEF_STYLE, 'italic', DEF_STYLE])



def assert_heading(heading, text, style_name):
    """Assert that the provided heading matches the text and style name provided"""
    assert heading.text == text
    assert heading.style.name == style_name

def assert_text_runs(text_run, text, run_styles=None):
    """Assert that the provided text has the content and styles provided"""
    assert text_run.text == text
    if not run_styles:
        assert len(text_run.runs) == 1
        assert text_run.runs[0].style.name == DEF_STYLE
    else:
        assert len(text_run.runs) == len(run_styles)
