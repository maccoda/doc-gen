# doc-gen

Documentation generator of DOCX files from Markdown.

The initial concept of this project was to address the issue where formal
documentation was required in Microsoft Word documents but this added large
overhead in maintaining the document. Considering it is commonplace to have the
code and project documented easily with Markdown within the code repository
itself it seemed only logical that we create something that is able use this
content and apply it to whatever formal template is required.

It is very early days for this project and currently it can only do a 1-to-1
conversion of the Markdown file to DOCX.

Future Goals:
- Build DOCX from a template file (either DOCX or Markdown defined not sure yet)
- Handle multiple Markdown files
    - Define an ordering to the multiple Markdown files so that it does not all
      need to be in a single file but is written and viewed in the correct
      sections
    - File names could map to headings
