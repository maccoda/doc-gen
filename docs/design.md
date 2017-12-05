# Design Decisions

## Template issue

It appears that am unable to correctly set the document styles to the default
styles when reading in from a document thus making it difficult for performing
the appending.

New idea will be rather than write that document in place we will construct a
new document and rebuild the template in that one to avoid the styling issues.
This works as the default `Document` contains all of the default styles provided
by Microsoft Word.

## Generating within templates

The simplest way to get this working is to develop a convention. The problem
trying to be solved here is that it is very simple to append to a template
document but the more likely use case will be to insert content into certain
sections of the template. So as to not make this tool at all complex we will
define rules for the definition of the template and the markdown files.

### Initial convention

This initial convention is to consider the simple situation whereby there is
only a single field in the document and there is only a single markdown file.
All the generator will look for is a Heading element of a certain name and will
add the generated documents into that position.

To continue on from the initial template example this keyword will be
**Content**

### Future ideas

In the likely case there are more than one section of the template wishing to
populate a better alternative will be needed. Two possible ideas come to mind:

1. The markdown file name must match the heading name
1. The first header in the markdown file must match the heading

Typically in Markdown convention these two things will be the same as one will
typically give the H1 heading of the Markdown file the same name of the file. So
I am inclined to use the second point as the convention. Whilst completing the
initial insertion functionality will consider this for future proofing.
