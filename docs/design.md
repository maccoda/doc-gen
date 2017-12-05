# Design Decisions

## Template issue

It appears that am unable to correctly set the document styles to the default
styles when reading in from a document thus making it difficult for performing
the appending.

New idea will be rather than write that document in place we will construct a
new document and rebuild the template in that one to avoid the styling issues.
This works as the default `Document` contains all of the default styles provided
by Microsoft Word.
