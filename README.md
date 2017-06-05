# picExtractor
Extracts pictures from .docx and .xlsx files

Uses python to extract picture files from word documents. Pictures will be extracted into the directory where the documents are located. Pictures will be named as "<name of the file> - <picture number in document>"

Issues:

Currently configured for a Linux file structure, will need to be modified to work on Windows. Use \\\\ to get a single \ litteral in your strings.

Might not get all the pictures? I had some issues pulling 100% of the pictures when testing (only happens sometime). Give me feedback, if something wonky (technical term) happens.
