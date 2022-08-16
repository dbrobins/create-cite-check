# Cite-Check Report Creator for Law Journals

This is a Visual Basic (for Applications, VBA) program to help with law journal cite-checking.

## Setup

Create a document with a table with two rows and two columns. Leave the first columns blank (the footnote number will be inserted there) and in the second  columns, insert something like "TEXT: " (include a space after the colon) in the first row and "FOOTNOTE TEXT: " in the second. It's only important that there's a colon and space in those cells; anything else, and formatting, will be cloned/preserved when content is inserted or rows are added.

Edit the document's macros (may require saving as a ".docm") and paste the citecheck.vba file contents into the editor. Save the document.


## Use

1.	Open the article to be cite-checked and select the main body text (with footnote flags in it) that you have been assigned. The selection should start just past the last footnote before your assignment, and should include at the end the last footnote you were assigned.

2.	Save-As the report document (on your computer, not a shared/cloud location) according to the journal conventions (e.g., `<author last name>_<footnote #-#>_CC Report_<your last name>`) (keep the “.docm” extension as-is). It's only critical to have "CC" and "Report" in the name.

3.	Go to the View menu (ribbon) at the top and click on Macro. A dialog will open and the only macro, “CreateCiteCheck” should be highlighted. Click “Run.”

4.	The table in the report document should be extended with your assigned text/footnotes.


## Author

David Robins, software engineer and now law student.
