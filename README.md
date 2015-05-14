# Email2PDF
Converts an email to pdf

Uses Outlook Rule to trigger a script.
This creates a doc file from the email and also calls a vbs container for a batch script (batchdoc2pdf.cmd).
This deletes the doc files after running them through doc2pdf.vbs which uses Word to convert them to pdfs.

*Due to opening Word, this script is slow*

I'd really like to change this to be able to run with pandoc in the future.

**I borrowed a lot from Rob van der Woude.
You can find a lot of his excellent work on his site http://www.robvanderwoude.com/**
