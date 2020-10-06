# pdf-splitter
Python Program for splitting multi page PDF documents into single pages or merging multiple pdfs into one.

This program can be used via commandline with the following flags:

"--pdf" Use the --pdf flag to enter a PDF file to split into multiple pages 
        or enter multiple pdf documents to merge into a single pdf (Required).
        
"--merge" Use this flag to split pages into one pdf (Optional).

"--pages"  Select specific pages or page ranges to split (ie 1 3-5 7) (Optional)


Split PDF into single file command line example:

`pdf_handler.exe --pdf "c:/path/to/pdf_file.pdf" --merge --pages 1 3-5 7`


Split PDF file into single pages:

`pdf_handler.exe --pdf "c:/path/to/pdf_file.pdf" --pages 1 3-5 7`


Merge multiple PDFs into one:

***Note. PDFs will merge based on order they are entered. The output file will appear as "merged.pdf" in the directory of the first PDF in the list***

`pdf_handler.exe --pdf "c:/path/to/pdf_file1.pdf" "c:/path/to/pdf_file2.pdf"`



The PDF Handler program also has a basic user interface built in. Download the .exe file and open to run.


# Using pyinstaller to compile
Use the following command to create an executable/compiled program for the PdfManager

'''bash
pyinstaller pdfmanager/main.py --onefile --noconsole --name pdfmanager
'''

The program will be created in the *./dist* folder as **pdfmanager.exe**. It can be moved and run from any directory on the system.