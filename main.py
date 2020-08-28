"""
Utility for splitting multi page PDFs into single page files
Developed By: Zackary King
Created: August 27, 2020
Updated By:
Updated On:
"""

import os
from PyPDF4 import PdfFileReader, PdfFileWriter


def split_pdf(pdf_path, start_page=None, end_page=None):
    """Takes a PDF file and splits it into a page range if given one"""
    if os.path.isfile(pdf_path):
        pdf = PdfFileReader(pdf_path)
        pdf_dir, file_name = os.path.split(pdf_path)
        pdf_name = os.path.splitext(file_name)[0]

        if start_page is None:
            start_page = 1

        if end_page is None:
            end_page = pdf.getNumPages()
            
        for page in range(pdf.getNumPages()):
            try:
                if page >= start_page and page <= end_page:
                    pdf = PdfFileReader(pdf_path)
                    pdf_writer = PdfFileWriter()
                    pdf_writer.addPage(pdf.getPage(page))

                    output_filename = f'{pdf_name}{page}.pdf'
                    output_pdf_file = os.path.join(pdf_dir, output_filename)
                    with open(output_pdf_file, 'wb') as output_pdf:
                        pdf_writer.write(output_pdf)
                if page > end_page:
                    break
            except Exception as e:
                print("Failed to split PDF page {}. ERROR: {}".format(page, e))
        return
    else:
        print("This is not a PDF File: {}".format(pdf_path))

if __name__ == '__main__':
    split_pdf('./test_data/docs-pdf/howto-argparse.pdf', start_page=1, end_page=5)
