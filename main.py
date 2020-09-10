"""
Utility for splitting multi page PDFs into single page files
Developed By: Zackary King
Created: August 27, 2020
Updated By: Zackary King 
Updated On: August 28, 2020
"""

import os
import sys
import argparse
import tkinter
from tkinter import Tk, filedialog, messagebox, ttk
from PyPDF4 import PdfFileReader, PdfFileWriter


class PdfHandler(object):
    def __init__(self):
        self.pdf_paths = []
        self.pdf_path = None
        self.start_page = None
        self.end_page = None
        self.page_range = None
        self.pages_array = []
        self.merge = None
        self.pdf = None
        self.page_count = None
        self.split_files = []
        self.pdf_writer = PdfFileWriter()

    def error_message(self, message):
        print("Error: {}".format(message))

    def complete(self):
        self.pdf_writer = None
        self.pdf = None

    def execute_handler(self):
        # if only 1 pdf - split into single pages
        if len(self.pdf_paths) == 0:
            message = 'No PDFs were selected.'
            self.error_message(message)
            assert Exception(message)
        elif len(self.pdf_paths) == 1:
            if not self.page_range:
                self.list_pages(self.pdf_paths[0])
            else:
                self.parse_pages(self.page_range)
                self.get_bounding_pages()
            self.split_pdf()
        else:
            message = 'Multiple PDFs detected. Waiting on merge function'
            self.error_message(message)
            assert Exception(message)

    def page_num_prefix(self, current_page):
        zero_pad_len = len(str(self.page_count))
        padding_len = zero_pad_len - len(str(current_page))
        count = 0
        page_num_str = ''
        while count < padding_len:
            page_num_str += str(0)
            count += 1
        page_num_str += str(current_page)
        return page_num_str

    def cast_to_page_int(self, page_num_str):
        try:
            page_num = int(page_num_str)
            return page_num
        except ValueError:
            message = "Only use whole numbers as page numbers."
            self.error_message(message)
            assert Exception(message)

    def list_pages(self, pdf):
        for page in range(PdfFileReader(pdf).getNumPages()):
            self.pages_array.append(page + 1)
        return

    def get_bounding_pages(self):
        if not self.start_page:
            self.start_page = self.pages_array[0]
        if not self.end_page:
            self.end_page = self.pages_array[-1]

    def parse_pages(self, page_range_input):
        """
        Take user input with page selection and ranges and create
        an array of acceptable pages to split/merge
        """
        def range_to_pages(start_page, end_page):
            x = start_page
            while x <= end_page:
                self.pages_array.append(x)
                x += 1

        page_range_input = page_range_input.replace(' ', '')
        page_input_array = page_range_input.split(',')
        for page_input in page_input_array:
            if len(page_input) == 0:
                return
            elif '-' in page_input:
                page_range = page_input.split('-')
                if len(page_range) > 2:
                    message = 'Please enter an appropriate page range'
                    self.error_message(message)
                    assert Exception(message)
                else:
                    range_start = self.cast_to_page_int(page_range[0])
                    range_end = self.cast_to_page_int(page_range[1])
                    range_to_pages(range_start, range_end)
            else:
                page_number_int = self.cast_to_page_int(page_input)
                self.pages_array.append(page_number_int)
        return

    def write_pdf(self, output_file_name):
        """Write PDF to Disk"""
        with open(output_file_name, 'wb') as output_pdf:
            self.pdf_writer.write(output_pdf)

    def split_pdf(self):
        """Takes a PDF file and splits it into a page range if given one"""
        self.pdf_path = self.pdf_paths[0]
        if os.path.isfile(self.pdf_path):
            self.pdf = PdfFileReader(self.pdf_path)
            self.page_count = self.pdf.getNumPages()
            pdf_dir, file_name = os.path.split(self.pdf_path)
            pdf_name = os.path.splitext(file_name)[0]
            
            self.get_bounding_pages()
            for page in range(self.pdf.getNumPages()):
                pdf_page = page + 1
                if pdf_page in self.pages_array or len(self.pages_array) == 0:
                    try:
                        if pdf_page >= self.start_page and pdf_page <= self.end_page:
                            self.pdf = PdfFileReader(self.pdf_path)
                            self.pdf_writer.addPage(self.pdf.getPage(page))

                            if not self.merge:
                                output_filename = f'{pdf_name}_{self.page_num_prefix(pdf_page)}.pdf'
                                output_pdf_file = os.path.join(pdf_dir, output_filename)
                                self.write_pdf(output_pdf_file)
                                self.pdf_writer = PdfFileWriter()
                                
                        if pdf_page == self.end_page or pdf_page == self.pdf.getNumPages():
                            if self.merge:
                                output_filename = f'{pdf_name}_merged.pdf'
                                output_pdf_file = os.path.join(pdf_dir, output_filename)
                                self.write_pdf(output_pdf_file)
                            break
                    except Exception as e:
                        message = "Failed to split PDF page {}. ERROR: {}".format(page, e)
                        self.error_message(message)
            self.complete()
        else:
            message = "This is not a PDF File: {}".format(self.pdf_path)
            self.error_message(message)
            self.complete()
            assert Exception(message)


class PdfMergeUI(PdfHandler):
    def __init__(self):
        super().__init__()
        self.ui = Tk()
        self.setup_ui()
    
    def submit_callback(self):
        """Actions to complete after submit button clicked"""
        page_range_input = self.page_range.get()
        self.parse_pages(page_range_input)
        if len(self.pages_array) == 0:
            self.error_message("Please enter pdf pages to merge")
        self.split_pdf()

    def error_message(self, message):
        '''Error message as popup'''
        popup = Tk()
        popup.wm_title("!")
        label = ttk.Label(popup, text=message, font=("Verdana", 10))
        label.pack(side="top", fill="x", pady=10)
        exit_btn = ttk.Button(popup, text="Okay", command=popup.destroy)
        exit_btn.pack()
        popup.mainloop()
    
    def setup_ui(self):
        """Setup user interface with TKinter"""

        def select_pdf():
            """Function for opening file selector for selecting single PDF"""
            self.ui.filename =  filedialog.askopenfilename(
                initialdir="/", title="Select file",
                filetypes=(("pdf","*.pdf"),("all files","*.*"))
            )
            self.filename = self.ui.filename
            self.pdf_paths.append(self.ui.filename)

        def select_pdfs():
            """Function for opening file selector for multiple PDFs"""
            self.ui.files_list =  filedialog.askopenfilenames(
                initialdir="/", title="Select file",
                filetypes=(("pdf","*.pdf"),("all files","*.*"))
            )
            self.merge_files = self.ui.files_list

        def combobox_select(event=None):
            """Change type of file selector based on combobox selection"""
            if event:
                if event.widget.get() == "Merge PDFs":
                    self.select_pdf['command'] = select_pdfs
                    self.select_pdf['text'] = "Select PDF Files"
                else:
                    self.select_pdf['command'] = select_pdf
                    self.select_pdf['text'] = "Select PDF File"
                    if event.widget.get() == 'Split to Single PDF':
                        self.merge = True
                    else:
                        self.merge = False

        label = tkinter.Label(text="Select PDF Operation")

        # Combobox for selecting how PDFs will be processed
        self.combo_box = ttk.Combobox(
            self.ui,
            values=['Split to Multiple PDFs', 'Split to Single PDF', 'Merge PDFs'],
        )
        self.combo_box.set("Split to Multiple PDFs")
        self.combo_box.pack()
        self.combo_box.bind('<<ComboboxSelected>>', combobox_select)

        self.select_pdf = tkinter.Button(
            self.ui, text="Select PDF", command=select_pdf
        )

        # Add text input for selecting page ranges
        self.page_range_label = tkinter.Label(
            text="Enter Page Selection/Range \n (i.e. 1,2,4-6,10)",
            height=2
            )
        self.page_range = tkinter.Entry(self.ui, textvariable=tkinter.StringVar())
        self.page_range.pack()

        self.submit_button = tkinter.Button(
            self.ui, text="Start", command=self.submit_callback, bg="green"
        )

        label.place(x=5, y=10)
        self.combo_box.place(x=10, y=30)
        self.select_pdf.place(x=10, y=60)
        self.page_range_label.place(x=10, y=90)
        self.page_range.place(x=10, y=130)
        self.submit_button.place(x=50, y=170)
        self.ui.mainloop()

if __name__ == '__main__':
    """If program is run with commandline flags, parses arguments, otherwise use TKinter"""
    parser = argparse.ArgumentParser(description='Split Multi page PDF into single pages.')
    parser.add_argument('--pdf', type=str, nargs='+', required=False, help='Path to PDF file')
    parser.add_argument(
        '--merge', dest='merge', action='store_true',
        help='Merge Pages into one PDF'
        )
    parser.add_argument(
        '--pages', type=str, nargs='?', required=False,
        help='Select Pages (ie 1 3-5 7)'
        )
    args = parser.parse_args()
    pdf = args.pdf
    if pdf:
        for pdf_file in pdf:
            if os.path.splitext(pdf_file)[1].lower() != '.pdf':
                raise Exception("Only include PDF Files")
            if not os.path.isfile(pdf_file):
                raise Exception("Ensure PDF files exist")

    merge = args.merge
    pages = args.pages

    if not pdf:
        PdfMergeUI()
    else:
        pdf_handler = PdfHandler()
        pdf_handler.page_range = pages
        pdf_handler.pdf_paths = pdf
        pdf_handler.merge = merge
        if len(pdf) == 1:
            pdf_handler.execute_handler()
    print("Complete")
