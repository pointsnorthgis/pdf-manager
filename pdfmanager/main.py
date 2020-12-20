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
from tkinter import Tk, filedialog, messagebox, ttk, Toplevel
from PyPDF4 import PdfFileReader, PdfFileWriter
import PIL
from PIL import ImageTk
import fitz


class PdfHandler(object):
    def __init__(self):
        self.pdf_paths = []  # List of incoming PDFs
        self.pdf_path = None  # Path to open PDF
        self.start_page = None  # First page to split
        self.end_page = None  # Last page to split
        self.page_range = None  # String input of pages to split
        self.pages_array = []  # List of pages to split
        self.merge = False  # Flag for splitting pages into one file
        self.pdf = None  # PDF File Reader
        self.page_count = None  # Number of pages in split PDF
        self.pdf_writer = PdfFileWriter()  # PDF File Writer
        self.pdf_image = None
        self.edited_pages = {}

    def error_message(self, message):
        print("Error: {}".format(message))

    def complete(self):
        return

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
            self.merge_pdf()

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

    def pdf_page_image(self, pdf_page=0):
        doc = fitz.open(self.pdf_path)
        page = doc.loadPage(pdf_page)
        pix = page.getPixmap()
        self.pdf_image = PIL.Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return self.pdf_image

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
        self.complete()

    def write_pdf(self, output_file_name):
        """Write PDF to Disk"""
        with open(output_file_name, 'wb') as output_pdf:
            self.pdf_writer.write(output_pdf)

    def merge_pdf(self):
        base_path = os.path.split(self.pdf_paths[0])[0]
        output_filename = os.path.join(base_path, 'merged.pdf')
        for pdf in self.pdf_paths:
            self.pdf = PdfFileReader(pdf)
            for page_num in range(0, self.pdf.getNumPages()):
                page = self.pdf.getPage(page_num)
                self.pdf_writer.addPage(page)
        self.write_pdf(output_filename)
        return

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
    """TKinter user interface for splitting/merging PDFs"""
    def __init__(self):
        super().__init__()
        self.current_page = 0
        self.ui = Tk(className=' PDF Manager')
        self.setup_ui()
    
    def submit_callback(self):
        """Actions to complete after submit button clicked"""
        if len(self.pdf_paths) == 0:
            self.error_message("Please select PDF(s) to process.")
        if len(self.pdf_paths) == 1:
            page_range_input = self.page_range.get()
            self.parse_pages(page_range_input)
            if len(self.pages_array) == 0:
                self.error_message("Please enter pdf pages to merge")
            self.split_pdf()
        else:
            self.merge_pdf()
        self.complete_message()

    def complete_message(self):
        '''complete message as popup'''
        popup = Tk()
        popup.wm_title("!")
        label = ttk.Label(popup, text="Proccessing Complete", font=("Verdana", 10))
        label.pack(side="top", fill="x", pady=10)
        exit_btn = ttk.Button(popup, text="Okay", command=popup.destroy)
        exit_btn.pack()
        popup.mainloop()

    def init_pdf_image(self, pdf_page=0):
        '''Get PDF Page as image and display in GUI'''
        self.pdf_page_image(pdf_page)
        rotate = 0
        if pdf_page in self.edited_pages.keys():
            if "rotate" in self.edited_pages[pdf_page].keys():
                rotate = self.edited_pages[pdf_page]["rotate"]
        if rotate != 0:
            self.pdf_image = self.pdf_image.rotate(rotate)
        self.pdf_image = ImageTk.PhotoImage(self.pdf_image)
        self.pdf_canvas.create_image((0,0), image=self.pdf_image, anchor='nw')

    def next_pdf_page(self):
        if self.current_page < self.end_page:
            self.current_page += 1
        self.init_pdf_image(pdf_page=self.current_page)

    def previous_pdf_page(self):
        if self.current_page >= self.start_page:
            self.current_page = self.current_page - 1
        self.init_pdf_image(pdf_page=self.current_page)

    def rotate_pdf_ccw(self):
        '''Rotate PDF Page Clockwise'''
        page = self.current_page
        if page in self.edited_pages.keys():
            if "rotate" in self.edited_pages[page].keys():
                self.edited_pages[page]["rotate"] += 90
            else:
                self.edited_pages[page]["rotate"] = 90
        else:
            self.edited_pages[page] = {"rotate": 90}
        rotation = self.edited_pages[page]["rotate"]
        self.init_pdf_image(pdf_page=page)
        return

    def rotate_pdf_cw(self):
        '''Rotate PDF Page Counter Clockwise'''
        page = self.current_page
        if page in self.edited_pages.keys():
            if "rotate" in self.edited_pages[page].keys():
                self.edited_pages[page]["rotate"] += -90
            else:
                self.edited_pages[page]["rotate"] = -90
        else:
            self.edited_pages[page] = {"rotate": -90}
        rotation = self.edited_pages[page]["rotate"]
        self.init_pdf_image(pdf_page=page)
        return

    def save_pdf_edit(self):
        '''Save edits to PDF File'''
        return

    def edit_pdf_page(self):
        '''Child window with buttons for rotating pdf pages'''
        # If a valid PDF is not selected, do not open window
        if not self.pdf_path or not os.path.isfile(self.pdf_path):
            return
        self.pdf_window = Toplevel()
        pdf_window = self.pdf_window
        pdf_window.wm_title('Edit PDF')
        pdf_window.geometry('1000x1000')
        rotate_ccw_btn = tkinter.Button(pdf_window, text="Rotate CCW", command=self.rotate_pdf_ccw)
        rotate_ccw_btn.pack()
        previous_page_btn = tkinter.Button(pdf_window, text="Previous Page", command=self.previous_pdf_page)
        previous_page_btn.pack()
        next_page_btn = tkinter.Button(pdf_window, text="Next Page", command=self.next_pdf_page)
        next_page_btn.pack()
        rotate_cw_btn = tkinter.Button(pdf_window, text="Rotate CW", command=self.rotate_pdf_cw)
        rotate_cw_btn.pack()
        save_btn = tkinter.Button(pdf_window, text="Save", command=self.save_pdf_edit)
        save_btn.pack()
        cancel_btn = tkinter.Button(pdf_window, text="Cancel", bg="red", command=self.close_pdf_edit)
        cancel_btn.pack()
        self.pdf_canvas = tkinter.Canvas(pdf_window, height=1000, width=1000, bg="grey")
        self.pdf_canvas.pack()

        rotate_ccw_btn.place(x=10, y=10)
        previous_page_btn.place(x=100, y=10)
        next_page_btn.place(x=200, y=10)
        rotate_cw_btn.place(x=280, y=10)
        save_btn.place(x=360, y=10)
        cancel_btn.place(x=410, y=10)
        self.pdf_canvas.place(x=10, y=50)

        # Get PDF Properties
        self.list_pages(self.pdf_path)
        self.get_bounding_pages()
        self.init_pdf_image()
        pdf_window.mainloop()
        return

    def error_message(self, message):
        '''Error message as popup'''
        popup = Tk()
        popup.wm_title("!")
        label = ttk.Label(popup, text=message, font=("Verdana", 10))
        label.pack(side="top", fill="x", pady=10)
        exit_btn = ttk.Button(popup, text="Okay", command=popup.destroy)
        exit_btn.pack()
        popup.mainloop()

    def close(self):
        self.ui.destroy()

    def close_pdf_edit(self):
        self.pdf_window.destroy()
    
    def setup_ui(self):
        """Setup user interface with TKinter"""

        def select_pdf():
            """Function for opening file selector for selecting single PDF"""
            self.ui.filename =  filedialog.askopenfilename(
                initialdir="./", title="Select file",
                filetypes=(("pdf","*.pdf"),("all files","*.*"))
            )
            self.filename = self.ui.filename
            self.pdf_paths = []
            self.pdf_paths.append(self.filename)
            self.pdf_path = self.filename
            label = os.path.split(self.filename)[1]
            if len(label) > 50:
                label = '...' + label[-50:]
            self.file_name_label.config(text=label)
            
        def select_pdfs():
            """Function for opening file selector for multiple PDFs"""
            self.ui.files_list =  filedialog.askopenfilenames(
                initialdir="./", title="Select file",
                filetypes=(("pdf","*.pdf"),("all files","*.*"))
            )
            self.pdf_paths = self.ui.files_list
            populate_listbox()

        def clear_listbox():
            """Clear items in listbox"""
            list_length = self.pdf_select_list.size()
            self.pdf_select_list.delete(0, list_length)

        def populate_listbox():
            """Add selected PDF paths to listbox"""
            clear_listbox()
            for f in self.pdf_paths:
                self.pdf_select_list.insert('end', f)

        def show_frame_content(frame):
            """Show all UI content in a TK Frame"""
            for widget in frame.winfo_children():
                if widget.widgetName == 'scrollbar':
                    widget.pack(fill='x', side='bottom')
                else:
                    widget.pack()

        def hide_frame_content(frame):
            """Hide all UI content in a TK Frame"""
            for widget in frame.winfo_children():
                widget.pack_forget()

        def combobox_select(event=None):
            """Change type of file selector based on combobox selection"""
            if event:
                if event.widget.get() == "Merge PDFs":
                    self.select_pdf['command'] = select_pdfs
                    self.select_pdf['text'] = "Select PDF Files"
                    hide_frame_content(self.page_range_frame)
                    hide_frame_content(self.file_name_frame)
                    hide_frame_content(self.edit_pdf_btn_frame)
                    show_frame_content(self.list_frame)
                else:
                    self.select_pdf['command'] = select_pdf
                    self.select_pdf['text'] = "Select PDF File"
                    if event.widget.get() == 'Split to Single PDF':
                        self.merge = True
                    else:
                        self.merge = False
                    self.list_frame.lower()
                    hide_frame_content(self.list_frame)

                    if event.widget.get() != "Edit PDF":
                        hide_frame_content(self.edit_pdf_btn_frame)
                        show_frame_content(self.page_range_frame)
                    if event.widget.get() == "Edit PDF":
                        hide_frame_content(self.page_range_frame)
                        show_frame_content(self.edit_pdf_btn_frame)
                    
                    show_frame_content(self.file_name_frame)

        def listbox_select(event=None):
            """Get the currently selected listbox item"""
            if event:
                print(event.widget.get(event.widget.curselection()))
            return

        self.ui.geometry('300x500')

        label = tkinter.Label(text="Select PDF Operation")

        # Combobox for selecting how PDFs will be processed
        self.combo_box = ttk.Combobox(
            self.ui,
            values=['Split to Multiple PDFs', 'Split to Single PDF', 'Merge PDFs', 'Edit PDF'],
        )
        self.combo_box.set("Split to Multiple PDFs")
        self.combo_box.pack()
        self.combo_box.bind('<<ComboboxSelected>>', combobox_select)

        self.select_pdf = tkinter.Button(
            self.ui, text="Select PDF", command=select_pdf
        )

        self.file_name_frame = tkinter.Frame(self.ui)
        self.file_name_label = tkinter.Label(
            self.file_name_frame,
            text="No PDF Selected",
            height=2
        )
        self.file_name_label.pack()

        # List of PDFs to merge
        self.list_frame = tkinter.Frame(self.ui)
        
        scroll_bar = tkinter.Scrollbar(
            self.list_frame, orient='horizontal'
        )

        self.pdf_select_list = tkinter.Listbox(
            self.list_frame, selectmode='single',
            xscrollcommand=scroll_bar.set,
            width=40, height=10
        )
        self.pdf_select_list.bind('<<ListboxSelect>>', listbox_select)
        self.pdf_select_list.config()
        scroll_bar.config(command=self.pdf_select_list.xview)

        # Add frame, label and text input for selecting page ranges
        self.page_range_frame = tkinter.Frame(self.ui)
        page_range_label = tkinter.Label(
            self.page_range_frame,
            text="Enter Page Selection/Range \n (i.e. 1,2,4-6,10)",
            height=2
        )
        page_range_label.pack()
        
        self.page_range_text = tkinter.Entry(self.page_range_frame, textvariable=tkinter.StringVar())
        self.page_range_text.pack(fill='y')

        self.edit_pdf_btn_frame = tkinter.Frame(self.ui)
        edit_pdf_btn = tkinter.Button(
            self.edit_pdf_btn_frame, text="Edit PDF",
            command=self.edit_pdf_page
        )
        edit_pdf_btn.pack_forget()

        self.submit_button = tkinter.Button(
            self.ui, text="Start", command=self.submit_callback, bg="green"
        )

        self.close_button = tkinter.Button(
            self.ui, text="Close", command=self.close, bg="red"
        )

        # self.canvas = tkinter.Canvas(height=100, width=100)
        # self.canvas.pack()

        label.place(x=5, y=10)
        self.combo_box.place(x=10, y=30)
        self.select_pdf.place(x=10, y=60)
        self.file_name_frame.place(x=10, y=85)
        self.page_range_frame.place(x=10, y=115)
        self.edit_pdf_btn_frame.place(x=10, y=115)
        self.list_frame.place(x=10, y=195)
        self.submit_button.place(x=100, y=380)
        self.close_button.place(x=140, y=380)
        # self.canvas.place(x=10, y=410)
        self.ui.mainloop()

if __name__ == '__main__':
    """If program is run with commandline flags, parses arguments, otherwise use TKinter"""
    parser = argparse.ArgumentParser(description='Split Multi page PDF into single pages.')
    parser.add_argument(
        '--pdf', type=str, nargs='+', required=False,
        help='''
        Use the --pdf flag to enter a PDF file to split into multiple pages 
        or enter multiple pdf documents to merge into a single pdf.
        '''
        )
    parser.add_argument(
        '--merge', dest='merge', action='store_true',
        help='''
        Use this flag to split pages into one pdf.
            '''
        )
    parser.add_argument(
        '--pages', type=str, nargs='?', required=False,
        help='Select specific pages or page ranges to split (ie 1 3-5 7)'
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
        pdf_handler.execute_handler()
    print("Complete")
