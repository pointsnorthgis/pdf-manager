import unittest
import PIL
from pdfmanager.main import PdfMergeUI, PdfHandler


class TestPdfEditor(unittest.TestCase):

    def setUp(self):
        self.test_pdf = 'test_data/docs-pdf/faq.pdf'
        self.pdf_handler = PdfHandler()
        self.pdf_handler.pdf_paths = [self.test_pdf]
        self.pdf_handler.pdf_path = self.test_pdf

    def test_pdf_to_image(self):
        pdf_image = self.pdf_handler.pdf_page_image()

        # Test initial PDF page image
        assert isinstance(pdf_image, PIL.Image.Image)
        
        # Test PDF Page rotation
        pdf_image = self.pdf_handler.pdf_page_image(pdf_page=1, rotate=90)
        assert isinstance(pdf_image, PIL.Image.Image)
