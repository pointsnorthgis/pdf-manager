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
        self.assertIsInstance(pdf_image, PIL.Image.Image)

    
    def test_pdf_rotate(self):
        self.pdf_handler.current_page = 0

        # Test Counter Clockwise Rotation
        self.pdf_handler.rotate_pdf_page(self.pdf_handler.current_page, 'ccw')
        self.assertGreater(len(self.pdf_handler.edited_pages), 0)
        self.assertNotEqual(
            self.pdf_handler.edited_pages[self.pdf_handler.current_page]['rotate'], 0
            )

        # Test Clockwise Rotation
        self.pdf_handler.rotate_pdf_page(self.pdf_handler.current_page, 'cw')
        self.assertEqual(
            self.pdf_handler.edited_pages[self.pdf_handler.current_page]['rotate'], 0
            )

    def test_save_pdf(self):
        return
