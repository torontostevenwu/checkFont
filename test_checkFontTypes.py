import unittest
from checkFontTypes import get_fonts_from_docx, get_fonts_from_pdf

class TestFontExtraction(unittest.TestCase):
    def test_docx_fonts(self):
        expected_fonts = {"Calibri", "Cambria"}
        # The get_fonts_from_docx function expects a file object opened in binary mode,
        # similar to how streamlit provides the uploaded file.
        with open("sample.docx", "rb") as f:
            extracted_fonts = get_fonts_from_docx(f)
        self.assertEqual(extracted_fonts, expected_fonts)

    def test_pdf_fonts(self):
        # Common names from reportlab are "Times-Roman" and "Helvetica" (for Arial).
        # This will be verified and adjusted if needed after the first test run.
        expected_fonts = {"Times-Roman", "Helvetica"}
        with open("sample.pdf", "rb") as f:
            extracted_fonts = get_fonts_from_pdf(f)
        self.assertEqual(extracted_fonts, expected_fonts)

if __name__ == '__main__':
    unittest.main()
