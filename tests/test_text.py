import unittest
from modules.ppt_text_extraction import extract_text_from_ppt

class TestTextExtraction(unittest.TestCase):
    def test_extract_text(self):
        file_path = "tests/sample.pptx"
        result = extract_text_from_ppt(file_path)
        self.assertTrue(len(result) > 0)
        self.assertIn("Sample Slide Text", result[0])
