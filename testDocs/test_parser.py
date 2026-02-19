import unittest as ut
from docxParser import DocumentParser
from docx import Document
from pathlib import Path

from default_settings import *

class TestDocParser(ut.TestCase):
    """
    Unit tests for docxParser
    """

    # Empty docx
    # Testing comment count
    def test_empty(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testEmpty.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments), 0, "Comment count incorrect")

    # Docx with 20 paragraphs of text
    # 5 correct by manually applying styles, 5 correct by establishing a new style, 10 incorrect
    # Testing comment count
    def test_text_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testText.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments), 10, "Comment count incorrect")

    # Docx with 10 paragraphs of text
    # Testing amount of mistakes in first 5 incorrect paragraphs
    def test_text_mistake_count_1(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testText.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments[0]), 5, "Mistake count incorrect")

    # Docx with 10 paragraphs of text
    # Testing amount of mistakes in last 5 incorrect paragraphs
    def test_text_mistake_count_2(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testText.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments[6]), 5, "Mistake count incorrect")


    # Docx with 2 tables & 2 table names
    # 1 correct, 1 incorrect
    # Testing comment count
    def test_tables_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testTables.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments), 3, "Comment count incorrect")

    # Docx with 2 tables & 2 table names
    # 1 correct, 1 incorrect
    # Testing name mistake count
    def test_tables_name_mistake_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testTables.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments[1]), 6, "Mistake count incorrect")

    # Docx with 2 tables & 2 table names
    # 1 correct, 1 incorrect
    # Testing table mistake count
    def test_tables_table_mistake_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testTables.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments[2]), 3, "Mistake count incorrect")

    # Docx with 2 images & 2 image names
    # 1 correct, 1 incorrect
    # Testing comment count
    def test_images_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testImages.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments), 2, "Comment count incorrect")

    # Docx with 2 images & 2 image names
    # 1 correct, 1 incorrect
    # Testing name mistake count
    def test_images_name_mistake_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testImages.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments[0]), 3, "Mistake count incorrect")

    # Docx with 2 images & 2 image names
    # 1 correct, 1 incorrect
    # Testing table mistake count
    def test_images_table_mistake_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testImages.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments[1]), 4, "Mistake count incorrect")

    # Docx with text, table, image
    # Testing comment count
    def test_all_comments_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testTextTableImage.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments), 4, "Comment count incorrect")

    def test_margins_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testmargins.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments), 11, "Comment count incorrect")

    def test_list_count(self):
        parser = DocumentParser()
        parser.set_settings(default_text_checklist, default_heading1_checklist, default_heading2_checklist,
                            default_heading3_checklist, default_table_name_checklist, default_table_headings_checklist,
                            default_table_text_checklist, default_list_checklist, default_margins_checklist,
                            default_image_checklist, default_image_name_checklist)
        file = Path(__file__).with_name('testlist.docx')
        comments = parser.parse_document(file)
        self.assertEqual(len(comments), 3, "Comment count incorrect")


if __name__ == '__main__':
    ut.main()
