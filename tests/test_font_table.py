from unittest import TestCase

from docx2css.ooxml.package import OpcPackage


class TestFontTable(TestCase):

    def setUp(self):
        package = OpcPackage('test_files/no_theme.docx')
        self.font_table = package.font_table

    def test_font_name(self):
        font = self.font_table.get_font('Times New Roman')
        self.assertIsNotNone(font.name)
        self.assertEqual('Times New Roman', font.name)

    def test_alt_name(self):
        font = self.font_table.get_font('Liberation Sans')
        self.assertEqual('Arial', font.alt_name)

    def test_no_alt_name(self):
        font = self.font_table.get_font('Symbol')
        self.assertIsNone(font.alt_name)

    def test_family(self):
        font = self.font_table.get_font('Liberation Sans')
        self.assertEqual('swiss', font.family)

    def test_css_generic_family(self):
        font = self.font_table.get_font('Liberation Sans')
        self.assertEqual('sans-serif', font.css_generic_family)

    def test_css_family(self):
        font = self.font_table.get_font('Liberation Sans')
        expected = ('Liberation Sans', 'Arial', 'sans-serif')
        self.assertEqual(expected, font.css_family)

    def test_css_family_no_alt_name(self):
        font = self.font_table.get_font('Times New Roman')
        expected = ('Times New Roman', 'serif')
        self.assertEqual(expected, font.css_family)
