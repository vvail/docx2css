from unittest import TestCase

from docx2css.ooxml.package import OpcPackage


class TestTheme(TestCase):

    def setUp(self):
        parser = OpcPackage('test_files/character/docx/character_styles.docx')
        self.theme = parser.theme

    def test_colors(self):
        """There should 12 colors in a theme"""
        self.assertEqual(12, len(self.theme.colors))

    def test_colors_no_theme(self):
        package = OpcPackage('test_files/no_theme.docx')
        theme = package.theme
        self.assertEqual(0, len(theme.colors))

    def test_get_font(self):
        package = OpcPackage('test_files/character/docx/character_styles.docx')
        theme = package.theme
        self.assertEqual('Calibri Light', theme.get_font('majorAscii'))
        self.assertEqual('Calibri Light', theme.get_font('majorHAnsi'))
        self.assertEqual('', theme.get_font('majorBidi'))
        self.assertEqual('', theme.get_font('majorEastAsia'))
        self.assertEqual('Calibri', theme.get_font('minorAscii'))
        self.assertEqual('Calibri', theme.get_font('minorHAnsi'))
        self.assertEqual('', theme.get_font('minorBidi'))
        self.assertEqual('', theme.get_font('minorEastAsia'))

    def test_get_font_no_theme(self):
        package = OpcPackage('test_files/no_theme.docx')
        theme = package.theme
        self.assertEqual(None, theme.get_font('majorAscii'))
        self.assertEqual(None, theme.get_font('majorHAnsi'))
        self.assertEqual(None, theme.get_font('majorBidi'))
        self.assertEqual(None, theme.get_font('majorEastAsia'))
        self.assertEqual(None, theme.get_font('minorAscii'))
        self.assertEqual(None, theme.get_font('minorHAnsi'))
        self.assertEqual(None, theme.get_font('minorBidi'))
        self.assertEqual(None, theme.get_font('minorEastAsia'))
