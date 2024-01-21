from unittest import TestCase

import cssutils

from docx2css.css.serializers import FACTORY
from docx2css.ooxml.package import OpcPackage
from docx2css.ooxml.parsers import DocxParser

cssutils.ser.prefs.indentClosingBrace = False
cssutils.ser.prefs.omitLastSemicolon = False


def get_page_style(filename):
    parser = DocxParser(filename)
    return parser.parse_page_style()


class SectionParserTestCase(TestCase):

    def test_height_letter(self):
        style = get_page_style('test_files/numbering/docx/requete.docx')
        self.assertEqual(11, style.page_height.inches)

    def test_height_legal_landscape(self):
        style = get_page_style('test_files/sections/docx/legal_landscape.docx')
        self.assertEqual(8.5, style.page_height.inches)

    def test_width_letter(self):
        style = get_page_style('test_files/numbering/docx/requete.docx')
        self.assertEqual(8.5, style.page_width.inches)

    def test_width_legal_landscape(self):
        style = get_page_style('test_files/sections/docx/legal_landscape.docx')
        self.assertEqual(14, style.page_width.inches)

    def test_orientation_portrait(self):
        style = get_page_style('test_files/numbering/docx/requete.docx')
        self.assertEqual('portrait', style.page_orientation)

    def test_orientation_landscape(self):
        style = get_page_style('test_files/sections/docx/legal_landscape.docx')
        self.assertEqual('landscape', style.page_orientation)

    def test_margin_top_1in(self):
        style = get_page_style('test_files/numbering/docx/requete.docx')
        self.assertEqual(1, style.margin_top.inches)

    def test_margin_right_1in(self):
        style = get_page_style('test_files/numbering/docx/requete.docx')
        self.assertEqual(1, style.margin_right.inches)

    def test_margin_bottom_1in(self):
        style = get_page_style('test_files/numbering/docx/requete.docx')
        self.assertEqual(1, style.margin_bottom.inches)

    def test_margin_left_1in(self):
        style = get_page_style('test_files/numbering/docx/requete.docx')
        self.assertEqual(1, style.margin_left.inches)


class PageSizeSerializerTestCase(TestCase):

    def compare_style(self, docx_filename, css_filename):
        with open(css_filename, 'r') as css_file:
            expected = css_file.read()
            style = get_page_style(docx_filename)
            serializer = FACTORY.get_block_serializer(style)
            css_stylesheet = cssutils.css.CSSStyleSheet()
            for rule in serializer.css_style_rules():
                css_stylesheet.add(rule)
            result = css_stylesheet.cssText.decode('utf-8')
            self.assertEqual(expected, result)

    def test_css_print_letter(self):
        # There shouldn't be a tab before the last brace, but there it is...
        self.compare_style('test_files/numbering/docx/requete.docx',
                           'test_files/sections/css/requete_print.css')

    def test_css_print_legal(self):
        self.compare_style('test_files/sections/docx/legal_landscape.docx',
                           'test_files/sections/css/legal_landscape_print.css')
