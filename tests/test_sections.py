from unittest import TestCase

import cssutils

from docx2css.ooxml.package import OpcPackage


cssutils.ser.prefs.indentClosingBrace = False
cssutils.ser.prefs.omitLastSemicolon = False


class PageSizeTestCase(TestCase):

    def compare_style(self, css_style_rule, css_filename):
        with open(css_filename, 'r') as css_file:
            expected = css_file.read()
            css_text = css_style_rule.cssText
            print(css_text)
            self.assertEqual(expected, css_text)

    def test_height_letter(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        section = package.sections[-1]
        page_size = section.page_size
        self.assertEqual(11, page_size.height)

    def test_height_legal_landscape(self):
        package = OpcPackage('test_files/sections/docx/legal_landscape.docx')
        section = package.sections[-1]
        page_size = section.page_size
        self.assertEqual(8.5, page_size.height)

    def test_width_letter(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        section = package.sections[-1]
        page_size = section.page_size
        self.assertEqual(8.5, page_size.width)

    def test_width_legal_landscape(self):
        package = OpcPackage('test_files/sections/docx/legal_landscape.docx')
        section = package.sections[-1]
        page_size = section.page_size
        self.assertEqual(14, page_size.width)

    def test_orientation_portrait(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        section = package.sections[-1]
        page_size = section.page_size
        self.assertEqual('portrait', page_size.orientation)

    def test_orientation_landscape(self):
        package = OpcPackage('test_files/sections/docx/legal_landscape.docx')
        section = package.sections[-1]
        page_size = section.page_size
        self.assertEqual('landscape', page_size.orientation)

    def test_margin_top_1in(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        section = package.sections[-1]
        margin = section.margins.top
        self.assertEqual(1, margin)

    def test_margin_right_1in(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        section = package.sections[-1]
        margin = section.margins.right
        self.assertEqual(1, margin)

    def test_margin_bottom_1in(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        section = package.sections[-1]
        margin = section.margins.bottom
        self.assertEqual(1, margin)

    def test_margin_left_1in(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        section = package.sections[-1]
        margin = section.margins.left
        self.assertEqual(1, margin)

    def test_css_print_letter(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        section = package.sections[-1]
        # There shouldn't be a tab before the last brace, but there it is...
        self.compare_style(section.css_style_rule_print(),
                           'test_files/sections/css/requete_print.css')

    def test_css_print_legal(self):
        package = OpcPackage('test_files/sections/docx/legal_landscape.docx')
        section = package.sections[-1]
        self.compare_style(section.css_style_rule_print(),
                           'test_files/sections/css/legal_landscape_print.css')

    def test_css_screen_letter(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        section = package.sections[-1]
        # There shouldn't be a tab before the last brace, but there it is...
        self.compare_style(section.css_style_rule_screen(),
                           'test_files/sections/css/requete_screen.css')
