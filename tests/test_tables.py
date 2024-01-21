import logging
from unittest import TestCase

import cssutils
from lxml import etree

from docx2css.api import (
    Border,
    TableCellProperties,
    TableConditionalFormatting,
    TableRowProperties,
    TableStyle,
)
from docx2css.stylesheet import Stylesheet
from docx2css.css.serializers import (
    CssStylesheetSerializer,
    CssTableSerializer,
    FACTORY
)
from docx2css.ooxml import opc_parser
from docx2css.ooxml.parsers import DocxParser
from docx2css.utils import AutoLength, CssUnit, Percentage


cssutils.ser.prefs.indentClosingBrace = False
cssutils.ser.prefs.omitLastSemicolon = False
logging.basicConfig(format='%(filename)s:%(lineno)d %(message)s',
                    level=logging.DEBUG)


def load_docx_styles(filename):
    styles = {}
    parser = DocxParser(filename)
    stylesheet = parser.opc_package.styles
    for style in stylesheet.values():
        table_style = parser.parse_docx_table_style(style)
        styles[style.name] = table_style
    return styles


def load_xml_fragment(filename):
    """Load style located in XML fragment instead of docx file"""
    with open(filename) as file:
        xml = etree.fromstring(file.read(), opc_parser)
        parser = DocxParser('')
        return parser.parse_docx_table_style(xml)


class TestTableProperties(TestCase):

    files = (
        'table_bold.docx',
        'table_properties.docx',
    )
    css_files_location = None
    docx_files_location = 'test_files/tables/docx/'
    fragments_location = 'test_files/tables/docx/fragments/'

    def xml_style(self, name):
        """Load style located in XML fragment instead of docx file"""
        filename = f'{self.fragments_location}{name}.xml'
        return load_xml_fragment(filename)

    def setUp(self):
        self.styles = {}
        self.xml_elements = {}
        for file in self.files:
            parser = DocxParser(f'{self.docx_files_location}{file}')
            stylesheet = parser.opc_package.styles
            for style in stylesheet.values():
                self.xml_elements[style.name] = style
                table_style = parser.parse_docx_table_style(style)
                self.styles[style.name] = table_style

    def test_bold(self):
        style = self.styles['table-bold']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertTrue(style.bold)
        self.assertIsNone(style.italics)

    def test_alignment_none(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.alignment)

    def test_table_alignment_center(self):
        style = self.styles['table-align-center']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual('center', style.alignment)

    def test_table_alignment_right(self):
        style = self.styles['table-align-right']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual('end', style.alignment)

    def test_table_background_color_none(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.background_color)

    def test_table_background_color_red(self):
        """Use a fragment because Word UI puts the shading at the row
        level instead of table
        """
        style = self.xml_style('table-fill-red')
        self.assertEqual('#FF0000', style.background_color)

    def test_table_border_no_borders(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.border_bottom)

    def test_table_border_bottom_none(self):
        style = self.styles['table-borders-inside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.border_bottom)

    def test_table_border_bottom_05pt(self):
        style = self.styles['table-borders-outside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(0.5, style.border_bottom.width.pt)
        self.assertEqual('solid', style.border_bottom.style)
        self.assertIsNone(style.border_bottom.color)

    def test_table_border_inside_horizontal_none(self):
        style = self.styles['table-borders-outside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.border_inside_horizontal)

    def test_table_border_inside_horizontal_05pt(self):
        style = self.styles['table-borders-inside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(0.5, style.border_inside_horizontal.width.pt)
        self.assertEqual('solid', style.border_inside_horizontal.style)
        self.assertIsNone(style.border_inside_horizontal.color)

    def test_table_border_inside_vertical_none(self):
        style = self.styles['table-borders-outside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.border_inside_vertical)

    def test_table_border_inside_vertical_05pt(self):
        style = self.styles['table-borders-inside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(0.5, style.border_inside_vertical.width.pt)
        self.assertEqual('solid', style.border_inside_vertical.style)
        self.assertIsNone(style.border_inside_vertical.color)

    def test_table_border_left_none(self):
        style = self.styles['table-borders-inside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.border_left)

    def test_table_border_left_05pt(self):
        style = self.styles['table-borders-outside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(0.5, style.border_left.width.pt)
        self.assertEqual('solid', style.border_left.style)
        self.assertIsNone(style.border_left.color)

    def test_table_border_right_none(self):
        style = self.styles['table-borders-inside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.border_right)

    def test_table_border_right_05pt(self):
        style = self.styles['table-borders-outside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(0.5, style.border_right.width.pt)
        self.assertEqual('solid', style.border_right.style)
        self.assertIsNone(style.border_right.color)

    def test_table_border_top_none(self):
        style = self.styles['table-borders-inside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.border_top)

    def test_table_border_top_05pt(self):
        style = self.styles['table-borders-outside']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(0.5, style.border_top.width.pt)
        self.assertEqual('solid', style.border_top.style)
        self.assertIsNone(style.border_top.color)

    def test_table_cell_margin_bottom_none(self):
        style = self.styles['table-align-right']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.cell_padding_bottom)

    def test_table_cell_margin_bottom_0_dxa(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNotNone(style.cell_padding_bottom)
        self.assertEqual(0, style.cell_padding_bottom)

    def test_table_cell_margin_left_none(self):
        style = self.styles['table-align-right']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.cell_padding_left)

    def test_table_cell_margin_left_108_dxa(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNotNone(style.cell_padding_left)
        self.assertEqual(108, style.cell_padding_left.twips)

    def test_table_cell_margin_right_none(self):
        style = self.styles['table-align-right']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.cell_padding_right)

    def test_table_cell_margin_right_108_dxa(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNotNone(style.cell_padding_right)
        self.assertEqual(108, style.cell_padding_right.twips)

    def test_table_cell_margin_top_none(self):
        style = self.styles['table-align-right']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.cell_padding_top)

    def test_table_cell_margin_top_0_dxa(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNotNone(style.cell_padding_top)
        self.assertEqual(0, style.cell_padding_top)

    def test_table_cell_margins_01in(self):
        style = self.styles['table-cell-margins-01-inch']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(0.1, style.cell_padding_bottom.inches)
        self.assertEqual(0.1, style.cell_padding_left.inches)
        self.assertEqual(0.1, style.cell_padding_right.inches)
        self.assertEqual(0.1, style.cell_padding_top.inches)

    def test_table_cell_spacing_none(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.cell_spacing)

    def test_table_cell_spacing_01in(self):
        style = self.styles['table-cell-spacing-01in']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(0.1, style.cell_spacing.inches)

    def test_table_layout_none(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.layout)

    def test_table_layout_fixed(self):
        style = self.xml_style('table-layout-fixed')
        self.assertEqual('fixed', style.layout)

    def test_table_layout_auto(self):
        style = self.xml_style('table-layout-auto')
        self.assertEqual('auto', style.layout)

    def test_table_left_indent_none(self):
        style = self.styles['table-bold']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.indent)

    def test_table_left_indent_0_dxa(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(0, style.indent)

    def test_table_left_indent_1in(self):
        style = self.styles['table-leftindent-1in']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(1, style.indent.inches)

    def test_table_row_band_size_none(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.row_band_size)

    def test_table_row_band_size_2(self):
        style = self.styles['table-band-size-2']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(2, style.row_band_size)

    def test_table_col_band_size_none(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.col_band_size)

    def test_table_col_band_size_2(self):
        style = self.styles['table-band-size-2']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertEqual(2, style.col_band_size)

    def test_table_width_none(self):
        style = self.styles['Normal Table']
        docx_style = self.xml_elements[style.name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        self.assertIsNone(style.width)

    def test_table_width_auto(self):
        style = self.xml_style('table-width-auto')
        self.assertIsInstance(style.width, AutoLength)
        self.assertEqual(0, style.width)

    def test_table_width_25pct(self):
        style = self.xml_style('table-width-25pct')
        self.assertEqual(25, style.width.pct)

    def test_table_width_4in(self):
        style = self.xml_style('table-width-4in')
        self.assertEqual(4, style.width.inches)


class TestRowProperties(TestCase):

    docx_file_location = 'test_files/tables/docx/row_properties.docx'
    fragments_location = 'test_files/tables/docx/fragments/'

    def setUp(self):
        self.styles = {}
        parser = DocxParser(self.docx_file_location)
        stylesheet = parser.opc_package.styles
        for style in stylesheet.values():
            table_style = parser.parse_docx_table_style(style)
            self.styles[style.name] = table_style

    def test_alignment_center(self):
        style = self.styles['table-align-center']
        self.assertEqual('center', style.default_row.alignment)

    def test_alignment_right(self):
        style = self.styles['table-align-right']
        self.assertEqual('end', style.default_row.alignment)

    def test_row_cell_spacing_01in(self):
        filename = f'{self.fragments_location}table-row-cell-spacing-01in.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.1, style.default_row.cell_spacing.inches)

    def test_row_min_height_05in(self):
        filename = f'{self.fragments_location}table-row-height-min-05in.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.5, style.default_row.min_height.inches)

    def test_row_height_05in(self):
        filename = f'{self.fragments_location}table-row-height-05in.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.5, style.default_row.height.inches)

    def test_row_is_header_true(self):
        style = self.styles['table-header-true']
        self.assertTrue(style.default_row.is_header)

    def test_row_is_header_false(self):
        style = self.styles['table-header-false']
        self.assertFalse(style.default_row.is_header)

    def test_row_split_true(self):
        """Can split row"""
        style = self.styles['table-row-can-split']
        self.assertTrue(style.default_row.split)

    def test_row_split_false(self):
        """Can't split row"""
        style = self.styles['table-row-cant-split']
        self.assertFalse(style.default_row.split)


class TestCellProperties(TestCase):
    fragments_location = 'test_files/tables/docx/fragments/'

    def test_table_cell_background_color_none(self):
        filename = f'{self.fragments_location}table-cell-width-33pct.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.background_color)

    def test_table_cell_background_color_red(self):
        filename = f'{self.fragments_location}table-cell-fill-red.xml'
        style = load_xml_fragment(filename)
        self.assertEqual('#FF0000', style.default_cell.background_color)

    def test_table_cell_border_no_borders(self):
        filename = f'{self.fragments_location}table-cell-width-33pct.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.border_bottom)

    def test_table_cell_border_bottom_none(self):
        filename = f'{self.fragments_location}table-cell-borders-inside.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.border_bottom)

    def test_table_cell_border_bottom_05pt(self):
        filename = f'{self.fragments_location}table-cell-borders-outside.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.5, style.default_cell.border_bottom.width.pt)
        self.assertEqual('solid', style.default_cell.border_bottom.style)
        self.assertIsNone(style.default_cell.border_bottom.color)

    def test_table_cell_border_inside_horizontal_none(self):
        filename = f'{self.fragments_location}table-cell-borders-outside.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.border_inside_horizontal)

    def test_table_cell_border_inside_horizontal_05pt(self):
        filename = f'{self.fragments_location}table-cell-borders-inside.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.5, style.default_cell.border_inside_horizontal.width.pt)
        self.assertEqual('solid', style.default_cell.border_inside_horizontal.style)
        self.assertIsNone(style.default_cell.border_inside_horizontal.color)

    def test_table_cell_border_inside_vertical_none(self):
        filename = f'{self.fragments_location}table-cell-borders-outside.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.border_inside_vertical)

    def test_table_cell_border_inside_vertical_05pt(self):
        filename = f'{self.fragments_location}table-cell-borders-inside.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.5, style.default_cell.border_inside_vertical.width.pt)
        self.assertEqual('solid', style.default_cell.border_inside_vertical.style)
        self.assertIsNone(style.default_cell.border_inside_vertical.color)

    def test_table_cell_border_left_none(self):
        filename = f'{self.fragments_location}table-cell-borders-inside.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.border_left)

    def test_table_cell_border_left_05pt(self):
        filename = f'{self.fragments_location}table-cell-borders-outside.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.5, style.default_cell.border_left.width.pt)
        self.assertEqual('solid', style.default_cell.border_left.style)
        self.assertIsNone(style.default_cell.border_left.color)

    def test_table_cell_border_right_none(self):
        filename = f'{self.fragments_location}table-cell-borders-inside.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.border_right)

    def test_table_cell_border_right_05pt(self):
        filename = f'{self.fragments_location}table-cell-borders-outside.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.5, style.default_cell.border_right.width.pt)
        self.assertEqual('solid', style.default_cell.border_right.style)
        self.assertIsNone(style.default_cell.border_right.color)

    def test_table_cell_border_top_none(self):
        filename = f'{self.fragments_location}table-cell-borders-inside.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.border_top)

    def test_table_cell_border_top_05pt(self):
        filename = f'{self.fragments_location}table-cell-borders-outside.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.5, style.default_cell.border_top.width.pt)
        self.assertEqual('solid', style.default_cell.border_top.style)
        self.assertIsNone(style.default_cell.border_top.color)

    def test_cell_colspan_default(self):
        filename = f'{self.fragments_location}table-cell-width-33pct.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.colspan)

    def test_cell_colspan_2(self):
        filename = f'{self.fragments_location}table-cell-colspan-2.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(2, style.default_cell.colspan)

    def test_cell_fit_text_none(self):
        filename = f'{self.fragments_location}table-cell-colspan-2.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.fit_text)

    def test_cell_fit_text_true(self):
        filename = f'{self.fragments_location}table-cell-fit-text.xml'
        style = load_xml_fragment(filename)
        self.assertTrue(style.default_cell.fit_text)

    def test_table_cell_margin_bottom_none(self):
        filename = f'{self.fragments_location}table-cell-colspan-2.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.padding_bottom)

    def test_table_cell_margin_bottom_0_dxa(self):
        filename = f'{self.fragments_location}table-cell-margins-normal.xml'
        style = load_xml_fragment(filename)
        self.assertIsNotNone(style.default_cell.padding_bottom)
        self.assertEqual(0, style.default_cell.padding_bottom)

    def test_table_cell_margin_left_none(self):
        filename = f'{self.fragments_location}table-cell-colspan-2.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.padding_left)

    def test_table_cell_margin_left_108_dxa(self):
        filename = f'{self.fragments_location}table-cell-margins-normal.xml'
        style = load_xml_fragment(filename)
        self.assertIsNotNone(style.default_cell.padding_left)
        self.assertEqual(108, style.default_cell.padding_left.twips)

    def test_table_cell_margin_right_none(self):
        filename = f'{self.fragments_location}table-cell-colspan-2.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.padding_right)

    def test_table_cell_margin_right_108_dxa(self):
        filename = f'{self.fragments_location}table-cell-margins-normal.xml'
        style = load_xml_fragment(filename)
        self.assertIsNotNone(style.default_cell.padding_right)
        self.assertEqual(108, style.default_cell.padding_right.twips)

    def test_table_cell_margin_top_none(self):
        filename = f'{self.fragments_location}table-cell-colspan-2.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.padding_top)

    def test_table_cell_margin_top_0_dxa(self):
        filename = f'{self.fragments_location}table-cell-margins-normal.xml'
        style = load_xml_fragment(filename)
        self.assertIsNotNone(style.default_cell.padding_top)
        self.assertEqual(0, style.default_cell.padding_top)

    def test_table_cell_margins_01in(self):
        filename = f'{self.fragments_location}table-cell-margins-01-inch.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.1, style.default_cell.padding_bottom.inches)
        self.assertEqual(0.1, style.default_cell.padding_left.inches)
        self.assertEqual(0.1, style.default_cell.padding_right.inches)
        self.assertEqual(0.1, style.default_cell.padding_top.inches)

    def test_cell_valign_top(self):
        filename = f'{self.fragments_location}table-cell-valign-top.xml'
        style = load_xml_fragment(filename)
        self.assertEqual('top', style.default_cell.valign)

    def test_cell_valign_center(self):
        filename = f'{self.fragments_location}table-cell-valign-center.xml'
        style = load_xml_fragment(filename)
        self.assertEqual('center', style.default_cell.valign)

    def test_cell_valign_bottom(self):
        filename = f'{self.fragments_location}table-cell-valign-bottom.xml'
        style = load_xml_fragment(filename)
        self.assertEqual('bottom', style.default_cell.valign)

    def test_cell_width_33pct(self):
        filename = f'{self.fragments_location}table-cell-width-33pct.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(33, style.default_cell.width.pct)

    def test_cell_width_05in(self):
        filename = f'{self.fragments_location}table-cell-width-05in.xml'
        style = load_xml_fragment(filename)
        self.assertEqual(0.5, style.default_cell.width.inches)

    def test_cell_wrap_default(self):
        filename = f'{self.fragments_location}table-cell-width-05in.xml'
        style = load_xml_fragment(filename)
        self.assertIsNone(style.default_cell.wrap_text)

    def test_cell_wrap_nowrap(self):
        filename = f'{self.fragments_location}table-cell-nowrap.xml'
        style = load_xml_fragment(filename)
        self.assertFalse(style.default_cell.wrap_text)


class TestTableConditionalFormatting(TestCase):

    @classmethod
    def setUpClass(cls):
        docx_file = 'test_files/tables/docx/tables_conditional.docx'
        cls.styles = load_docx_styles(docx_file)

    def test_table_odd_banded_rows(self):
        style = self.styles['table-banded-rows']
        self.assertTrue(style.odd_rows.bold)

    def test_table_even_banded_rows(self):
        style = self.styles['table-banded-rows']
        self.assertTrue(style.even_rows.italics)
        self.assertTrue(style.even_rows.default_row.is_header)

    def test_table_odd_banded_cols(self):
        style = self.styles['table-banded-cols']
        result = style.odd_columns.default_cell.background_color
        self.assertEqual('#FFFF00', result)

    def test_table_even_banded_cols(self):
        style = self.styles['table-banded-cols']
        result = style.even_columns.default_cell.background_color
        self.assertEqual('#00B050', result)

    def test_table_banded_cols_border_top(self):
        style = self.styles['table-banded-cols-border-top']
        self.assertIsNone(style.odd_columns.border_top)
        self.assertIsNotNone(style.odd_columns.default_cell.border_top)

    def test_first_column(self):
        style = self.styles['table-first-col']
        spacing = CssUnit(7, 'twip')
        self.assertEqual('end', style.first_column.alignment)
        self.assertEqual(spacing, style.first_column.cell_spacing)
        self.assertTrue(style.first_column.italics)
        self.assertTrue(style.first_column.default_row.is_header)
        self.assertEqual(spacing, style.first_column.default_row.cell_spacing)
        self.assertEqual('end', style.first_column.default_row.alignment)

    def test_first_row(self):
        style = self.styles['table-first-row']
        self.assertTrue(style.first_row.italics)
        self.assertFalse(style.first_row.default_row.split)
        self.assertFalse(style.first_row.default_cell.wrap_text)

    def test_last_column(self):
        style = self.styles['table-last-col']
        self.assertTrue(style.last_column.italics)

    def test_last_row(self):
        style = self.styles['table-last-row']
        self.assertTrue(style.last_row.italics)

    def test_top_left_cell(self):
        style = self.styles['table-top-left-cell']
        cell = style.top_left_cell.default_cell
        self.assertEqual('#FFFF00', cell.background_color)

    def test_top_right_cell(self):
        style = self.styles['table-top-right-cell']
        self.assertTrue(style.top_right_cell.italics)

    def test_bottom_left_cell(self):
        style = self.styles['table-bottom-left-cell']
        self.assertTrue(style.bottom_left_cell.italics)

    def test_bottom_right_cell(self):
        style = self.styles['table-bottom-right-cell']
        self.assertTrue(style.bottom_right_cell.italics)

    def test_whole_table(self):
        pass


def print_style_properties(style):
    print('---------------------------------------------')
    print(f'Testing style "{style.id}" with following properties:')
    for prop in style.properties():
        print(f'   {prop.name} = {prop.value}')


class CssSerializerTestHarness(TestCase):

    css_files_location = 'test_files/tables/css/'

    def compare_style(self, style, css_filename):
        with open(f'{self.css_files_location}{css_filename}', 'r') as css_file:
            expected = css_file.read()
            stylesheet = Stylesheet()
            stylesheet.add_style(style)
            serializer = CssStylesheetSerializer(stylesheet, FACTORY)
            serializer.include_media_rules = False
            result = serializer.serialize()
            print_style_properties(style)
            print('\nResult:')
            print(result)
            print('---------------------------------------------')
            self.assertEqual(expected, result)


class TestCssTableSerializer(CssSerializerTestHarness):
    fragments_location = 'test_files/tables/docx/fragments/'

    def test_table_alignment_center(self):
        style = TableStyle(
            id='table-alignment-center',
            name='table-alignment-center',
            alignment='center',
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_alignment_right(self):
        style = TableStyle(
            id='table-alignment-right',
            name='table-alignment-right',
            alignment='end',
        )
        self.compare_style(style, f'{style.id}.css')

    def test_background_color_red(self):
        filename = 'table-fill-red'
        style_location = f'{self.fragments_location}{filename}.xml'
        style = load_xml_fragment(style_location)
        self.compare_style(style, f'{filename}.css')

    def test_table_borders_0(self):
        name = 'table-borders-0'
        style = TableStyle(
            id=name,
            name=name,
            border_bottom=Border(width=CssUnit(0)),
            border_left=Border(width=CssUnit(0)),
            border_right=Border(width=CssUnit(0)),
            border_top=Border(width=CssUnit(0)),
        )
        self.compare_style(style, f'{style.id}.css')
        
    def test_table_border_bottom_05pt(self):
        name = 'table-border-bottom-05pt'
        style = TableStyle(
            id=name,
            name=name,
            border_bottom=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_border_inside_vertical_05pt(self):
        name = 'table-border-inside-vertical-05pt'
        style = TableStyle(
            id=name,
            name=name,
            border_bottom=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
            border_inside_vertical=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_border_inside_horizontal_05pt(self):
        name = 'table-border-inside-horizontal-05pt'
        style = TableStyle(
            id=name,
            name=name,
            border_bottom=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
            border_inside_horizontal=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_border_inside_05pt(self):
        name = 'table-border-inside-05pt'
        style = TableStyle(
            id=name,
            name=name,
            border_inside_vertical=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
            border_inside_horizontal=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_border_left_05pt(self):
        name = 'table-border-left-05pt'
        style = TableStyle(
            id=name,
            name=name,
            border_left=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_border_right_05pt(self):
        name = 'table-border-right-05pt'
        style = TableStyle(
            id=name,
            name=name,
            border_right=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_border_top_05pt(self):
        name = 'table-border-top-05pt'
        style = TableStyle(
            id=name,
            name=name,
            border_top=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_cell_padding_bottom_0_dxa(self):
        name = 'table-cell-padding-bottom-0'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_bottom=CssUnit(0),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_cell_padding_left_120_dxa(self):
        name = 'table-cell-padding-left-120dxa'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_left=CssUnit(120, 'twip'),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_cell_padding_right_120_dxa(self):
        name = 'table-cell-padding-right-120dxa'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_right=CssUnit(120, 'twip'),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_cell_padding_top_0_dxa(self):
        name = 'table-cell-padding-top-0'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_top=CssUnit(0),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_cell_paddings_01in(self):
        name = 'table-cell-paddings-01in'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_bottom=CssUnit(0.1, 'in'),
            cell_padding_left=CssUnit(0.1, 'in'),
            cell_padding_right=CssUnit(0.1, 'in'),
            cell_padding_top=CssUnit(0.1, 'in'),
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_spacing_0(self):
        style = TableStyle(
            id='table-cell-spacing-0',
            name='table-cell-spacing-0',
            cell_spacing=CssUnit(0, 'pt')
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_spacing_12pt(self):
        style = TableStyle(
            id='table-cell-spacing-12pt',
            name='table-cell-spacing-12pt',
            cell_spacing=CssUnit(12, 'pt')
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_left_indent_1in(self):
        style = TableStyle(
            id='table-left-indent-1in',
            name='table-left-indent-1in',
            indent=CssUnit(1, 'in')
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_layout_auto(self):
        style = TableStyle(
            id='table-layout-auto',
            name='table-layout-auto',
            layout='auto',
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_layout_fixed(self):
        style = TableStyle(
            id='table-layout-fixed',
            name='table-layout-fixed',
            layout='fixed',
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_width_auto(self):
        style = TableStyle(
            id='table-width-auto',
            name='table-width-auto',
            width=AutoLength()
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_width_25pct(self):
        style = TableStyle(
            id='table-width-25pct',
            name='table-width-25pct',
            width=Percentage(25)
        )
        self.compare_style(style, f'{style.id}.css')

    def test_table_width_4in(self):
        style = TableStyle(
            id='table-width-4in',
            name='table-width-4in',
            width=CssUnit(4, 'in')
        )
        self.compare_style(style, f'{style.id}.css')


class TestCssTableCellSerializer(CssSerializerTestHarness):
    fragments_location = 'test_files/tables/docx/fragments/'

    def test_cell_background_color_red(self):
        filename = 'table-cell-fill-red'
        style_location = f'{self.fragments_location}{filename}.xml'
        style = load_xml_fragment(style_location)
        self.compare_style(style, f'{filename}.css')

    def test_cell_borders_0(self):
        default_cell = TableCellProperties(
            border_bottom=Border(width=CssUnit(0)),
            border_left=Border(width=CssUnit(0)),
            border_right=Border(width=CssUnit(0)),
            border_top=Border(width=CssUnit(0)),
        )
        name = 'cell-borders-0'
        style = TableStyle(
            id=name,
            name=name,
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_border_bottom_05pt(self):
        default_cell = TableCellProperties(
            border_bottom=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        name = 'cell-border-bottom-05pt'
        style = TableStyle(
            id=name,
            name=name,
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_border_inside_vertical_05pt(self):
        default_cell = TableCellProperties(
            border_bottom=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
            border_inside_vertical=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        name = 'cell-border-inside-vertical-05pt'
        style = TableStyle(
            id=name,
            name=name,
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_border_inside_horizontal_05pt(self):
        default_cell = TableCellProperties(
            border_bottom=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
            border_inside_horizontal=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        name = 'cell-border-inside-horizontal-05pt'
        style = TableStyle(
            id=name,
            name=name,
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_border_inside_05pt(self):
        default_cell = TableCellProperties(
            border_inside_vertical=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
            border_inside_horizontal=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        name = 'cell-border-inside-05pt'
        style = TableStyle(
            id=name,
            name=name,
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_border_left_05pt(self):
        default_cell = TableCellProperties(
            border_left=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        name = 'cell-border-left-05pt'
        style = TableStyle(
            id=name,
            name=name,
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_border_right_05pt(self):
        default_cell = TableCellProperties(
            border_right=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        name = 'cell-border-right-05pt'
        style = TableStyle(
            id=name,
            name=name,
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_border_top_05pt(self):
        default_cell = TableCellProperties(
            border_top=Border(
                width=CssUnit(0.5, 'pt'),
                style='solid',
            ),
        )
        name = 'cell-border-top-05pt'
        style = TableStyle(
            id=name,
            name=name,
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_padding_bottom_0_dxa(self):
        default_cell = TableCellProperties(
            padding_bottom=CssUnit(0),
        )
        name = 'table-cell-padding-bottom-0'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_bottom=CssUnit(6, 'pt'),
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_padding_left_120_dxa(self):
        default_cell = TableCellProperties(
            padding_left=CssUnit(120, 'twip'),
        )
        name = 'table-cell-padding-left-120dxa'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_left=CssUnit(60, 'twip'),
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_padding_right_120_dxa(self):
        default_cell = TableCellProperties(
            padding_right=CssUnit(120, 'twip'),
        )
        name = 'table-cell-padding-right-120dxa'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_right=CssUnit(60, 'twip'),
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_padding_top_0_dxa(self):
        default_cell = TableCellProperties(
            padding_top=CssUnit(0),
        )
        name = 'table-cell-padding-top-0'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_top=CssUnit(120, 'twip'),
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_paddings_01in(self):
        default_cell = TableCellProperties(
            padding_bottom=CssUnit(0.1, 'in'),
            padding_left=CssUnit(0.1, 'in'),
            padding_right=CssUnit(0.1, 'in'),
            padding_top=CssUnit(0.1, 'in'),
        )
        name = 'table-cell-paddings-01in'
        style = TableStyle(
            id=name,
            name=name,
            cell_padding_bottom=CssUnit(0.2, 'in'),
            cell_padding_left=CssUnit(0.2, 'in'),
            cell_padding_right=CssUnit(0.2, 'in'),
            cell_padding_top=CssUnit(0.2, 'in'),
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_valign_top(self):
        default_cell = TableCellProperties(
            valign='top',
        )
        style = TableStyle(
            id='cell-valign-top',
            name='cell-valign-top',
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_valign_center(self):
        default_cell = TableCellProperties(
            valign='center',
        )
        style = TableStyle(
            id='cell-valign-center',
            name='cell-valign-center',
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_valign_bottom(self):
        default_cell = TableCellProperties(
            valign='bottom',
        )
        style = TableStyle(
            id='cell-valign-bottom',
            name='cell-valign-bottom',
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_width_33pct(self):
        default_cell = TableCellProperties(
            width=Percentage(33),
        )
        style = TableStyle(
            id='cell-width-33pct',
            name='cell-width-33pct',
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_width_05in(self):
        default_cell = TableCellProperties(
            width=CssUnit(0.5, 'in'),
        )
        style = TableStyle(
            id='cell-width-05in',
            name='cell-width-05in',
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')

    def test_cell_wrap_text_true(self):
        default_cell = TableCellProperties(
            wrap_text=True,
        )
        style = TableStyle(
            id='cell-wrap-text-true',
            name='cell-wrap-text-true',
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')
    
    def test_cell_wrap_text_false(self):
        default_cell = TableCellProperties(
            wrap_text=False,
        )
        style = TableStyle(
            id='cell-wrap-text-false',
            name='cell-wrap-text-false',
            default_cell=default_cell,
        )
        self.compare_style(style, f'{style.id}.css')


class TestCssTableRowSerializer(CssSerializerTestHarness):

    def test_row_is_header_true(self):
        default_row = TableRowProperties(
            is_header=True
        )
        name = 'row-is-header-true'
        style = TableStyle(
            id=name,
            name=name,
            default_row=default_row
        )
        self.compare_style(style, f'{style.id}.css')

    def test_row_is_header_false(self):
        default_row = TableRowProperties(
            is_header=False
        )
        name = 'row-is-header-false'
        style = TableStyle(
            id=name,
            name=name,
            default_row=default_row
        )
        self.compare_style(style, f'{style.id}.css')

    def test_row_split_true(self):
        default_row = TableRowProperties(
            split=True
        )
        name = 'row-split-true'
        style = TableStyle(
            id=name,
            name=name,
            default_row=default_row
        )
        self.compare_style(style, f'{style.id}.css')

    def test_row_split_false(self):
        default_row = TableRowProperties(
            split=False
        )
        name = 'row-split-false'
        style = TableStyle(
            id=name,
            name=name,
            default_row=default_row
        )
        self.compare_style(style, f'{style.id}.css')

    def test_row_height_1in(self):
        default_row = TableRowProperties(
            min_height=CssUnit(1, 'in'),
        )
        name = 'row-height-1in'
        style = TableStyle(
            id=name,
            name=name,
            default_row=default_row
        )
        self.compare_style(style, f'{style.id}.css')

    def test_row_min_height_1in(self):
        default_row = TableRowProperties(
            min_height=CssUnit(1, 'in'),
        )
        name = 'row-min-height-1in'
        style = TableStyle(
            id=name,
            name=name,
            default_row=default_row
        )
        self.compare_style(style, f'{style.id}.css')


class TestCssTableConditionalFormatting(CssSerializerTestHarness):

    @classmethod
    def setUpClass(cls):
        docx_file = 'test_files/tables/docx/tables_conditional.docx'
        cls.styles = load_docx_styles(docx_file)

    def test_odd_rows_selector(self):
        name = 'table-odd-rows'
        table = TableStyle(
            id=name,
            name=name,
        )
        selector = 'table.table-odd-rows tr:nth-child(2n+1)'
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.odd_row_selector())

    def test_even_rows_selector(self):
        name = 'even-rows'
        table = TableStyle(
            id=name,
            name=name,
        )
        selector = 'table.even-rows tr:nth-child(2n+2)'
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.even_row_selector())

    def test_2_odd_rows_selector(self):
        name = 'odd-rows'
        table = TableStyle(
            id=name,
            name=name,
            row_band_size=2,
        )
        selector = ('table.odd-rows tr:nth-child(4n+1), '
                    'table.odd-rows tr:nth-child(4n+2)')
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.odd_row_selector())

    def test_2_even_rows_selector(self):
        name = 'even-rows'
        table = TableStyle(
            id=name,
            name=name,
            row_band_size=2,
        )
        selector = ('table.even-rows tr:nth-child(4n+3), '
                    'table.even-rows tr:nth-child(4n+4)')
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.even_row_selector())

    def test_3_odd_rows_selector(self):
        name = 'odd-rows'
        table = TableStyle(
            id=name,
            name=name,
            row_band_size=3,
        )
        selector = ('table.odd-rows tr:nth-child(6n+1), '
                    'table.odd-rows tr:nth-child(6n+2), '
                    'table.odd-rows tr:nth-child(6n+3)')
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.odd_row_selector())

    def test_3_even_rows_selector(self):
        name = 'even-rows'
        table = TableStyle(
            id=name,
            name=name,
            row_band_size=3,
        )
        selector = ('table.even-rows tr:nth-child(6n+4), '
                    'table.even-rows tr:nth-child(6n+5), '
                    'table.even-rows tr:nth-child(6n+6)')
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.even_row_selector())

    def test_odd_cols_selector(self):
        name = 'table-odd-cols'
        table = TableStyle(
            id=name,
            name=name,
        )
        selector = 'table.table-odd-cols tr td:nth-child(2n+1)'
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.odd_column_selector())

    def test_even_cols_selector(self):
        name = 'even-cols'
        table = TableStyle(
            id=name,
            name=name,
        )
        selector = 'table.even-cols tr td:nth-child(2n+2)'
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.even_column_selector())

    def test_2_odd_cols_selector(self):
        name = 'odd-cols'
        table = TableStyle(
            id=name,
            name=name,
            col_band_size=2,
        )
        selector = ('table.odd-cols tr td:nth-child(4n+1), '
                    'table.odd-cols tr td:nth-child(4n+2)')
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.odd_column_selector())

    def test_2_even_cols_selector(self):
        name = 'even-cols'
        table = TableStyle(
            id=name,
            name=name,
            col_band_size=2,
        )
        selector = ('table.even-cols tr td:nth-child(4n+3), '
                    'table.even-cols tr td:nth-child(4n+4)')
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.even_column_selector())

    def test_3_odd_cols_selector(self):
        name = 'odd-cols'
        table = TableStyle(
            id=name,
            name=name,
            col_band_size=3,
        )
        selector = ('table.odd-cols tr td:nth-child(6n+1), '
                    'table.odd-cols tr td:nth-child(6n+2), '
                    'table.odd-cols tr td:nth-child(6n+3)')
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.odd_column_selector())

    def test_3_even_cols_selector(self):
        name = 'even-cols'
        table = TableStyle(
            id=name,
            name=name,
            col_band_size=3,
        )
        selector = ('table.even-cols tr td:nth-child(6n+4), '
                    'table.even-cols tr td:nth-child(6n+5), '
                    'table.even-cols tr td:nth-child(6n+6)')
        serializer = CssTableSerializer(table, FACTORY)
        self.assertEqual(selector, serializer.even_column_selector())

    def test_table_banded_rows(self):
        name = 'table-banded-rows'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_table_banded_rows_default_cell(self):
        name = 'table-banded-rows-valign-center'
        style = TableStyle(
            id=name,
            name=name,
            odd_rows=TableConditionalFormatting(
                bold=True,
                default_cell=TableCellProperties(
                    valign='center',
                ),
                default_row=TableRowProperties(
                    min_height=CssUnit(1, 'in'),
                ),
            ),
        )
        self.compare_style(style, f'{name}.css')

    def test_table_2_banded_rows_default_cell(self):
        name = 'table-banded-rows-wrap-text'
        style = TableStyle(
            id=name,
            name=name,
            row_band_size=2,
            even_rows=TableConditionalFormatting(
                bold=True,
                default_cell=TableCellProperties(
                    wrap_text=True,
                ),
                default_row=TableRowProperties(
                    min_height=CssUnit(1, 'in'),
                ),
            ),
        )
        self.compare_style(style, f'{name}.css')

    def test_table_banded_cols(self):
        name = 'table-banded-cols'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_table_banded_cols_border_bottom(self):
        name = 'table-banded-cols-border-bottom'
        style = self.styles[name]
        self.assertIsNone(style.odd_columns.border_bottom)
        self.compare_style(style, f'{name}.css')

    def test_table_banded_cols_border_left(self):
        name = 'table-banded-cols-border-left'
        style = self.styles[name]
        self.assertIsNone(style.odd_columns.border_left)
        self.compare_style(style, f'{name}.css')

    def test_table_banded_cols_border_right(self):
        name = 'table-banded-cols-border-right'
        style = self.styles[name]
        self.assertIsNone(style.odd_columns.border_right)
        self.compare_style(style, f'{name}.css')

    def test_table_banded_cols_border_top(self):
        name = 'table-banded-cols-border-top'
        style = self.styles[name]
        self.assertIsNone(style.odd_columns.border_top)
        self.compare_style(style, f'{name}.css')

    def test_table_banded_cols_border_inside_horizontal(self):
        name = 'table-banded-cols-border-inside-horizontal'
        style = self.styles[name]
        self.assertIsNone(style.odd_columns.border_inside_horizontal)
        self.compare_style(style, f'{name}.css')

    def test_table_odd_column_inside_vertical_selector_1(self):
        table = TableStyle(
            id='table',
            name='table',
            col_band_size=1,
        )
        serializer = CssTableSerializer(table, FACTORY)
        expected = ''
        self.assertEqual(expected, serializer.column_inside_vertical_selector())

    def test_table_odd_column_inside_vertical_selector_2(self):
        table = TableStyle(
            id='table',
            name='table',
            col_band_size=2,
        )
        serializer = CssTableSerializer(table, FACTORY)
        expected = (
            'table.table tr td:nth-child(4n+1) + td:nth-child(4n+2)'
        )
        self.assertEqual(expected, serializer.column_inside_vertical_selector())

    def test_table_odd_column_inside_vertical_selector_3(self):
        table = TableStyle(
            id='table',
            name='table',
            col_band_size=3,
        )
        serializer = CssTableSerializer(table, FACTORY)
        expected = (
            'table.table tr td:nth-child(6n+1) + td:nth-child(6n+2), '
            'table.table tr td:nth-child(6n+2) + td:nth-child(6n+3)'
        )
        self.assertEqual(expected, serializer.column_inside_vertical_selector())

    def test_table_even_column_inside_vertical_selector_1(self):
        table = TableStyle(
            id='table',
            name='table',
            col_band_size=1,
        )
        serializer = CssTableSerializer(table, FACTORY)
        expected = ''
        result = serializer.column_inside_vertical_selector(odd=False)
        self.assertEqual(expected, result)

    def test_table_even_column_inside_vertical_selector_2(self):
        table = TableStyle(
            id='table',
            name='table',
            col_band_size=2,
        )
        serializer = CssTableSerializer(table, FACTORY)
        expected = (
            'table.table tr td:nth-child(4n+3) + td:nth-child(4n+4)'
        )
        result = serializer.column_inside_vertical_selector(odd=False)
        self.assertEqual(expected, result)

    def test_table_even_column_inside_vertical_selector_3(self):
        table = TableStyle(
            id='table',
            name='table',
            col_band_size=3,
        )
        serializer = CssTableSerializer(table, FACTORY)
        expected = (
            'table.table tr td:nth-child(6n+4) + td:nth-child(6n+5), '
            'table.table tr td:nth-child(6n+5) + td:nth-child(6n+6)'
        )
        result = serializer.column_inside_vertical_selector(odd=False)
        self.assertEqual(expected, result)

    def test_table_banded_cols_border_inside_vertical(self):
        name = 'table-banded-cols-border-inside-vertical'
        style = self.styles[name]
        self.assertIsNone(style.odd_columns.border_inside_vertical)
        self.compare_style(style, f'{name}.css')

    def test_table_banded_rows_border_outside(self):
        name = 'table-banded-rows-border-outside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_table_banded_rows_border_inside(self):
        name = 'table-banded-rows-border-inside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_table_top_left_cell_selector(self):
        name = 'table'
        table = TableStyle(
            id=name,
            name=name,
        )
        style = CssTableSerializer(table, FACTORY)
        expected = 'table.table tr:first-of-type td:first-of-type'
        self.assertEqual(expected, style.top_left_cell_selector())

    def test_table_top_right_cell_selector(self):
        name = 'table'
        table = TableStyle(
            id=name,
            name=name,
        )
        style = CssTableSerializer(table, FACTORY)
        expected = 'table.table tr:first-of-type td:last-of-type'
        self.assertEqual(expected, style.top_right_cell_selector())

    def test_table_bottom_left_cell_selector(self):
        name = 'table'
        table = TableStyle(
            id=name,
            name=name,
        )
        style = CssTableSerializer(table, FACTORY)
        expected = 'table.table tr:last-of-type td:first-of-type'
        self.assertEqual(expected, style.bottom_left_cell_selector())

    def test_table_bottom_right_cell_selector(self):
        name = 'table'
        table = TableStyle(
            id=name,
            name=name,
        )
        style = CssTableSerializer(table, FACTORY)
        expected = 'table.table tr:last-of-type td:last-of-type'
        self.assertEqual(expected, style.bottom_right_cell_selector())

    def test_first_column(self):
        name = 'table-first-col'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_first_column_border_outside(self):
        name = 'table-first-col-border-outside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_first_column_border_inside(self):
        name = 'table-first-col-border-inside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_first_row(self):
        name = 'table-first-row'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_first_row_border_outside(self):
        name = 'table-first-row-border-outside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_first_row_border_inside(self):
        name = 'table-first-row-border-inside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_last_column(self):
        name = 'table-last-col'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_last_column_border_outside(self):
        name = 'table-last-col-border-outside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_last_column_border_inside(self):
        name = 'table-last-col-border-inside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_last_row(self):
        name = 'table-last-row'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_last_row_border_outside(self):
        name = 'table-last-row-border-outside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_last_row_border_inside(self):
        name = 'table-last-row-border-inside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_top_left_cell(self):
        name = 'table-top-left-cell'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_top_left_cell_border_outside(self):
        name = 'table-top-left-cell-border-outside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_top_left_cell_border_inside(self):
        name = 'table-top-left-cell-border-inside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_top_right_cell(self):
        name = 'table-top-right-cell'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_top_right_cell_border_outside(self):
        name = 'table-top-right-cell-border-outside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_top_right_cell_border_inside(self):
        name = 'table-top-right-cell-border-inside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_bottom_left_cell(self):
        name = 'table-bottom-left-cell'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_bottom_left_cell_border_outside(self):
        name = 'table-bottom-left-cell-border-outside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_bottom_left_cell_border_inside(self):
        name = 'table-bottom-left-cell-border-inside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_bottom_right_cell(self):
        name = 'table-bottom-right-cell'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_bottom_right_cell_border_outside(self):
        name = 'table-bottom-right-cell-border-outside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')

    def test_bottom_right_cell_border_inside(self):
        name = 'table-bottom-right-cell-border-inside'
        style = self.styles[name]
        self.compare_style(style, f'{name}.css')
