import logging
from unittest import TestCase
import warnings

import cssutils
from lxml import etree

from docx2css.css.serializers import FACTORY, CssStylesheetSerializer
from docx2css.ooxml.package import OpcPackage
from docx2css.ooxml.styles import Styles
from docx2css.ooxml.parsers import DocxParser
from docx2css.utils import CSSColor, CssUnit


cssutils.ser.prefs.indentClosingBrace = False
cssutils.ser.prefs.omitLastSemicolon = False
logging.basicConfig(format='%(filename)s:%(lineno)d %(message)s',
                    level=logging.DEBUG)


def print_styles_tree(docx_filename):
    parser = OpcPackage(docx_filename)
    styles = parser.styles
    root_styles = (
        s for s in styles.values() if s.parent_id is None
    )
    print(f'Styles in {docx_filename}:')
    print(styles.print_styles_tree(root_styles))


class TestHarness(TestCase):

    files = ()
    css_files_location = None
    docx_files_location = None
    parse_numbering_instances = False

    @classmethod
    def setUpClass(cls):
        cls.styles = {}
        cls.api_styles = {}
        # cls.stylesheet = docx2css.stylesheet.Stylesheet()
        for file in cls.files:
            parser = DocxParser(f'{cls.docx_files_location}{file}')
            docx_styles = parser.opc_package.styles
            if cls.parse_numbering_instances:
                parser.parse_numbering()
            for docx_style in docx_styles.values():
                # Sanity check
                if docx_style.name in cls.styles.keys():
                    msg = f'ERROR: Style "{docx_style.name}" already exists!'
                    warnings.warn(msg)
                cls.styles[docx_style.name] = docx_style
                api_style = parser.parse_docx_style(docx_style)
                if api_style:
                    cls.api_styles[api_style.id] = api_style
                #     cls.stylesheet.add_style(api_style)

    def single_style_css(self, style_id):
        """Get the CSS text equivalent to the api_style provided as arg"""
        api_style = self.api_styles.get(style_id)
        serializer = FACTORY.get_block_serializer(api_style)
        css_stylesheet = cssutils.css.CSSStyleSheet()
        for rule in serializer.css_style_rules():
            css_stylesheet.add(rule)
        return css_stylesheet.cssText.decode('utf-8')

    def style_test_harness(self, style_name):
        docx_style = self.styles[style_name]
        print(etree.tostring(docx_style, pretty_print=True).decode('utf-8'))
        css_file = f'{self.css_files_location}{"".join(style_name.split())}.css'
        self.compare_style(docx_style.id, css_file)

    def compare_style(self, style_id, css_filename):
        with open(css_filename, 'r') as css_file:
            expected = css_file.read()
            css_text = self.single_style_css(style_id)
            print(css_text)
            self.assertEqual(expected, css_text)

    def style_color_test_harness(self,
                                 style_name,
                                 rule_name='color',
                                 tolerance=2):
        """
        Test the font color by making sure the RGB hex value for each
        component is within a certain tolerance (2 by default)

        :param rule_name:
        :param style_name:
        :param tolerance:
        :return:
        """
        style = self.styles[style_name]
        print(etree.tostring(style, pretty_print=True).decode('utf-8'))
        api_style = self.api_styles.get(style.id)
        serializer = FACTORY.get_block_serializer(api_style)
        with open(f'{self.css_files_location}{"".join(style_name.split())}.css', 'r') as css_file:
            expected = cssutils.parseString(css_file.read()).cssRules[0]
            expected_color = expected.style[rule_name]
            result = list(serializer.css_style_rules())[0]
            result_color = result.style[rule_name]

            # Compare selector
            self.assertEqual(expected.selectorText, result.selectorText)

            # Compare properties
            self.assertEqual(len(list(expected.style)), len(list(result.style)))
            for eprop, rprop in zip(expected.style, result.style):
                if eprop.name != rule_name:
                    self.assertEqual(eprop.cssText, rprop.cssText)

            # Compare color
            er, eg, eb = CSSColor.split_rgb(expected_color[1:])
            rr, rg, rb = CSSColor.split_rgb(result_color[1:])
            self.assertAlmostEqual(int(er, 16), int(rr, 16), delta=tolerance)
            self.assertAlmostEqual(int(eg, 16), int(rg, 16), delta=tolerance)
            self.assertAlmostEqual(int(eb, 16), int(rb, 16), delta=tolerance)


class StylesTestCase(TestCase):

    def test_iter(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        styles = Styles(package)
        for s in styles:
            print(s)
        self.assertEqual(27, len(styles))
        self.assertEqual('allegations-L1', styles['allegations-L1'].id)

    def test_hierarchy(self):
        docx_files_location = 'test_files/character/docx/'
        parser = DocxParser(f'{docx_files_location}character_styles.docx')
        docx_styles = parser.opc_package.styles

        parent_id = 'boldcharstyle'
        parent = docx_styles[parent_id]
        child1_id = 'bolditaliccharstyle'
        child2_id = 'notboldcharstyle'
        child1 = docx_styles[child1_id]
        child2 = docx_styles[child2_id]
        self.assertEqual(2, len(parent.children_styles))
        self.assertEqual(parent_id, child1.parent_id)
        self.assertEqual(parent_id, child2.parent_id)

        api_parent = parser.parse_docx_style(parent)
        parser.parse_docx_style(child1)
        parser.parse_docx_style(child2)

        self.assertEqual(2, len(api_parent.children))


class CharacterStylesTestCase(TestHarness):
    files = (
        'character_styles.docx',
        'border_character_styles.docx',
        'font_character_styles.docx',
        'highlight_char_styles.docx',
        'legacy_char_styles.docx',
        'color_char_styles.docx',
    )
    css_files_location = 'test_files/character/css/'
    docx_files_location = 'test_files/character/docx/'

    ##########################################################################
    #                                                                        #
    # Bold                                                                   #
    #                                                                        #
    ##########################################################################
    def test_bold_char_style(self):
        """Bold is active"""
        self.style_test_harness('bold_char_style')

    def test_not_bold_char_style(self):
        """Child of bold_char_style. Bold is toggled to off"""
        self.style_test_harness('not_bold_char_style')

    ##########################################################################
    #                                                                        #
    # Border Character Styles                                                #
    #                                                                        #
    ##########################################################################
    def test_border_char_style(self):
        self.style_test_harness('border_char_style')

    def test_border_rgb_char_style(self):
        self.style_test_harness('border_rgb_char_style')

    def test_border_shadow_char_style(self):
        self.style_test_harness('border_shadow_char_style')

    def test_border_thinThickSmallGap_3d_char_style(self):
        self.style_test_harness('border_thinThickSmallGap_3d_char_style')

    def test_border_1pt_char_style(self):
        self.style_test_harness('border_1pt_char_style')

    def test_border_theme_char_style(self):
        self.style_test_harness('border_theme_char_style')

    def test_border_theme_lighter_char_style(self):
        self.style_color_test_harness('border_theme_lighter_char_style',
                                      rule_name='border-color')

    def test_border_theme_darker_char_style(self):
        self.style_color_test_harness('border_theme_darker_char_style',
                                      rule_name='border-color')

    def test_no_border_char_style(self):
        self.style_test_harness('no_border_char_style')

    def test_border_dashDotStroked_char_style(self):
        self.style_test_harness('border_dashDotStroked_char_style')

    def test_border_dashSmallGap_char_style(self):
        self.style_test_harness('border_dashSmallGap_char_style')

    def test_border_dotDash_char_style(self):
        self.style_test_harness('border_dotDash_char_style')

    def test_border_dotDotDash_char_style(self):
        self.style_test_harness('border_dotDotDash_char_style')

    def test_border_doubleWave_char_style(self):
        self.style_test_harness('border_doubleWave_char_style')

    def test_border_nil_char_style(self):
        self.style_test_harness('border_nil_char_style')

    def test_border_thick_char_style(self):
        self.style_test_harness('border_thick_char_style')

    def test_border_wave_char_style(self):
        self.style_test_harness('border_wave_char_style')

    def test_border_thickThinLargeGap_char_style(self):
        self.style_test_harness('border_thickThinLargeGap_char_style')

    def test_border_thickThinMediumGap_char_style(self):
        self.style_test_harness('border_thickThinMediumGap_char_style')

    def test_border_thickThinSmallGap_char_style(self):
        self.style_test_harness('border_thickThinSmallGap_char_style')

    def test_border_thinThickLargeGap_char_style(self):
        self.style_test_harness('border_thinThickLargeGap_char_style')

    def test_border_thinThickMediumGap_char_style(self):
        self.style_test_harness('border_thinThickMediumGap_char_style')

    def test_border_thinThickSmallGap_char_style(self):
        self.style_test_harness('border_thinThickSmallGap_char_style')

    def test_border_thinThickThinLargeGap_char_style(self):
        self.style_test_harness('border_thinThickThinLargeGap_char_style')

    def test_border_thinThickThinMediumGap_char_style(self):
        self.style_test_harness('border_thinThickThinMediumGap_char_style')

    def test_border_thinThickThinSmallGap_char_style(self):
        self.style_test_harness('border_thinThickThinSmallGap_char_style')

    def test_border_triple_char_style(self):
        self.style_test_harness('border_triple_char_style')

    def test_border_dashed_char_style(self):
        self.style_test_harness('border_dashed_char_style')

    def test_border_dotted_char_style(self):
        self.style_test_harness('border_dotted_char_style')

    def test_border_double_char_style(self):
        self.style_test_harness('border_double_char_style')

    def test_border_inset_char_style(self):
        self.style_test_harness('border_inset_char_style')

    def test_border_outset_char_style(self):
        self.style_test_harness('border_outset_char_style')

    def test_border_threeDEmboss_char_style(self):
        self.style_test_harness('border_threeDEmboss_char_style')

    def test_border_threeDEngrave_char_style(self):
        self.style_test_harness('border_threeDEngrave_char_style')

    ##########################################################################
    #                                                                        #
    # Emboss                                                                 #
    #                                                                        #
    ##########################################################################
    def test_emboss_char_style(self):
        self.style_test_harness('emboss_char_style')

    def test_not_emboss_char_style(self):
        self.style_test_harness('not_emboss_char_style')

    ##########################################################################
    #                                                                        #
    # Fonts                                                                  #
    #                                                                        #
    ##########################################################################
    def test_font_arial_char_style(self):
        self.style_test_harness('font_arial_char_style')

    def test_font_timesNewRoman_char_style(self):
        self.style_test_harness('font_timesNewRoman_char_style')

    def test_font_body_char_style(self):
        self.style_test_harness('font_body_char_style')

    def test_font_hansi_char_style(self):
        self.style_test_harness('font_hansi_char_style')

    ##########################################################################
    #                                                                        #
    # Imprint                                                                #
    #                                                                        #
    ##########################################################################
    def test_imprint_char_style(self):
        self.style_test_harness('imprint_char_style')

    def test_not_imprint_char_style(self):
        self.style_test_harness('not_imprint_char_style')

    ##########################################################################
    #                                                                        #
    # Italic                                                                 #
    #                                                                        #
    ##########################################################################
    def test_italic_char_style(self):
        """Italic is active"""
        self.style_test_harness('italic_char_style')

    def test_not_italic_char_style(self):
        """Child of italic_char_style. Italic is toggled to off"""
        self.style_test_harness('not_italic_char_style')

    def test_bold_italic_char_style(self):
        """Child of bold_char_style with italic added"""
        self.style_test_harness('bold_italic_char_style')

    ##########################################################################
    #                                                                        #
    # Caps                                                                   #
    #                                                                        #
    ##########################################################################
    def test_caps_char_style(self):
        """Text in uppercase"""
        self.style_test_harness('caps_char_style')

    def test_not_caps_char_style(self):
        """Child of caps_char_style. Uppercase it toggled off"""
        self.style_test_harness('not_caps_char_style')

    ##########################################################################
    #                                                                        #
    # Font Color                                                             #
    #                                                                        #
    ##########################################################################
    def test_color_rgb_char_style(self):
        self.style_test_harness('color_rgb_char_style')

    def test_color_theme_accent2_char_style(self):
        self.style_test_harness('color_theme_accent2_char_style')

    def test_color_theme_accent2_25darker_char_style(self):
        self.style_color_test_harness('color_theme_accent2_25darker_char_style')

    def test_color_theme_accent2_50darker_char_style(self):
        self.style_color_test_harness('color_theme_accent2_50darker_char_style')

    def test_color_theme_accent2_40lighter_char_style(self):
        self.style_color_test_harness('color_theme_accent2_40lighter_char_style')

    def test_color_theme_accent2_60lighter_char_style(self):
        self.style_color_test_harness('color_theme_accent2_60lighter_char_style')

    def test_color_theme_accent2_80lighter_char_style(self):
        self.style_color_test_harness('color_theme_accent2_80lighter_char_style')

    def test_color_theme_accent1_25darker_char_style(self):
        self.style_color_test_harness('color_theme_accent1_25darker_char_style')

    def test_color_theme_accent1_50darker_char_style(self):
        self.style_color_test_harness('color_theme_accent1_50darker_char_style')

    def test_color_theme_accent1_40lighter_char_style(self):
        self.style_color_test_harness('color_theme_accent1_40lighter_char_style')

    def test_color_theme_accent1_60lighter_char_style(self):
        self.style_color_test_harness('color_theme_accent1_60lighter_char_style')

    def test_color_theme_accent1_80lighter_char_style(self):
        self.style_color_test_harness('color_theme_accent1_80lighter_char_style')

    ##########################################################################
    #                                                                        #
    # dStrike                                                                #
    #                                                                        #
    ##########################################################################
    def test_dstrike_char_style(self):
        self.style_test_harness('dstrike_char_style')

    def test_not_dstrike_char_style(self):
        self.style_test_harness('not_dstrike_char_style')

    ##########################################################################
    #                                                                        #
    # Strike                                                                 #
    #                                                                        #
    ##########################################################################
    def test_strike_char_style(self):
        self.style_test_harness('strike_char_style')

    def test_not_strike_char_style(self):
        self.style_test_harness('not_strike_char_style')

    ##########################################################################
    #                                                                        #
    # Font Size                                                              #
    #                                                                        #
    ##########################################################################
    def test_font_14pt_char_style(self):
        self.style_test_harness('font_14pt_char_style')

    def test_font_17hlfpt_char_style(self):
        self.style_test_harness('font_17hlfpt_char_style')

    ##########################################################################
    #                                                                        #
    # Highlight                                                              #
    #                                                                        #
    ##########################################################################
    def test_highlight_darkMagenta_char_style(self):
        self.style_test_harness('highlight_darkMagenta_char_style')

    def test_no_highlight_char_style(self):
        self.style_test_harness('no_highlight_char_style')

    ##########################################################################
    #                                                                        #
    # Kerning                                                                #
    #                                                                        #
    ##########################################################################
    def test_kerning_12pt_char_style(self):
        self.style_test_harness('kerning_12pt_char_style')

    def test_no_kerning_char_style(self):
        self.style_test_harness('no_kerning_char_style')

    ##########################################################################
    #                                                                        #
    # Outline                                                                #
    #                                                                        #
    ##########################################################################
    def test_outline_char_style(self):
        self.style_test_harness('outline_char_style')

    def test_not_outline_char_style(self):
        self.style_test_harness('not_outline_char_style')

    ##########################################################################
    #                                                                        #
    # Position                                                               #
    #                                                                        #
    ##########################################################################
    def test_position_lowered_3pt_char_style(self):
        self.style_test_harness('position_lowered_3pt_char_style')

    def test_position_raised_3pt_char_style(self):
        self.style_test_harness('position_raised_3pt_char_style')

    def test_position_normal_char_style(self):
        self.style_test_harness('position_normal_char_style')

    def test_position_lowered_3pt_superscript_char_style(self):
        self.style_test_harness('position_lowered_3pt_superscript_char_style')

    ##########################################################################
    #                                                                        #
    # Shadow                                                                 #
    #                                                                        #
    ##########################################################################
    def test_shadow_char_style(self):
        self.style_test_harness('shadow_char_style')

    def test_not_shadow_char_style(self):
        self.style_test_harness('not_shadow_char_style')

    ##########################################################################
    #                                                                        #
    # Shading                                                                #
    #                                                                        #
    ##########################################################################
    def test_shading_blue_char_style(self):
        self.style_test_harness('shading_blue_char_style')

    def test_no_shading_char_style(self):
        self.style_test_harness('no_shading_char_style')

    def test_highlight_darkMagenta_shading_yellow_char_style(self):
        self.style_test_harness('highlight_darkMagenta_shading_yellow_char_style')

    def test_shading_theme_color_char_style(self):
        self.style_color_test_harness('shading_theme_color_char_style',
                                      'background-color')

    def test_shading_theme_shade_char_style(self):
        self.style_color_test_harness('shading_theme_shade_char_style',
                                      'background-color')

    def test_shading_theme_tint_char_style(self):
        self.style_color_test_harness('shading_theme_tint_char_style',
                                      'background-color')

    ##########################################################################
    #                                                                        #
    # Small Caps                                                             #
    #                                                                        #
    ##########################################################################
    def test_small_caps_char_style(self):
        self.style_test_harness('small_caps_char_style')

    ##########################################################################
    #                                                                        #
    # Character Spacing                                                      #
    #                                                                        #
    ##########################################################################
    def test_spacing_expanded_2pt_char_style(self):
        self.style_test_harness('spacing_expanded_2pt_char_style')

    def test_spacing_condensed_3pt_char_style(self):
        self.style_test_harness('spacing_condensed_3pt_char_style')

    def test_spacing_normal_char_style(self):
        self.style_test_harness('spacing_normal_char_style')

    ##########################################################################
    #                                                                        #
    # Underline                                                              #
    #                                                                        #
    ##########################################################################
    def test_underline_char_style(self):
        self.style_test_harness('underline_char_style')

    def test_not_underline_char_style(self):
        self.style_test_harness('not_underline_char_style')

    def test_underline_dashdotdotheavy_char_style(self):
        self.style_test_harness('underline_dashdotdotheavy_char_style')

    def test_underline_dashdotheavy_char_style(self):
        self.style_test_harness('underline_dashdotheavy_char_style')

    def test_underline_dashedheavy_char_style(self):
        self.style_test_harness('underline_dashedheavy_char_style')

    def test_underline_dashlong_char_style(self):
        self.style_test_harness('underline_dashlong_char_style')

    def test_underline_dashlongheavy_char_style(self):
        self.style_test_harness('underline_dashlongheavy_char_style')

    def test_underline_dotdash_char_style(self):
        self.style_test_harness('underline_dotdash_char_style')

    def test_underline_dotdotdash_char_style(self):
        self.style_test_harness('underline_dotdotdash_char_style')

    def test_underline_dottedheavy_char_style(self):
        self.style_test_harness('underline_dottedheavy_char_style')

    def test_underline_thick_char_style(self):
        self.style_test_harness('underline_thick_char_style')

    def test_underline_wavydouble_char_style(self):
        self.style_test_harness('underline_wavydouble_char_style')

    def test_underline_wavyheavy_char_style(self):
        self.style_test_harness('underline_wavyheavy_char_style')

    def test_underline_words_char_style(self):
        self.style_test_harness('underline_words_char_style')

    def test_underline_dash_char_style(self):
        self.style_test_harness('underline_dash_char_style')

    def test_underline_dotted_char_style(self):
        self.style_test_harness('underline_dotted_char_style')

    def test_underline_double_char_style(self):
        self.style_test_harness('underline_double_char_style')

    def test_underline_wave_char_style(self):
        self.style_test_harness('underline_wave_char_style')

    def test_strike_not_dstrike_underline_char_style(self):
        self.style_test_harness('strike_not_dstrike_underline_char_style')

    def test_strike_not_dstrike_not_underline_char_style(self):
        self.style_test_harness('strike_not_dstrike_not_underline_char_style')

    def test_not_strike_dstrike_underline_char_style(self):
        self.style_test_harness('not_strike_dstrike_underline_char_style')

    def test_not_strike_dstrike_not_underline_char_style(self):
        self.style_test_harness('not_strike_dstrike_not_underline_char_style')

    def test_not_strike_not_dstrike_underline_char_style(self):
        self.style_test_harness('not_strike_not_dstrike_underline_char_style')

    def test_not_strike_not_dstrike_not_underline_char_style(self):
        self.style_test_harness('not_strike_not_dstrike_not_underline_char_style')

    def test_strike_not_dstrike_underline_wave_char_style(self):
        self.style_test_harness('strike_not_dstrike_underline_wave_char_style')

    ##########################################################################
    #                                                                        #
    # Underline Colors                                                       #
    #                                                                        #
    ##########################################################################
    def test_underline_rgb_char_style(self):
        self.style_test_harness('underline_rgb_char_style')

    def test_underline_theme_char_style(self):
        self.style_test_harness('underline_theme_char_style')

    def test_underline_themetint_char_style(self):
        self.style_test_harness('underline_themetint_char_style')

    def test_underline_hsl_char_style(self):
        self.style_test_harness('underline_hsl_char_style')

    ##########################################################################
    #                                                                        #
    # Vanish                                                                 #
    #                                                                        #
    ##########################################################################
    def test_vanish_char_style(self):
        self.style_test_harness('vanish_char_style')

    def test_not_vanish_char_style(self):
        self.style_test_harness('not_vanish_char_style')

    ##########################################################################
    #                                                                        #
    # vertAlign                                                              #
    #                                                                        #
    ##########################################################################
    def test_superscript_char_style(self):
        self.style_test_harness('superscript_char_style')

    def test_not_superscript_char_style(self):
        self.style_test_harness('not_superscript_char_style')

    def test_subscript_char_style(self):
        self.style_test_harness('subscript_char_style')

    def test_not_subscript_char_style(self):
        self.style_test_harness('not_subscript_char_style')


class ParagraphStylesTestCase(TestHarness):
    files = (
        'paragraph_styles.docx',
        'border_paragraph_styles.docx',
    )
    css_files_location = 'test_files/paragraph/css/'
    docx_files_location = 'test_files/paragraph/docx/'

    ##########################################################################
    #                                                                        #
    # Bold                                                                   #
    #                                                                        #
    ##########################################################################
    def test_bold_paragraph(self):
        self.style_test_harness('bold_paragraph')

    ##########################################################################
    #                                                                        #
    # Borders                                                                #
    #                                                                        #
    ##########################################################################
    def test_border_all_solid_auto_05pt_paragraph(self):
        self.style_test_harness('border_all_solid_auto_05pt_paragraph')

    ##########################################################################
    #                                                                        #
    # Indent                                                                 #
    #                                                                        #
    ##########################################################################
    def test_indent_firstline_05in_paragraph(self):
        self.style_test_harness('indent_firstline_05in_paragraph')

    def test_indent_hanging_044in_paragraph(self):
        self.style_test_harness('indent_hanging_044in_paragraph')

    def test_indent_left_02in_negative_paragraph(self):
        self.style_test_harness('indent_left_02in_negative_paragraph')

    def test_indent_left_05in_paragraph(self):
        self.style_test_harness('indent_left_05in_paragraph')

    def test_indent_right_03in_negative_paragraph(self):
        self.style_test_harness('indent_right_03in_negative_paragraph')

    def test_indent_right_04in_paragraph(self):
        self.style_test_harness('indent_right_04in_paragraph')

    def test_indent_left_05in_right_02in_mirror_paragraph(self):
        self.style_test_harness('indent_left_05in_right_02in_mirror_paragraph')

    ##########################################################################
    #                                                                        #
    # Justify                                                                #
    #                                                                        #
    ##########################################################################
    def test_justify_left_paragraph(self):
        self.style_test_harness('justify_left_paragraph')

    def test_justify_center_paragraph(self):
        self.style_test_harness('justify_center_paragraph')

    def test_justify_right_paragraph(self):
        self.style_test_harness('justify_right_paragraph')

    def test_justify_justify_paragraph(self):
        self.style_test_harness('justify_justify_paragraph')

    ##########################################################################
    #                                                                        #
    # Line Spacing                                                           #
    #                                                                        #
    ##########################################################################
    def test_line_spacing_15lines_paragraph(self):
        self.style_test_harness('line_spacing_15lines_paragraph')

    def test_line_spacing_atleast_14pt_paragraph(self):
        self.style_test_harness('line_spacing_atleast_14pt_paragraph')

    def test_line_spacing_double_paragraph(self):
        self.style_test_harness('line_spacing_double_paragraph')

    def test_line_spacing_exactly_16pt_paragraph(self):
        self.style_test_harness('line_spacing_exactly_16pt_paragraph')

    def test_line_spacing_multiple_3_paragraph(self):
        self.style_test_harness('line_spacing_multiple_3_paragraph')

    def test_line_spacing_multiple_109_paragraph(self):
        self.style_test_harness('line_spacing_multiple_109_paragraph')

    def test_line_spacing_single_paragraph(self):
        self.style_test_harness('line_spacing_single_paragraph')

    ##########################################################################
    #                                                                        #
    # Paragraph Shading                                                      #
    #                                                                        #
    ##########################################################################
    def test_shading_rgb_paragraph(self):
        self.style_test_harness('shading_rgb_paragraph')

    ##########################################################################
    #                                                                        #
    # Paragraph Spacing                                                      #
    #                                                                        #
    ##########################################################################
    def test_spacing_after_12pt_paragraph(self):
        self.style_test_harness('spacing_after_12pt_paragraph')

    def test_spacing_after_155pt_paragraph(self):
        self.style_test_harness('spacing_after_155pt_paragraph')

    def test_spacing_after_auto_before_18pt_paragraph(self):
        self.style_test_harness('spacing_after_auto_before_18pt_paragraph')

    def test_spacing_before_18pt_paragraph(self):
        self.style_test_harness('spacing_before_18pt_paragraph')

    def test_spacing_before_auto_after_6pt_paragraph(self):
        self.style_test_harness('spacing_before_auto_after_6pt_paragraph')

    def test_spacing_auto_fontsize_72pt_paragraph(self):
        self.style_test_harness('spacing_auto_fontsize_72pt_paragraph')

    ##########################################################################
    #                                                                        #
    # Pagination Control                                                     #
    #                                                                        #
    ##########################################################################
    def test_pagination_keep_lines_together_on_paragraph(self):
        self.style_test_harness('pagination_keep_lines_together_on_paragraph')

    def test_pagination_keep_lines_together_off_paragraph(self):
        self.style_test_harness('pagination_keep_lines_together_off_paragraph')

    def test_pagination_keep_with_next_on_paragraph(self):
        self.style_test_harness('pagination_keep_with_next_on_paragraph')

    def test_pagination_keep_with_next_off_paragraph(self):
        self.style_test_harness('pagination_keep_with_next_off_paragraph')

    def test_pagination_page_break_before_on_paragraph(self):
        self.style_test_harness('pagination_page_break_before_on_paragraph')

    def test_pagination_page_break_before_off_paragraph(self):
        self.style_test_harness('pagination_page_break_before_off_paragraph')

    def test_pagination_widow_control_off_paragraph(self):
        self.style_test_harness('pagination_widow_control_off_paragraph')

    def test_pagination_widow_control_on_paragraph(self):
        self.style_test_harness('pagination_widow_control_on_paragraph')


class DocDefaultsTestCase(TestCase):
    css_files_location = 'test_files/paragraph/css/'

    def setUp(self):
        filename = 'test_files/paragraph/docx/normal_paragraph_styles.docx'
        parser = DocxParser(filename)
        self.body_style = parser.parse().body_style

    def test_parse_doc_defaults(self):
        self.assertEqual(CssUnit(11, 'pt'), self.body_style.font_size)
        self.assertEqual('Calibri, sans-serif', self.body_style.font_family)
        self.assertEqual(1.08, round(self.body_style.line_height, 2))
        self.assertEqual(CssUnit(8, 'pt'), self.body_style.margin_bottom)

    def test_normal_paragraph(self):
        css_file = f'{self.css_files_location}body_defaults.css'

        with open(css_file, 'r') as css_file:
            expected = css_file.read()

        serializer = FACTORY.get_block_serializer(self.body_style)
        css_stylesheet = cssutils.css.CSSStyleSheet()
        for rule in serializer.css_style_rules():
            css_stylesheet.add(rule)
        result = css_stylesheet.cssText.decode('utf-8')

        self.assertEqual(expected, result)


class RequeteTestCase(TestCase):
    maxDiff = None

    def compare_documents(self, docx_source, expected_css):
        parser = DocxParser(docx_source)
        stylesheet = parser.parse()
        serializer = CssStylesheetSerializer(stylesheet)
        css = serializer.serialize()
        # stylesheet.preferences['simulate_printed_page'] = True
        # css = stylesheet.cssText
        with open(expected_css, mode='r', encoding='utf-8') as expected:
            self.assertEqual(expected.read(), css)

    def test_endos(self):
        self.compare_documents('test_files/endos.docx',
                               'test_files/endos.css')

    def test_contract(self):
        self.compare_documents('test_files/contrat.docx',
                               'test_files/contrat.css')

    def test_enveloppe(self):
        self.compare_documents('test_files/enveloppe.docx',
                               'test_files/enveloppe.css')

    def test_fax(self):
        print_styles_tree('test_files/fax.docx')
        self.compare_documents('test_files/fax.docx',
                               'test_files/fax.css')

    def test_labels(self):
        self.compare_documents('test_files/labels.docx',
                               'test_files/labels.css')

    def test_memo(self):
        self.compare_documents('test_files/memo.docx',
                               'test_files/memo.css')

    def test_requete(self):
        self.compare_documents('test_files/numbering/docx/requete.docx',
                               'test_files/numbering/css/requete.css')

    def test_resolution(self):
        self.compare_documents('test_files/resolution.docx',
                               'test_files/resolution.css')

    def test_normal_paragraph_selector(self):
        parser = DocxParser('test_files/numbering/docx/requete.docx')
        stylesheet = parser.parse()
        style = stylesheet.paragraph_styles['']
        block_serializer = FACTORY.get_block_serializer(style)
        expected = 'p, h1, h2, h3'
        self.assertEqual(expected, block_serializer.css_selector())

    def test_no_space_hierarchy(self):
        parser = DocxParser('test_files/numbering/docx/requete.docx')
        stylesheet = parser.parse()
        style = stylesheet.paragraph_styles['no-space']
        block_serializer = FACTORY.get_block_serializer(style)
        expected = 'p.no-space, p.no-space-center-bold'
        self.assertEqual(expected, block_serializer.css_selector())


class CharacterStylesParserTestCase(TestCase):
    files = (
        'character_styles.docx',
        'border_character_styles.docx',
        'font_character_styles.docx',
        'highlight_char_styles.docx',
        'legacy_char_styles.docx',
        'color_char_styles.docx',
    )
    css_files_location = 'test_files/character/css/'
    docx_files_location = 'test_files/character/docx/'

    @classmethod
    def setUpClass(cls):
        cls.docx_styles = {}
        for file in cls.files:
            cls.parser = DocxParser(f'{cls.docx_files_location}{file}')
            for docx_style in cls.parser.opc_package.styles.values():
                cls.docx_styles[docx_style.name] = docx_style

    def get_parsed_style_by_name(self, style_name):
        docx_style = self.docx_styles[style_name]
        return self.parser.parse_docx_style(docx_style)

    def test_doc_defaults(self):
        parser = DocxParser('test_files/paragraph/docx/normal_paragraph_styles.docx')
        body_style = parser.parse().body_style
        self.assertEqual(CssUnit(11, 'pt'), body_style.font_size)
        self.assertEqual('Calibri, sans-serif', body_style.font_family)
        self.assertEqual(1.08, round(body_style.line_height, 2))
        self.assertEqual(CssUnit(8, 'pt'), body_style.margin_bottom)

    ##########################################################################
    #                                                                        #
    # Bold                                                                   #
    #                                                                        #
    ##########################################################################
    def test_bold_char_style(self):
        """Bold is active"""
        style = self.get_parsed_style_by_name('bold_char_style')
        self.assertTrue(style.bold)

    def test_not_bold_char_style(self):
        """Child of bold_char_style. Bold is toggled to off"""
        style = self.get_parsed_style_by_name('not_bold_char_style')
        self.assertFalse(style.bold)

    ##########################################################################
    #                                                                        #
    # Border Character Styles                                                #
    #                                                                        #
    ##########################################################################
    def test_border_char_style(self):
        style = self.get_parsed_style_by_name('border_char_style')
        self.assertEqual(CssUnit(0.5, 'pt'), style.border.width)
        self.assertEqual('solid', style.border.style)

    def test_border_rgb_char_style(self):
        style = self.get_parsed_style_by_name('border_rgb_char_style')
        # border-style: solid;
        # border-width: 0.5pt;
        # border-color: #F00;
        self.assertEqual('solid', style.border.style)
        self.assertEqual(CssUnit(0.5, 'pt'), style.border.width)
        self.assertEqual('#FF0000', style.border.color)

    def test_border_shadow_char_style(self):
        style = self.get_parsed_style_by_name('border_shadow_char_style')
        # border-style: solid;
        # border-width: 0.5pt;
        # box-shadow: 0.5pt 0.5pt;
        self.assertEqual('solid', style.border.style)
        self.assertEqual(CssUnit(0.5, 'pt'), style.border.width)
        self.assertTrue(style.border.shadow)

    def test_border_thinThickSmallGap_3d_char_style(self):
        style = self.get_parsed_style_by_name('border_thinThickSmallGap_3d_char_style')
        #     border-style: double;
        #     border-width: 3pt;
        self.assertEqual('double', style.border.style)
        self.assertEqual(CssUnit(3, 'pt'), style.border.width)

    def test_border_1pt_char_style(self):
        style = self.get_parsed_style_by_name('border_1pt_char_style')
        #     border-style: solid;
        #     border-width: 1pt;
        self.assertEqual('solid', style.border.style)
        self.assertEqual(CssUnit(1, 'pt'), style.border.width)

    def test_border_theme_char_style(self):
        style = self.get_parsed_style_by_name('border_theme_char_style')
        #     border-style: solid;
        #     border-width: 0.25pt;
        #     border-color: #4472c4;
        self.assertEqual('solid', style.border.style)
        self.assertEqual(CssUnit(0.25, 'pt'), style.border.width)
        self.assertEqual('#4472c4', style.border.color)

    def test_no_border_char_style(self):
        style = self.get_parsed_style_by_name('no_border_char_style')
        #     border-style: none;
        self.assertEqual('none', style.border.style)
        self.assertIsNone(style.border.color)
        self.assertEqual(0, style.border.padding)
        self.assertFalse(style.border.shadow)
        self.assertEqual(0, style.border.width)

    def test_border_dashDotStroked_char_style(self):
        style = self.get_parsed_style_by_name('border_dashDotStroked_char_style')
        #     border-style: dashed;
        #     border-width: 3pt;
        self.assertEqual('dashed', style.border.style)
        self.assertEqual(CssUnit(3, 'pt'), style.border.width)

    ##########################################################################
    #                                                                        #
    # Emboss                                                                 #
    #                                                                        #
    ##########################################################################
    def test_emboss_char_style(self):
        style = self.get_parsed_style_by_name('emboss_char_style')
        self.assertTrue(style.emboss)

    def test_not_emboss_char_style(self):
        style = self.get_parsed_style_by_name('not_emboss_char_style')
        self.assertFalse(style.emboss)

    ##########################################################################
    #                                                                        #
    # Fonts                                                                  #
    #                                                                        #
    ##########################################################################
    def test_font_arial_char_style(self):
        style = self.get_parsed_style_by_name('font_arial_char_style')
        #     font-family: Arial, sans-serif;
        self.assertEqual('Arial, sans-serif', style.font_family)

    def test_font_timesNewRoman_char_style(self):
        style = self.get_parsed_style_by_name('font_timesNewRoman_char_style')
        #     font-family: "Times New Roman", serif;
        self.assertEqual('"Times New Roman", serif', style.font_family)

    def test_font_body_char_style(self):
        style = self.get_parsed_style_by_name('font_body_char_style')
        #     font-family: Calibri, sans-serif;
        self.assertEqual('Calibri, sans-serif', style.font_family)

    def test_font_hansi_char_style(self):
        style = self.get_parsed_style_by_name('font_hansi_char_style')
        #     font-family: "Cooper Black", "Times New Roman", serif;
        self.assertEqual('"Cooper Black", "Times New Roman", serif',
                         style.font_family)

    ##########################################################################
    #                                                                        #
    # Imprint                                                                #
    #                                                                        #
    ##########################################################################
    def test_imprint_char_style(self):
        style = self.get_parsed_style_by_name('imprint_char_style')
        self.assertTrue(style.imprint)

    def test_not_imprint_char_style(self):
        style = self.get_parsed_style_by_name('not_imprint_char_style')
        self.assertFalse(style.imprint)

    ##########################################################################
    #                                                                        #
    # Italic                                                                 #
    #                                                                        #
    ##########################################################################
    def test_italic_char_style(self):
        """Italic is active"""
        style = self.get_parsed_style_by_name('italic_char_style')
        self.assertTrue(style.italics)

    def test_not_italic_char_style(self):
        """Child of italic_char_style. Italic is toggled to off"""
        style = self.get_parsed_style_by_name('not_italic_char_style')
        self.assertFalse(style.italics)

    def test_bold_italic_char_style(self):
        """Child of bold_char_style with italic added"""
        style = self.get_parsed_style_by_name('bold_italic_char_style')
        self.assertTrue(style.italics)
        self.assertTrue(style.bold)

    ##########################################################################
    #                                                                        #
    # Caps                                                                   #
    #                                                                        #
    ##########################################################################
    def test_caps_char_style(self):
        """Text in uppercase"""
        style = self.get_parsed_style_by_name('caps_char_style')
        self.assertTrue(style.all_caps)

    def test_not_caps_char_style(self):
        """Child of caps_char_style. Uppercase it toggled off"""
        style = self.get_parsed_style_by_name('not_caps_char_style')
        self.assertFalse(style.all_caps)

    ##########################################################################
    #                                                                        #
    # Font Color                                                             #
    #                                                                        #
    ##########################################################################
    def test_color_rgb_char_style(self):
        style = self.get_parsed_style_by_name('color_rgb_char_style')
        self.assertEqual('#FF0000', style.font_color)

    def test_color_theme_accent2_char_style(self):
        style = self.get_parsed_style_by_name('color_theme_accent2_char_style')
        self.assertEqual('#ed7d31', style.font_color)

    ##########################################################################
    #                                                                        #
    # dStrike                                                                #
    #                                                                        #
    ##########################################################################
    def test_dstrike_char_style(self):
        style = self.get_parsed_style_by_name('dstrike_char_style')
        self.assertTrue(style.double_strike)

    def test_not_dstrike_char_style(self):
        style = self.get_parsed_style_by_name('not_dstrike_char_style')
        self.assertFalse(style.double_strike)

    ##########################################################################
    #                                                                        #
    # Strike                                                                 #
    #                                                                        #
    ##########################################################################
    def test_strike_char_style(self):
        style = self.get_parsed_style_by_name('strike_char_style')
        self.assertTrue(style.strike)

    def test_not_strike_char_style(self):
        style = self.get_parsed_style_by_name('not_strike_char_style')
        self.assertFalse(style.strike)

    ##########################################################################
    #                                                                        #
    # Font Size                                                              #
    #                                                                        #
    ##########################################################################
    def test_font_14pt_char_style(self):
        style = self.get_parsed_style_by_name('font_14pt_char_style')
        self.assertEqual(CssUnit(14, 'pt'), style.font_size)

    def test_font_17hlfpt_char_style(self):
        style = self.get_parsed_style_by_name('font_17hlfpt_char_style')
        self.assertEqual(CssUnit(8.5, 'pt'), style.font_size)

    ##########################################################################
    #                                                                        #
    # Highlight                                                              #
    #                                                                        #
    ##########################################################################
    def test_highlight_darkMagenta_char_style(self):
        style = self.get_parsed_style_by_name('highlight_darkMagenta_char_style')
        self.assertEqual('darkmagenta', style.highlight)

    def test_no_highlight_char_style(self):
        style = self.get_parsed_style_by_name('no_highlight_char_style')
        self.assertEqual('none', style.highlight)

    ##########################################################################
    #                                                                        #
    # Kerning                                                                #
    #                                                                        #
    ##########################################################################
    def test_kerning_12pt_char_style(self):
        style = self.get_parsed_style_by_name('kerning_12pt_char_style')
        self.assertTrue(style.font_kerning)

    def test_no_kerning_char_style(self):
        style = self.get_parsed_style_by_name('no_kerning_char_style')
        self.assertFalse(style.font_kerning)

    ##########################################################################
    #                                                                        #
    # Outline                                                                #
    #                                                                        #
    ##########################################################################
    def test_outline_char_style(self):
        style = self.get_parsed_style_by_name('outline_char_style')
        self.assertTrue(style.outline)

    def test_not_outline_char_style(self):
        style = self.get_parsed_style_by_name('not_outline_char_style')
        self.assertFalse(style.outline)

    ##########################################################################
    #                                                                        #
    # Position                                                               #
    #                                                                        #
    ##########################################################################
    def test_position_lowered_3pt_char_style(self):
        style = self.get_parsed_style_by_name('position_lowered_3pt_char_style')
        self.assertEqual(CssUnit(-3, 'pt'), style.position)

    def test_position_raised_3pt_char_style(self):
        style = self.get_parsed_style_by_name('position_raised_3pt_char_style')
        self.assertEqual(CssUnit(3, 'pt'), style.position)

    def test_position_normal_char_style(self):
        style = self.get_parsed_style_by_name('position_normal_char_style')
        self.assertEqual(CssUnit(0, 'pt'), style.position)

    def test_position_lowered_3pt_superscript_char_style(self):
        style = self.get_parsed_style_by_name('position_lowered_3pt_superscript_char_style')
        self.assertEqual(CssUnit(-3, 'pt'), style.position)

    ##########################################################################
    #                                                                        #
    # Shadow                                                                 #
    #                                                                        #
    ##########################################################################
    def test_shadow_char_style(self):
        style = self.get_parsed_style_by_name('shadow_char_style')
        self.assertTrue(style.shadow)

    def test_not_shadow_char_style(self):
        style = self.get_parsed_style_by_name('not_shadow_char_style')
        self.assertFalse(style.shadow)

    ##########################################################################
    #                                                                        #
    # Shading                                                                #
    #                                                                        #
    ##########################################################################
    def test_shading_blue_char_style(self):
        style = self.get_parsed_style_by_name('shading_blue_char_style')
        self.assertEqual('#0070C0', style.background_color)

    def test_no_shading_char_style(self):
        style = self.get_parsed_style_by_name('no_shading_char_style')
        self.assertEqual('', style.background_color)

    def test_highlight_darkMagenta_shading_yellow_char_style(self):
        style = self.get_parsed_style_by_name('highlight_darkMagenta_shading_yellow_char_style')
        self.assertEqual('#FF0000', style.background_color)

    ##########################################################################
    #                                                                        #
    # Small Caps                                                             #
    #                                                                        #
    ##########################################################################
    def test_small_caps_char_style(self):
        style = self.get_parsed_style_by_name('small_caps_char_style')
        self.assertTrue(style.small_caps)

    ##########################################################################
    #                                                                        #
    # Character Spacing                                                      #
    #                                                                        #
    ##########################################################################
    def test_spacing_expanded_2pt_char_style(self):
        style = self.get_parsed_style_by_name('spacing_expanded_2pt_char_style')
        self.assertEqual(CssUnit(2, 'pt'), style.letter_spacing)

    def test_spacing_condensed_3pt_char_style(self):
        style = self.get_parsed_style_by_name('spacing_condensed_3pt_char_style')
        self.assertEqual(CssUnit(-3, 'pt'), style.letter_spacing)

    def test_spacing_normal_char_style(self):
        style = self.get_parsed_style_by_name('spacing_normal_char_style')
        self.assertEqual(CssUnit(0, 'pt'), style.letter_spacing)

    ##########################################################################
    #                                                                        #
    # Underline                                                              #
    #                                                                        #
    ##########################################################################
    def test_underline_char_style(self):
        style = self.get_parsed_style_by_name('underline_char_style')
        #     text-decoration-line: underline;
        #     text-decoration-style: solid;
        self.assertEqual('solid', style.underline.style)
        self.assertEqual(style.underline.UNDERLINE, style.underline.line)

    def test_not_underline_char_style(self):
        style = self.get_parsed_style_by_name('not_underline_char_style')
        #     text-decoration-line: none;
        self.assertEqual(0, style.underline.line)

    ##########################################################################
    #                                                                        #
    # Underline Colors                                                       #
    #                                                                        #
    ##########################################################################
    def test_underline_rgb_char_style(self):
        style = self.get_parsed_style_by_name('underline_rgb_char_style')
        self.assertEqual('#FF0000', style.underline.color)

    def test_underline_theme_char_style(self):
        style = self.get_parsed_style_by_name('underline_theme_char_style')
        self.assertEqual('#4472c4', style.underline.color)

    def test_underline_themetint_char_style(self):
        style = self.get_parsed_style_by_name('underline_themetint_char_style')
        self.assertEqual('#8eaadb', style.underline.color)

    def test_underline_hsl_char_style(self):
        style = self.get_parsed_style_by_name('underline_hsl_char_style')
        self.assertEqual('#178D79', style.underline.color)

    ##########################################################################
    #                                                                        #
    # Vanish                                                                 #
    #                                                                        #
    ##########################################################################
    def test_vanish_char_style(self):
        style = self.get_parsed_style_by_name('vanish_char_style')
        self.assertTrue(style.visible)

    def test_not_vanish_char_style(self):
        style = self.get_parsed_style_by_name('not_vanish_char_style')
        self.assertFalse(style.visible)

    ##########################################################################
    #                                                                        #
    # vertAlign                                                              #
    #                                                                        #
    ##########################################################################
    def test_superscript_char_style(self):
        style = self.get_parsed_style_by_name('superscript_char_style')
        self.assertEqual('superscript', style.vertical_align)

    def test_not_superscript_char_style(self):
        style = self.get_parsed_style_by_name('not_superscript_char_style')
        self.assertEqual('baseline', style.vertical_align)

    def test_subscript_char_style(self):
        style = self.get_parsed_style_by_name('subscript_char_style')
        self.assertEqual('subscript', style.vertical_align)

    def test_not_subscript_char_style(self):
        style = self.get_parsed_style_by_name('not_subscript_char_style')
        self.assertEqual('baseline', style.vertical_align)


class ParagraphStylesParserTestCase(TestCase):
    files = (
        'paragraph_styles.docx',
        'border_paragraph_styles.docx',
    )
    css_files_location = 'test_files/paragraph/css/'
    docx_files_location = 'test_files/paragraph/docx/'

    @classmethod
    def setUpClass(cls):
        cls.docx_styles = {}
        for file in cls.files:
            cls.parser = DocxParser(f'{cls.docx_files_location}{file}')
            for docx_style in cls.parser.opc_package.styles.values():
                cls.docx_styles[docx_style.name] = docx_style

    def get_parsed_style_by_name(self, style_name):
        docx_style = self.docx_styles[style_name]
        return self.parser.parse_docx_style(docx_style)

    def test_normal_paragraph(self):
        style = self.get_parsed_style_by_name('Normal')
        self.assertEqual('', style.name)
        self.assertEqual('', style.id)
        style = self.get_parsed_style_by_name('bold_paragraph')
        self.assertIsNotNone(style.parent)

    ##########################################################################
    #                                                                        #
    # Bold                                                                   #
    #                                                                        #
    ##########################################################################
    def test_bold_paragraph(self):
        style = self.get_parsed_style_by_name('bold_paragraph')
        self.assertTrue(style.bold)

    ##########################################################################
    #                                                                        #
    # Borders                                                                #
    #                                                                        #
    ##########################################################################
    def test_border_all_solid_auto_05pt_paragraph(self):
        style = self.get_parsed_style_by_name('border_all_solid_auto_05pt_paragraph')

        def test_border(api_style, direction, values):
            border = getattr(api_style, f'border_{direction}')
            self.assertEqual(values[0], border.style)
            self.assertEqual(values[1], border.width)
            self.assertEqual(values[2], border.padding)

        #     border-bottom-style: solid;
        #     border-bottom-width: 0.5pt;
        #     padding-bottom: 1pt;
        test_border(style, 'bottom',
                    ('solid', CssUnit(0.5, 'pt'), CssUnit(1, 'pt')))
        #     border-left-style: solid;
        #     border-left-width: 0.5pt;
        #     padding-left: 4pt;
        test_border(style, 'left',
                    ('solid', CssUnit(0.5, 'pt'), CssUnit(4, 'pt')))
        #     border-top-style: solid;
        #     border-top-width: 0.5pt;
        #     padding-top: 1pt;
        test_border(style, 'top',
                    ('solid', CssUnit(0.5, 'pt'), CssUnit(1, 'pt')))
        #     border-right-style: solid;
        #     border-right-width: 0.5pt;
        #     padding-right: 4pt;
        test_border(style, 'right',
                    ('solid', CssUnit(0.5, 'pt'), CssUnit(4, 'pt')))

    ##########################################################################
    #                                                                        #
    # Indent                                                                 #
    #                                                                        #
    ##########################################################################
    def test_indent_firstline_05in_paragraph(self):
        style = self.get_parsed_style_by_name('indent_firstline_05in_paragraph')
        self.assertEqual(CssUnit(0.5, 'in'), style.text_indent)

    def test_indent_hanging_044in_paragraph(self):
        style = self.get_parsed_style_by_name('indent_hanging_044in_paragraph')
        #     margin-left: 0.44in;
        #     text-indent: -0.44in;
        self.assertEqual(-0.44, round(style.text_indent.inches, 2))
        self.assertEqual(0.44, round(style.margin_left.inches, 2))

    def test_indent_left_02in_negative_paragraph(self):
        style = self.get_parsed_style_by_name('indent_left_02in_negative_paragraph')
        #     margin-left: -0.2in;
        self.assertEqual(CssUnit(-0.2, 'in'), style.margin_left)

    def test_indent_left_05in_paragraph(self):
        style = self.get_parsed_style_by_name('indent_left_05in_paragraph')
        self.assertEqual(CssUnit(0.5, 'in'), style.margin_left)

    def test_indent_right_03in_negative_paragraph(self):
        style = self.get_parsed_style_by_name('indent_right_03in_negative_paragraph')
        self.assertEqual(CssUnit(-0.3, 'in'), style.margin_right)

    def test_indent_right_04in_paragraph(self):
        style = self.get_parsed_style_by_name('indent_right_04in_paragraph')
        self.assertEqual(CssUnit(0.4, 'in'), style.margin_right)

    def test_indent_left_05in_right_02in_mirror_paragraph(self):
        style = self.get_parsed_style_by_name('indent_left_05in_right_02in_mirror_paragraph')
        self.assertEqual(CssUnit(0.5, 'in'), style.margin_left)
        self.assertEqual(CssUnit(0.2, 'in'), style.margin_right)

    ##########################################################################
    #                                                                        #
    # Justify                                                                #
    #                                                                        #
    ##########################################################################
    def test_justify_left_paragraph(self):
        style = self.get_parsed_style_by_name('justify_left_paragraph')
        self.assertEqual('start', style.text_align)

    def test_justify_center_paragraph(self):
        style = self.get_parsed_style_by_name('justify_center_paragraph')
        self.assertEqual('center', style.text_align)

    def test_justify_right_paragraph(self):
        style = self.get_parsed_style_by_name('justify_right_paragraph')
        self.assertEqual('end', style.text_align)

    def test_justify_justify_paragraph(self):
        style = self.get_parsed_style_by_name('justify_justify_paragraph')
        self.assertEqual('justify', style.text_align)

    ##########################################################################
    #                                                                        #
    # Line Spacing                                                           #
    #                                                                        #
    ##########################################################################
    def test_line_spacing_15lines_paragraph(self):
        style = self.get_parsed_style_by_name('line_spacing_15lines_paragraph')
        self.assertEqual(1.5, style.line_height)

    def test_line_spacing_atleast_14pt_paragraph(self):
        style = self.get_parsed_style_by_name('line_spacing_atleast_14pt_paragraph')
        self.assertEqual(CssUnit(14, 'pt'), style.line_height)

    def test_line_spacing_double_paragraph(self):
        style = self.get_parsed_style_by_name('line_spacing_double_paragraph')
        self.assertEqual(2, style.line_height)

    def test_line_spacing_exactly_16pt_paragraph(self):
        style = self.get_parsed_style_by_name('line_spacing_exactly_16pt_paragraph')
        self.assertEqual(CssUnit(16, 'pt'), style.line_height)

    def test_line_spacing_multiple_3_paragraph(self):
        style = self.get_parsed_style_by_name('line_spacing_multiple_3_paragraph')
        self.assertEqual(3, style.line_height)

    def test_line_spacing_multiple_109_paragraph(self):
        style = self.get_parsed_style_by_name('line_spacing_multiple_109_paragraph')
        self.assertEqual(1.09, round(style.line_height, 2))

    def test_line_spacing_single_paragraph(self):
        style = self.get_parsed_style_by_name('line_spacing_single_paragraph')
        self.assertEqual(1, style.line_height)

    ##########################################################################
    #                                                                        #
    # Paragraph Shading                                                      #
    #                                                                        #
    ##########################################################################
    def test_shading_rgb_paragraph(self):
        style = self.get_parsed_style_by_name('shading_rgb_paragraph')
        self.assertEqual('#FF0000', style.background_color)

    ##########################################################################
    #                                                                        #
    # Paragraph Spacing                                                      #
    #                                                                        #
    ##########################################################################
    def test_spacing_after_12pt_paragraph(self):
        style = self.get_parsed_style_by_name('spacing_after_12pt_paragraph')
        self.assertEqual(CssUnit(12, 'pt'), style.margin_bottom)

    def test_spacing_after_155pt_paragraph(self):
        style = self.get_parsed_style_by_name('spacing_after_155pt_paragraph')
        self.assertEqual(CssUnit(15.5, 'pt'), style.margin_bottom)

    def test_spacing_after_auto_before_18pt_paragraph(self):
        style = self.get_parsed_style_by_name('spacing_after_auto_before_18pt_paragraph')
        self.assertIsNone(style.margin_bottom)
        self.assertEqual(CssUnit(18, 'pt'), style.margin_top)

    def test_spacing_before_18pt_paragraph(self):
        style = self.get_parsed_style_by_name('spacing_before_18pt_paragraph')
        self.assertEqual(CssUnit(18, 'pt'), style.margin_top)

    def test_spacing_before_auto_after_6pt_paragraph(self):
        style = self.get_parsed_style_by_name('spacing_before_auto_after_6pt_paragraph')
        self.assertEqual(CssUnit(6, 'pt'), style.margin_bottom)
        self.assertIsNone(style.margin_top)

    def test_spacing_auto_fontsize_72pt_paragraph(self):
        style = self.get_parsed_style_by_name('spacing_auto_fontsize_72pt_paragraph')
        self.assertIsNone(style.margin_bottom)
        self.assertIsNone(style.margin_top)
        self.assertEqual(CssUnit(72, 'pt'), style.font_size)

    ##########################################################################
    #                                                                        #
    # Pagination Control                                                     #
    #                                                                        #
    ##########################################################################
    def test_pagination_keep_lines_together_on_paragraph(self):
        style = self.get_parsed_style_by_name('pagination_keep_lines_together_on_paragraph')
        self.assertTrue(style.keep_together)

    def test_pagination_keep_lines_together_off_paragraph(self):
        style = self.get_parsed_style_by_name('pagination_keep_lines_together_off_paragraph')
        self.assertFalse(style.keep_together)

    def test_pagination_keep_with_next_on_paragraph(self):
        style = self.get_parsed_style_by_name('pagination_keep_with_next_on_paragraph')
        self.assertTrue(style.keep_with_next)

    def test_pagination_keep_with_next_off_paragraph(self):
        style = self.get_parsed_style_by_name('pagination_keep_with_next_off_paragraph')
        self.assertFalse(style.keep_with_next)

    def test_pagination_page_break_before_on_paragraph(self):
        style = self.get_parsed_style_by_name('pagination_page_break_before_on_paragraph')
        self.assertTrue(style.page_break_before)

    def test_pagination_page_break_before_off_paragraph(self):
        style = self.get_parsed_style_by_name('pagination_page_break_before_off_paragraph')
        self.assertFalse(style.page_break_before)

    def test_pagination_widow_control_off_paragraph(self):
        style = self.get_parsed_style_by_name('pagination_widow_control_off_paragraph')
        self.assertFalse(style.widows_control)

    def test_pagination_widow_control_on_paragraph(self):
        style = self.get_parsed_style_by_name('pagination_widow_control_on_paragraph')
        self.assertTrue(style.widows_control)
