from unittest import TestCase
import warnings

import cssutils
from lxml import etree

from docx2css import open_docx
from docx2css.ooxml.package import OpcPackage
from docx2css.ooxml.styles import Styles
from docx2css.stylesheet import Stylesheet
from docx2css.utils import CSSColor

cssutils.ser.prefs.indentClosingBrace = False
cssutils.ser.prefs.omitLastSemicolon = False


class TestHarness(TestCase):

    files = ()
    css_files_location = None
    docx_files_location = None

    def setUp(self):
        self.styles = {}
        for file in self.files:
            parser = OpcPackage(f'{self.docx_files_location}{file}')
            stylesheet = parser.styles
            for style in stylesheet.values():
                # Sanity check
                if style.name in self.styles.keys():
                    msg = f'ERROR: Style "{style.name}" already exists!'
                    warnings.warn(msg)
                self.styles[style.name] = style

    def print_styles_tree(self, style_type):
        for file in self.files:
            parser = OpcPackage(f'{self.docx_files_location}{file}')
            stylesheet = parser.styles
            style_list = None
            if style_type == 'character':
                style_list = stylesheet.character_styles.values()
            elif style_type == 'paragraph':
                style_list = stylesheet.paragraph_styles.values()
            root_styles = (
                s for s in style_list if s.parent_id is None
            )
            print(f'Styles in {file}:')
            print(stylesheet.print_styles_tree(root_styles))

    def style_test_harness(self, style_name):
        style = self.styles[style_name]
        css_file = f'{self.css_files_location}{"".join(style_name.split())}.css'
        print(etree.tostring(style, pretty_print=True).decode('utf-8'))
        self.compare_style(style.css_style_rule(), css_file)

    def compare_style(self, css_style_rule, css_filename):
        with open(css_filename, 'r') as css_file:
            expected = css_file.read()
            css_text = css_style_rule.cssText
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
        print(style.css_style_rule().cssText)
        with open(f'{self.css_files_location}{"".join(style_name.split())}.css', 'r') as css_file:
            expected = cssutils.parseString(css_file.read()).cssRules[0]
            expected_color = expected.style[rule_name]
            result = style.css_style_rule()
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

    def test_print_tree(self):
        self.print_styles_tree('paragraph')

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


class DocDefaultsTestCase(TestHarness):
    files = (
        'normal_paragraph_styles.docx',
    )
    css_files_location = 'test_files/paragraph/css/'
    docx_files_location = 'test_files/paragraph/docx/'

    def test_normal_paragraph(self):
        css_file = f'{self.css_files_location}body_defaults.css'
        package = OpcPackage('test_files/paragraph/docx/normal_paragraph_styles.docx')
        stylesheet = Stylesheet(package)
        body = stylesheet.css_body_style()
        self.compare_style(body, css_file)


class RequeteTestCase(TestCase):

    def test_requete(self):
        css = open_docx('test_files/numbering/docx/requete.docx').cssText
        with open('test_files/numbering/css/requete.css', 'r') as expected:
            self.assertEqual(expected.read(), css)

    def test_normal_paragraph_selector(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        p = package.styles['Normal']
        expected = 'p, h1, h2, h3'
        self.assertEqual(expected, p.css_selector())
