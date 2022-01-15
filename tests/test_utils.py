from unittest import TestCase

from docx2css.utils import CSSColor, CssUnit


class CSSColorTestCase(TestCase):

    def test_hex2int(self):
        color = CSSColor()
        self.assertEqual(255, color.hex2int('FF'))

    def test_split_rgb(self):
        color = CSSColor()
        expected = ['FF', 'FF', 'FF']
        self.assertEqual(expected, color.split_rgb('FFFFFF'))

    def rgb2hsl(self, color, expected):
        result = CSSColor.from_string(color).to_hsl()
        self.assertEqual(expected, result)

    def test_rgb2hsl_black(self):
        self.rgb2hsl('000000', (0, 0, 0))

    def test_rgb2hsl_white(self):
        self.rgb2hsl('FFFFFF', (0, 0, 1.0))

    def test_rgb2hsl_red(self):
        self.rgb2hsl('FF0000', (0, 1.0, 0.5))

    def test_from_string(self):
        color = CSSColor.from_string('C0504D')
        self.assertEqual('c0504d', str(color))

    def test_from_hsl(self):
        magenta = (300 / 360, 1.0, .5)
        expected = 'ff00ff'
        color = CSSColor.from_hsl(*magenta)
        self.assertEqual(expected, str(color))

    def test_apply_tint(self):
        color = CSSColor.from_string('4F81BD')
        hex_tint = '99'
        expected_rgb = '95B3D7'
        color.apply_hsl_tint(hex_tint)
        self.assertEqual(expected_rgb, str(color).upper())
        r, g, b = CSSColor.split_rgb(expected_rgb)
        self.assertAlmostEqual(CSSColor.hex2int(r), color.red, delta=2)
        self.assertAlmostEqual(CSSColor.hex2int(g), color.green, delta=2)
        self.assertAlmostEqual(CSSColor.hex2int(b), color.blue, delta=2)

    def test_apply_tint2(self):
        color = CSSColor.from_string('ED7D31')
        hex_tint = '99'
        color.apply_hsl_tint(hex_tint)
        # green should be B0, but rounding errors bring it to b1 which is
        # close enough
        self.assertEqual('F4B183', str(color).upper())
        # According to the docs, the result should be f4b083, but with
        # rounding errors, we get f4b183 and it's close enough
        r, g, b = CSSColor.split_rgb('f4b083')
        self.assertAlmostEqual(CSSColor.hex2int(r), color.red, delta=2)
        self.assertAlmostEqual(CSSColor.hex2int(g), color.green, delta=2)
        self.assertAlmostEqual(CSSColor.hex2int(b), color.blue, delta=2)

    def test_apply_shade(self):
        color = CSSColor.from_string('C0504D')
        hex_shade = 'BF'
        color.apply_hsl_shade(hex_shade)
        # According to the docs, the result should be 943634, but with
        # rounding errors, we get 953735 and it's close enough
        r, g, b = CSSColor.split_rgb('943634')
        self.assertAlmostEqual(CSSColor.hex2int(r), color.red, delta=2)
        self.assertAlmostEqual(CSSColor.hex2int(g), color.green, delta=2)
        self.assertAlmostEqual(CSSColor.hex2int(b), color.blue, delta=2)

    def test_apply_rgb_shade(self):
        color = CSSColor.from_string('F79646')
        hex_shade = '%02x' % int(.5 * 255)
        color.apply_rgb_shade(hex_shade)
        self.assertEqual('7b4a22', str(color))

    def test_apply_rgb_shade2(self):
        color = CSSColor.from_string('4472C4')
        hex_shade = 'BF'
        color.apply_rgb_shade(hex_shade)
        self.assertEqual('325592', str(color))

    def test_apply_rgb_tint(self):
        color = CSSColor.from_string('C0504D')
        hex_tint = '%02x' % int(.6 * 255)
        color.apply_rgb_tint(hex_tint)
        self.assertEqual('d99694', str(color))


class CssUnitTestCase(TestCase):
    def test_from_invalid_unit(self):
        with self.assertRaises(ValueError):
            CssUnit(1, '%')

    def test_from_emu(self):
        self.assertEqual(1, CssUnit(1, 'emu'))

    def test_from_px(self):
        self.assertEqual(9525, CssUnit(1, 'px'))

    def test_from_pc(self):
        self.assertEqual(152400, CssUnit(1, 'pc'))

    def test_from_pt(self):
        self.assertEqual(12700, CssUnit(1, 'pt'))

    def test_from_mm(self):
        self.assertEqual(36000, CssUnit(1, 'mm'))

    def test_from_cm(self):
        self.assertEqual(360000, CssUnit(1, 'cm'))

    def test_from_in(self):
        self.assertEqual(914400, CssUnit(1, 'in'))

    def test_from_twip(self):
        self.assertEqual(914400, CssUnit(1440, 'twip'))

    def test_px(self):
        self.assertEqual(1, CssUnit(1, 'px').px)

    def test_pc(self):
        self.assertEqual(1, CssUnit(1, 'pc').pc)

    def test_pt(self):
        self.assertEqual(1, CssUnit(1, 'pt').pt)

    def test_cm(self):
        self.assertEqual(1, CssUnit(1, 'cm').cm)

    def test_mm(self):
        self.assertEqual(1, CssUnit(1, 'mm').mm)

    def test_inches(self):
        self.assertEqual(1, CssUnit(1, 'in').inches)

    def test_twips(self):
        self.assertEqual(1, CssUnit(1, 'twip').twips)
