from collections import namedtuple
import colorsys
import textwrap
from dataclasses import dataclass, fields

KeyValueProperty = namedtuple('KeyValueProperty', ['name', 'value'])


@dataclass
class PropertyContainer:

    def properties(self):
        exclude = ('name', 'id', 'parent', 'parent_id', 'children', 'type')

        def accept(f):
            result = f.name not in exclude and getattr(self, f.name) is not None
            return result
        return (KeyValueProperty(f.name, getattr(self, f.name))
                for f in fields(self) if accept(f))


class CSSColor:
    
    def __init__(self, red=0, green=0, blue=0):
        self.red = red
        self.green = green
        self.blue = blue

    @classmethod
    def hex2int(cls, hex_number):
        """Returns the int represented by the hex number string"""
        return int(hex_number, base=16)

    @classmethod
    def split_rgb(cls, rgb):
        """Split the three parts of an RGB color code

        :returns: 3-tuple (R, G, B)
        """
        return textwrap.wrap(rgb, 2)

    def to_hsl(self):
        h, l, s = colorsys.rgb_to_hls(self.red / 255, self.green / 255, self.blue / 255)
        return h, s, l

    def apply_hsl_shade(self, hex_value):
        """
        Apply shade to the color as per the border specs in the ooxml
        reference.

        :param hex_value: String representing a Hex value of the shade
        to apply. The string will be converted to a base 10 and divided
        by 255 to get a percentage.
        """
        shade = self.hex2int(hex_value) / 255
        h, s, l = self.to_hsl()
        l2 = l * shade
        self.red, self.green, self.blue = (round(x * 255)
                                           for x in colorsys.hls_to_rgb(h, l2, s))

    def apply_hsl_tint(self, hex_value):
        """
        Apply a tint as per the border specs in the ooxml reference
        :param hex_value:
        """
        tint = self.hex2int(hex_value) / 255
        h, s, l = self.to_hsl()
        l2 = l * tint + (1 - tint)
        self.red, self.green, self.blue = (round(x*255)
                                           for x in colorsys.hls_to_rgb(h, l2, s))

    def apply_rgb_shade(self, hex_percentage):
        """
        Apply a shade percentage to each of the RGB components of the
        color. The formula used is defined at sect. 17.3.2.6 of the
        docx specs (font color).

        This method returns a new instance without modifying the current
        object properties

        :param hex_percentage: String representing a Hex value of the shade
        to apply. The string will be converted to a base 10 and divided
        by 255 to get a percentage.
        :return: CSSColor instance
        """
        shade = self.hex2int(hex_percentage) / 255
        def transform(x): return int(shade * x)
        self.red = transform(self.red)
        self.green = transform(self.green)
        self.blue = transform(self.blue)

    def apply_rgb_tint(self, hex_percentage):
        """
        Apply a tint percentage on each of the RGB components of the
        color. The formula is defined at sect 17.3.2.6 of the docx
        specs (font color).

        :param hex_percentage: String representing a Hex value of the tint
        to apply. The string will be converted to a base 10 and divided
        by 255 to get a percentage.
        """
        tint = self.hex2int(hex_percentage) / 255
        def transform(x): return int((1 - tint) * (255 - x) + x)
        self.red = transform(self.red)
        self.green = transform(self.green)
        self.blue = transform(self.blue)

    @classmethod
    def from_hsl(cls, h, s, l):
        r, g, b = (round(x*255) for x in colorsys.hls_to_rgb(h, l, s))
        return cls(red=r, green=g, blue=b)

    @classmethod
    def from_string(cls, rgb):
        """
        Create a new instance using a RGB hex String (with or without '#')
        
        :param rgb: RGB hex String
        :return: New CssColor instance
        """
        if rgb[0] == '#':
            rgb = rgb[1:]
        r, g, b = map(lambda x: cls.hex2int(x), cls.split_rgb(rgb))
        return cls(red=r, green=g, blue=b)

    def __str__(self):
        return '%02x%02x%02x' % (self.red, self.green, self.blue)


class CssUnit(int):
    _EMUS_PER_INCH = 914400
    _EMUS_PER_CM = 360000
    _EMUS_PER_MM = 36000
    _EMUS_PER_PT = 12700
    _EMUS_PER_TWIP = 635
    _EMUS_PER_PX = _EMUS_PER_INCH / 96
    _EMUS_PER_PC = _EMUS_PER_PT * 12

    def __new__(cls, value, unit='emu'):
        return int.__new__(cls, cls.to_emu(value, unit))

    @classmethod
    def to_emu(cls, value, unit):
        if unit == 'px':
            return float(value) * float(cls._EMUS_PER_INCH) / 96
        elif unit == 'pc':
            return float(value) * float(cls._EMUS_PER_PT) * 12
        elif unit == 'pt':
            return float(value) * float(cls._EMUS_PER_PT)
        elif unit == 'mm':
            return float(value) * float(cls._EMUS_PER_MM)
        elif unit == 'cm':
            return float(value) * float(cls._EMUS_PER_CM)
        elif unit == 'in':
            return float(value) * float(cls._EMUS_PER_INCH)
        elif unit == 'twip':
            return float(value) * float(cls._EMUS_PER_TWIP)
        elif unit == 'emu':
            return value
        else:
            raise ValueError(
                f'{unit} is not a valid unit. '
                'Choices are: px, pc, pt, mm, cm, in, twip, emu'
            )

    @property
    def px(self):
        return self / float(self._EMUS_PER_PX)

    @property
    def pc(self):
        return self / float(self._EMUS_PER_PC)

    @property
    def pt(self):
        return self / float(self._EMUS_PER_PT)

    @property
    def cm(self):
        return self / float(self._EMUS_PER_CM)

    @property
    def mm(self):
        return self / float(self._EMUS_PER_MM)

    @property
    def inches(self):
        return self / float(self._EMUS_PER_INCH)

    @property
    def twips(self):
        return int(round(self / float(self._EMUS_PER_TWIP)))

    def to(self, unit):
        if unit == 'px':
            return self.px
        elif unit == 'pc':
            return self.pc
        elif unit == 'pt':
            return self.pt
        elif unit == 'mm':
            return self.mm
        elif unit == 'cm':
            return self.cm
        elif unit == 'in':
            return self.inches
        else:
            return self


class AutoLength(CssUnit):

    def __new__(cls, value=0):
        return int.__new__(cls, value)


class Percentage(CssUnit):

    def __new__(cls, value):
        return int.__new__(cls, value * 100)

    @property
    def pct(self):
        return self / 100
