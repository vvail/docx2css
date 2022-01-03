import colorsys
import textwrap


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
        print('applying shade:', shade, 'on', self)
        print('before:', self.red, self.green, self.blue)
        def transform(x): return int(shade * x)
        self.red = transform(self.red)
        self.green = transform(self.green)
        self.blue = transform(self.blue)
        print('After:', self.red, self.green, self.blue)

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
        print('applying tint:', tint)
        print('before:', self.red, self.green, self.blue)
        def transform(x): return int((1 - tint) * (255 - x) + x)
        self.red = transform(self.red)
        self.green = transform(self.green)
        self.blue = transform(self.blue)
        print('After:', self.red, self.green, self.blue)

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
