from lxml import etree

from docx2css.ooxml import drawingml
from docx2css.ooxml.constants import CONTENT_TYPE, NAMESPACES as NS
from docx2css.ooxml.simple_types import ST_Theme


class Theme:

    def __init__(self, opc_package):
        self.colors = {}
        self.fonts = {key: None for key in ST_Theme}
        # A theme is optional in the package. There might not be one
        try:
            part = opc_package.parts[CONTENT_TYPE.THEME]
            self.unmarshall_colors(part)
            self.unmarshall_fonts(part)
        except KeyError:
            pass

    def get_color(self, color_name):
        """Return the color associated with the name provided.
        The name must be one of the following values:

        * dk1
        * lt1
        * dk2
        * lt2
        * accent1
        * accent2
        * accent3
        * accent4
        * accent5
        * accent6
        * hlink
        * folHlink

        :returns: Color Hex Code or None if the color is not defined (or
        most likely there is no theme)
        """
        return self.colors.get(color_name, None)

    def get_font(self, font_name):
        """Return font name defined for _font_name_ or None if undefined
        The font_name must be one of the following values:

        * majorAscii
        * majorBidi
        * majorEastAsi
        * majorHAnsi
        * minorAscii
        * minorBidi
        * minorEastAsi
        * minorHAnsi
        """
        return self.fonts.get(font_name, None)

    def unmarshall_colors(self, part):
        color_scheme = part.findall('.//a:clrScheme/a:*', namespaces=NS)
        for color in color_scheme:
            self.colors[etree.QName(color).localname] = color[0].rgb_value

    def unmarshall_fonts(self, part):
        font_scheme = part.find('.//a:fontScheme', namespaces=NS)
        for main_type in ('major', 'minor'):
            el = font_scheme.find(f'./a:{main_type}Font', namespaces=NS)
            latin = el.find('./a:latin', namespaces=NS).get('typeface'),
            ea = el.find('./a:ea', namespaces=NS).get('typeface'),
            cs = el.find('./a:cs', namespaces=NS).get('typeface'),
            self.fonts[f'{main_type}Ascii'] = latin[0]
            self.fonts[f'{main_type}HAnsi'] = latin[0]
            self.fonts[f'{main_type}Bidi'] = cs[0]
            self.fonts[f'{main_type}EastAsia'] = ea[0]


@drawingml('sysClr')
class SystemColor(etree.ElementBase):

    @property
    def rgb_value(self):
        return self.get('lastClr')


@drawingml('srgbClr')
class RGBColor(etree.ElementBase):

    @property
    def rgb_value(self):
        return self.get('val')
