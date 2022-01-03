from lxml import etree

from docx2css.ooxml import w, wordml
from docx2css.ooxml.constants import CONTENT_TYPE, NAMESPACES as NS
from docx2css.ooxml.simple_types import ST_FontFamily


class FontTable:

    def __init__(self, opc_package):
        self.fonts = {}
        try:
            part = opc_package.parts[CONTENT_TYPE.FONTS]
            self._unmarshall_fonts(part)
        except KeyError:
            pass

    def _unmarshall_fonts(self, font_table):
        for font in font_table:
            if isinstance(font, Font):
                self.fonts[font.name] = font

    def get_font(self, font_name):
        return self.fonts.get(font_name, None)


@wordml('font')
class Font(etree.ElementBase):

    @property
    def name(self):
        return self.get(w('name'))

    @property
    def alt_name(self):
        element = self.find('w:altName', namespaces=NS)
        return element.get(w('val')) if element is not None else None

    @property
    def family(self):
        element = self.find('w:family', namespaces=NS)
        return element.get(w('val')) if element is not None else None

    @property
    def css_family(self):
        """Returns a tuple of font names appropriate for CSS font-family
        property, including altName and family
        """
        values = (self.name, self.alt_name, self.css_generic_family)
        return tuple(v for v in values if v is not None)

    @property
    def css_generic_family(self):
        """Returns the font's generic family as a valid CSS value, or
        None if the family is not defined
        """
        return ST_FontFamily.css_value(self.family)
