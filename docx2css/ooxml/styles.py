from abc import ABC, abstractmethod
from collections.abc import Mapping

from lxml import etree

from docx2css.api import Border, TextDecoration
from docx2css.ooxml import w, wordml
from docx2css.ooxml.constants import CONTENT_TYPE, NAMESPACES
from docx2css.ooxml.simple_types import (
    ST_Border, ST_FontFamily, ST_Underline, ST_Jc
)
from docx2css.utils import AutoLength, CSSColor, CssUnit, Percentage


class Styles(Mapping):

    def __init__(self, opc_package):
        self.__styles__ = {}
        self.opc_package = opc_package
        styles_part = opc_package.parts[CONTENT_TYPE.STYLES]
        for style in (s for s in styles_part if isinstance(s, DocxStyle)):
            style.styles = self
            if style.name != 'Default Paragraph Font':
                self.__styles__[style.id] = style
        self.doc_defaults = styles_part.find(w('docDefaults'))
        self.doc_defaults.styles = self

    def __getitem__(self, k):
        return self.__styles__[k]

    def __len__(self) -> int:
        return len(self.__styles__)

    def __iter__(self):
        return iter(self.__styles__)

    @property
    def character_styles(self):
        return self.get_styles_by_type('character')

    @property
    def paragraph_styles(self):
        return self.get_styles_by_type('paragraph')

    @property
    def numbering_styles(self):
        return self.get_styles_by_type('numbering')

    def get_styles_by_type(self, style_type):
        return {k: v for k, v in self.__styles__.items()
                if v.type == style_type}

    def print_styles_tree(self, styles=None, level=0):
        s = ''
        for style in styles:
            s += f'|{level * 3 * "-"} {style.name}\n'
            s += self.print_styles_tree(style.children_styles, level+1)
        return s


@wordml('style')
class DocxStyle(etree.ElementBase):

    @property
    def name(self):
        return self.find(w('name')).get(w('val'))

    @property
    def id(self):
        return self.get(w('styleId'))

    @property
    def type(self):
        return self.get(w('type'))

    @property
    def parent_id(self):
        el = self.find(w('basedOn'))
        if el is not None:
            return el.get(w('val'))

    @property
    def parent(self):
        """Return the parent style or None"""
        return self.styles.get(self.parent_id, None)

    @property
    def styles(self):
        return getattr(self, '_styles', None)

    @styles.setter
    def styles(self, styles):
        setattr(self, '_styles', styles)

    @property
    def children_styles(self):
        if not hasattr(self, '_children_styles'):
            children = [s for s in self.styles.values()
                        if s.parent_id == self.id]
            setattr(self, '_children_styles', children)
        return self._children_styles


class RPrProxy(etree.ElementBase):

    @property
    def all_caps(self):
        element = self.find('.//w:caps', namespaces=NAMESPACES)
        if isinstance(element, AllCapsProperty):
            return element.prop_value

    @property
    def background_color(self):
        element = self.find('.//w:shd', namespaces=NAMESPACES)
        if isinstance(element, ShadingProperty):
            return element.prop_value

    @property
    def bold(self):
        element = self.find('.//w:b', namespaces=NAMESPACES)
        if isinstance(element, BoldProperty):
            return element.prop_value

    @property
    def border(self):
        element = self.find('.//w:bdr', namespaces=NAMESPACES)
        if isinstance(element, BorderProperty):
            return element.prop_value

    @property
    def double_strike(self):
        element = self.find('.//w:dstrike', namespaces=NAMESPACES)
        if isinstance(element, DStrikeProperty):
            return element.prop_value

    @property
    def emboss(self):
        element = self.find('.//w:emboss', namespaces=NAMESPACES)
        if isinstance(element, EmbossProperty):
            return element.prop_value

    @property
    def font_color(self):
        element = self.find('.//w:color', namespaces=NAMESPACES)
        if isinstance(element, FontColorProperty):
            return element.prop_value

    @property
    def font_family(self):
        element = self.find('.//w:rFonts', namespaces=NAMESPACES)
        if element is not None:
            return element.prop_value

    @property
    def font_kerning(self):
        element = self.find('.//w:kern', namespaces=NAMESPACES)
        if isinstance(element, FontKerningProperty):
            return element.prop_value

    @property
    def font_size(self):
        element = self.find('.//w:sz', namespaces=NAMESPACES)
        if element is not None:
            return element.prop_value

    @property
    def highlight(self):
        element = self.find('.//w:highlight', namespaces=NAMESPACES)
        if isinstance(element, HighlightProperty):
            return element.prop_value

    @property
    def imprint(self):
        element = self.find('.//w:imprint', namespaces=NAMESPACES)
        if isinstance(element, ImprintProperty):
            return element.prop_value

    @property
    def italics(self):
        element = self.find('.//w:i', namespaces=NAMESPACES)
        if isinstance(element, ItalicProperty):
            return element.prop_value

    @property
    def letter_spacing(self):
        element = self.find('.//w:spacing', namespaces=NAMESPACES)
        if isinstance(element, FontSpacingProperty):
            return element.letter_spacing()

    @property
    def outline(self):
        element = self.find('.//w:outline', namespaces=NAMESPACES)
        if isinstance(element, OutlineProperty):
            return element.prop_value

    @property
    def position(self):
        element = self.find('.//w:position', namespaces=NAMESPACES)
        if isinstance(element, PositionProperty):
            return element.prop_value

    @property
    def shadow(self):
        element = self.find('.//w:shadow', namespaces=NAMESPACES)
        if isinstance(element, ShadowProperty):
            return element.prop_value

    @property
    def small_caps(self):
        element = self.find('.//w:smallCaps', namespaces=NAMESPACES)
        if isinstance(element, SmallCapsProperty):
            return element.prop_value

    @property
    def strike(self):
        element = self.find('.//w:strike', namespaces=NAMESPACES)
        if isinstance(element, StrikeProperty):
            return element.prop_value

    @property
    def underline(self):
        element = self.find('.//w:u', namespaces=NAMESPACES)
        if isinstance(element, UnderlineProperty):
            return element.prop_value

    @property
    def vertical_align(self):
        element = self.find('.//w:vertAlign', namespaces=NAMESPACES)
        if isinstance(element, VerticalAlignProperty):
            return element.prop_value

    @property
    def visible(self):
        element = self.find('.//w:vanish', namespaces=NAMESPACES)
        if isinstance(element, VanishProperty):
            return element.prop_value


class PPrProxy(etree.ElementBase):

    @property
    def border_bottom(self):
        element = self.find('.//w:pBdr/w:bottom', namespaces=NAMESPACES)
        if isinstance(element, BorderBottomProperty):
            return element.prop_value

    @property
    def border_left(self):
        element = self.find('.//w:pBdr/w:left', namespaces=NAMESPACES)
        if isinstance(element, BorderLeftProperty):
            return element.prop_value

    @property
    def border_top(self):
        element = self.find('.//w:pBdr/w:top', namespaces=NAMESPACES)
        if isinstance(element, BorderTopProperty):
            return element.prop_value

    @property
    def border_right(self):
        element = self.find('.//w:pBdr/w:right', namespaces=NAMESPACES)
        if isinstance(element, BorderRightProperty):
            return element.prop_value

    @property
    def keep_together(self):
        element = self.find('.//w:keepLines', namespaces=NAMESPACES)
        if isinstance(element, KeepLinesTogetherProperty):
            return element.prop_value

    @property
    def keep_with_next(self):
        element = self.find('.//w:keepNext', namespaces=NAMESPACES)
        if isinstance(element, KeepWithNextParagraphProperty):
            return element.prop_value

    @property
    def line_height(self):
        element = self.find('.//w:spacing', namespaces=NAMESPACES)
        if isinstance(element, FontSpacingProperty):
            return element.line_height()

    @property
    def margin_left(self):
        element = self.find('.//w:ind', namespaces=NAMESPACES)
        if isinstance(element, IndentProperty):
            return element.margin_left()

    @property
    def margin_right(self):
        element = self.find('.//w:ind', namespaces=NAMESPACES)
        if isinstance(element, IndentProperty):
            return element.margin_right()

    @property
    def space_after(self):
        element = self.find('.//w:spacing', namespaces=NAMESPACES)
        if isinstance(element, FontSpacingProperty):
            return element.margin_bottom()

    @property
    def space_before(self):
        element = self.find('.//w:spacing', namespaces=NAMESPACES)
        if isinstance(element, FontSpacingProperty):
            return element.margin_top()

    @property
    def numbering_instance_id(self):
        element = self.find('.//w:numPr', namespaces=NAMESPACES)
        if isinstance(element, NumberingProperty):
            return element.id

    @property
    def numbering_instance_level(self):
        element = self.find('.//w:numPr', namespaces=NAMESPACES)
        if isinstance(element, NumberingProperty):
            return element.level

    @property
    def page_break_before(self):
        element = self.find('.//w:pageBreakBefore', namespaces=NAMESPACES)
        if isinstance(element, PageBreakBeforeProperty):
            return element.prop_value

    @property
    def text_align(self):
        element = self.find('.//w:jc', namespaces=NAMESPACES)
        if isinstance(element, JustificationProperty):
            return element.prop_value

    @property
    def text_indent(self):
        element = self.find('.//w:ind', namespaces=NAMESPACES)
        if isinstance(element, IndentProperty):
            return element.text_indent()

    @property
    def widows_control(self):
        element = self.find('.//w:widowControl', namespaces=NAMESPACES)
        if isinstance(element, WidowControlProperty):
            return element.prop_value


class DocxCharacterStyle(DocxStyle, RPrProxy):
    pass


class DocxParagraphStyle(DocxStyle, PPrProxy, RPrProxy):
    pass


class DocxNumberingStyle(DocxStyle, PPrProxy):
    pass


class DocxTableStyle(DocxStyle):
    pass


@wordml('docDefaults')
class DocDefaults(RPrProxy, PPrProxy):

    @property
    def styles(self):
        return getattr(self, '_styles', None)

    @styles.setter
    def styles(self, styles):
        setattr(self, '_styles', styles)


class DocxPropertyAdapter(etree.ElementBase, ABC):

    @property
    @abstractmethod
    def prop_name(self):
        pass

    @property
    @abstractmethod
    def prop_value(self):
        pass

    @property
    def docx_parser(self):
        return getattr(self, '_docx_parser')

    @docx_parser.setter
    def docx_parser(self, package):
        setattr(self, '_docx_parser', package)

    def get_measure(self):
        """Parse a TblWidth as a Measure"""
        unit = self.get(w('type'))
        value = self.get(w('w'))
        if unit == 'auto':
            return AutoLength()
        elif unit == 'dxa':
            return CssUnit(value, 'twip')
        elif unit == 'nil':
            return CssUnit(0)
        elif unit == 'pct':
            return Percentage(int(value) / 50)
        else:
            raise ValueError(f'Unit "{unit}" is invalid!')

    def get_boolean_attribute(self, name):
        """Get the boolean value of an attribute or None if the
        attribute doesn't exist.
        """
        attribute_value = self.get(w(name))
        if attribute_value is not None:
            return not attribute_value.lower() in ('false', '0')
        else:
            return None

    def get_toggle_property(self, name):
        """Parse a toggle (boolean) property from a child 'name' of the
        xml element.

        :returns: bool value or None if the child 'name' doesn't exist
        """
        prop = self.getparent().find(w(name))
        if prop is None:
            return None
        attribute_value = prop.get(w('val'))
        if attribute_value:
            return not attribute_value.lower() in ('false', '0')
        else:
            return True

    def get_opc_package(self):
        ancestors = list(self.iterancestors(w('style'), w('docDefaults')))
        if len(ancestors):
            styles = ancestors[0].styles
            package = styles.opc_package
        else:
            numbering = list(self.iterancestors(w('abstractNum')))
            numbering_part = numbering[0].numbering_part
            package = numbering_part.opc_package
        return package

    def get_theme(self):
        package = self.get_opc_package()
        return package.theme

    def get_font_table(self):
        package = self.get_opc_package()
        return package.font_table


class ColorPropertyAdapter(DocxPropertyAdapter, ABC):

    color_attribute = 'val'
    theme_color_attribute = 'themeColor'
    theme_shade_attribute = 'themeShade'
    theme_tint_attribute = 'themeTint'

    def get_color(self):
        theme_color = self.get(w(self.theme_color_attribute))
        if theme_color is not None:
            theme = self.get_theme()
            color = CSSColor.from_string(theme.get_color(theme_color))
            shade = self.get(w(self.theme_shade_attribute))
            if shade is not None:
                color.apply_hsl_shade(shade)
            tint = self.get(w(self.theme_tint_attribute))
            if tint is not None:
                color.apply_rgb_tint(tint)
        else:
            color = self.get(w(self.color_attribute))
        return f'#{color}' if color is not None and color != 'auto' else ''


@wordml('caps')
class AllCapsProperty(DocxPropertyAdapter):
    prop_name = 'all_caps'

    @property
    def prop_value(self):
        return self.get_toggle_property('caps')


@wordml('b')
class BoldProperty(DocxPropertyAdapter):
    prop_name = 'bold'

    @property
    def prop_value(self):
        return self.get_toggle_property('b')


@wordml('bdr')
class BorderProperty(ColorPropertyAdapter):
    prop_name = 'border'
    color_attribute = 'color'
    direction = ''  # Direction (top, left, bottom, right) of the border

    @property
    def prop_value(self):
        return Border(
            color=self.color,
            padding=self.padding,
            shadow=self.shadow,
            style=self.style,
            width=self.width
        )

    @property
    def color(self):
        """Get the optional color of the border taking in consideration
        that the color can be defined with different attributes such as
        'color', 'themeColor', 'themeTint' or 'themeShade'
        Returns a #Hex code, or None if the color is undefined or if its
        value is set to 'auto'
        """
        color = self.get_color()
        return color if color and color != 'auto' else None

    @property
    def padding(self):
        """Get the padding that shall be used to place this border on
        the parent object
        """
        space = self.get(w('space'))
        if space is None:
            return None
        return CssUnit(int(space), 'pt')

    @property
    def shadow(self):
        """Specifies whether this border should be modified to create
        the appearance of a shadow."""
        attribute_value = self.get(w('shadow'))
        if attribute_value:
            return not attribute_value.lower() in ('false', '0')
        else:
            return False

    @property
    def style(self) -> str:
        """Get the line style as a string.

        Possible values:
            * none;
            * dotted;
            * dashed;
            * solid;
            * double;
            * groove;
            * ridge;
            * inset;
            * outset;
        """
        value = self.get(w('val'))
        return ST_Border.css_value(value)

    @property
    def width(self):
        """Get the width of the border

        :returns: CssUnit
        """
        width = self.get(w('sz'))
        # The 'sz' attribute is in 8th of a pt.
        if width is None:
            return None
        return CssUnit(int(width) / 8, 'pt')


@wordml('bottom')
class BorderBottomProperty(BorderProperty):
    prop_name = 'border_bottom'
    direction = 'bottom'


@wordml('left')
class BorderLeftProperty(BorderProperty):
    prop_name = 'border_left'
    direction = 'left'


@wordml('right')
class BorderRightProperty(BorderProperty):
    prop_name = 'border_right'
    direction = 'right'


@wordml('top')
class BorderTopProperty(BorderProperty):
    prop_name = 'border_top'
    direction = 'top'


@wordml('dstrike')
class DStrikeProperty(DocxPropertyAdapter):
    prop_name = 'double_strike'

    @property
    def prop_value(self):
        return self.get_toggle_property('dstrike')


@wordml('emboss')
class EmbossProperty(DocxPropertyAdapter):
    prop_name = 'emboss'

    @property
    def prop_value(self):
        return self.get_toggle_property('emboss')


@wordml('color')
class FontColorProperty(ColorPropertyAdapter):
    prop_name = 'font_color'

    @property
    def prop_value(self):
        return self.get_color()


@wordml('kern')
class FontKerningProperty(DocxPropertyAdapter):
    prop_name = 'font_kerning'

    @property
    def prop_value(self):
        value = int(self.get(w('val'))) / 2  # Value is in half-points
        return value != 0


@wordml('rFonts')
class FontProperty(DocxPropertyAdapter):
    prop_name = 'font_family'

    def _wrap(self, font_name):
        """Wrap a font name in quotes if it has a space in it"""
        if font_name is not None:
            return f'"{font_name}"' if ' ' in font_name else font_name

    def _get_theme_font_or_font_value(self, font_name):
        """Get the theme font associated with the theme or return the
        same value if it's not a theme color
        """
        theme = self.get_theme()
        if font_name in theme.fonts:
            font_name = theme.get_font(font_name)
        return font_name

    def _get_font_from_font_table(self, font_name):
        font_table = self.get_font_table()
        font = font_table.get_font(font_name)
        if font is not None:
            return font.css_family
        return font_name,

    @property
    def prop_value(self):
        # What we want to do here is have a set of the fonts, but at the
        # same time, we want to keep the order so it's easier to use a
        # dict because the order is guaranteed
        fonts = {}
        # Theme values take precedence over explicit values, so we
        # favour the former
        attributes = (
            self.get(w('hAnsiTheme')) or self.get(w('hAnsi')),
            self.get(w('asciiTheme')) or self.get(w('ascii')),
            self.get(w('eastAsiaTheme')) or self.get(w('eastAsia')),
            self.get(w('cstheme')) or self.get(w('cs')),
        )
        for attribute in attributes:
            font_name = self._get_theme_font_or_font_value(attribute)
            if font_name:
                for f in self._get_font_from_font_table(font_name):
                    fonts[self._wrap(f)] = None
        # Push the generic family at the end. This happens when different
        # fonts are specified, and they are found in the font table
        for generic in ST_FontFamily.docx2css.values():
            if generic in fonts:
                value = fonts.pop(generic)
                fonts[generic] = value
        return ', '.join(fonts.keys())


@wordml('spacing')
class FontSpacingProperty(DocxPropertyAdapter):
    """
    This element can represent the font spacing property, the line height,
    or the margins between paragraphs.
    """

    @property
    def prop_name(self):
        if self.get(w('val')) is not None:
            return 'letter_spacing'
        else:
            return 'line_height', 'margin_top', 'margin_bottom'

    @property
    def prop_value(self):
        if self.prop_name == 'letter_spacing':
            value = self.get(w('val'))
            return CssUnit(int(value), 'twip')  # Value is in 20th of a point
        else:
            return self.line_height(), self.margin_top(), self.margin_bottom()

    def letter_spacing(self):
        value = self.get(w('val'))
        if value is not None:
            return CssUnit(int(value), 'twip')  # Value is in 20th of a point

    def line_height(self):
        height = self.get(w('line'))
        rule = self.get(w('lineRule'))
        if height is not None:
            if rule in ('atLeast', 'exact'):
                return CssUnit(int(height), 'twip')
            elif rule == 'auto':
                # Height is 240th of a line
                return int(height) / 240
                # return AutoLength(height)

    def margin_bottom(self):
        after = self.get(w('after'))
        auto = self.get_boolean_attribute('afterAutospacing')
        if auto is not True and after is not None:
            # after = int(after) / 20  # Value is in 20th of a point
            # return f'{after:.2f}pt'
            return CssUnit(after, 'twip')

    def margin_top(self):
        before = self.get(w('before'))
        auto = self.get_boolean_attribute('beforeAutospacing')
        if auto is not True and before is not None:
            # before = int(before) / 20  # Value is in 20th of a point
            # return f'{before:.2f}pt'
            return CssUnit(before, 'twip')


@wordml('highlight')
class HighlightProperty(DocxPropertyAdapter):
    prop_name = 'highlight'

    @property
    def prop_value(self):
        return self.get(w('val')).lower()


@wordml('imprint')
class ImprintProperty(DocxPropertyAdapter):
    prop_name = 'imprint'

    @property
    def prop_value(self):
        return self.get_toggle_property('imprint')


@wordml('ind')
class IndentProperty(DocxPropertyAdapter):
    prop_name = (
        'margin_left',
        'margin_right',
        'text_indent',
    )

    @property
    def prop_value(self):
        return (
            self.margin_left(),
            self.margin_right(),
            self.text_indent(),
        )

    def first_line_indent(self):
        first_line = self.get(w('firstLine'))
        if first_line is not None:
            return CssUnit(int(first_line), 'twip')

    def hanging_indent(self):
        hanging = self.get(w('hanging'))
        if hanging is not None:
            return CssUnit(-1 * int(hanging), 'twip')

    def text_indent(self):
        indents = (self.first_line_indent(), self.hanging_indent())
        return next((x for x in indents if x is not None), None)

    def margin_left(self):
        left = self.get(w('start')) or self.get(w('left'))
        if left is not None:
            return CssUnit(int(left), 'twip')

    def margin_right(self):
        right = self.get(w('end')) or self.get(w('right'))
        if right is not None:
            return CssUnit(int(right), 'twip')


@wordml('i')
class ItalicProperty(DocxPropertyAdapter):
    prop_name = 'italics'

    @property
    def prop_value(self):
        return self.get_toggle_property('i')


@wordml('jc')
class JustificationProperty(DocxPropertyAdapter):

    @property
    def prop_name(self):
        if self.getparent().tag == w('pPr'):
            return 'text_align'
        else:
            return 'alignment'

    @property
    def prop_value(self):
        attr_value = self.get(w('val'))
        return ST_Jc.css_value(attr_value)


@wordml('keepLines')
class KeepLinesTogetherProperty(DocxPropertyAdapter):
    prop_name = 'break_inside'

    @property
    def prop_value(self):
        return self.get_toggle_property('keepLines')


@wordml('keepNext')
class KeepWithNextParagraphProperty(DocxPropertyAdapter):
    prop_name = 'break_after'

    @property
    def prop_value(self):
        return self.get_toggle_property('keepNext')


@wordml('numPr')
class NumberingProperty(etree.ElementBase):

    @property
    def id(self):
        element = self.find(w('numId'))
        if element is not None:
            value = int(element.get(w('val')))
            # numbering = self.styles.opc_package.numbering
            # return numbering.numbering_instances[value]
            return value

    @property
    def level(self):
        element = self.find(w('ilvl'))
        return int(element.get(w('val'))) if element is not None else None


@wordml('outline')
class OutlineProperty(DocxPropertyAdapter):
    prop_name = 'outline'

    @property
    def prop_value(self):
        return self.get_toggle_property('outline')


@wordml('pageBreakBefore')
class PageBreakBeforeProperty(DocxPropertyAdapter):
    prop_name = 'break_before'

    @property
    def prop_value(self):
        return self.get_toggle_property('pageBreakBefore')


@wordml('position')
class PositionProperty(DocxPropertyAdapter):
    prop_name = 'position'

    @property
    def prop_value(self):
        value = int(self.get(w('val'))) / 2  # Value in half-points
        return CssUnit(value, 'pt')


@wordml('shd')
class ShadingProperty(ColorPropertyAdapter):
    prop_name = 'background_color'
    color_attribute = 'fill'
    theme_color_attribute = 'themeFill'
    theme_shade_attribute = 'themeFillShade'
    theme_tint_attribute = 'themeFillTint'

    @property
    def prop_value(self):
        return self.get_color()


@wordml('shadow')
class ShadowProperty(DocxPropertyAdapter):
    prop_name = 'shadow'

    @property
    def prop_value(self):
        return self.get_toggle_property('shadow')


@wordml('smallCaps')
class SmallCapsProperty(DocxPropertyAdapter):
    prop_name = 'small_caps'

    @property
    def prop_value(self):
        return self.get_toggle_property('smallCaps')


@wordml('strike')
class StrikeProperty(DocxPropertyAdapter):
    prop_name = 'strike'

    @property
    def prop_value(self):
        return self.get_toggle_property('strike')


@wordml('sz')
class SizeProperty(DocxPropertyAdapter):
    prop_name = 'font_size'

    @property
    def prop_value(self):
        sz = int(self.get(w('val')))
        return CssUnit(sz / 2, 'pt')


@wordml('u')
class UnderlineProperty(ColorPropertyAdapter):
    color_attribute = 'color'
    prop_name = 'underline'

    @property
    def prop_value(self):
        color = self.get_color()
        style = ST_Underline.css_value(self.get(w('val')))
        value = TextDecoration(color=color, style=style)
        if style != 'none':
            value.add_line(TextDecoration.UNDERLINE)
        return value


@wordml('vanish')
class VanishProperty(DocxPropertyAdapter):
    prop_name = 'visible'

    @property
    def prop_value(self):
        return self.get_toggle_property('vanish')


@wordml('vertAlign')
class VerticalAlignProperty(DocxPropertyAdapter):
    prop_name = 'vertical_align'

    @property
    def prop_value(self):
        return self.get(w('val'))


@wordml('widowControl')
class WidowControlProperty(DocxPropertyAdapter):
    prop_name = 'widows'

    @property
    def prop_value(self):
        return self.get_toggle_property('widowControl')
