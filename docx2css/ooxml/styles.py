import re
from abc import ABC, abstractmethod
from collections.abc import Mapping

import cssutils
from lxml import etree

from docx2css.ooxml import w, wordml
from docx2css.ooxml.constants import CONTENT_TYPE, NAMESPACES
from docx2css.ooxml.simple_types import (
    ST_Border, ST_FontFamily, ST_Underline, ST_Jc
)
from docx2css.utils import CSSColor


class Styles(Mapping):

    def __init__(self, opc_package):
        self.__styles__ = {}
        self.opc_package = opc_package
        styles_part = opc_package.parts[CONTENT_TYPE.STYLES]
        for style in (s for s in styles_part if isinstance(s, DocxStyle)):
            style.stylesheet = self
            if style.name != 'Default Paragraph Font':
                self.__styles__[style.id] = style

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
        return self.stylesheet.get(self.parent_id, None)

    @property
    def stylesheet(self):
        return getattr(self, '_stylesheet', None)

    @stylesheet.setter
    def stylesheet(self, stylesheet):
        setattr(self, '_stylesheet', stylesheet)

    @property
    def numbering(self):
        element = self.find('.//w:numPr', namespaces=NAMESPACES)
        if element is not None:
            element.styles = self.stylesheet
        return element

    def get_numbering_definition(self):
        """Return the definition of the numbering, meaning the method
        will lookup the numId and return the actual AbstractNumbering.
        The numId can be inherited from the parent style.
        """
        definition = self.numbering.definition
        if definition is None:
            return self.parent.get_numbering_definition()
        else:
            return definition

    @property
    def children_styles(self):
        if not hasattr(self, '_children_styles'):
            children = [s for s in self.stylesheet.values()
                        if s.parent_id == self.id]
            setattr(self, '_children_styles', children)
        return self._children_styles

    def css_current_selector(self):
        """Get the selector for the current style only"""
        return f"{self.css_selector_prefix}.{self.id}"

    def css_selector(self):
        """Get the CSS selector for this style, including all the
        children
        """
        names = [self.css_current_selector()]

        for child in self.children_styles:
            names.append(child.css_selector())
        return ', '.join(names)

    @property
    def css_selector_prefix(self):
        return ''

    @property
    def css_properties(self):
        for d in self.iterdescendants():
            if isinstance(d, CssPropertyAdapter):
                yield d

    def css_style_declaration(self):
        css_style = cssutils.css.CSSStyleDeclaration()
        package = self.stylesheet.opc_package
        for prop in self.css_properties:
            prop.set_css_style(css_style, package)
        return css_style

    def css_style_rule(self):
        if not hasattr(self, '_css_style_rule'):
            css_style = self.css_style_declaration()
            rule = cssutils.css.CSSStyleRule(self.css_selector(), style=css_style)
            setattr(self, '_css_style_rule', rule)
        return getattr(self, '_css_style_rule')

    def css_style_rules(self):
        return self.css_style_rule(),


class CharacterStyle(DocxStyle):

    @property
    def css_selector_prefix(self):
        return 'span'


class ParagraphStyle(DocxStyle):

    def css_current_selector(self):
        class_name = f'.{self.id}'
        prefix = self.css_selector_prefix
        if class_name == '.Normal' or re.match('h[1-6]', prefix):
            return prefix
        else:
            return f"{self.css_selector_prefix}{class_name}"

    def css_selector(self):
        current_selector = self.css_current_selector()
        names = [current_selector]

        def not_p(s):
            return s.parent_id == 'Normal' and not s.css_selector_prefix == 'p'

        # Treat Normal style a bit differently
        if current_selector == 'p':
            children = filter(not_p, self.children_styles)
        else:
            children = self.children_styles

        for child in children:
            names.append(child.css_selector())
        return ', '.join(names)

    @property
    def css_selector_prefix(self):
        class_name = ''.join(self.name.split())
        regex = re.match('heading([1-6])', class_name)
        if regex:
            return f'h{regex.group(1)}'
        return 'p'

    def css_style_rule(self):
        if not hasattr(self, '_css_style_rule'):
            css_style = self.css_style_declaration()
            if self.numbering is not None:
                self.pull_margin_left(css_style)
                self.pull_text_indent(css_style)
                self.adjust_for_indent(css_style)

                # Also, it is necessary to reset the appropriate counters
                # since it cannot be done at on the pseudo :before element
                level = self.get_numbering_level_for_style()
                css_style['counter-reset'] = level.css_counter_resets()
            rule = cssutils.css.CSSStyleRule(self.css_selector(), style=css_style)
            setattr(self, '_css_style_rule', rule)
        return getattr(self, '_css_style_rule')

    def css_style_rules(self):
        if self.numbering is not None:
            return (
                self.css_numbering_style_rule(),
                self.css_style_rule(),
            )
        else:
            return self.css_style_rule(),

    def get_numbering_level_for_style(self):
        return self.get_numbering_definition().get_level_for_paragraph(self.id)

    def css_numbering_style_declaration(self):
        if self.numbering is not None:
            return self.get_numbering_level_for_style().css_style_declaration()

    def css_numbering_style_rule(self):
        return self.get_numbering_level_for_style().css_style_rule()

    def adjust_for_indent(self, css_style):
        """
        Certain properties must be set depending on whether there is a
        hanging indent, or a first line indent.
        """
        text_indent = css_style.getProperty('text-indent')
        text_indent_value = text_indent.propertyValue[0].value
        if text_indent_value < 0:
            # Negative text-indent is a hanging indent
            numbering_level = self.get_numbering_level_for_style()
            if numbering_level.suffix == 'tab':
                css_style['text-indent'] = ''
        else:
            # Positive text-indent is a first line indent
            css_style['text-indent'] = ''

    def pull_margin_left(self, css_style):
        """
        If there is a numbering property, it might necessary to grab the
        margin-left if none is currently set
        :param css_style:
        :return:
        """
        self.pull_property_from_numbering_style(css_style, 'margin-left')

    def pull_text_indent(self, css_style):
        """
        If there is a numbering property, it might be necessary to grab
        the text-indent if none is currently set
        :param css_style:
        :return:
        """
        self.pull_property_from_numbering_style(css_style, 'text-indent')

    def pull_property_from_numbering_style(self, css_style, prop_name):
        numbering_style = self.css_numbering_style_declaration()
        p_property = css_style.getProperty(prop_name)
        n_property = numbering_style.getProperty(prop_name)
        if p_property is None and n_property is not None:
            css_style[prop_name] = n_property.propertyValue.cssText


class NumberingStyle(DocxStyle):

    def css_style_rules(self):
        # TODO: Handle numbering styles
        return []


class TableStyle(DocxStyle):

    def css_style_rules(self):
        # TODO: Handle table styles
        return []


STYLE_MAPPING = {
    'character': CharacterStyle,
    'numbering': NumberingStyle,
    'paragraph': ParagraphStyle,
    'table': TableStyle,
}


@wordml('docDefaults')
class DocDefaults(etree.ElementBase):

    def css_properties(self):
        for d in self.iterdescendants():
            if isinstance(d, CssPropertyAdapter):
                yield d

    def css_style_declaration(self):
        css_style = cssutils.css.CSSStyleDeclaration()
        package = getattr(self, 'package', None)
        for prop in self.css_properties():
            prop.set_css_style(css_style, package)
        return css_style


class CssPropertyAdapter(etree.ElementBase, ABC):

    def get_boolean(self, attribute_name, default=True):
        """Return the value of an attribute as a boolean
        A default value can be supplied in case the attribute is non-existent.
        """
        attribute_value = self.get_string(attribute_name)
        if attribute_value:
            return not attribute_value.lower() in ('false', '0')
        return default

    def get_string(self, attribute_name, default=None):
        """Return the value of an attribute as a string
        If a default value is provided, it is returned if the attribute
        does not exist, or it doesn't have any value
        """
        return self.get(w(attribute_name)) or default

    @abstractmethod
    def set_css_style(self, style_rule, opc_package=None):
        pass


class ToggleProperty(CssPropertyAdapter):

    @property
    @abstractmethod
    def css_name(self):
        pass

    @property
    @abstractmethod
    def css_value(self):
        pass

    @property
    @abstractmethod
    def css_none_value(self):
        pass

    def get_value(self):
        return self.get_boolean('val')

    def set_css_style(self, style_rule, opc_package=None):
        style_rule[self.css_name] = self.css_value if self.get_value() else self.css_none_value


class ComplexToggleProperty(ToggleProperty, ABC):

    def set_css_style(self, style_rule, opc_package=None):
        for name, value, none_value in zip(self.css_name, self.css_value, self.css_none_value):
            existing = style_rule[name]
            # existing = style_rule.getPropertyValue(name)
            new_value = value if self.get_value() else none_value

            # Some docx elements are mutually exclusive, such as strike and
            # dstrike. In this case, both elements will be present and one
            # will be toggled on while the other will be toggled off. It is
            # important not to override the property that is toggled on.
            #
            # For example, if strike is parsed first, the CSS value will be
            # 'line-through'. When dstrike is parsed later on, it must not
            # set the value to 'none'
            #
            # Additionally, some properties must coexist such as strike
            # and u which are both represented as the CSS property
            # 'text-decoration-line'. If both strike and u are toggled on,
            # the CSS value needs to be 'line-through underline'
            if not existing or existing == none_value:
                style_rule[name] = new_value
            elif new_value != none_value and self.can_coexists_with(existing):
                style_rule[name] = ' '.join((existing, new_value))

    def can_coexists_with(self, existing_value):
        return False


class ColorPropertyAdapter(CssPropertyAdapter, ABC):

    color_attribute = 'val'
    theme_color_attribute = 'themeColor'
    theme_shade_attribute = 'themeShade'
    theme_tint_attribute = 'themeTint'

    def get_color(self):
        theme_color = self.get(w(self.theme_color_attribute))
        if theme_color is not None:
            color = CSSColor.from_string(self.theme.get_color(theme_color))
            shade = self.get(w(self.theme_shade_attribute))
            if shade is not None:
                color.apply_hsl_shade(shade)
            tint = self.get(w(self.theme_tint_attribute))
            if tint is not None:
                color.apply_rgb_tint(tint)
        else:
            color = self.get(w(self.color_attribute))
        return f'#{color}' if color is not None and color != 'auto' else ''

    def set_css_style(self, style_rule, opc_package=None):
        self.theme = opc_package.theme
        super().set_css_style(style_rule, opc_package)


@wordml('caps')
class AllCapsProperty(ToggleProperty):
    css_name = 'text-transform'
    css_value = 'uppercase'
    css_none_value = 'none'


@wordml('b')
class BoldProperty(ToggleProperty):
    css_name = 'font-weight'
    css_value = 'bold'
    css_none_value = 'normal'


@wordml('bdr')
class BorderProperty(ColorPropertyAdapter):
    color_attribute = 'color'
    direction = ''  # Direction (top, left, bottom, right) of the border

    def css_border_color(self):
        """Get the optional color atttribute of the border.
        Returns a #Hex code, or an empty string if the color is undefined
        or if its value is set to 'auto'
        """
        color = self.get_string('color')
        return f"#{color}" if color and color != 'auto' else ''

    def css_border_shadow(self):
        """Get the value of the box-shadow if the attribute shadow is set to
        true in the docx.
        The color is never defined because Word seem to make all shadows black
        """
        shadow = self.get_boolean('shadow', None)
        width = self.css_border_width()
        return f'{width} {width}' if width and shadow else ''

    def css_border_style(self):
        """Get the CSS value for this element.
        The 'val' attribute is required and will be always be present.
        The border type will be translated using an ST class
        """
        value = self.get_string('val')
        return ST_Border.css_value(value)

    def css_border_width(self):
        """Get the CSS border width. The 'sz' attribute is in 8th of a pt.
        """
        width = self.get_string('sz')
        return f'{int(width) / 8:.2f}pt' if width is not None else ''

    def css_padding(self):
        """
        Add the padding corresponding to space attribute
        """
        space = self.get(w('space'))
        if space is not None:
            padding = int(space)
            return f'{padding}pt' if padding else ''

    def set_css_style(self, style_rule, opc_package=None):
        # The following attributes are not supported:
        # frame
        super().set_css_style(style_rule, opc_package)
        direction = f'-{self.direction}' if self.direction else self.direction
        style_property_name = f'border{direction}-style'
        width_property_name = f'border{direction}-width'
        color_property_name = f'border{direction}-color'
        padding_property_name = f'padding{direction}'
        style = self.css_border_style()
        style_rule[style_property_name] = style
        if style != 'none':
            style_rule[width_property_name] = self.css_border_width()
            style_rule[color_property_name] = self.get_color()
            style_rule[padding_property_name] = self.css_padding()
            style_rule['box-shadow'] = self.css_border_shadow()


@wordml('bottom')
class BorderBottomProperty(BorderProperty):
    direction = 'bottom'


@wordml('left')
class BorderLeftProperty(BorderProperty):
    direction = 'left'


@wordml('right')
class BorderRightProperty(BorderProperty):
    direction = 'right'


@wordml('top')
class BorderTopProperty(BorderProperty):
    direction = 'top'


@wordml('dstrike')
class DStrikeProperty(ComplexToggleProperty):
    css_name = (
        'text-decoration-line',
        'text-decoration-style',
    )
    css_value = (
        'line-through',
        'double',
    )
    css_none_value = (
        'none',
        '',
    )

    def can_coexists_with(self, existing_value):
        return existing_value == 'underline'


@wordml('emboss')
class EmbossProperty(ComplexToggleProperty):
    css_name = ('text-shadow',)
    css_value = (
        '-1px -1px 0 rgba(255,255,255,0.3), 1px 1px 0 rgba(0,0,0,0.8)',
    )
    css_none_value = ('unset',)


@wordml('color')
class FontColorProperty(ColorPropertyAdapter):

    def set_css_style(self, style_rule, opc_package=None):
        super().set_css_style(style_rule, opc_package)
        style_rule['color'] = self.get_color()


@wordml('kern')
class FontKerningProperty(CssPropertyAdapter):

    def set_css_style(self, style_rule, opc_package=None):
        value = int(self.get(w('val'))) / 2  # Value is in half-points
        style_rule['font-kerning'] = 'normal' if value != 0 else 'auto'


@wordml('rFonts')
class FontProperty(CssPropertyAdapter):

    def _wrap(self, font_name):
        """Wrap a font name in quotes if it has a space in it"""
        if font_name is not None:
            return f'"{font_name}"' if ' ' in font_name else font_name

    def _get_theme_font_or_font_value(self, font_name, theme):
        """Get the theme font associated with the theme or return the
        same value if it's not a theme color
        """
        if font_name in theme.fonts:
            font_name = theme.get_font(font_name)
        return font_name

    def _get_font_from_font_table(self, font_name, font_table):
        font = font_table.get_font(font_name)
        if font is not None:
            return font.css_family
        return font_name,

    def set_css_style(self, style_rule, opc_package=None):
        theme = opc_package.theme
        font_table = opc_package.font_table
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
            font_name = self._get_theme_font_or_font_value(attribute, theme)
            if font_name:
                for f in self._get_font_from_font_table(font_name, font_table):
                    fonts[self._wrap(f)] = None
        # Push the generic family at the end. This happens when different
        # fonts are specified, and they are found in the font table
        for generic in ST_FontFamily.docx2css.values():
            if generic in fonts:
                value = fonts.pop(generic)
                fonts[generic] = value
        style_rule['font-family'] = ', '.join(fonts.keys())


@wordml('spacing')
class FontSpacingProperty(CssPropertyAdapter):
    """
    This element can represent the font spacing property, the line height,
    or the margins between paragraphs.
    """

    def css_letter_spacing(self):
        value = self.get(w('val'))
        if value is not None:
            value = int(value) / 20  # Value is in 20th of a point
            return f'{value}pt'

    def css_line_height(self):
        height = self.get(w('line'))
        rule = self.get(w('lineRule'))
        if height is not None:
            if rule in ('atLeast', 'exact'):
                # Height is in 20th of a point
                height = int(height) / 20
                return f'{height}pt'
            elif rule == 'auto':
                # Height is 240th of a line
                height = int(height) / 240
                return f'{height:.2f}'

    def css_margin_bottom(self):
        after = self.get(w('after'))
        auto = self.get_boolean('afterAutospacing', False)
        if not auto and after is not None:
            after = int(after) / 20  # Value is in 20th of a point
            return f'{after:.2f}pt'

    def css_margin_top(self):
        before = self.get(w('before'))
        auto = self.get_boolean('beforeAutospacing', False)
        if not auto and before is not None:
            before = int(before) / 20  # Value is in 20th of a point
            return f'{before:.2f}pt'

    def set_css_style(self, style_rule, opc_package=None):
        style_rule['letter-spacing'] = self.css_letter_spacing()
        style_rule['line-height'] = self.css_line_height()
        style_rule['margin-bottom'] = self.css_margin_bottom()
        style_rule['margin-top'] = self.css_margin_top()


@wordml('highlight')
class HighlightProperty(ComplexToggleProperty):
    css_name = (
        'background-color',
    )
    css_none_value = (
        'unset',
    )

    @property
    def css_value(self):
        value = self.get(w('val'))
        if value == 'none':
            value = self.css_none_value[0]
        return value.lower(),


@wordml('imprint')
class ImprintProperty(ComplexToggleProperty):
    css_name = ('text-shadow',)
    css_value = (
        '0 1px 0 rgba(255,255,255,0.3), 0 -1px 0 rgba(0,0,0,0.7)',
    )
    css_none_value = ('unset',)


@wordml('ind')
class IndentProperty(CssPropertyAdapter):

    def css_first_line_indent(self):
        first_line = self.get(w('firstLine'))
        if first_line is not None:
            first_line = int(first_line) / 20 / 72
            return f'{first_line:.2f}in'

    def css_hanging_indent(self):
        hanging = self.get(w('hanging'))
        if hanging is not None:
            hanging = int(hanging) / 20 / 72
            return f'{-1 * hanging:.2f}in'

    def css_margin_left(self):
        left = self.get(w('start')) or self.get(w('left'))
        if left is not None:
            # Value is in 20th of a point and converted in inches
            left = int(left) / 20 / 72
            return f'{left:.2f}in'

    def css_margin_right(self):
        right = self.get(w('end')) or self.get(w('right'))
        if right is not None:
            # Value is in 20th of a point and converted in inches
            right = int(right) / 20 / 72
            return f'{right:.2f}in'

    def set_css_style(self, style_rule, opc_package=None):
        style_rule['margin-left'] = self.css_margin_left()
        style_rule['margin-right'] = self.css_margin_right()
        text_indent = self.css_hanging_indent() or self.css_first_line_indent()
        style_rule['text-indent'] = text_indent


@wordml('i')
class ItalicProperty(ToggleProperty):
    css_name = 'font-style'
    css_value = 'italic'
    css_none_value = 'normal'


@wordml('jc')
class JustificationProperty(CssPropertyAdapter):

    def set_css_style(self, style_rule, opc_package=None):
        value = self.get(w('val'))
        style_rule['text-align'] = ST_Jc.css_value(value)


@wordml('keepLines')
class KeepLinesTogetherProperty(ToggleProperty):
    css_name = 'break-inside'
    css_value = 'avoid'
    css_none_value = 'unset'


@wordml('keepNext')
class KeepWithNextParagraphProperty(ToggleProperty):
    css_name = 'break-after'
    css_value = 'avoid'
    css_none_value = 'unset'


@wordml('numPr')
class NumberingProperty(etree.ElementBase):

    @property
    def definition(self):
        element = self.find(w('numId'))
        if element is not None:
            value = int(element.get(w('val')))
            numbering = self.styles.opc_package.get_numbering()
            return numbering.numbering_instances[value]

    @property
    def level(self):
        element = self.find(w('ilvl'))
        return int(element.get(w('val'))) if element is not None else None


@wordml('outline')
class OutlineProperty(ComplexToggleProperty):
    css_name = (
        '-webkit-text-stroke',
        '-webkit-text-fill-color',
    )
    css_value = (
        '1px',
        '#fff',
    )
    css_none_value = (
        'unset',
        'unset',
    )


@wordml('pageBreakBefore')
class PageBreakBeforeProperty(ToggleProperty):
    css_name = 'break-before'
    css_value = 'page'
    css_none_value = 'unset'


@wordml('position')
class PositionProperty(ComplexToggleProperty):
    css_name = (
        'vertical-align',
    )
    css_none_value = (
        'baseline',
    )

    @property
    def css_value(self):
        value = int(self.get(w('val'))) / 2  # Value in half-points
        return f'{value}pt',


@wordml('shd')
class ShadingProperty(ColorPropertyAdapter):
    css_name = 'background-color'
    css_none_value = 'unset'
    color_attribute = 'fill'
    theme_color_attribute = 'themeFill'
    theme_shade_attribute = 'themeFillShade'
    theme_tint_attribute = 'themeFillTint'

    def set_css_style(self, style_rule, opc_package=None):
        super().set_css_style(style_rule, opc_package)
        existing = style_rule[self.css_name]
        color = self.get_color() or self.css_none_value

        # It's important to check that the CSS property does not already have
        # a value. If there is one already, leave it alone, unless it's the
        # none value
        if not existing or existing == self.css_none_value:
            style_rule[self.css_name] = color


@wordml('shadow')
class ShadowProperty(ComplexToggleProperty):
    css_name = ('text-shadow',)
    css_value = ('1px 1px 2px',)
    css_none_value = ('unset',)


@wordml('smallCaps')
class SmallCapsProperty(ToggleProperty):
    css_name = 'font-variant-caps'
    css_value = 'small-caps'
    css_none_value = 'normal'


@wordml('strike')
class StrikeProperty(ComplexToggleProperty):
    css_name = (
        'text-decoration-line',
        'text-decoration-style',
    )
    css_value = (
        'line-through',
        'solid',
    )
    css_none_value = (
        'none',
        '',
    )

    def can_coexists_with(self, existing_value):
        return existing_value == 'underline'


@wordml('sz')
class SizeProperty(CssPropertyAdapter):

    def get_value(self):
        return int(self.get(w('val')))

    def set_css_style(self, style_rule, opc_package=None):
        style_rule['font-size'] = f"{self.get_value() / 2}pt"


@wordml('u')
class UnderlineProperty(ColorPropertyAdapter, ComplexToggleProperty):
    color_attribute = 'color'
    css_name = (
        'text-decoration-line',
        'text-decoration-style',
        'text-decoration-color',
    )
    css_none_value = (
        'none',
        '',
        '',
    )

    @property
    def css_value(self):
        value = self.get(w('val'))
        has_underline = value != self.css_none_value[0]
        color = self.get_color() if has_underline else self.css_none_value[1]
        style = ST_Underline.css_value(value) if has_underline else ''
        return (
            'underline' if value != self.css_none_value[0] else value,
            style,
            color,
        )

    def can_coexists_with(self, existing_value):
        return existing_value == 'line-through'


@wordml('vanish')
class VanishProperty(ToggleProperty):
    css_name = 'visibility'
    css_value = 'hidden'
    css_none_value = 'visible'


@wordml('vertAlign')
class VerticalAlignProperty(ComplexToggleProperty):
    css_name = (
        'vertical-align',
        'font-size',
    )
    css_none_value = (
        '',
        ''
    )

    @property
    def css_value(self):
        value = self.get(w('val'))
        align = 'baseline'
        if value == 'superscript':
            align = 'super'
        elif value == 'subscript':
            align = 'sub'
        return (
            align,
            'smaller' if align != 'baseline' else ''
        )


@wordml('widowControl')
class WidowControlProperty(ComplexToggleProperty):
    css_name = (
        'widows',
        'orphans',
    )
    # Widow control is usually on, so values are inverted
    css_value = (
        'unset',
        'unset',
    )
    css_none_value = (
        '0',
        '0',
    )
