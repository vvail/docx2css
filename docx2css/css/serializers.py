from abc import ABC, abstractmethod, ABCMeta
import re

import cssutils

from docx2css import api
from docx2css.stylesheet import Stylesheet
from docx2css.utils import (
    AutoLength,
    CssUnit,
    Percentage,
)


class CssPropertySerializer(ABC):

    def __init__(self, block_serializer: 'CssBlockSerializer', property_value):
        self.serializer = block_serializer
        self.property_value = property_value

    @abstractmethod
    def set_css_style(self, style_rule: cssutils.css.CSSStyleDeclaration):
        pass


class CssTablePropertySerializer(CssPropertySerializer, ABC):

    def __init__(self, block_serializer: 'CssTableSerializer', property_value):
        super().__init__(block_serializer, property_value)
        self.serializer = block_serializer


class CssSerializerFactory:

    def __init__(self):
        self.block_serializers = {}
        self.property_serializers = {}

    def register(self, prop_name, serializer_class):
        self.property_serializers[prop_name] = serializer_class

    def register_block_serializer(self, block_class, serializer_class):
        self.block_serializers[block_class] = serializer_class

    def get_block_serializer(self, block_value):
        block_class = block_value.__class__
        creator = self.block_serializers.get(block_class)
        if not creator:
            raise ValueError(f'No serializer registered for "{block_class}"')
        return creator(block_value, self)

    def get_property_serializer(self, style_serializer, prop_name, prop_value):
        creator = self.property_serializers.get(prop_name)
        if not creator:
            raise ValueError(f'No serializer registered for "{prop_name}"')
        return creator(style_serializer, prop_value)

    # def get_property_serializers(self,
    #                              style_serializer: 'CssBlockSerializer',
    #                              property_container: PropertyContainer):
    #     """Get the serializers associated with each of the properties of
    #     the property_container object. If none is provide, the default
    #     will be self.style
    #     """
    #     for prop in property_container.properties():
    #         yield self.get_property_serializer(style_serializer, prop)


########################################################################
#                                                                      #
# Text Formatting Serializers                                          #
#                                                                      #
########################################################################

class ToggleMixin(CssPropertySerializer, ABC):

    @property
    @abstractmethod
    def css_name(self):
        pass

    @property
    @abstractmethod
    def css_true(self):
        pass

    @property
    @abstractmethod
    def css_false(self):
        pass

    def set_css_style(self, style_rule):
        if self.property_value is True:
            value = self.css_true
        else:
            value = self.css_false
        style_rule[self.css_name] = value


class ComplexToggleMixin(ToggleMixin, ABC):

    def set_css_style(self, style_rule):
        css_name = self.css_name
        if isinstance(css_name, str):
            css_name = (css_name,)
        css_true = self.css_true
        if isinstance(css_true, str):
            css_true = (css_true,)
        css_false = self.css_false
        if isinstance(css_false, str):
            css_false = (css_false,)
        for name, value, none_value in zip(css_name, css_true, css_false):
            existing = style_rule[name]
            # existing = style_rule.getPropertyValue(name)
            new_value = value if self.property_value else none_value

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
        pass


class AllCapsSerializer(ToggleMixin):
    css_name = 'text-transform'
    css_true = 'uppercase'
    css_false = 'none'


class BackgroundColorSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        existing = style_rule['background-color']
        style_rule['background-color'] = self.property_value
        # It's important to check that the CSS property does not already have
        # a value. If there is one already, leave it alone, unless it's the
        # none value
        if not existing or existing == 'unset':
            style_rule['background-color'] = self.property_value or 'unset'


class BoldSerializer(ToggleMixin):
    css_name = 'font-weight'
    css_true = 'bold'
    css_false = 'normal'


class DoubleStrikeSerializer(ComplexToggleMixin):
    css_name = (
        'text-decoration-line',
        'text-decoration-style',
    )
    css_true = (
        'line-through',
        'double',
    )
    css_false = (
        'none',
        '',
    )

    def can_coexists_with(self, existing_value):
        return existing_value == 'underline'


class EmbossSerializer(ComplexToggleMixin):
    css_name = 'text-shadow',
    css_true = '-1px -1px 0 rgba(255,255,255,0.3), 1px 1px 0 rgba(0,0,0,0.8)'
    css_false = 'unset'


class FontColorSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['color'] = self.property_value


class FontFamilySerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['font-family'] = self.property_value


class FontKerningSerializer(ToggleMixin):
    css_name = 'font-kerning'
    css_true = 'normal'
    css_false = 'auto'


class FontSizeSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['font-size'] = f'{self.property_value.pt}pt'


class HighlightSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        value = self.property_value
        if value == 'none':
            value = 'unset'
        style_rule['background-color'] = value


class ImprintSerializer(ComplexToggleMixin):
    css_name = 'text-shadow'
    css_true = '0 1px 0 rgba(255,255,255,0.3), 0 -1px 0 rgba(0,0,0,0.7)'
    css_false = 'unset'


class ItalicsSerializer(ToggleMixin):
    css_name = 'font-style'
    css_true = 'italic'
    css_false = 'normal'


class LetterSpacingSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['letter-spacing'] = f'{self.property_value.pt}pt'


class OutlineSerializer(ComplexToggleMixin):
    css_name = (
        '-webkit-text-stroke',
        '-webkit-text-fill-color',
    )
    css_true = (
        '1px',
        '#fff',
    )
    css_false = (
        'unset',
        'unset',
    )


class PositionSerializer(ComplexToggleMixin):
    css_name = 'vertical-align'
    css_false = 'baseline'

    @property
    def css_true(self):
        return f'{self.property_value.pt}pt'


class ShadowSerializer(ComplexToggleMixin):
    css_name = 'text-shadow'
    css_true = '1px 1px 2px'
    css_false = 'unset'


class SmallCapsSerializer(ToggleMixin):
    css_name = 'font-variant-caps'
    css_true = 'small-caps'
    css_false = 'normal'


class StrikeSerializer(ComplexToggleMixin):
    css_name = (
        'text-decoration-line',
        'text-decoration-style',
    )
    css_true = (
        'line-through',
        'solid',
    )
    css_false = (
        'none',
        '',
    )

    def can_coexists_with(self, existing_value):
        return existing_value == 'underline'


class UnderlineSerializer(ComplexToggleMixin):
    css_name = (
        'text-decoration-line',
        'text-decoration-style',
        'text-decoration-color',
    )
    css_false = (
        'none',
        '',
        '',
    )

    @property
    def css_true(self):
        has_underline = self.property_value.has_line(api.TextDecoration.UNDERLINE)
        return (
            'underline' if has_underline else 'none',
            self.property_value.style if has_underline else '',
            self.property_value.color if has_underline else '',
        )

    def can_coexists_with(self, existing_value):
        return existing_value == 'line-through'


class VerticalAlignSerializer(ComplexToggleMixin):
    css_name = (
        'vertical-align',
        'font-size',
    )
    css_false = (
        '',
        ''
    )

    @property
    def css_true(self):
        value = self.property_value
        if value == 'superscript':
            css_true = (
                'super',
                'smaller'
            )
        elif value == 'subscript':
            css_true = (
                'sub',
                'smaller'
            )
        else:
            css_true = (
                'baseline',
                ''
            )
        return css_true


class VisibleSerializer(ToggleMixin):
    css_name = 'visibility'
    css_true = 'hidden'
    css_false = 'visible'


########################################################################
#                                                                      #
# Paragraph Formatting Serializers                                     #
#                                                                      #
########################################################################

class BreakAfterSerializer(ToggleMixin):
    css_name = 'break-after'
    css_true = 'avoid'
    css_false = 'unset'


class BreakBeforeSerializer(ToggleMixin):
    css_name = 'break-before'
    css_true = 'page'
    css_false = 'unset'


class BreakInsideSerializer(ToggleMixin):
    css_name = 'break-inside'
    css_true = 'avoid'
    css_false = 'unset'


class CounterSerializer(CssPropertySerializer):

    def css_counter(self, counter=None):
        if counter is None:
            counter = self.property_value
        if counter.style != 'none':
            return f'counter({counter.name}, {counter.style})'

    def css_counter_name(self):
        return self.property_value.name

    def css_counter_content(self):
        counter = self.property_value
        if counter.style == 'none':
            return
        elif counter.style == '':
            # When the content is a bullet, there should only be one
            # character in the level_text string, and it might not be
            # printable. Therefore, it is best to escape it
            return fr'"\005C {ord(counter.text):04x}"'
        contents = []
        tokens = re.split(r'({.*?})', counter.text)
        for token in (t for t in tokens if t):
            regex = re.match(r'{(.*?)}', token)
            if regex:
                c = counter.counter_list.counters[regex.group(1)]
                contents.append(self.css_counter(c))
            else:
                contents.append(f'"{token}"')
        if counter.suffix == 'space':
            contents.append(r'"\005C 00A0"')
        return ' '.join(filter(lambda x: x is not None, contents))

    def css_counter_resets(self):
        """
        Get a space-separated list of counters to reset at this level
        :return: String
        """
        return ' '.join(sorted(self.property_value.restart))

    def handle_margin_left(self, css_style_before, css_style):
        counter = self.property_value
        paragraph = self.serializer.style
        margins = (paragraph.margin_left, counter.margin_left)
        paragraph_margin_left = next((x for x in margins if x is not None), None)
        if paragraph_margin_left is not None:
            css_style['margin-left'] = f'{paragraph_margin_left.inches:.2f}in'

    def handle_text_indent(self, css_style_before, css_style):
        counter = self.property_value
        paragraph = self.serializer.style
        indents = (paragraph.text_indent, counter.text_indent)
        text_indent = next((x for x in indents if x is not None), None)
        if text_indent is None:
            return
        css_style_before['text-indent'] = f'{text_indent.inches:.2f}in'
        css_style_before['margin-left'] = ''
        if text_indent < 0:
            if counter.suffix == 'tab':
                css_style['text-indent'] = ''
                css_style_before['display'] = 'inline-block'
            else:
                css_style['text-indent'] = f'{text_indent.inches:.2f}in'
                css_style_before['text-indent'] = ''
        else:
            if counter.suffix == 'tab':
                css_style['text-indent'] = ''
                css_style_before['margin-right'] = f'{text_indent.inches:.2f}in'
                css_style_before['display'] = 'inline-block'
            else:
                css_style['text-indent'] = ''
                css_style_before['display'] = 'inline-block'

    def serialize_properties(self, css_style, properties):
        for k, v in properties:
            if k == 'counter':
                continue
            factory = self.serializer.factory
            serializer = factory.get_property_serializer(self.serializer, k, v)
            serializer.set_css_style(css_style)

    def set_css_style_before(self, css_style: cssutils.css.CSSStyleDeclaration):
        text_fields = self.property_value.text_properties(True)
        paragraph_style = self.serializer.style
        paragraph_fields = paragraph_style.paragraph_properties(True, False)
        self.serialize_properties(css_style, text_fields)
        self.serialize_properties(css_style, paragraph_fields)
        css_style['content'] = self.css_counter_content()
        css_style['counter-increment'] = self.css_counter_name()
        css_style['text-align'] = self.property_value.justification
        return css_style

    def set_css_style(self, style_rule: cssutils.css.CSSStyleDeclaration):
        before_selector = f'{self.serializer.css_current_selector()}:before'
        before_rule = self.serializer.get_or_create_rule(before_selector)
        self.set_css_style_before(before_rule)
        self.handle_margin_left(before_rule, style_rule)
        self.handle_text_indent(before_rule, style_rule)
        style_rule['counter-reset'] = self.css_counter_resets()


class LineHeightSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        height = self.property_value
        if isinstance(height, CssUnit):
            style_rule['line-height'] = f'{height.pt}pt'
        else:
            style_rule['line-height'] = f'{height:.2f}'


class MarginSerializer(CssPropertySerializer):
    direction = ''

    def value(self):
        return f'{self.property_value.inches:.2f}in'

    def set_css_style(self, style_rule):
        style_rule[f'margin-{self.direction}'] = self.value()


class MarginVerticalSerializer(MarginSerializer):

    def value(self):
        return f'{self.property_value.pt:.2f}pt'


class MarginBottomSerializer(MarginVerticalSerializer):
    direction = 'bottom'


class MarginLeftSerializer(MarginSerializer):
    direction = 'left'


class MarginRightSerializer(MarginSerializer):
    direction = 'right'


class MarginTopSerializer(MarginVerticalSerializer):
    direction = 'top'


class TextAlignSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['text-align'] = self.property_value


class TextIndentSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['text-indent'] = f'{self.property_value.inches:.2f}in'


class WidowsSerializer(ComplexToggleMixin):
    css_name = (
        'widows',
        'orphans',
    )
    # Widow control is usually on, so values are inverted
    css_true = (
        'unset',
        'unset',
    )
    css_false = (
        '0',
        '0',
    )


########################################################################
#                                                                      #
# Borders                                                              #
#                                                                      #
########################################################################

class BorderSerializer(CssPropertySerializer):
    direction = ''  # Direction (top, left, bottom, right) of the border

    def css_border_shadow(self):
        """Get the value of the box-shadow if the attribute shadow is set to
        true in the docx.
        The color is never defined because Word seem to make all shadows black
        """
        width = self.css_border_width()
        return f'{width} {width}' if width and self.property_value.shadow else ''

    def css_border_width(self):
        """Get the CSS border width. The 'sz' attribute is in 8th of a pt.
        """
        width = self.property_value.width
        return f'{width.pt:.2f}pt' if width is not None else ''

    def css_padding(self):
        """
        Add the padding corresponding to space attribute
        """
        padding = self.property_value.padding
        return f'{padding.pt}pt' if padding else ''

    def set_border_rule(self, style_rule: cssutils.css.CSSStyleDeclaration):
        direction = f'-{self.direction}' if self.direction else self.direction
        style_property_name = f'border{direction}-style'
        width_property_name = f'border{direction}-width'
        color_property_name = f'border{direction}-color'
        padding_property_name = f'padding{direction}'
        style = self.property_value.style
        style_rule[style_property_name] = style
        if style != 'none':
            style_rule[width_property_name] = self.css_border_width()
            style_rule[color_property_name] = self.property_value.color
            style_rule[padding_property_name] = self.css_padding()
            style_rule['box-shadow'] = self.css_border_shadow()

    def set_css_style(self, style_rule):
        # Always collapse the borders at the table level
        if isinstance(self.serializer, CssTableSerializer):
            table_selector = self.serializer.css_selector()
            table_css_rule = self.serializer.get_or_create_rule(table_selector)
            table_css_rule['border-collapse'] = 'collapse'
        self.set_border_rule(style_rule)


class BorderBottomSerializer(BorderSerializer):
    direction = 'bottom'


class BorderInsideHorizontalSerializer(BorderSerializer, CssTablePropertySerializer):
    inside_border_selector_suffix = 'td'
    direction = 'bottom'

    @property
    def inside_border_selector(self):
        attr_name = '_inside_border_selector'
        if not hasattr(self, attr_name):
            suffix = self.inside_border_selector_suffix
            setattr(self, attr_name, self.serializer.css_selector(suffix=suffix))
        return getattr(self, attr_name)

    @inside_border_selector.setter
    def inside_border_selector(self, selector):
        setattr(self, '_inside_border_selector', selector)

    def set_inside_border(self):
        selector = self.inside_border_selector
        css_style = self.serializer.get_or_create_rule(selector)
        self.set_border_rule(css_style)

    def set_css_style(self, style_rule):
        # Inside horizontal border can be set at the default cell level,
        # therefore, we need to get the table style rule in order to set
        # the border-collapse and border-style at the table level
        # instead of the cell
        table_selector = self.serializer.css_selector()
        table_rule = self.serializer.get_or_create_rule(table_selector)
        # Reorder the children of the style rule so that border-style:
        # hidden will be first
        current_children = list(table_rule.children())
        for child in current_children:
            table_rule.removeProperty(child.name)
        table_rule['border-collapse'] = 'collapse'
        table_rule['border-style'] = 'hidden'
        for child in current_children:
            table_rule.setProperty(child)
        self.set_inside_border()


class BorderInsideVerticalSerializer(BorderInsideHorizontalSerializer):
    direction = 'left'
    inside_border_selector_suffix = 'td + td'
    inside_border_last_selector_suffix = None


class BorderLeftSerializer(BorderSerializer):
    direction = 'left'


class BorderRightSerializer(BorderSerializer):
    direction = 'right'


class BorderTopSerializer(BorderSerializer):
    direction = 'top'


class RowHeightSerializer(CssPropertySerializer):
    
    def set_css_style(self, style_rule):
        style_rule['height'] = f'{self.property_value.inches:.2f}in'


########################################################################
#                                                                      #
# Table Serializers                                                    #
#                                                                      #
########################################################################

class RowIsHeaderSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        # This serializer doesn't do anything because table-header-group
        # needs to be applied to the parent element (thead or tbody) of
        # the row and not to the row itself
        # if self.property.value is True:
        #     style_rule['display'] = 'table-header-group'
        pass


class RowSplitSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        if self.property_value is True:
            style_rule['break-inside'] = 'avoid'
        else:
            style_rule['break-inside'] = 'auto'


class TableAlignmentSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        if self.property_value == 'center':
            style_rule['margin-left'] = 'auto'
            style_rule['margin-right'] = 'auto'
        elif self.property_value == 'end':
            style_rule['margin-left'] = 'auto'


class TableCellPaddingSerializer(CssTablePropertySerializer):
    direction = ''  # Direction (top, left, bottom, right) of the border

    def set_css_style(self, style_rule):
        selector = self.serializer.css_selector(suffix='td')
        style_rule = self.serializer.get_or_create_rule(selector)
        direction = f'-{self.direction}' if self.direction else self.direction
        style_rule[f'padding{direction}'] = f'{self.property_value.pt}pt'


class TableCellPaddingBottomSerializer(TableCellPaddingSerializer):
    direction = 'bottom'


class TableCellPaddingLeftSerializer(TableCellPaddingSerializer):
    direction = 'left'


class TableCellPaddingRightSerializer(TableCellPaddingSerializer):
    direction = 'right'


class TableCellPaddingTopSerializer(TableCellPaddingSerializer):
    direction = 'top'


class TableCellSpacingSerializer(CssPropertySerializer):
    
    def set_css_style(self, style_rule):
        value = self.property_value.pt
        if value == 0:
            style_rule['border-spacing'] = 'unset'
        else:
            style_rule['border-spacing'] = f'{value:.2f}pt'


class TableCellVerticalAlignSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        value = self.property_value
        if value == 'center':
            value = 'middle'
        style_rule['vertical-align'] = value


class TableCellWrapTextSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        value = 'normal' if self.property_value is True else 'nowrap'
        style_rule['white-space'] = value


class TableIndentSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['margin-left'] = f'{self.property_value.inches:.2f}in'


class TableLayoutSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['table-layout'] = self.property_value


class TableWidthSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        if isinstance(self.property_value, AutoLength):
            style_rule['width'] = 'auto'
        elif isinstance(self.property_value, Percentage):
            style_rule['width'] = f'{self.property_value.pct:.2f}%'
        else:
            style_rule['width'] = f'{self.property_value.inches:.2f}in'


class TableConditionalFormatting(CssTablePropertySerializer, ABC):

    @abstractmethod
    def cell_selector(self):
        pass

    @abstractmethod
    def row_selector(self):
        pass

    @abstractmethod
    def border_bottom_selector(self):
        pass

    @abstractmethod
    def border_left_selector(self):
        pass

    @abstractmethod
    def border_right_selector(self):
        pass

    @abstractmethod
    def border_top_selector(self):
        pass

    @abstractmethod
    def border_inside_horizontal_selector(self):
        pass

    @abstractmethod
    def border_inside_vertical_selector(self):
        pass

    def serialize_single_property(self, prop_tuple, css_rule):
        factory = self.serializer.factory
        handler = factory.get_property_serializer(self.serializer, *prop_tuple)
        handler.set_css_style(css_rule)

    def serialize_border_inside_horizontal(self, border_property):
        selector = self.border_inside_horizontal_selector()
        factory = self.serializer.factory
        border_tuple = 'border_inside_horizontal', border_property
        handler = factory.get_property_serializer(self.serializer, *border_tuple)
        handler.inside_border_selector = selector
        css_rule = self.serializer.get_or_create_rule(selector)
        handler.set_css_style(css_rule)

    def serialize_border_inside_vertical(self, border_property):
        selector = self.border_inside_vertical_selector()
        factory = self.serializer.factory
        border_tuple = 'border_inside_vertical', border_property
        handler = factory.get_property_serializer(self.serializer, *border_tuple)
        handler.inside_border_selector = selector
        css_rule = self.serializer.get_or_create_rule(selector)
        handler.set_css_style(css_rule)

    def set_css_style(self, style_rule):
        selector = self.cell_selector()
        table_properties = self.property_value.table_properties(True)
        self.serializer.serialize_properties(selector, table_properties)
        default_cell = self.property_value.default_cell
        default_cell_css_rule = self.serializer.get_or_create_rule(selector)
        default_row = self.property_value.default_row
        if default_row:
            row_properties = default_row.table_row_properties(True)
            self.serializer.serialize_properties(self.row_selector(), row_properties)
        if default_cell:
            cell_properties = default_cell.table_cell_properties(True)
            for k, v in cell_properties:
                if k == 'border_inside_horizontal':
                    self.serialize_border_inside_horizontal(v)
                elif k == 'border_inside_vertical':
                    self.serialize_border_inside_vertical(v)
                elif k.startswith('border'):
                    func = getattr(self, f'{k}_selector')
                    css_rule = self.serializer.get_or_create_rule(func())
                    self.serialize_single_property((k, v), css_rule)
                else:
                    self.serialize_single_property((k, v), default_cell_css_rule)


class OddRowsSerializer(TableConditionalFormatting):

    def border_bottom_selector(self):
        return ' '.join((self.serializer.last_odd_row_selector(), 'td'))

    def border_left_selector(self):
        return self.serializer.odd_row_selector(suffix=' td:first-of-type')

    def border_right_selector(self):
        return self.serializer.odd_row_selector(suffix=' td:last-of-type')

    def border_top_selector(self):
        return ' '.join((self.serializer.first_odd_row_selector(), 'td'))

    def border_inside_horizontal_selector(self):
        return self.serializer.row_inside_horizontal_selector()

    def border_inside_vertical_selector(self):
        return self.serializer.odd_row_selector(suffix=' td + td')

    def cell_selector(self):
        return self.serializer.odd_row_selector(suffix=' td')

    def row_selector(self):
        return self.serializer.odd_row_selector()


class EvenRowsSerializer(TableConditionalFormatting):

    def border_bottom_selector(self):
        return ' '.join((self.serializer.last_even_row_selector(), 'td'))

    def border_left_selector(self):
        return self.serializer.even_row_selector(suffix=' td:first-of-type')

    def border_right_selector(self):
        return self.serializer.even_row_selector(suffix=' td:last-of-type')

    def border_top_selector(self):
        return ' '.join((self.serializer.first_even_row_selector(), 'td'))

    def border_inside_horizontal_selector(self):
        return self.serializer.row_inside_horizontal_selector(odd=False)

    def border_inside_vertical_selector(self):
        return self.serializer.even_row_selector(suffix=' td + td')

    def cell_selector(self):
        return self.serializer.even_row_selector(suffix=' td')

    def row_selector(self):
        return self.serializer.even_row_selector()


class OddColumnsSerializer(TableConditionalFormatting):

    def cell_selector(self):
        return self.serializer.odd_column_selector()

    def row_selector(self):
        return None
    
    def border_bottom_selector(self):
        return self.serializer.odd_column_selector(row='tr:last-of-type')

    def border_left_selector(self):
        return self.serializer.first_odd_column_selector()

    def border_right_selector(self):
        return self.serializer.last_odd_column_selector()

    def border_top_selector(self):
        return self.serializer.odd_column_selector(row='tr:first-of-type')

    def border_inside_horizontal_selector(self):
        return self.cell_selector()

    def border_inside_vertical_selector(self):
        return self.serializer.column_inside_vertical_selector()


class EvenColumnsSerializer(OddColumnsSerializer):

    def cell_selector(self):
        return self.serializer.even_column_selector()

    def border_bottom_selector(self):
        return self.serializer.even_column_selector(row='tr:last-of-type')

    def border_left_selector(self):
        return self.serializer.first_even_column_selector()

    def border_right_selector(self):
        return self.serializer.last_even_column_selector()

    def border_top_selector(self):
        return self.serializer.even_column_selector(row='tr:first-of-type')

    def border_inside_vertical_selector(self):
        self.serializer.column_inside_vertical_selector(odd=False)


class SingleColumnMixin(TableConditionalFormatting, ABC):

    def border_left_selector(self):
        return self.cell_selector()

    def border_right_selector(self):
        return self.cell_selector()

    def border_inside_vertical_selector(self):
        return ''

    def serialize_border_inside_vertical(self, border_property):
        # Doesn't make sense to have inside vertical borders in a single
        # column
        pass


class FirstColumnSerializer(SingleColumnMixin):

    def cell_selector(self):
        return self.serializer.first_column_selector()

    def row_selector(self):
        return self.serializer.css_selector(suffix='tr:first-of-type')

    def border_bottom_selector(self):
        return self.serializer.bottom_left_cell_selector()

    def border_top_selector(self):
        return self.serializer.top_left_cell_selector()

    def border_inside_horizontal_selector(self):
        suffix = 'tr:not(:last-of-type) td:first-of-type'
        return self.serializer.css_selector(suffix=suffix)


class LastColumnSerializer(SingleColumnMixin):

    def cell_selector(self):
        return self.serializer.last_column_selector()

    def row_selector(self):
        return self.serializer.css_selector(suffix='tr:last-of-type')

    def border_bottom_selector(self):
        return self.serializer.bottom_right_cell_selector()

    def border_top_selector(self):
        return self.serializer.top_right_cell_selector()

    def border_inside_horizontal_selector(self):
        suffix = 'tr:not(:last-of-type) td:last-of-type'
        return self.serializer.css_selector(suffix=suffix)


class SingleRowMixin(TableConditionalFormatting, ABC):

    def border_bottom_selector(self):
        return self.cell_selector()

    def border_top_selector(self):
        return self.cell_selector()

    def border_inside_horizontal_selector(self):
        return ''

    def serialize_border_inside_horizontal(self, border_property):
        # Doesn't make sense to have an inside horizontal border in a
        # single row
        pass


class FirstRowSerializer(SingleRowMixin):

    def cell_selector(self):
        return self.serializer.css_selector(suffix='tr:first-of-type td')

    def row_selector(self):
        return self.serializer.first_row_selector()

    def border_left_selector(self):
        return self.serializer.top_left_cell_selector()

    def border_right_selector(self):
        return self.serializer.top_right_cell_selector()

    def border_inside_vertical_selector(self):
        return self.serializer.css_selector(suffix='tr:first-of-type td + td')


class LastRowSerializer(SingleRowMixin):

    def cell_selector(self):
        return self.serializer.css_selector(suffix='tr:last-of-type td')

    def row_selector(self):
        return self.serializer.last_row_selector()

    def border_left_selector(self):
        return self.serializer.bottom_left_cell_selector()

    def border_right_selector(self):
        return self.serializer.bottom_right_cell_selector()

    def border_inside_vertical_selector(self):
        return self.serializer.css_selector(suffix='tr:last-of-type td + td')


class SingleCellMixin(TableConditionalFormatting, ABC):

    def row_selector(self):
        return self.cell_selector()

    def border_bottom_selector(self):
        return self.cell_selector()

    def border_left_selector(self):
        return self.cell_selector()

    def border_right_selector(self):
        return self.cell_selector()

    def border_top_selector(self):
        return self.cell_selector()

    def border_inside_horizontal_selector(self):
        return ''

    def border_inside_vertical_selector(self):
        return ''

    def serialize_border_inside_horizontal(self, border_property):
        # Doesn't make sense to have inside borders in a single cell
        pass

    def serialize_border_inside_vertical(self, border_property):
        # Doesn't make sense to have inside borders in a single cell
        pass


class BottomLeftCellSerializer(SingleCellMixin):

    def cell_selector(self):
        return self.serializer.bottom_left_cell_selector()


class BottomRightCellSerializer(SingleCellMixin):

    def cell_selector(self):
        return self.serializer.bottom_right_cell_selector()


class TopLeftCellSerializer(SingleCellMixin):

    def cell_selector(self):
        return self.serializer.top_left_cell_selector()


class TopRightCellSerializer(SingleCellMixin):

    def cell_selector(self):
        return self.serializer.top_right_cell_selector()


class NoopSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        pass


FACTORY = CssSerializerFactory()
FACTORY.register('alignment', TableAlignmentSerializer)
FACTORY.register('all_caps', AllCapsSerializer)
FACTORY.register('background_color', BackgroundColorSerializer)
FACTORY.register('bold', BoldSerializer)
FACTORY.register('border', BorderSerializer)
FACTORY.register('border_bottom', BorderBottomSerializer)
FACTORY.register('border_inside_horizontal', BorderInsideHorizontalSerializer)
FACTORY.register('border_inside_vertical', BorderInsideVerticalSerializer)
FACTORY.register('border_left', BorderLeftSerializer)
FACTORY.register('border_right', BorderRightSerializer)
FACTORY.register('border_top', BorderTopSerializer)
FACTORY.register('bottom_left_cell', BottomLeftCellSerializer)
FACTORY.register('bottom_right_cell', BottomRightCellSerializer)
FACTORY.register('cell_padding_bottom', TableCellPaddingBottomSerializer)
FACTORY.register('cell_padding_left', TableCellPaddingLeftSerializer)
FACTORY.register('cell_padding_right', TableCellPaddingRightSerializer)
FACTORY.register('cell_padding_top', TableCellPaddingTopSerializer)
FACTORY.register('cell_spacing', TableCellSpacingSerializer)
FACTORY.register('col_band_size', NoopSerializer)
FACTORY.register('counter', CounterSerializer)
FACTORY.register('double_strike', DoubleStrikeSerializer)
FACTORY.register('emboss', EmbossSerializer)
FACTORY.register('even_columns', EvenColumnsSerializer)
FACTORY.register('even_rows', EvenRowsSerializer)
FACTORY.register('first_column', FirstColumnSerializer)
FACTORY.register('first_row', FirstRowSerializer)
FACTORY.register('font_color', FontColorSerializer)
FACTORY.register('font_family', FontFamilySerializer)
FACTORY.register('font_kerning', FontKerningSerializer)
FACTORY.register('font_size', FontSizeSerializer)
FACTORY.register('height', RowHeightSerializer)
FACTORY.register('highlight', HighlightSerializer)
FACTORY.register('imprint', ImprintSerializer)
FACTORY.register('indent', TableIndentSerializer)
FACTORY.register('is_header', RowIsHeaderSerializer)
FACTORY.register('italics', ItalicsSerializer)
FACTORY.register('keep_together', BreakInsideSerializer)
FACTORY.register('keep_with_next', BreakAfterSerializer)
FACTORY.register('last_column', LastColumnSerializer)
FACTORY.register('last_row', LastRowSerializer)
FACTORY.register('layout', TableLayoutSerializer)
FACTORY.register('letter_spacing', LetterSpacingSerializer)
FACTORY.register('line_height', LineHeightSerializer)
FACTORY.register('margin_bottom', MarginBottomSerializer)
FACTORY.register('margin_left', MarginLeftSerializer)
FACTORY.register('margin_right', MarginRightSerializer)
FACTORY.register('margin_top', MarginTopSerializer)
FACTORY.register('min_height', RowHeightSerializer)
FACTORY.register('odd_columns', OddColumnsSerializer)
FACTORY.register('odd_rows', OddRowsSerializer)
FACTORY.register('outline', OutlineSerializer)
FACTORY.register('padding_bottom', TableCellPaddingBottomSerializer)
FACTORY.register('padding_left', TableCellPaddingLeftSerializer)
FACTORY.register('padding_right', TableCellPaddingRightSerializer)
FACTORY.register('padding_top', TableCellPaddingTopSerializer)
FACTORY.register('page_break_before', BreakBeforeSerializer)
FACTORY.register('position', PositionSerializer)
FACTORY.register('row_band_size', NoopSerializer)
FACTORY.register('shadow', ShadowSerializer)
FACTORY.register('small_caps', SmallCapsSerializer)
FACTORY.register('split', RowSplitSerializer)
FACTORY.register('strike', StrikeSerializer)
FACTORY.register('text_align', TextAlignSerializer)
FACTORY.register('text_indent', TextIndentSerializer)
FACTORY.register('top_left_cell', TopLeftCellSerializer)
FACTORY.register('top_right_cell', TopRightCellSerializer)
FACTORY.register('underline', UnderlineSerializer)
FACTORY.register('valign', TableCellVerticalAlignSerializer)
FACTORY.register('vertical_align', VerticalAlignSerializer)
FACTORY.register('visible', VisibleSerializer)
FACTORY.register('widows_control', WidowsSerializer)
FACTORY.register('width', TableWidthSerializer)
FACTORY.register('wrap_text', TableCellWrapTextSerializer)
FACTORY.register('default_cell', NoopSerializer)
FACTORY.register('default_row', NoopSerializer)


class CssStylesheetSerializer:
    include_media_rules = True
    initialize_counters_in_body = True

    def __init__(self, stylesheet: Stylesheet, factory: CssSerializerFactory = None):
        self.stylesheet = stylesheet
        self.factory = factory if factory else FACTORY
        self._css_stylesheet = None

    @property
    def css_stylesheet(self):
        self._css_stylesheet = cssutils.css.CSSStyleSheet()
        self._serialize_css()
        return self._css_stylesheet

    def serialize(self):
        return self.css_stylesheet.cssText.decode('utf-8')

    def _add_rules(self, rule):
        """Add a set of rules to the CSSStylesheet"""
        for r in rule:
            self._css_stylesheet.add(r)

    def _serialize_css(self):
        if self.include_media_rules:
            self.serialize_page_style()
        body_style = self.stylesheet.body_style
        root_counters = self.css_root_counters()
        if root_counters:
            body_style.counter = api.Counter(restart=root_counters, text='')
        serializer = self.factory.get_block_serializer(body_style)
        self._add_rules(serializer.css_style_rules())
        from itertools import chain
        all_styles = chain(
            self.stylesheet.span_styles.values(),
            self.stylesheet.paragraph_styles.values(),
            self.stylesheet.table_styles.values()
        )
        for style in all_styles:
            serializer = self.factory.get_block_serializer(style)
            if serializer is not None:
                self._add_rules(serializer.css_style_rules())

    def css_root_counters(self):
        """Return a sorted set of all counters if
        self.initialize_all_counters_in_body is True, otherwise, the set
        will consist only of those counters that aren't restarted.
        """
        all_counters = set()
        restarted_counters = set()
        for style in self.stylesheet.paragraph_styles.values():
            if hasattr(style, 'counter') and style.counter is not None:
                counter = style.counter
                counter_name = counter.name
                if counter.start != 1:
                    counter_name += f' {counter.start - 1}'
                all_counters.add(counter_name)
                restarted_counters.update(counter.restart)
        if self.initialize_counters_in_body:
            return all_counters
        else:
            return all_counters - restarted_counters

    def serialize_page_style(self):
        page_style = self.stylesheet.page_style
        serializer = self.factory.get_block_serializer(page_style)
        self._add_rules(serializer.css_style_rules())


class CssBlockSerializer(ABC, metaclass=ABCMeta):

    def __init__(self, style: api.BaseStyle, factory: CssSerializerFactory):
        self.__style_rules = {}
        self.factory = factory
        self.style = style

    @property
    @abstractmethod
    def css_selector_prefix(self):
        pass

    @abstractmethod
    def _serialize(self):
        pass

    def get_or_create_rule(self, selector):
        rule = self.__style_rules.get(selector)
        if rule is None:
            rule = cssutils.css.CSSStyleDeclaration()
            self.__style_rules[selector] = rule
        return rule

    def serialize_properties(self, selector: str, properties):
        """Serialize all properties known to match a selector for the
        container used as argument
        """
        css_style = self.get_or_create_rule(selector)
        for prop_tuple in properties:
            serializer = self.factory.get_property_serializer(self, *prop_tuple)
            serializer.set_css_style(css_style)
        # for prop in self.factory.get_property_serializers(self, container):
        #     prop.set_css_style(css_style)

    def css_current_selector(self):
        """Get the selector for the current style only"""
        return f"{self.css_selector_prefix}.{self.style.id}"

    def css_selector(self, suffix=''):
        """Get the CSS selector for this style, including all the
        children with an optional suffix appended
        """
        names = [' '.join((self.css_current_selector(), suffix))]

        for child in self.get_style_children():
            block_serializer = self.factory.get_block_serializer(child)
            names.append(block_serializer.css_selector(suffix=suffix))
        return ', '.join(names)

    def get_style_children(self):
        return self.style.children

    def css_style_rules(self):
        self._serialize()
        return (cssutils.css.CSSStyleRule(k, style=v)
                for k, v in self.__style_rules.items())


class CssPageSerializer(CssBlockSerializer):
    css_selector_prefix = ''

    def __init__(self, style: api.PageStyle, factory: CssSerializerFactory):
        super().__init__(style, factory)
        self.style = style

    def _css_margin_value(self):
        top = self.style.margin_top.inches
        right = self.style.margin_right.inches
        bottom = self.style.margin_bottom.inches
        left = self.style.margin_left.inches
        return f'{top}in {right}in {bottom}in {left}in'

    def css_style_declaration_print(self):
        css_style = cssutils.css.CSSStyleDeclaration()
        css_style['size'] = (f'{self.style.page_width.inches}in '
                             f'{self.style.page_height.inches}in')
        css_style['margin'] = self._css_margin_value()
        return css_style

    def css_style_declaration_screen(self):
        css_style = cssutils.css.CSSStyleDeclaration()
        # Adjust max-width to margins
        max_width = self.style.page_width - self.style.margin_left - self.style.margin_right
        css_style['max-width'] = f'{CssUnit(max_width).inches}in'
        css_style['margin'] = '1em auto'
        css_style['padding'] = self._css_margin_value()
        return css_style

    def css_style_rule_screen(self):
        screen = cssutils.css.CSSMediaRule('screen')
        body_style = self.css_style_declaration_screen()
        body_rule = cssutils.css.CSSStyleRule('body', body_style)
        screen.add(body_rule)
        return screen

    def _serialize(self):
        pass

    def css_style_rules(self):
        css_style = self.css_style_declaration_print()
        return (
            cssutils.css.CSSPageRule(style=css_style),
            self.css_style_rule_screen()
        )


class CssBodySerializer(CssBlockSerializer):
    css_selector_prefix = 'body'

    def __init__(self, style: api.BodyStyle, factory: CssSerializerFactory):
        super().__init__(style, factory)
        self.style = style

    def css_selector(self, suffix=''):
        return self.css_selector_prefix

    def css_current_selector(self):
        return self.css_selector_prefix

    def _serialize(self):
        body_properties = self.style.paragraph_properties(True)
        self.serialize_properties(self.css_selector(), body_properties)


class CssSpanSerializer(CssBlockSerializer):
    css_selector_prefix = 'span'

    def __init__(self, style: api.SpanStyle, factory: CssSerializerFactory):
        super().__init__(style, factory)
        self.style = style

    def _serialize(self):
        span_properties = self.style.text_properties(True)
        self.serialize_properties(self.css_selector(), span_properties)


class CssParagraphSerializer(CssBlockSerializer):

    def __init__(self, style: api.ParagraphStyle, factory: CssSerializerFactory):
        super().__init__(style, factory)
        self.style = style

    def css_current_selector(self):
        class_name = f'{self.style.id}'
        prefix = self.css_selector_prefix
        if class_name == '' or re.match('h[1-6]', prefix):
            return prefix
        else:
            return f"{self.css_selector_prefix}.{class_name}"

    def css_selector(self, suffix=''):
        current_selector = self.css_current_selector()
        names = [current_selector]

        def not_p(s):
            ser = self.factory.get_block_serializer(s)
            return s.parent_id == '' and not ser.css_selector_prefix == 'p'

        # Treat Normal style a bit differently
        if current_selector == 'p':
            children = filter(not_p, self.style.children)
        else:
            children = self.style.children

        for child in children:
            serializer = self.factory.get_block_serializer(child)
            names.append(serializer.css_selector())
        return ', '.join(names)

    @property
    def css_selector_prefix(self):
        class_name = ''.join(self.style.name.split())
        regex = re.match('heading([1-6])', class_name)
        if regex:
            return f'h{regex.group(1)}'
        return 'p'

    def _serialize(self):
        self.get_or_create_rule(f'{self.css_current_selector()}:before')
        # self.serialize_properties(self.css_selector(), self.style)
        css_style = self.get_or_create_rule(self.css_selector())
        counter = None
        for k, v in self.style.paragraph_properties(active=True):
            if k == 'counter':
                counter = v
                continue
            serializer = self.factory.get_property_serializer(self, k, v)
            serializer.set_css_style(css_style)
        if counter:
            serializer = self.factory.get_property_serializer(self, 'counter', counter)
            serializer.set_css_style(css_style)


class CssTableSerializer(CssBlockSerializer):
    css_selector_prefix = 'table'

    def __init__(self, style: api.TableStyle, factory: CssSerializerFactory):
        super().__init__(style, factory)
        self.style = style

    def css_current_selector(self):
        """Get the selector for the current style only"""
        return self.style.qualified_id

    def get_style_children(self):
        if self.css_current_selector() == 'table':
            return []
        else:
            return self.style.children

    def default_cell_css_selector(self):
        """Get the CSS selector for the default cell"""
        return self.css_selector(suffix='td')

    def _col_row_selectors(self, column=True, odd=True, row='tr', suffix=''):
        if column:
            n = self.style.col_band_size or 1
            element = f'{row} td'
        else:
            n = self.style.row_band_size or 1
            element = f'{row}'
        bands = 2 * n
        suffixes = []
        start = 1 if odd else n + 1
        stop = n + 1 if odd else bands + 1
        for x in range(start, stop):
            current = ''.join((f'{element}:nth-child({bands}n+{x})', suffix))
            suffixes.append(self.css_selector(suffix=current))
        return suffixes

    def column_inside_vertical_selector(self, odd=True):
        n = self.style.col_band_size or 1
        bands = 2 * n
        suffixes = []
        start = 1 if odd else n + 1
        stop = n if odd else bands
        for x in range(start, stop):
            td1 = f'tr td:nth-child({bands}n+{x})'
            td2 = f'td:nth-child({bands}n+{x + 1})'
            current = ' + '.join((td1, td2))
            suffixes.append(self.css_selector(suffix=current))
        return ', '.join(suffixes)

    def col_row_selector(self, column=True, odd=True, row='tr', suffix=''):
        suffixes = self._col_row_selectors(column, odd, row, suffix)
        return ', '.join(suffixes)

    def odd_row_selector(self, suffix=''):
        """Get the CSS selector for odd rows"""
        return self.col_row_selector(column=False, suffix=suffix)

    def even_row_selector(self, suffix=''):
        """Get the CSS selector for even rows"""
        return self.col_row_selector(column=False, odd=False, suffix=suffix)

    def odd_column_selector(self, row='tr', suffix=''):
        """Get the CSS selector for odd columns"""
        return self.col_row_selector(row=row, suffix=suffix)

    def first_odd_column_selector(self):
        return self._col_row_selectors()[0]

    def last_odd_column_selector(self):
        return self._col_row_selectors()[-1]

    def even_column_selector(self, row='tr', suffix=''):
        """Get the CSS selector for even columns"""
        return self.col_row_selector(odd=False, row=row, suffix=suffix)

    def first_even_column_selector(self):
        return self._col_row_selectors(odd=False)[0]

    def last_even_column_selector(self):
        return self._col_row_selectors(odd=False)[-1]

    def first_odd_row_selector(self):
        return self._col_row_selectors(column=False)[0]

    def last_odd_row_selector(self):
        return self._col_row_selectors(column=False)[-1]

    def first_even_row_selector(self):
        return self._col_row_selectors(column=False, odd=False)[0]

    def last_even_row_selector(self):
        return self._col_row_selectors(column=False, odd=False)[-1]

    def row_inside_horizontal_selector(self, odd=True):
        suffixes = self._col_row_selectors(column=False, odd=odd, suffix=' td')
        return ', '.join(suffixes[:-1])

    def top_left_cell_selector(self):
        return self.css_selector(suffix='tr:first-of-type td:first-of-type')

    def top_right_cell_selector(self):
        return self.css_selector(suffix='tr:first-of-type td:last-of-type')

    def bottom_left_cell_selector(self):
        return self.css_selector(suffix='tr:last-of-type td:first-of-type')

    def bottom_right_cell_selector(self):
        return self.css_selector(suffix='tr:last-of-type td:last-of-type')

    def first_row_selector(self):
        return self.css_selector(suffix='tr:first-of-type')

    def last_row_selector(self):
        return self.css_selector(suffix='tr:last-child')

    def first_column_selector(self):
        return self.css_selector(suffix='tr td:first-of-type')
    
    def last_column_selector(self):
        return self.css_selector(suffix='tr td:last-of-type')

    def _serialize(self):
        # Table properties
        table_props = self.style.table_properties()
        self.serialize_properties(self.css_selector(), table_props)

        # Conditional formatting
        conditional_props = self.style.table_conditional_formatting_properties()
        self.serialize_properties(self.css_selector(), conditional_props)

        # Default row
        selector = self.css_selector(suffix='tr')
        default_row = self.style.default_row
        if default_row is not None:
            row_properties = default_row.table_row_properties(True)
            self.serialize_properties(selector, row_properties)

        # Default cell
        selector = self.default_cell_css_selector()
        default_cell = self.style.default_cell
        if default_cell is not None:
            cell_properties = default_cell.table_cell_properties(True)
            self.serialize_properties(selector, cell_properties)


FACTORY.register_block_serializer(api.PageStyle, CssPageSerializer)
FACTORY.register_block_serializer(api.BodyStyle, CssBodySerializer)
FACTORY.register_block_serializer(api.SpanStyle, CssSpanSerializer)
FACTORY.register_block_serializer(api.ParagraphStyle, CssParagraphSerializer)
FACTORY.register_block_serializer(api.TableStyle, CssTableSerializer)
