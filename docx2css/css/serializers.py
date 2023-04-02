from abc import ABC, abstractmethod

import cssutils

from docx2css.api import Stylesheet, TableStyle
from docx2css.utils import (
    AutoLength,
    KeyValueProperty,
    Percentage,
    PropertyContainer,
)


class CssPropertySerializer(ABC):

    def __init__(self, style: 'CssTableSerializer', prop: KeyValueProperty):
        self.style = style
        self.property = prop

    @abstractmethod
    def set_css_style(self, style_rule: cssutils.css.CSSStyleRule):
        pass


class CssPropertySerializerFactory:

    def __init__(self):
        self.creators = {}

    def register(self, prop_name, serializer_class):
        self.creators[prop_name] = serializer_class

    def get_serializer(self, style_serializer, prop):
        creator = self.creators.get(prop.name)
        if not creator:
            raise ValueError(f'No serializer registered for "{prop.name}"')
        return creator(style_serializer, prop)

    def get_serializers(self,
                        style_serializer: 'CssTableSerializer',
                        property_container: PropertyContainer):
        """Get the serializers associated with each of the properties of
        the property_container object. If none is provide, the default
        will be self.style
        """
        for prop in property_container.properties():
            yield self.get_serializer(style_serializer, prop)


########################################################################
#                                                                      #
# Text Formatting Serializers                                          #
#                                                                      #
########################################################################

class BoldSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['font-weight'] = 'bold' if self.property.value else 'normal'


class ItalicsSerializer(CssPropertySerializer):
    
    def set_css_style(self, style_rule):
        style_rule['font-style'] = 'italic' if self.property.value else 'normal'


class BackgroundColorSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        existing = style_rule['background-color']
        style_rule['background-color'] = self.property.value
        # It's important to check that the CSS property does not already have
        # a value. If there is one already, leave it alone, unless it's the
        # none value
        if not existing or existing == 'unset':
            style_rule['background-color'] = self.property.value


class BorderSerializer(CssPropertySerializer):
    direction = ''  # Direction (top, left, bottom, right) of the border

    def css_border_shadow(self):
        """Get the value of the box-shadow if the attribute shadow is set to
        true in the docx.
        The color is never defined because Word seem to make all shadows black
        """
        width = self.css_border_width()
        return f'{width} {width}' if width and self.property.value.shadow else ''

    def css_border_width(self):
        """Get the CSS border width. The 'sz' attribute is in 8th of a pt.
        """
        width = self.property.value.width
        return f'{width.pt:.2f}pt' if width is not None else ''

    def css_padding(self):
        """
        Add the padding corresponding to space attribute
        """
        padding = self.property.value.padding
        return f'{padding.pt}pt' if padding else ''

    def set_border_rule(self, style_rule: cssutils.css.CSSStyleDeclaration):
        direction = f'-{self.direction}' if self.direction else self.direction
        style_property_name = f'border{direction}-style'
        width_property_name = f'border{direction}-width'
        color_property_name = f'border{direction}-color'
        padding_property_name = f'padding{direction}'
        style = self.property.value.style
        style_rule[style_property_name] = style
        if style != 'none':
            style_rule[width_property_name] = self.css_border_width()
            style_rule[color_property_name] = self.property.value.color
            style_rule[padding_property_name] = self.css_padding()
            style_rule['box-shadow'] = self.css_border_shadow()

    def set_css_style(self, style_rule):
        # Always collapse the borders at the table level
        table_selector = self.style.css_selector()
        table_css_rule = self.style.get_or_create_rule(table_selector)
        table_css_rule['border-collapse'] = 'collapse'
        self.set_border_rule(style_rule)


class BorderBottomSerializer(BorderSerializer):
    direction = 'bottom'


class BorderInsideHorizontalSerializer(BorderSerializer):
    inside_border_selector_suffix = 'td'
    direction = 'bottom'

    @property
    def inside_border_selector(self):
        attr_name = '_inside_border_selector'
        if not hasattr(self, attr_name):
            suffix = self.inside_border_selector_suffix
            setattr(self, attr_name, self.style.css_selector(suffix=suffix))
        return getattr(self, attr_name)

    @inside_border_selector.setter
    def inside_border_selector(self, selector):
        setattr(self, '_inside_border_selector', selector)

    def set_inside_border(self):
        selector = self.inside_border_selector
        css_style = self.style.get_or_create_rule(selector)
        self.set_border_rule(css_style)

    def set_css_style(self, style_rule):
        # Inside horizontal border can be set at the default cell level,
        # therefore, we need to get the table style rule in order to set
        # the border-collapse and border-style at the table level
        # instead of the cell
        table_selector = self.style.css_selector()
        table_rule = self.style.get_or_create_rule(table_selector)
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


class BorderLeftSerializer(BorderSerializer):
    direction = 'left'


class BorderRightSerializer(BorderSerializer):
    direction = 'right'


class BorderTopSerializer(BorderSerializer):
    direction = 'top'


class RowHeightSerializer(CssPropertySerializer):
    
    def set_css_style(self, style_rule):
        style_rule['height'] = f'{self.property.value.inches:.2f}in'


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
        if self.property.value is True:
            style_rule['break-inside'] = 'avoid'
        else:
            style_rule['break-inside'] = 'auto'


class TableAlignmentSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        if self.property.value == 'center':
            style_rule['margin-left'] = 'auto'
            style_rule['margin-right'] = 'auto'
        elif self.property.value == 'end':
            style_rule['margin-left'] = 'auto'


class TableCellPaddingSerializer(CssPropertySerializer):
    direction = ''  # Direction (top, left, bottom, right) of the border

    def set_css_style(self, style_rule):
        selector = self.style.css_selector(suffix='td')
        style_rule = self.style.get_or_create_rule(selector)
        direction = f'-{self.direction}' if self.direction else self.direction
        style_rule[f'padding{direction}'] = f'{self.property.value.pt}pt'


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
        value = self.property.value.pt
        if value == 0:
            style_rule['border-spacing'] = 'unset'
        else:
            style_rule['border-spacing'] = f'{value:.2f}pt'


class TableCellVerticalAlignSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        value = self.property.value
        if value == 'center':
            value = 'middle'
        style_rule['vertical-align'] = value


class TableCellWrapTextSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        value = 'normal' if self.property.value is True else 'nowrap'
        style_rule['white-space'] = value


class TableIndentSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['margin-left'] = f'{self.property.value.inches:.2f}in'


class TableLayoutSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        style_rule['table-layout'] = self.property.value


class TableWidthSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        if isinstance(self.property.value, AutoLength):
            style_rule['width'] = 'auto'
        elif isinstance(self.property.value, Percentage):
            style_rule['width'] = f'{self.property.value.pct:.2f}%'
        else:
            style_rule['width'] = f'{self.property.value.inches:.2f}in'


class TableConditionalFormatting(CssPropertySerializer, ABC):

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

    def serialize_single_property(self, prop, css_rule):
        handler = factory.get_serializer(self.style, prop)
        handler.set_css_style(css_rule)

    def serialize_border_inside_horizontal(self, border_property):
        selector = self.border_inside_horizontal_selector()
        handler = factory.get_serializer(self.style, border_property)
        handler.inside_border_selector = selector
        css_rule = self.style.get_or_create_rule(selector)
        handler.set_css_style(css_rule)

    def serialize_border_inside_vertical(self, border_property):
        selector = self.border_inside_vertical_selector()
        handler = factory.get_serializer(self.style, border_property)
        handler.inside_border_selector = selector
        css_rule = self.style.get_or_create_rule(selector)
        handler.set_css_style(css_rule)

    def set_css_style(self, style_rule):
        selector = self.cell_selector()
        self.style.serialize_properties(selector, self.property.value)
        default_cell = self.property.value.default_cell
        default_cell_css_rule = self.style.get_or_create_rule(selector)
        default_row = self.property.value.default_row
        if default_row:
            self.style.serialize_properties(self.row_selector(), default_row)
        if default_cell:
            for prop in default_cell.properties():
                if prop.name == 'border_inside_horizontal':
                    self.serialize_border_inside_horizontal(prop)
                elif prop.name == 'border_inside_vertical':
                    self.serialize_border_inside_vertical(prop)
                elif prop.name.startswith('border'):
                    func = getattr(self, f'{prop.name}_selector')
                    css_rule = self.style.get_or_create_rule(func())
                    self.serialize_single_property(prop, css_rule)
                else:
                    self.serialize_single_property(prop, default_cell_css_rule)


class OddRowsSerializer(TableConditionalFormatting):

    def border_bottom_selector(self):
        return ' '.join((self.style.last_odd_row_selector(), 'td'))

    def border_left_selector(self):
        return self.style.odd_row_selector(suffix=' td:first-of-type')

    def border_right_selector(self):
        return self.style.odd_row_selector(suffix=' td:last-of-type')

    def border_top_selector(self):
        return ' '.join((self.style.first_odd_row_selector(), 'td'))

    def border_inside_horizontal_selector(self):
        return self.style.row_inside_horizontal_selector()

    def border_inside_vertical_selector(self):
        return self.style.odd_row_selector(suffix=' td + td')

    def cell_selector(self):
        return self.style.odd_row_selector(suffix=' td')

    def row_selector(self):
        return self.style.odd_row_selector()


class EvenRowsSerializer(TableConditionalFormatting):

    def border_bottom_selector(self):
        return ' '.join((self.style.last_even_row_selector(), 'td'))

    def border_left_selector(self):
        return self.style.even_row_selector(suffix=' td:first-of-type')

    def border_right_selector(self):
        return self.style.even_row_selector(suffix=' td:last-of-type')

    def border_top_selector(self):
        return ' '.join((self.style.first_even_row_selector(), 'td'))

    def border_inside_horizontal_selector(self):
        return self.style.row_inside_horizontal_selector(odd=False)

    def border_inside_vertical_selector(self):
        return self.style.even_row_selector(suffix=' td + td')

    def cell_selector(self):
        return self.style.even_row_selector(suffix=' td')

    def row_selector(self):
        return self.style.even_row_selector()


class OddColumnsSerializer(TableConditionalFormatting):

    def cell_selector(self):
        return self.style.odd_column_selector()

    def row_selector(self):
        return None
    
    def border_bottom_selector(self):
        return self.style.odd_column_selector(row='tr:last-of-type')

    def border_left_selector(self):
        return self.style.first_odd_column_selector()

    def border_right_selector(self):
        return self.style.last_odd_column_selector()

    def border_top_selector(self):
        return self.style.odd_column_selector(row='tr:first-of-type')

    def border_inside_horizontal_selector(self):
        return self.cell_selector()

    def border_inside_vertical_selector(self):
        return self.style.column_inside_vertical_selector()


class EvenColumnsSerializer(OddColumnsSerializer):

    def cell_selector(self):
        return self.style.even_column_selector()

    def border_bottom_selector(self):
        return self.style.even_column_selector(row='tr:last-of-type')

    def border_left_selector(self):
        return self.style.first_even_column_selector()

    def border_right_selector(self):
        return self.style.last_even_column_selector()

    def border_top_selector(self):
        return self.style.even_column_selector(row='tr:first-of-type')

    def border_inside_vertical_selector(self):
        self.style.column_inside_vertical_selector(odd=False)


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
        return self.style.first_column_selector()

    def row_selector(self):
        return self.style.css_selector(suffix='tr:first-of-type')

    def border_bottom_selector(self):
        return self.style.bottom_left_cell_selector()

    def border_top_selector(self):
        return self.style.top_left_cell_selector()

    def border_inside_horizontal_selector(self):
        suffix = 'tr:not(:last-of-type) td:first-of-type'
        return self.style.css_selector(suffix=suffix)


class LastColumnSerializer(SingleColumnMixin):

    def cell_selector(self):
        return self.style.last_column_selector()

    def row_selector(self):
        return self.style.css_selector(suffix='tr:last-of-type')

    def border_bottom_selector(self):
        return self.style.bottom_right_cell_selector()

    def border_top_selector(self):
        return self.style.top_right_cell_selector()

    def border_inside_horizontal_selector(self):
        suffix = 'tr:not(:last-of-type) td:last-of-type'
        return self.style.css_selector(suffix=suffix)


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
        return self.style.css_selector(suffix='tr:first-of-type td')

    def row_selector(self):
        return self.style.first_row_selector()

    def border_left_selector(self):
        return self.style.top_left_cell_selector()

    def border_right_selector(self):
        return self.style.top_right_cell_selector()

    def border_inside_vertical_selector(self):
        return self.style.css_selector(suffix='tr:first-of-type td + td')


class LastRowSerializer(SingleRowMixin):

    def cell_selector(self):
        return self.style.css_selector(suffix='tr:last-of-type td')

    def row_selector(self):
        return self.style.last_row_selector()

    def border_left_selector(self):
        return self.style.bottom_left_cell_selector()

    def border_right_selector(self):
        return self.style.bottom_right_cell_selector()

    def border_inside_vertical_selector(self):
        return self.style.css_selector(suffix='tr:last-of-type td + td')


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
        return self.style.bottom_left_cell_selector()


class BottomRightCellSerializer(SingleCellMixin):

    def cell_selector(self):
        return self.style.bottom_right_cell_selector()


class TopLeftCellSerializer(SingleCellMixin):

    def cell_selector(self):
        return self.style.top_left_cell_selector()


class TopRightCellSerializer(SingleCellMixin):

    def cell_selector(self):
        return self.style.top_right_cell_selector()


class NoopSerializer(CssPropertySerializer):

    def set_css_style(self, style_rule):
        pass


factory = CssPropertySerializerFactory()
factory.register('alignment', TableAlignmentSerializer)
factory.register('background_color', BackgroundColorSerializer)
factory.register('bold', BoldSerializer)
# factory.register('border', NoopSerializer)
factory.register('border_bottom', BorderBottomSerializer)
factory.register('border_inside_horizontal', BorderInsideHorizontalSerializer)
factory.register('border_inside_vertical', BorderInsideVerticalSerializer)
factory.register('border_left', BorderLeftSerializer)
factory.register('border_right', BorderRightSerializer)
factory.register('border_top', BorderTopSerializer)
factory.register('bottom_left_cell', BottomLeftCellSerializer)
factory.register('bottom_right_cell', BottomRightCellSerializer)
factory.register('cell_padding_bottom', TableCellPaddingBottomSerializer)
factory.register('cell_padding_left', TableCellPaddingLeftSerializer)
factory.register('cell_padding_right', TableCellPaddingRightSerializer)
factory.register('cell_padding_top', TableCellPaddingTopSerializer)
factory.register('cell_spacing', TableCellSpacingSerializer)
factory.register('col_band_size', NoopSerializer)
factory.register('even_columns', EvenColumnsSerializer)
factory.register('even_rows', EvenRowsSerializer)
factory.register('first_column', FirstColumnSerializer)
factory.register('first_row', FirstRowSerializer)
factory.register('height', RowHeightSerializer)
factory.register('indent', TableIndentSerializer)
factory.register('is_header', RowIsHeaderSerializer)
factory.register('italics', ItalicsSerializer)
factory.register('last_column', LastColumnSerializer)
factory.register('last_row', LastRowSerializer)
factory.register('layout', TableLayoutSerializer)
factory.register('min_height', RowHeightSerializer)
factory.register('odd_columns', OddColumnsSerializer)
factory.register('odd_rows', OddRowsSerializer)
factory.register('padding_bottom', TableCellPaddingBottomSerializer)
factory.register('padding_left', TableCellPaddingLeftSerializer)
factory.register('padding_right', TableCellPaddingRightSerializer)
factory.register('padding_top', TableCellPaddingTopSerializer)
factory.register('row_band_size', NoopSerializer)
factory.register('split', RowSplitSerializer)
factory.register('top_left_cell', TopLeftCellSerializer)
factory.register('top_right_cell', TopRightCellSerializer)
factory.register('valign', TableCellVerticalAlignSerializer)
factory.register('width', TableWidthSerializer)
factory.register('wrap_text', TableCellWrapTextSerializer)
factory.register('default_cell', NoopSerializer)
factory.register('default_row', NoopSerializer)


class CssStylesheetSerializer:

    def __init__(self, stylesheet: Stylesheet):
        self.stylesheet = stylesheet
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
        for style in self.stylesheet.styles.values():
            if style.type == 'table':
                serializer = CssTableSerializer(style)
                self._add_rules(serializer.css_style_rules())


class CssTableSerializer:
    css_selector_prefix = 'table'

    def __init__(self, style: TableStyle):
        self.style = style
        self.__style_rules = {}

    def get_or_create_rule(self, selector):
        rule = self.__style_rules.get(selector)
        if rule is None:
            rule = cssutils.css.CSSStyleDeclaration()
            self.__style_rules[selector] = rule
        return rule

    def serialize_properties(self, selector: str, container: PropertyContainer):
        """Serialize all properties known to match a selector for the
        container used as argument
        """
        css_style = self.get_or_create_rule(selector)
        for prop in factory.get_serializers(self, container):
            prop.set_css_style(css_style)

    def css_current_selector(self):
        """Get the selector for the current style only"""
        class_name = f'.{self.style.id}'
        if class_name == '.TableNormal':
            class_name = ''
        return ''.join((self.css_selector_prefix, class_name))

    def css_selector(self, style=None, suffix=''):
        """Get the CSS selector for this style, including all the
        children with an optional suffix appended
        """
        style = self.style if style is None else style
        names = [' '.join((self.css_current_selector(), suffix))]

        for child in style.children:
            names.append(self.css_selector(child))
        return ', '.join(names)

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

    def css_style_rules(self):
        # Table properties
        self.serialize_properties(self.css_selector(), self.style)

        # Default row
        selector = self.css_selector(suffix='tr')
        default_row = self.style.default_row
        if default_row is not None:
            self.serialize_properties(selector, self.style.default_row)

        # Default cell
        selector = self.default_cell_css_selector()
        default_cell = self.style.default_cell
        if default_cell is not None:
            self.serialize_properties(selector, self.style.default_cell)

        return (cssutils.css.CSSStyleRule(k, style=v)
                for k, v in self.__style_rules.items())
