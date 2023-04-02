from abc import ABC
import logging

from docx2css import api
from docx2css.ooxml import w, wordml
from docx2css.ooxml.styles import BorderProperty, DocxPropertyAdapter
from docx2css.utils import AutoLength, CssUnit, Percentage


logger = logging.getLogger(__name__)


def _get_toggle_property(xml_element, name):
    """Parse a toggle (boolean) property from a child 'name' of the xml
    element.

    :returns: boolean value or None if the child 'name' does not exist
    """
    prop = xml_element.find(w(name))
    if prop is None:
        return None
    attribute_value = prop.get(w('val'))
    if attribute_value:
        return not attribute_value.lower() in ('false', '0')
    else:
        return True


def _get_measure(xml_element):
    """Parse a TblWidth as a Measure"""
    if xml_element is None:
        return None
    unit = xml_element.get(w('type'))
    value = xml_element.get(w('w'))
    if unit == 'auto':
        return AutoLength()
    elif unit == 'dxa':
        return CssUnit(value, 'twip')
    elif unit == 'nil':
        return CssUnit(0)
    elif unit == 'pct':
        return Percentage(int(value)/50)
    else:
        raise ValueError(f'Unit "{unit}" is invalid!')


########################################################################
#                                                                      #
# Table Properties                                                     #
#                                                                      #
########################################################################

class TableCellMargin(DocxPropertyAdapter, ABC):

    @property
    def prop_value(self):
        return _get_measure(self)


class TableCellMarginBottom(TableCellMargin):
    prop_name = 'padding_bottom'


class TableCellMarginLeft(TableCellMargin):
    prop_name = 'padding_left'


class TableCellMarginRight(TableCellMargin):
    prop_name = 'padding_right'


class TableCellMarginTop(TableCellMargin):
    prop_name = 'padding_top'


@wordml('tblCellSpacing')
class TableCellSpacing(DocxPropertyAdapter):
    prop_name = 'cell_spacing'

    @property
    def prop_value(self):
        return _get_measure(self)


@wordml('tblStyleColBandSize')
class TableColBandSize(DocxPropertyAdapter):
    prop_name = 'col_band_size'

    @property
    def prop_value(self):
        return int(self.get(w('val')))


@wordml('tblLayout')
class TableLayout(DocxPropertyAdapter):
    prop_name = 'layout'

    @property
    def prop_value(self):
        layout = self.get(w('type'))
        return layout if layout == 'fixed' else 'auto'


@wordml('tblInd')
class TableIndent(DocxPropertyAdapter):
    prop_name = 'indent'

    @property
    def prop_value(self):
        return _get_measure(self)


@wordml('tblStyleRowBandSize')
class TableRowBandSize(DocxPropertyAdapter):
    prop_name = 'row_band_size'

    @property
    def prop_value(self):
        return int(self.get(w('val')))


@wordml('tblW')
class TableWidth(DocxPropertyAdapter):

    prop_name = 'width'

    @property
    def prop_value(self):
        return _get_measure(self)


########################################################################
#                                                                      #
# Row Properties                                                       #
#                                                                      #
########################################################################

@wordml('trPr')
class TableRowProperties(DocxPropertyAdapter):
    prop_name = 'default_row'

    @property
    def prop_value(self):
        prop = api.TableRowProperties()
        _parse_descendants(self, prop)
        return prop


@wordml('cantSplit')
class TableRowCantSplit(DocxPropertyAdapter):
    prop_name = 'split'

    @property
    def prop_value(self):
        return not _get_toggle_property(self.getparent(), 'cantSplit')


@wordml('tblHeader')
class TableRowHeader(DocxPropertyAdapter):
    prop_name = 'is_header'

    @property
    def prop_value(self):
        return _get_toggle_property(self.getparent(), 'tblHeader')


@wordml('trHeight')
class TableRowHeight(DocxPropertyAdapter):

    @property
    def prop_name(self):
        # Word UI sometimes provides a trHeight element without an hRule
        # attribute and this is supposed to mean 'atLeast'
        rule = self.get(w('hRule')) or 'atLeast'
        if rule == 'exact':
            return 'height'
        elif rule == 'atLeast':
            return 'min_height'

    @property
    def prop_value(self):
        value = self.get(w('val'))
        return CssUnit(value, 'twip')


########################################################################
#                                                                      #
# Cell Properties                                                      #
#                                                                      #
########################################################################

@wordml('tcPr')
class TableCellProperties(DocxPropertyAdapter):
    prop_name = 'default_cell'

    @property
    def prop_value(self):
        prop = api.TableCellProperties()
        _parse_descendants(self, prop)
        return prop


@wordml('gridSpan')
class TableCellColspan(DocxPropertyAdapter):
    prop_name = 'colspan'

    @property
    def prop_value(self):
        return int(self.get(w('val')))


@wordml('tcFitText')
class TableCellFitText(DocxPropertyAdapter):
    prop_name = 'fit_text'

    @property
    def prop_value(self):
        return _get_toggle_property(self.getparent(), 'tcFitText')


@wordml('vAlign')
class TableCellVerticalAlign(DocxPropertyAdapter):
    prop_name = 'valign'

    @property
    def prop_value(self):
        return self.get(w('val'))


@wordml('tcW')
class TableCellWidth(DocxPropertyAdapter):
    prop_name = 'width'

    @property
    def prop_value(self):
        return _get_measure(self)


@wordml('noWrap')
class TableCellWrapText(DocxPropertyAdapter):
    prop_name = 'wrap_text'

    @property
    def prop_value(self):
        return not _get_toggle_property(self.getparent(), 'noWrap')


########################################################################
#                                                                      #
# Conditional Table Style Formatting                                   #
#                                                                      #
########################################################################

@wordml('tblStylePr')
class TableConditionalFormatting(DocxPropertyAdapter):

    @property
    def prop_name(self):
        type_mapping = {
            'band1Horz': 'odd_rows',
            'band1Vert': 'odd_columns',
            'band2Horz': 'even_rows',
            'band2Vert': 'even_columns',
            'firstCol': 'first_column',
            'firstRow': 'first_row',
            'lastCol': 'last_column',
            'lastRow': 'last_row',
            'neCell': 'top_right_cell',
            'nwCell': 'top_left_cell',
            'seCell': 'bottom_right_cell',
            'swCell': 'bottom_left_cell',
            'wholeTable': 'whole_table',
        }
        return type_mapping.get(self.get(w('type')))

    @property
    def prop_value(self):
        prop = api.TableConditionalFormatting()
        # _parse_descendants(self, prop)
        _parse_partial_table(self, prop)
        return prop


@wordml('insideH')
class BorderInsideHorizontalProperty(BorderProperty):
    prop_name = 'border_inside_horizontal'
    direction = 'inside-horizontal'


@wordml('insideV')
class BorderInsideVerticalProperty(BorderProperty):
    prop_name = 'border_inside_vertical'
    direction = 'inside-vertical'


def _parse_property_adapter(element, style):
    if isinstance(element, DocxPropertyAdapter):
        prop_name = element.prop_name
        # Handle the case where a descendant of tblPr can be either
        # a table border or a default cell margin (padding) and
        if element.getparent().tag == w('tblCellMar'):
            prop_name = f'cell_{prop_name}'
        logger.debug(f'   Found adapter {type(element)}')
        prop_value = element.prop_value
        logger.debug(f'{8 * " "}{prop_name} = {prop_value}')
        setattr(style, prop_name, prop_value)


def _parse_descendants(xml_element, style):
    if xml_element is None:
        return
    for d in xml_element.iterdescendants():
        _parse_property_adapter(d, style)


def _parse_partial_table(xml_element, style):
    """Parse a table style or a table conditional formatting element"""
    run_properties = xml_element.find(w('rPr'))
    _parse_descendants(run_properties, style)

    # TODO: Paragraph props

    table_properties = xml_element.find(w('tblPr'))
    _parse_descendants(table_properties, style)

    row_properties = xml_element.find(w('trPr'))
    _parse_property_adapter(row_properties, style)

    cell_properties = xml_element.find(w('tcPr'))
    _parse_property_adapter(cell_properties, style)


def parse_docx_table_style(xml_element):
    if xml_element.type != 'table':
        return
    logger.debug(f'Parsing table style "{xml_element.name}"')
    from docx2css.api import TableStyle

    style = TableStyle(name=xml_element.name, id=xml_element.id)

    _parse_partial_table(xml_element, style)

    conditional_formats = xml_element.findall(w('tblStylePr'))
    for conditional_format in conditional_formats:
        _parse_property_adapter(conditional_format, style)

    return style
