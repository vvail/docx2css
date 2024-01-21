from abc import ABC

from docx2css import api
from docx2css.ooxml import w, wordml
from docx2css.ooxml.parsers import DocxParser
from docx2css.ooxml.styles import BorderProperty, DocxPropertyAdapter
from docx2css.utils import CssUnit


########################################################################
#                                                                      #
# Table Properties                                                     #
#                                                                      #
########################################################################

class TableCellMargin(DocxPropertyAdapter, ABC):

    @property
    def prop_value(self):
        return self.get_measure()


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
        return self.get_measure()


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
        return self.get_measure()


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
        return self.get_measure()


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
        self.docx_parser.parse_descendants(self, prop)
        return prop


@wordml('cantSplit')
class TableRowCantSplit(DocxPropertyAdapter):
    prop_name = 'split'

    @property
    def prop_value(self):
        return not self.get_toggle_property('cantSplit')


@wordml('tblHeader')
class TableRowHeader(DocxPropertyAdapter):
    prop_name = 'is_header'

    @property
    def prop_value(self):
        return self.get_toggle_property('tblHeader')


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
        self.docx_parser.parse_descendants(self, prop)
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
        return self.get_toggle_property('tcFitText')


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
        return self.get_measure()


@wordml('noWrap')
class TableCellWrapText(DocxPropertyAdapter):
    prop_name = 'wrap_text'

    @property
    def prop_value(self):
        return not self.get_toggle_property('noWrap')


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
        parser = DocxParser('')
        parser.parse_partial_table(self, prop)
        return prop


@wordml('insideH')
class BorderInsideHorizontalProperty(BorderProperty):
    prop_name = 'border_inside_horizontal'
    direction = 'inside-horizontal'


@wordml('insideV')
class BorderInsideVerticalProperty(BorderProperty):
    prop_name = 'border_inside_vertical'
    direction = 'inside-vertical'
