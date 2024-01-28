from docx2css.ooxml import ct, NAMESPACES
from docx2css.ooxml.styles import (
    DocxStyle,
    PPrMixin,
    RPrMixin,
)


########################################################################
#                                                                      #
# Row Properties                                                       #
#                                                                      #
########################################################################

class TableRowPropertiesProxy:
    cant_split = ct.Boolean('w:trPr/w:cantSplit')
    """This element specifies whether the contents within the current cell shall 
    be rendered on a single page. When displaying the contents of a table cell 
    (such as the table cells in ECMA-376), it is possible that a page break 
    would fall within the contents of a table cell, causing the contents of that 
    cell to be displayed across two different pages. If this property is set, 
    then all contents of a table row shall be rendered on the same page by 
    moving the start of the current row to the start of a new page if necessary. 
    If the contents of this table row cannot fit on a single page, then this row 
    shall start on a new page and flow onto multiple pages as necessary.

    If this element is not present, the default behavior is dictated by the 
    setting in the associated table style. If this property is not specified in 
    the style hierarchy, then this table row shall be allowed to split across 
    multiple pages.
    """

    cell_spacing = ct.TableMeasure('w:trPr/w:tblCellSpacing')
    """This element specifies the default table cell spacing (the spacing 
    between adjacent cells and the edges of the table) for all cells in the 
    parent row. If specified, this element specifies the minimum amount of space 
    which shall be left between all cells in the table including the width of 
    the table borders in the calculation. It is important to note that row-level
    cell spacing shall be added inside of the text margins, which shall be 
    aligned with the innermost starting edge of the text extents in a cell 
    without row-level indentation or cell spacing. Row- level cell spacing 
    shall not increase the width of the overall table.
    """

    height = ct.TwipMeasure('w:trPr/w:trHeight')
    """This element specifies the height of the current table row within the 
    current table. This height shall be used to determine the resulting height 
    of the table row, which can be absolute or relative (depending on its 
    attribute values).

    If omitted, then the table row shall automatically resize its height to the 
    height required by its contents (the equivalent of an hRule value of auto).
    """

    height_type = ct.TableRowHeightType('w:trPr/w:trHeight')
    """Specifies the meaning of the height specified for this table row. 

    The meaning of the value of the val attribute is defined based on the value 
    of the hRule attribute for this table row as follows:

        * If the value of hRule is auto, then the table row's height should be 
        automatically determined based on the height of its contents. The 
        h value is ignored.

        * If the value of hRule is atLeast, then the table row's height should 
        be at least the value the h attribute.

        * If the value of hRule is exact, then the table row's height should be 
        exactly the value of the h attribute.

    If this attribute is omitted, then its value shall be assumed to be auto.
    """

    is_header = ct.Boolean('w:trPr/w:tblHeader')
    """This element specifies that the current table row shall be repeated at 
    the top of each new page on which part of this table is displayed. This 
    gives this table row the behavior of a 'header' row on each of these pages. 
    This element can be applied to any number of rows at the top of the table 
    structure in order to generate multi-row table headers.

    If this element is omitted, this table row shall not be repeated on each 
    new page on which the table is displayed. As well, if this row is not 
    contiguously connected with the first row of the table (that is, if this 
    table row is not either the first row, or all rows between this row and the 
    first row are not marked as header rows) then this property shall be 
    ignored.
    """

    justification = ct.Justification('w:trPr/w:jc')
    """This element specifies the alignment of a single row in the parent table 
    with respect to the text margins in the current section. When a table is 
    placed in a WordprocessingML document that does not have the same width as
    the margins, this property is used to determine how a specific row in that 
    table is positioned with respect to those margins. The interpretation of 
    property is reversed if the parent table is right to left using the 
    bidiVisual element (§17.4.1).

    If this property is omitted on a table, then the justification shall be 
    determined by the default set of table properties on the parent table.
    """

    def __get__(self, instance, owner):
        self.instance = instance
        return self

    def find(self, *args, **kwargs):
        return self.instance.find(*args, **kwargs)


########################################################################
#                                                                      #
# Cell Properties                                                      #
#                                                                      #
########################################################################

class TableCellPropertiesProxy:
    background_color = ct.Shading('w:tcPr/w:shd')
    """This element specifies the shading which shall be applied to the extents 
    of the current table cell. Similarly to paragraph shading, this shading 
    shall be applied to the contents of the cell up to the cell borders, 
    regardless of the presence of text.
    """

    border_bottom = ct.BorderDescriptor('w:tcPr/w:tcBorders/w:bottom')
    """This element specifies the border which shall be displayed at the bottom 
    of the current table cell.
    """

    border_inside_horizontal = ct.BorderDescriptor('w:tcPr/w:tcBorders/w:insideH')
    """This element specifies the border which shall be displayed on all 
    interior horizontal edges of the current group of table cells. 
    
    Note: Although individual table cells have no concept of an internal edge, 
    which would render this property useless in most cases, it is used to 
    determine the cell borders to apply to a specific group of cells as part of 
    table conditional formatting in a table style, for example, the inside 
    horizontal edges on the set of cells in the first column.
    """

    border_inside_vertical = ct.BorderDescriptor('w:tcPr/w:tcBorders/w:insideV')
    """This element specifies the border which shall be displayed on all 
    interior vertical edges of the current group of table cells. 
    
    Note: Although individual table cells have no concept of an internal edge, 
    which would render this property useless in most cases, it is used to 
    determine the cell borders to apply to a specific group of cells as part of 
    table conditional formatting in a table style, for example, the inside 
    vertical edges on the set of cells in the header row.
    """

    border_left = ct.BorderDescriptor('w:tcPr/w:tcBorders/w:left')
    """This element specifies the border which shall be displayed on the leading 
    edge of the current table cell (left for LTR tables, right for RTL tables).
    """

    border_right = ct.BorderDescriptor('w:tcPr/w:tcBorders/w:right')
    """This element specifies the border which shall be displayed on the 
    trailing edge of the current table cell (right for LTR tables, left for RTL 
    tables).
    """

    border_top = ct.BorderDescriptor('w:tcPr/w:tcBorders/w:top')
    """This element specifies the border which shall be displayed at the top of 
    the current table cell.
    """

    fit_text = ct.Boolean('w:tcPr/w:tcFitText')
    """This element specifies that the contents of the current cell shall have 
    their inter-character spacing increased or reduced as necessary to fit the 
    width of the text extents of the current cell. This setting shall behave 
    identically to placing the contents of this paragraph in a run and using the 
    fitText element (§17.3.2.14), if the width provided on that element matched 
    the width of the current cell.
    
    If this element is omitted, then the text in this cell shall not be fit to 
    the current cell extents.
    """

    grid_span = ct.Integer('w:tcPr/w:gridSpan')
    """This element specifies the number of grid columns in the parent table's 
    table grid which shall be spanned by the current cell. This property allows 
    cells to have the appearance of being merged, as they span vertical 
    boundaries of other cells in the table.
    
    If this element is omitted, then the number of grid units spanned by this 
    cell shall be assumed to be one. If the number of grid units specified by 
    the val attribute exceeds the size of the table grid, then the table grid 
    shall be augmented as needed to create the number of grid columns required.
    """

    margin_bottom = ct.TableMeasure('w:tcPr/w:tcMar/w:bottom')
    """This element specifies the amount of space which shall be left between 
    the bottom extent of the cell contents and the border of all table cells 
    within  the parent table (or table row).
    """

    margin_left = ct.TableMeasure('w:tcPr/w:tcMar/w:left')
    """This element specifies the amount of space which shall be left between 
    the leading edge of the cell contents and the leading edge of all table 
    cells within the parent table (or table row).
    """

    margin_right = ct.TableMeasure('w:tcPr/w:tcMar/w:right')
    """This element specifies the amount of space which shall be present between 
    the trailing extent of the cell contents and the trailing border of all 
    table cells within the parent table (or table row).
    """

    margin_top = ct.TableMeasure('w:tcPr/w:tcMar/w:top')
    """This element specifies the amount of space which shall be left between 
    the top extent of the cell contents and the top border of all table cells 
    within  the parent table.
    """

    no_wrap = ct.Boolean('w:tcPr/w:noWrap')
    """This element specifies how this table cell shall be laid out when the 
    parent table is displayed in a document. This setting only affects the 
    behavior of the cell when the tblLayout for this row (§17.4.53; §17.4.54) 
    is set to use the auto algorithm.
    
    This setting shall be interpreted in the context of the tcW element 
    (§17.4.72) as follows:
    
        * If the table cell width has a type attribute value of fixed, then this 
        element specifies that that this table cell shall never be smaller than 
        that fixed value when other cells on the line are not at their absolute
        minimum width.
        
        * If the table cell width has a type attribute value of pct or auto, 
        then this element specifies that when running the auto fit algorithm, 
        the contents of that this table cell shall be treated as though they 
        have no breaking characters (the contents should be treated as a single 
        contiguous non-breaking string)
    
    If this element is omitted, then cell content shall be allowed to wrap (the 
    cell can be shrunk as needed if it is a fixed preferred width value, and the 
    contents shall be treated as having breaking characters if it is a 
    percentage or automatic width value).
    """

    valign = ct.VerticalJustification('w:tcPr/w:vAlign')
    """This element specifies the vertical alignment for text within the current 
    table cell.
    """

    width = ct.TableMeasure('w:tcPr/w:tcW')
    """This element specifies the preferred width for this table cell.
    
    Any width value of type pct for this element shall be calculated relative to 
    the overall width of the table.
    
    If this element is omitted, then the cell width shall be of type auto.
    """

    def __get__(self, instance, owner):
        self.instance = instance
        return self

    def find(self, *args, **kwargs):
        return self.instance.find(*args, **kwargs)


########################################################################
#                                                                      #
# Table Properties                                                     #
#                                                                      #
########################################################################

class TablePropertiesMixin:
    background_color = ct.Shading('w:tblPr/w:shd')
    """This element specifies the shading which shall be applied to the extents 
    the current table. Similarly to paragraph  shading, this shading shall be 
    applied to the contents of the tab up to the table borders, regardless of 
    the presence of text - unlike cell shading, table shading shall include any 
    cell padding. This property shall be superseded by any cell-level shading 
    via any table-level property exceptions (§17.4.31); or on any cell in this 
    row (§17.4.33).
    """

    border_bottom = ct.BorderDescriptor('w:tblPr/w:tblBorders/w:bottom')
    """This element specifies the border which shall be displayed at the bottom 
    of the current table.
    """

    border_inside_horizontal = ct.BorderDescriptor('w:tblPr/w:tblBorders/w:insideH')
    """This element specifies the border which shall be displayed on all 
    horizontal table cell borders which are not on an outmost edge of the parent 
    table (all horizontal borders which are not the topmost or bottommost 
    border).
    """

    border_inside_vertical = ct.BorderDescriptor('w:tblPr/w:tblBorders/w:insideV')
    """This element specifies the border which shall be displayed on all 
    vertical table cell borders which are not on an outmost edge of the parent 
    table (all horizontal borders which are not the leftmost or rightmost 
    border).
    """

    border_left = ct.BorderDescriptor('w:tblPr/w:tblBorders/w:left')
    """This element specifies the border which shall be displayed at the leading 
    edge of the current table (left for LTR tables, right for RTL tables).
    """

    border_right = ct.BorderDescriptor('w:tblPr/w:tblBorders/w:right')
    """This element specifies the border which shall be displayed at the 
    trailing edge of the current table (right for LTR tables, left for RTL 
    tables)."""

    border_top = ct.BorderDescriptor('w:tblPr/w:tblBorders/w:top')
    """This element specifies the border which shall be displayed at the top of 
    the current table.
    """

    cell_margin_bottom = ct.TableMeasure('w:tblPr/w:tblCellMar/w:bottom')
    """This element specifies the amount of space which shall be left between 
    the bottom extent of the cell contents and the border of all table cells 
    within  the parent table (or table row). This setting can be overridden by 
    the table cell bottom margin definition specified by the bottom element 
    contained within the table cell's properties (§17.4.2). 

    If this element is omitted, then it shall inherit the table cell margin from 
    the associated table style. If a bottom margin is never specified in the 
    style hierarchy, then this table shall have no bottom cell padding by 
    default (excepting individual cell overrides).
    """

    cell_margin_left = ct.TableMeasure('w:tblPr/w:tblCellMar/w:left')
    """This element specifies the amount of space which shall be left between 
    the leading edge of the cell contents and the leading edge of all table
    cells within the parent table (or table row). This setting can be overridden
    by the table cell leading margin definition specified by the start element
    contained within the table cell's properties (§17.4.36).

    If this element is omitted, then it shall inherit the table cell margin
    from the associated table style. If a leading margin is never specified in
    the style hierarchy, this table shall have 115 twentieths of a point
    (0.08 inches) left cell padding by default (excepting individual cell
    overrides).
    """

    cell_margin_right = ct.TableMeasure('w:tblPr/w:tblCellMar/w:right')
    """This element specifies the amount of space which shall be present between 
    the trailing extent of the cell contents and the trailing border of all 
    table cells within the parent table (or table row) . This setting can be 
    overridden by the table cell trailing margin definition specified by the end 
    element contained within the table cell's properties (§17.4.10).

    If this element is omitted, then it shall inherit the table cell margin from 
    the associated table style. If a trailing margin is never specified in the 
    style hierarchy, this table shall have 115 twentieths of a point 
    (0.08 inches) left cell padding by default (excepting individual cell 
    overrides).
    """

    cell_margin_top = ct.TableMeasure('w:tblPr/w:tblCellMar/w:top')
    """This element specifies the amount of space which shall be left between 
    the top extent of the cell contents and the top border of all table cells 
    within  the parent table. This setting can be overridden by the table cell 
    top margin definition specified by the top element contained within the 
    table cell's properties (§17.4.78).

    If this element is omitted, then it shall inherit the table cell margin from 
    the associated table style. If a top margin is never specified in the 
    style hierarchy, then this table shall have no top cell padding by default
    (excepting individual cell overrides).
    """

    cell_spacing = ct.TableMeasure('w:tblPr/w:tblCellSpacing')
    """This element specifies the default table cell spacing (the spacing 
    between adjacent cells and the edges of the table) for all cells in the 
    parent table. If specified, this element specifies the minimum amount of 
    space which shall be left between all cells in the table including the 
    width of the table borders in the calculation. This setting shall be 
    superseded by a table-level exception (§17.4.45) or the row cell spacing 
    value (§17.4.44) in that order. It is important to note that table-level 
    cell spacing shall be added outside of the text margins, which shall be
    aligned with the innermost starting edge of the text extents in a table 
    cell.

    If this element is omitted, then the table shall inherit the table cell 
    spacing from the associated table style. If table cell spacing is never 
    specified in the style hierarchy, no cell spacing shall be added to the 
    parent table.
    """

    col_band_size = ct.Integer('w:tblPr/w:tblStyleColBandSize')
    """This element specifies the number of columns which shall comprise each a 
    table style column band for this table style. This element determines how 
    many columns constitute each of the column bands for the current table, 
    allowing column band formatting to be applied to groups of columns (rather 
    than just single alternating columns) when the table is formatted.

    If this element is omitted, then the default number of columns in a single 
    column band shall be assumed to be 1.
    """

    indent = ct.TableMeasure('w:tblPr/w:tblInd')
    """This element specifies the indentation which shall be added before the 
    leading edge of the current table in the document (the left edge in a 
    left-to-right table, and the right edge in a right-to-left table). This 
    indentation should shift the table into the text margin by the specified 
    amount.
    """

    justification = ct.String('w:tblPr/w:jc')
    """This element specifies the alignment of the current table with respect to
    the text margins in the current section. When a table is placed in a
    WordprocessingML document that does not have the same width as the margins,
    this property is used to determine how the table is positioned with respect
    to those margins. The interpretation of property is reversed if the parent
    table is right to left using the bidiVisual element (§17.4.1).

    If this property is omitted on a table, then the justification shall be 
    determined by the associated table style. If this property is not specified 
    in the style hierarchy, then the table shall be left justified with zero 
    indentation from the leading margin (the left margin in a left-to-right 
    table or the right margin in a right-to-left table).
    """

    layout = ct.TableLayout('w:tblPr/w:tblLayout')
    """This element specifies the algorithm which shall be used to lay out the 
    contents of this table within the document. When a table is displayed in a 
    document, it can either be displayed using a fixed width or autofit layout 
    algorithm (each discussed in the simple type referenced by the val 
    attribute). 

    If this element is omitted, then the value of this element shall be assumed 
    to be auto.
    """

    row_band_size = ct.Integer('w:tblPr/w:tblStyleRowBandSize')
    """This element specifies the number of rows which shall comprise each a 
    table style row band for this table style. This element determines how many 
    rows constitute each of the row bands for the current table, allowing row 
    band formatting to be applied to groups of rows (rather than just single 
    alternating rows) when the table is
    formatted.

    If this element is omitted, then the default number of rows in a single row 
    band shall be assumed to be 1.
    """

    width = ct.TableMeasure('w:tblPr/w:tblW')
    """This element specifies the preferred width for this table. This 
    preferred width is used as part of the table layout algorithm specified by 
    the tblLayout element (§17.4.53; §17.4.54) - full description of the 
    algorithm in the ST_TblLayout simple type (§17.18.87). 

    If this element is omitted, then the cell width shall be of type auto. 
    """


class PartialTableMixin(TablePropertiesMixin, PPrMixin, RPrMixin):
    row_properties = TableRowPropertiesProxy()
    cell_properties = TableCellPropertiesProxy()


########################################################################
#                                                                      #
# Conditional Table Style Formatting                                   #
#                                                                      #
########################################################################

class TableConditionalFormattingProxy(PartialTableMixin):

    def __set_name__(self, owner, name):
        path_mapping = {
            'odd_rows': 'band1Horz',
            'odd_columns': 'band1Vert',
            'even_rows': 'band2Horz',
            'even_columns': 'band2Vert',
            'first_column': 'firstCol',
            'first_row': 'firstRow',
            'last_column': 'lastCol',
            'last_row': 'lastRow',
            'top_right_cell': 'neCell',
            'top_left_cell': 'nwCell',
            'bottom_right_cell': 'seCell',
            'bottom_left_cell': 'swCell',
            'whole_table': 'wholeTable',
        }
        self.base_path = path_mapping[name]

    def __get__(self, instance, owner):
        xpath_expr = f'./w:tblStylePr[@w:type="{self.base_path}"]'
        xpath_results = instance.xpath(xpath_expr, namespaces=NAMESPACES)
        if len(xpath_results):
            self.instance = xpath_results[0]
            return self
        else:
            return None

    def find(self, *args, **kwargs):
        return self.instance.find(*args, **kwargs)


class DocxTableStyle(PartialTableMixin, DocxStyle):

    # conditional formatting:
    whole_table = TableConditionalFormattingProxy()
    odd_columns = TableConditionalFormattingProxy()
    even_columns = TableConditionalFormattingProxy()
    odd_rows = TableConditionalFormattingProxy()
    even_rows = TableConditionalFormattingProxy()
    first_row = TableConditionalFormattingProxy()
    last_row = TableConditionalFormattingProxy()
    first_column = TableConditionalFormattingProxy()
    last_column = TableConditionalFormattingProxy()
    top_left_cell = TableConditionalFormattingProxy()
    top_right_cell = TableConditionalFormattingProxy()
    bottom_left_cell = TableConditionalFormattingProxy()
    bottom_right_cell = TableConditionalFormattingProxy()
