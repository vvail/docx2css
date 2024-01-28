from abc import ABC, abstractmethod
from dataclasses import dataclass, field, fields
from typing import Optional

from docx2css.utils import CssUnit


@dataclass
class Border:
    color: Optional[str] = None
    """Get the optional color of the border taking in consideration
    that the color can be defined with different attributes such as
    'color', 'themeColor', 'themeTint' or 'themeShade'
    Returns a #Hex code, or None if the color is undefined or if its
    value is set to 'auto'
    """

    padding: Optional[CssUnit] = None
    """Get the padding that shall be used to place this border on
    the parent object
    """

    shadow: Optional[bool] = None
    """Specifies whether this border should be modified to create
    the appearance of a shadow."""

    style: Optional[str] = None
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

    width: Optional[CssUnit] = None
    """Get the width of the border"""


@dataclass
class TextDecoration:
    UNDERLINE = 1
    LINE_THROUGH = 2
    line: int = 0
    color: Optional[str] = None
    style: Optional[str] = None

    def add_line(self, line_type):
        self.line ^= line_type

    def del_line(self, line_type):
        self.line &= ~line_type

    def has_line(self, line_type):
        return self.line & line_type == line_type


@dataclass
class TextFormatting:
    all_caps: Optional[bool] = None
    """Specifies that any lowercase characters in this text run shall be
    formatted for display only as their capital letter character 
    equivalents. This property does not affect any non-alphabetic 
    character in this run, and does not change the Unicode character for
    lowercase text, only the method in which it is displayed.
    """

    background_color: Optional[str] = None
    """Specifies the background color of the table.

    Return a hexadecimal color string, eg. '#FF00FF' or None
    """

    bold: Optional[bool] = None
    """Specifies whether the bold property shall be applied to all non-
    complex script characters
    """

    border: Optional[Border] = None
    """Specifies information about the border applied to the text in the
    current span.
    """

    double_strike: Optional[bool] = None
    """Specifies that the contents shall be displayed with two 
    horizontal lines through each character displayed on the line.
    
    This element shall not be present with the strike property, since 
    they are mutually exclusive in terms of appearance.
    """

    emboss: Optional[bool] = None
    """Specifies that the contents should be displayed as if embossed, 
    which makes text appear as if it is raised off the page in relief.
    
    This element shall not be present with either the imprint or outline
    properties , since they are mutually exclusive in terms of 
    appearance.
    """

    font_color: Optional[str] = None
    """Specifies the color of the font
    
    Returns a hexadecimal color string, eg. '#FF00FF' or None
    """

    font_family: Optional[str] = None
    """Get a comma-separated list of font faces"""

    font_kerning: Optional[bool] = None
    """Specifies whether font kerning shall be applied"""

    font_size: Optional[CssUnit] = None
    """Specifies the font size"""

    highlight: Optional[str] = None
    """Specifies a highlighting color which is applied as a background 
    behind the contents.

    If the content has any background color specified, then the 
    background color shall be superseded by the highlighting color when 
    the contents of are displayed.
    
    Value is the name of a color
    """

    imprint: Optional[bool] = None
    """Specifies that the contents should be displayed as if imprinted, 
    which makes text appear to be imprinted or pressed into page (also 
    referred to as 'engrave').
    
    This element shall not be present with either the emboss or outline 
    properties, since they are mutually exclusive in terms of 
    appearance.
    """

    italics: Optional[bool] = None
    """Specifies whether the italic property shall be applied.

    If this element is not present, the default value is to leave the 
    formatting applied at previous level in the style hierarchy.
    """

    letter_spacing: Optional[CssUnit] = None
    """Specifies the amount of character pitch which shall be added or 
    removed after each character before the following character is 
    rendered in the document.
    """

    outline: Optional[bool] = None
    """Specifies that the contents of this run should be displayed as if
    they have an outline, by drawing a one pixel wide border around the 
    inside and outside borders of each character glyph in the run.
    """

    position: Optional[CssUnit] = None
    """Specifies the vertical position of the text in relation to the
    baseline. Positive values for raised text and negative values for
    lowered text
    """

    shadow: Optional[bool] = None
    """Specifies that the contents of this run shall be displayed as if 
    each character has a shadow.
    """

    small_caps: Optional[bool] = None
    """Specifies that all small letter characters in this text run shall
    be formatted for display only as their capital letter character 
    equivalents in a font size two points smaller than the actual font 
    size specified for this text. This property does not affect any non-
    alphabetic character in this run, and does not change the Unicode
    character for lowercase text, only the method in which it is 
    displayed. If this font cannot be made two point smaller than the 
    current size, then it shall be displayed as the smallest possible 
    font size in capital letters.
    """

    strike: Optional[bool] = None
    """Specifies that the contents shall be displayed with a single 
    horizontal line through the center of the line.
    
    This element shall not be present with the double_strike property, 
    since they are mutually exclusive in terms of appearance.
    """

    underline: Optional[TextDecoration] = None
    """Specifies that the contents should be displayed along with an 
    underline appearing directly below the character height.
    """

    vertical_align: Optional[str] = None
    """Specifies the alignment which shall be applied to the contents in
    relation to the default appearance of the text. This allows the text
    to be repositioned as subscript or superscript.
    
    Possible values are:
        * baseline
        * superscript
        * subscript
    """

    visible: Optional[bool] = None
    """Specifies whether the contents shall be hidden from display at 
    display time in a document. Note: The setting should affect the 
    normal display of text, but an application can have settings to
    force hidden text to be displayed.
    """

    def text_properties(self, active=False):
        for f in sorted(fields(TextFormatting), key=lambda x: x.name):
            value = getattr(self, f.name)
            if active and value is None:
                continue
            else:
                yield f.name, value

    # Not implemented:
    #   * bCs (Complex Script Bold) §2.3.2.2
    #   * cs (Use Complex Script Formatting on Run) §2.3.2.6
    #   * eastAsianLayout (East Asian Typography Settings) §2.3.2.8
    #   * effect (Animated Text Effect) §2.3.2.9
    # em (Emphasis Mark) §2.3.2.10
    # fitText (Manual Run Width) §2.3.2.12
    #   * iCs (Complex Script Italics) §2.3.2.15
    #   * lang (Languages for Run Content) §2.3.2.18
    #   * noProof (Do Not Check Spelling or Grammar) §2.3.2.19
    #   * oMath (Office Open XML Math) §2.3.2.20
    #   * rPrChange (Revision Information for Run Properties) §2.13.5.32
    #   * rStyle (Referenced Character Style) §2.3.2.27
    #   * rtl (Right To Left Text) §2.3.2.28
    #   * snapToGrid (Use Document Grid Settings For Inter-Character Spacing) §2.3.2.32
    #   * specVanish (Paragraph Mark Is Always Hidden) §2.3.2.34
    #   * szCs (Complex Script Font Size) §2.3.2.37
    # w (Expanded/Compressed Text) §2.3.2.41
    #   * webHidden (Web Hidden Text) §2.3.2.42


@dataclass
class ParagraphFormatting(TextFormatting):
    border_bottom: Border = None
    border_left: Border = None
    border_right: Border = None
    border_top: Border = None

    keep_together: Optional[bool] = None
    """Specifies that when rendering in a paginated view, all lines are
    maintained on a single page whenever possible.
    """

    keep_with_next: Optional[bool] = None
    """Specifies that when rendering in a paginated view, the contents 
    are at least partly rendered on the same page as the following 
    paragraph whenever possible.
    """

    counter: Optional['Counter'] = None

    line_height: Optional[CssUnit] = None
    """Specifies the inter-line spacing which shall be applied to the 
    contents when it is displayed.
    
    If the value is NOT an instance of CssUnit (and is instead and int), 
    then the value of the line height must be multiplied by the font
    size.
    """

    margin_left: Optional[CssUnit] = None
    margin_right: Optional[CssUnit] = None
    margin_bottom: Optional[CssUnit] = None
    margin_top: Optional[CssUnit] = None

    page_break_before: Optional[bool] = None
    """Specifies that when rendering in a paginated view, the contents 
    are rendered on the start of a new page in the document.
    """

    text_align: Optional[str] = None
    text_indent: Optional[CssUnit] = None
    widows_control: Optional[bool] = None
    """Specifies whether a consumer shall prevent a single line of this 
    paragraph from being displayed on a separate page from the remaining
    content at display time by moving the line onto the following page.
    """

    def paragraph_properties(self, active=False, with_text_fields=True):
        all_fields = set(fields(ParagraphFormatting))
        if not with_text_fields:
            all_fields -= set(fields(TextFormatting))
        for f in sorted(all_fields, key=lambda x: x.name):
            value = getattr(self, f.name)
            if active and value is None:
                continue
            else:
                yield f.name, value

    # adjustRightInd (Automatically Adjust Right Indent When Using Document Grid) §2.3.1.1
    # autoSpaceDE (Automatically Adjust Spacing of Latin and East Asian Text) §2.3.1.2
    # autoSpaceDN (Automatically Adjust Spacing of East Asian Text and Numbers) §2.3.1.3
    # bidi (Right to Left Paragraph Layout) §2.3.1.6
    # cnfStyle (Paragraph Conditional Formatting) §2.3.1.8
    # contextualSpacing (Ignore Spacing Above and Below When Using Identical Styles) §2.3.1.9
    # divId (Associated HTML div ID) §2.3.1.10
    # framePr (Text Frame Properties) §2.3.1.11
    # ind (Paragraph Indentation) §2.3.1.12
    # jc (Paragraph Alignment) §2.3.1.13
    # keepLines (Keep All Lines On One Page) §2.3.1.14
    # keepNext (Keep Paragraph With Next Paragraph) §2.3.1.15
    # kinsoku (Use East Asian Typography Rules for First and Last Character per Line) §2.3.1.16
    # mirrorIndents (Use Left/Right Indents as Inside/Outside Indents) §2.3.1.18
    # numPr (NumberingPart Definition Instance Reference) §2.3.1.19
    # outlineLvl (Associated Outline Level) §2.3.1.20
    # overflowPunct (Allow Punctuation to Extent Past Text Extents) §2.3.1.21
    # pageBreakBefore (Start Paragraph on Next Page) §2.3.1.23
    # pBdr (Paragraph Borders) §2.3.1.24
    # pPrChange (Revision Information for Paragraph Properties) §2.13.5.31
    # pStyle (Referenced Paragraph Style) §2.3.1.27
    # rPr (Run Properties for the Paragraph Mark) §2.3.1.29
    # sectPr (Section Properties) §2.6.19
    # shd (Paragraph Shading) §2.3.1.31
    # snapToGrid (Use Document Grid Settings for Inter-Line Paragraph Spacing) §2.3.1.32
    # spacing (Spacing Between Lines and Above/Below Paragraph) §2.3.1.33
    # suppressAutoHyphens (Suppress Hyphenation for Paragraph) §2.3.1.34
    # suppressLineNumbers (Suppress Line Numbers for Paragraph) §2.3.1.35
    # suppressOverlap (Prevent Text Frames From Overlapping) §2.3.1.36
    # tabs (Set of Custom Tab Stops) §2.3.1.38
    # textAlignment (Vertical Character Alignment on Line) §2.3.1.39
    # textboxTightWrap (Allow Surrounding Paragraphs to Tight Wrap to Text Box Contents) §2.3.1.40
    # textDirection (Paragraph Text Flow Direction) §2.3.1.41
    # topLinePunct (Compress Punctuation at Start of a Line) §2.3.1.43
    # widowControl (Allow First/Last Line to Display on a Separate Page) §2.3.1.44
    # wordWrap (Allow Line Breaking At Character Level) §2.3.1.45


@dataclass
class TableProperties:
    alignment: Optional[str] = None
    """Specifies the alignment of the current table with respect to the 
    text margins. When the table does not have the same width as the 
    margins, this property is used to determine how the table is
    positioned with respect to those margins.

    If this property is not specified in the style hierarchy, then the 
    table shall be left justified with zero indentation from the leading
    margin.

    Possible values are defined docx2css.ooxml.simple_types.ST_Jc:
        * start
        * end
        * center
        * justify
    """

    background_color: Optional[str] = None
    """Specifies the background color of the table.
    
    Return a hexadecimal color string, eg. '#FF00FF' or None
    """

    border_bottom: Border = None
    border_inside_horizontal: Border = None
    border_inside_vertical: Border = None
    border_left: Border = None
    border_right: Border = None
    border_top: Border = None

    cell_padding_bottom: Optional[CssUnit] = None
    cell_padding_left: Optional[CssUnit] = None
    cell_padding_right: Optional[CssUnit] = None
    cell_padding_top: Optional[CssUnit] = None

    cell_spacing: Optional[CssUnit] = None
    """Specifies the default table cell spacing (the spacing between
    adjacent cells and the edges of the table) for all cells in the
    parent table. If specified, this element specifies the minimum
    amount of space which shall be left between all cells in the table 
    including the width of the table borders in the calculation.
    """

    col_band_size: Optional[int] = None
    """Specifies the number of columns which shall comprise each a table
    style column band for this table style. This element determines how 
    many columns constitute each of the column bands for the current 
    table, allowing column band formatting to be applied to groups of 
    columns (rather than just single alternating columns) when the table
    is formatted.

    Default value is 1.
    """

    indent: Optional[CssUnit] = None
    """Specifies the indentation which shall be added before the leading
    edge of the current table in the document (the left edge in a 
    left-to-right table, and the right edge in a right-to-left table). 
    This indentation should shift the table into the text margin by the 
    specified amount.
    """

    layout: Optional[str] = None
    """Specifies the algorithm which shall be used to lay out the
    contents of the table. It can either be displayed using a fixed
    width or autofit.

    If this element is omitted, then the value of this element shall be 
    assumed to be auto .

    Possible values:
        * fixed
        * auto
    """

    row_band_size: Optional[int] = None
    """Specifies the number of rows which shall comprise each a table 
    style row band for this table style. This element determines how 
    many rows constitute each of the row bands for the current table, 
    allowing row band formatting to be applied to groups of rows (rather
    than just single alternating rows) when the table is formatted.

    Default value is 1.
    """

    width: Optional[CssUnit] = None
    """Specifies the preferred width for this table.

    If this element is omitted, then the cell width shall be of type
    'auto'.
    """

    # Not implemented:
    #   * overlap
    #   * bidiVisual
    #   * look
    #   * table floating positioning


@dataclass
class TableRowProperties:
    alignment: Optional[str] = None
    """Specifies the alignment of a single row in the parent table with 
    respect to the text margins. When the table does not have the same 
    width as the margins, this property is used to determine how the 
    table is positioned with respect to those margins.

    If this property is omitted on a table, then the justification shall
    be determined by the default set of table properties on the parent 
    table.

    Possible values are defined docx2css.ooxml.simple_types.ST_Jc:
        * start
        * end
        * center
        * justify
    """

    split: Optional[bool] = None
    """Specifies whether the contents within the current cell shall be 
    rendered on a single page. If this property is False, then all 
    contents of a table row shall be rendered on the same page by moving
    the start of the current row to the start of a new page if 
    necessary. If the contents of this table row cannot fit on a single 
    page, then this row shall start on a new page and flow onto multiple
    pages as necessary.

    If this property is not specified in the style hierarchy, then this 
    table row shall be allowed to split across multiple pages.
    """

    cell_spacing: Optional[CssUnit] = None
    """Specifies the default table cell spacing (the spacing between
    adjacent cells and the edges of the table) for all cells in the
    parent row. If specified, this element specifies the minimum amount 
    of space which shall be left between all cells in the table 
    including the width of the table borders in the calculation.
    """

    height: Optional[CssUnit] = None
    """Specifies the exact height of the rows"""

    is_header: Optional[bool] = None
    """specifies that the current table row shall be repeated at the top
    of each new page on which part of this table is displayed. This 
    gives this table row the behavior of a 'header' row on each of these
    pages.
    """

    min_height: Optional[CssUnit] = None
    """Specifies the minimum height of the rows"""

    def table_row_properties(self, active=False):
        for f in sorted(fields(TableRowProperties), key=lambda x: x.name):
            value = getattr(self, f.name)
            if active and value is None:
                continue
            else:
                yield f.name, value

    # Not implemented:
    #   * cnfStyle (Table Row Conditional Formatting) §2.4.8
    #   * del (Deleted Table Row) §2.13.5.14
    #   * divId (Associated HTML div ID) §2.4.9
    #   * gridAfter (Grid Columns After Last Cell) §2.4.10
    #   * gridBefore (Grid Columns Before First Cell) §2.4.11
    #   * hidden (Hidden Table Row Marker) §2.4.14
    #   * ins (Inserted Table Row) §2.13.5.16
    #   * trPrChange (Revision Information for Table Row Properties) §2.13.5.39
    #   * wAfter (Preferred Width After Table Row) §2.4.82
    #   * wBefore (Preferred Width Before Table Row)


@dataclass
class TableCellProperties:
    background_color: Optional[str] = None
    """Specifies the default background color of the table cells.

    Return a hexadecimal color string, eg. '#FF00FF' or None
    """

    border_bottom: Border = None
    border_inside_horizontal: Border = None
    border_inside_vertical: Border = None
    border_left: Border = None
    border_right: Border = None
    border_top: Border = None

    colspan: Optional[int] = None
    """Specifies the number of grid columns in the parent table's table 
    grid which shall be spanned by the current cell. If this element is 
    omitted, then the number of grid units spanned by this cell shall be
    assumed to be one.
    """

    fit_text: Optional[bool] = None
    """Specifies that the contents of the current cell shall have their 
    inter-character spacing increased or reduced as necessary to fit the
    width of the text extents of the current cell.
    """

    padding_bottom: Optional[CssUnit] = None
    padding_left: Optional[CssUnit] = None
    padding_right: Optional[CssUnit] = None
    padding_top: Optional[CssUnit] = None

    valign: Optional[str] = None
    """Specifies the vertical alignment for text within the cells

    Possible values:
        * top
        * center
        * bottom
    """

    width: Optional[CssUnit] = None
    """Specifies the preferred width for this table cell"""

    wrap_text: Optional[bool] = None
    """Specifies whether the content of the cells shall be allowed to 
    wrap.
    """

    def table_cell_properties(self, active=False):
        for f in sorted(fields(TableCellProperties), key=lambda x: x.name):
            value = getattr(self, f.name)
            if active and value is None:
                continue
            else:
                yield f.name, value

    # Not implemented:
    #   * cellDel (Table Cell Deletion) §2.13.5.1
    #   * cellIns (Table Cell Insertion) §2.13.5.2
    #   * cellMerge (Vertically Merged/Split Table Cells) §2.13.5.3
    #   * cnfStyle (Table Cell Conditional Formatting) §2.4.7
    #   * hideMark (Ignore End Of Cell Marker In Row Height Calculation) §2.4.15
    #   * hMerge (Horizontally Merged Cell) §2.4.16
    #   * tcPrChange (Revision Information for Table Cell Properties) §2.13.5.38
    #   * textDirection (Table Cell Text Flow Direction) §2.4.69
    #   * vMerge (Vertically Merged Cell) §2.4.81


class BodyStyle(ParagraphFormatting):
    type = 'body'


@dataclass
class BaseStyle(ABC):

    name: str
    id: str
    parent: 'BaseStyle' = None
    parent_id: str = None
    children: list = field(default_factory=list)

    @property
    @abstractmethod
    def type(self):
        pass

    @property
    def qualified_id(self):
        return '.'.join(filter(lambda x: x, (self.type, self.id)))

    @property
    def qualified_name(self):
        return '.'.join(filter(lambda x: x, (self.type, self.name)))

    @property
    def qualified_parent_id(self):
        return '.'.join(filter(lambda x: x, (self.type, self.parent_id)))


@dataclass
class SpanStyle(TextFormatting, BaseStyle):
    type = 'span'


@dataclass
class ParagraphStyle(ParagraphFormatting, BaseStyle):
    type = 'p'


@dataclass
class TableConditionalFormatting(ParagraphFormatting, TableProperties):
    default_cell: TableCellProperties = None
    default_row: TableRowProperties = None

    def table_properties(self, active=True):
        exclude = ('name', 'id', 'parent', 'parent_id', 'children', 'type')
        for f in fields(self):
            value = getattr(self, f.name)
            if f.name in exclude or active and value is None:
                continue
            else:
                yield f.name, value


@dataclass
class TableStyle(TableConditionalFormatting, BaseStyle):
    type = 'table'
    # default_cell: TableCellProperties = None
    # default_row: TableRowProperties = None
    # conditional formatting:
    whole_table: Optional[TableConditionalFormatting] = None
    odd_columns: Optional[TableConditionalFormatting] = None
    even_columns: Optional[TableConditionalFormatting] = None
    odd_rows: Optional[TableConditionalFormatting] = None
    even_rows: Optional[TableConditionalFormatting] = None
    first_row: Optional[TableConditionalFormatting] = None
    last_row: Optional[TableConditionalFormatting] = None
    first_column: Optional[TableConditionalFormatting] = None
    last_column: Optional[TableConditionalFormatting] = None
    top_left_cell: Optional[TableConditionalFormatting] = None
    top_right_cell: Optional[TableConditionalFormatting] = None
    bottom_left_cell: Optional[TableConditionalFormatting] = None
    bottom_right_cell: Optional[TableConditionalFormatting] = None


@dataclass
class CounterList:
    id: str
    name: str
    counters: dict = field(default_factory=dict)


@dataclass
class Counter(ParagraphFormatting):
    counter_list: CounterList = None
    name: str = None
    style: str = 'decimal'
    start: int = 0
    text: str = None

    restart: set = field(default_factory=set)
    """Levels that are restarted at this level"""

    suffix: str = 'tab'
    """Specifies whether the contents should have 'nothing', a 'tab' or
    a 'space' appended.
    """

    justification: str = None


@dataclass
class PageStyle:
    type = 'page'
    margin_left: Optional[CssUnit] = None
    margin_right: Optional[CssUnit] = None
    margin_bottom: Optional[CssUnit] = None
    margin_top: Optional[CssUnit] = None
    page_height: Optional[CssUnit] = None
    page_orientation: str = 'portrait'
    page_width: Optional[CssUnit] = None
