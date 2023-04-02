from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional

import cssutils

from docx2css.utils import CssUnit, PropertyContainer


def getter(prop_name):
    def func(self):
        impl_getter = getattr(self, f'_get_{prop_name}')
        return getattr(self, f'_{prop_name}', impl_getter())
    return func


def setter(prop_name):
    return lambda self, value: setattr(self, f'_{prop_name}', value)


INCLUDE_PAGE_RULE = 'include_page_rule'
SIMULATE_PRINTED_PAGE = 'simulate_printed_page'


class PageStyle(ABC):

    height: CssUnit = property(
        lambda self: getter('height')(self),
        lambda self, value: setter('height')(self, value),
    )
    margin_bottom: CssUnit = property(
        lambda self: getter('margin_bottom')(self),
        lambda self, value: setter('margin_bottom')(self, value),
    )
    margin_left: CssUnit = property(
        lambda self: getter('margin_left')(self),
        lambda self, value: setter('margin_left')(self, value),
    )
    margin_right: CssUnit = property(
        lambda self: getter('margin_right')(self),
        lambda self, value: setter('margin_right')(self, value),
    )
    margin_top: CssUnit = property(
        lambda self: getter('margin_top')(self),
        lambda self, value: setter('margin_top')(self, value),
    )
    width: CssUnit = property(
        lambda self: getter('width')(self),
        lambda self, value: setter('width')(self, value),
    )

    @abstractmethod
    def _get_height(self) -> CssUnit:
        pass

    @abstractmethod
    def _get_margin_bottom(self) -> CssUnit:
        pass

    @abstractmethod
    def _get_margin_left(self) -> CssUnit:
        pass

    @abstractmethod
    def _get_margin_right(self) -> CssUnit:
        pass

    @abstractmethod
    def _get_margin_top(self) -> CssUnit:
        pass

    @abstractmethod
    def _get_width(self) -> CssUnit:
        pass

    def _css_margin_value(self):
        top = self.margin_top.inches
        right = self.margin_right.inches
        bottom = self.margin_bottom.inches
        left = self.margin_left.inches
        return f'{top}in {right}in {bottom}in {left}in'

    def css_style_declaration_print(self):
        css_style = cssutils.css.CSSStyleDeclaration()
        css_style['size'] = f'{self.width.inches}in {self.height.inches}in'
        css_style['margin'] = self._css_margin_value()
        return css_style

    def css_style_rule_print(self):
        css_style = self.css_style_declaration_print()
        return cssutils.css.CSSPageRule(style=css_style)

    def css_style_declaration_screen(self):
        css_style = cssutils.css.CSSStyleDeclaration()
        css_style['background-color'] = 'white'
        css_style['border'] = '1px darkgray solid'
        css_style['box-shadow'] = '1rem 0.5rem 1rem rgba(0,0,0,0.15)'
        # Adjust max-width to margins
        max_width = self.width - self.margin_left - self.margin_right
        css_style['max-width'] = f'{CssUnit(max_width).inches}in'
        css_style['margin'] = '1em auto'
        css_style['padding'] = self._css_margin_value()

        return css_style

    def css_style_rule_screen(self):
        screen = cssutils.css.CSSMediaRule('screen')

        html_style = cssutils.css.CSSStyleDeclaration()
        html_style['background-color'] = 'gainsboro'
        html_rule = cssutils.css.CSSStyleRule('html', html_style)
        screen.add(html_rule)

        body_style = self.css_style_declaration_screen()
        body_rule = cssutils.css.CSSStyleRule('body', body_style)
        screen.add(body_rule)
        return screen

    def css_style_rules(self, preferences=None):
        preferences = preferences or {}
        rules = []
        if preferences.get(INCLUDE_PAGE_RULE, True):
            rules.append(self.css_style_rule_print())
        if preferences.get(SIMULATE_PRINTED_PAGE, False):
            rules.append(self.css_style_rule_screen())

        return rules


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
class TextFormatting(PropertyContainer):
    bold: Optional[bool] = None
    """Specifies whether the bold property shall be applied to all non-
    complex script characters
    """

    italics: Optional[bool] = None
    """Specifies whether the italic property shall be applied.

    If this element is not present, the default value is to leave the 
    formatting applied at previous level in the style hierarchy.
    """

    # Not implemented:
    #   * bCs (Complex Script Bold) §2.3.2.2
    # bdr (Text Border) §2.3.2.3
    # caps (Display All Characters As Capital Letters) §2.3.2.4
    # color (Run Content Color) §2.3.2.5
    #   * cs (Use Complex Script Formatting on Run) §2.3.2.6
    # dstrike (Double Strikethrough) §2.3.2.7
    #   * eastAsianLayout (East Asian Typography Settings) §2.3.2.8
    #   * effect (Animated Text Effect) §2.3.2.9
    # em (Emphasis Mark) §2.3.2.10
    # emboss (Embossing) §2.3.2.11
    # fitText (Manual Run Width) §2.3.2.12
    # highlight (Text Highlighting) §2.3.2.13
    #   * iCs (Complex Script Italics) §2.3.2.15
    # imprint (Imprinting) §2.3.2.16
    # kern (Font Kerning) §2.3.2.17
    #   * lang (Languages for Run Content) §2.3.2.18
    #   * noProof (Do Not Check Spelling or Grammar) §2.3.2.19
    #   * oMath (Office Open XML Math) §2.3.2.20
    # outline (Display Character Outline) §2.3.2.21
    # position (Vertically Raised or Lowered Text) §2.3.2.22
    # rFonts (Run Fonts) §2.3.2.24
    #   * rPrChange (Revision Information for Run Properties) §2.13.5.32
    #   * rStyle (Referenced Character Style) §2.3.2.27
    # rtl (Right To Left Text) §2.3.2.28
    # shadow (Shadow) §2.3.2.29
    # shd (Run Shading) §2.3.2.30
    # smallCaps (Small Caps) §2.3.2.31
    #   * snapToGrid (Use Document Grid Settings For Inter-Character Spacing) §2.3.2.32
    # spacing (Character Spacing Adjustment) §2.3.2.33
    # specVanish (Paragraph Mark Is Always Hidden) §2.3.2.34
    # strike (Single Strikethrough) §2.3.2.35
    # sz (Font Size) §2.3.2.36
    #   * szCs (Complex Script Font Size) §2.3.2.37
    # u (Underline) §2.3.2.38
    # vanish (Hidden Text) §2.3.2.39
    # vertAlign (Subscript/Superscript Text) §2.3.2.40
    # w (Expanded/Compressed Text) §2.3.2.41
    #   * webHidden (Web Hidden Text) §2.3.2.42


@dataclass
class ParagraphFormatting(TextFormatting):
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
    # numPr (Numbering Definition Instance Reference) §2.3.1.19
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
    pass


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
class TableRowProperties(PropertyContainer):
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
class TableCellProperties(PropertyContainer):
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


class Stylesheet:

    def __init__(self, styles=None):
        self.styles = styles if styles is not None else {}

    def add_style(self, style):
        self.styles[style.id] = style
        if style.parent_id:
            parent = self.styles.get(style.parent_id, None)
            style.parent = parent
            parent.children.append(style)


@dataclass
class BaseStyle(PropertyContainer, ABC):

    name: str
    id: str
    parent: 'BaseStyle' = None
    parent_id: str = None
    children: list = field(default_factory=list)

    @property
    @abstractmethod
    def type(self):
        pass


@dataclass
class TableConditionalFormatting(ParagraphFormatting, TableProperties):
    default_cell: TableCellProperties = None
    default_row: TableRowProperties = None


@dataclass
class TableStyle(ParagraphFormatting, TableProperties, BaseStyle):
    type = 'table'
    default_cell: TableCellProperties = None
    default_row: TableRowProperties = None
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
