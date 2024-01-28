from collections.abc import Mapping

from lxml import etree

from docx2css.api import Border, TextDecoration
from docx2css.ooxml import ct, w, wordml
from docx2css.ooxml.constants import CONTENT_TYPE
from docx2css.ooxml.simple_types import (
    ST_Border
)
from docx2css.utils import CSSColor, CssUnit


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


class RPrMixin:

    all_caps: bool = ct.Boolean('w:rPr/w:caps')
    """This element specifies that any lowercase characters in this text run 
    shall be formatted for display only as their capital letter character 
    equivalents. 
    
    This element shall not be present with the smallCaps (§17.3.2.33) property 
    on the same run, since they are mutually exclusive in terms of appearance.
    """

    background_color: CSSColor = ct.Shading('w:rPr/w:shd')
    """This element specifies the shading applied to the contents of the run."""

    bold = ct.Boolean('w:rPr/w:b')
    """This element specifies whether the bold property shall be applied to all 
    non-complex script characters in the contents of this run when displayed in 
    a document.
    """

    border: Border = ct.BorderDescriptor('w:rPr/w:bdr')
    """This element specifies information about the border applied to the text 
    in the current run.
    
    The first piece of information specified by the bdr element is that the 
    current shall have a border when displayed. This information is specified 
    simply by the presence of the bdr element in run's properties.
    
    The second piece of information concerns the set of runs which share the 
    current run border. This is determined based on the attributes on the bdr 
    element. If the set of attribute values specifies on two adjacent runs is
    identical, then those two runs shall be considered to be part of the same 
    run border group and rendered within the same set of borders in the 
    document.
    """

    double_strike: bool = ct.Boolean('w:rPr/w:dstrike')
    """This element specifies that the contents of this run shall be displayed 
    with two horizontal lines through each character displayed on the line.
    
    This element shall not be present with the strike (§17.3.2.37) property on 
    the same run, since they are mutually exclusive in terms of appearance.
    """

    emboss: bool = ct.Boolean('w:rPr/w:emboss')
    """This element specifies that the contents of this run should be displayed 
    as if embossed, which makes text appear as if it is raised off the page in 
    relief.
    
    This element shall not be present with either the imprint (§17.3.2.18) or 
    outline (§17.3.2.23) properties on the same run, since they are mutually 
    exclusive in terms of appearance.
    """

    font_color: CSSColor = ct.Shading('w:rPr/w:color')
    """This element specifies the color which shall be used to display the 
    contents of this run in the document. 
    """

    font_family: str = ct.FontDescriptor('w:rPr/w:rFonts')
    """This element specifies the fonts which shall be used to display the text 
    contents of this run.
    """

    font_kerning: CssUnit = ct.HalfPointMeasure('w:rPr/w:kern')
    """This element specifies whether font kerning shall be applied to the 
    contents of this run. If it is specified, then kerning shall be 
    automatically adjusted when displaying characters in this run as needed.
    
    The val attribute specifies the smallest font size which shall have its 
    kerning automatically adjusted if this setting is specified. If the font 
    size in the sz element (§17.3.2.38) is smaller than this value, then no font 
    kerning shall be performed. 
    """

    font_size: CssUnit = ct.HalfPointMeasure('w:rPr/w:sz')

    highlight: str = ct.String('w:rPr/w:highlight')
    """This element specifies a highlighting color which is applied as a 
    background behind the contents of this run.
    
    If this run has any background shading specified using the shd element 
    (§17.3.2.32), then the background shading shall be superseded by the 
    highlighting color when the contents of this run are displayed.
    """

    imprint: bool = ct.Boolean('w:rPr/w:imprint')
    """This element specifies that the contents of this run should be displayed 
    as if imprinted, which makes text appear to be imprinted or pressed into 
    page (also referred to as 'engrave').
    
    This element shall not be present with either the emboss (§17.3.2.13) or 
    outline (§17.3.2.23) properties on the same run, since they are mutually 
    exclusive in terms of appearance.
    """

    italics: bool = ct.Boolean('w:rPr/w:i')
    """This element specifies whether the italic property should be applied to 
    all non-complex script characters in the contents of this run when displayed 
    in a document.
    """

    letter_spacing: CssUnit = ct.TwipMeasure('w:rPr/w:spacing')
    """This element specifies the amount of character pitch which shall be added 
    or removed after each character in this run before the following character 
    is rendered in the document. This property has an effect equivalent to
    the additional character pitched added by a document grid applied to the 
    contents of a run. 
    """

    outline: bool = ct.Boolean('w:rPr/w:outline')
    """This element specifies that the contents of this run should be displayed 
    as if they have an outline, by drawing a one pixel wide border around the 
    inside and outside borders of each character glyph in the run.
    
    This element shall not be present with either the emboss (§17.3.2.13) or 
    imprint (§17.3.2.18) properties on the same run, since they are mutually 
    exclusive in terms of appearance.
    """

    position: int = ct.HalfPointMeasure('w:rPr/w:position')
    """This element specifies the amount by which text shall be raised or 
    lowered for this run in relation to the default baseline of the surrounding 
    non-positioned text. This allows the text to be repositioned without 
    altering the font size of the contents.
    
    If the val attribute is positive, then the parent run shall be raised above 
    the baseline of the surrounding text by the specified number of half-points. 
    If the val attribute is negative, then the parent run shall be lowered below
    the baseline of the surrounding text by the specified number of half-points.
    """

    shadow: bool = ct.Boolean('w:rPr/w:shadow')
    """This element specifies that the contents of this run shall be displayed 
    as if each character has a shadow. For left-to-right text, the shadow is 
    beneath the text and to its right; for right-to-left text, the shadow is 
    beneath the text and to its left.
    
    This element shall not be present with either the emboss (§17.3.2.13) or 
    imprint (§17.3.2.18) properties on the same run, since they are mutually 
    exclusive in terms of appearance.
    """

    small_caps: bool = ct.Boolean('w:rPr/w:smallCaps')
    """This element specifies that all small letter characters in this text run 
    shall be formatted for display only as their capital letter character 
    equivalents in a font size two points smaller than the actual font size 
    specified for this text. This property does not affect any non-alphabetic 
    character in this run, and does not change the Unicode character for 
    lowercase text, only the method in which it is displayed. If this font 
    cannot be made two point smaller than the current size, then it shall be 
    displayed as the smallest possible font size in capital letters.
    
    This element shall not be present with the caps (§17.3.2.5) property on the 
    same run, since they are mutually exclusive in terms of appearance.
    """

    strike: bool = ct.Boolean('w:rPr/w:strike')
    """This element specifies that the contents of this run shall be displayed 
    with a single horizontal line through the center of the line. 
    
    This element shall not be present with the dstrike (§17.3.2.9) property on 
    the same run, since they are mutually exclusive in terms of appearance.
    """

    underline: TextDecoration = ct.UnderlineDescriptor('w:rPr/w:u')
    """This element specifies that the contents of this run should be displayed 
    along with an underline appearing directly below the character height (less 
    all spacing above and below the characters on the line).
    """

    vanish = ct.Boolean('w:rPr/w:vanish')
    """This element specifies whether the contents of this run shall be hidden 
    from display at display time in a document.
    """

    vertical_align = ct.VerticalJustification('w:rPr/w:vertAlign')
    """This element specifies the alignment which shall be applied to the 
    contents of this run in relation to the default appearance of the run's 
    text. This allows the text to be repositioned as subscript or superscript 
    without altering the font size of the run properties.
    """


class PPrMixin:
    background_color: CSSColor = ct.Shading('w:pPr/w:shd')
    """This element specifies the shading applied to the contents of the 
    paragraph.
    """

    border_bottom: Border = ct.BorderDescriptor('w:pPr/w:pBdr/w:bottom')
    """This element specifies the border which shall be displayed below a set of 
    paragraphs which have the same paragraph border settings.
    
    To determine if any two adjoining paragraphs shall have an individual top 
    and bottom border or a between border, the set of borders on the two 
    adjoining paragraphs are compared. If the border information on those two 
    paragraphs is different, then the first paragraph shall use its bottom 
    border and the following paragraph shall use its top border. Otherwise, the 
    between border is used. If this border specifies a space attribute, that
    value determines the space after the bottom of the text (ignoring any space 
    below) which should be left before this border is drawn, specified in 
    points.
    
    If this element is omitted on a given paragraph, its value is determined by 
    the setting previously set at any level of the style hierarchy (i.e. that 
    previous setting remains unchanged). If this setting is never specified in 
    the style hierarchy, then no between border shall be applied below identical 
    paragraphs.
    """

    border_left: Border = ct.BorderDescriptor('w:pPr/w:pBdr/w:left')
    """This element specifies the border which shall be displayed on the left 
    side of the page around the specified paragraph. This shall not change based 
    on the paragraph direction.
    
    To determine if any two adjoining paragraphs should have a left border which 
    spans the full line height or not, the left border shall be drawn between 
    the top border or between border at the top (whichever would be rendered for 
    the current paragraph), and the bottom border or between border at the 
    bottom (whichever would be rendered for the current paragraph).
    
    If this element is omitted on a given paragraph, its value is determined by 
    the setting previously set at any level of the style hierarchy (i.e. that 
    previous setting remains unchanged). If this setting is never specified in 
    the style hierarchy, then no left border shall be applied.
    """

    border_right: Border = ct.BorderDescriptor('w:pPr/w:pBdr/w:right')
    """This element specifies the border which shall be displayed on the right 
    side of the page around the specified paragraph. This shall not change based 
    on the paragraph direction.
    
    To determine if any two adjoining paragraphs should have a right border 
    which spans the full line height or not, the right border shall be drawn 
    between the top border or between border at the top (whichever would be
    rendered for the current paragraph), and the bottom border or between border 
    at the bottom (whichever would be rendered for the current paragraph).
    
    If this element is omitted on a given paragraph, its value is determined by 
    the setting previously set at any level of the style hierarchy (i.e. that 
    previous setting remains unchanged). If this setting is never specified in 
    the style hierarchy, then no right border shall be applied.
    """

    border_top: Border = ct.BorderDescriptor('w:pPr/w:pBdr/w:top')
    """This element specifies the  border which shall be displayed above a set 
    of paragraphs which have the same set of paragraph border settings.
    
    To determine if any two adjoining paragraphs shall have an individual top 
    and bottom border or a between border, the set of borders on the two 
    adjoining paragraphs are compared. If the border information on those two 
    paragraphs is identical for all possible paragraphs borders, then the 
    between border is displayed. Otherwise, the final paragraph shall use its 
    bottom border and the following paragraph shall use its top border,
    respectively. If this border specifies a space attribute, that value 
    determines the space above the text (ignoring any spacing above) which 
    should be left before this border is drawn, specified in points.
    
    If this element is omitted on a given paragraph, its value is determined by 
    the setting previously set at any level of the style hierarchy (i.e. that 
    previous setting remains unchanged). If this setting is never specified in 
    the style hierarchy, then no between border shall be applied above identical 
    paragraphs.
    """

    indent_left = ct.ParagraphIndentLeft('w:pPr/w:ind')
    """Specifies the indentation which shall be placed at the start of this 
    paragraph – between the left text margin for this paragraph and the left 
    edge of that paragraph's content in a left to right paragraph, and the right 
    text margin and the right edge of that paragraph's text in a right to left 
    paragraph. If the mirrorIndents property (§17.3.1.18) is specified for this 
    paragraph, then this indent is used for the inside page edge - the right 
    page edge for odd numbered pages and the left page edge for even numbered 
    pages.
    
    If this attribute is omitted, its value shall be assumed to be zero.
    
    All other values for this element are relative to the leading text margin, 
    Negative values are defined such that the text is moved past the text 
    margin, positive values move the text inside the text margin. This value can 
    be superseded for the first line only via use of the firstLine or hanging 
    attributes.
    """

    indent_right = ct.ParagraphIndentRight('w:pPr/w:ind')
    """Specifies the indentation which shall be placed at the end of this 
    paragraph – between the right text margin for this paragraph and the right 
    edge of that paragraph's content in a left to right paragraph, and the left 
    text margin and the left edge of that paragraph's text in a right to left 
    paragraph. 
    
    If this attribute is omitted, its value shall be assumed to be zero.
    
    All other values for this element are relative to the trailing text margin, 
    Negative values are defined such that the text is moved past the text 
    margin, positive values move the text inside the text margin.
    """

    keep_together: bool = ct.Boolean('w:pPr/w:keepLines')
    """This element specifies that when rendering this document in a page view, 
    all lines of this paragraph are maintained on a single page whenever 
    possible.
    
    This means that if the contents of the current paragraph would normally span 
    across two pages due to the placement of the paragraph's text, all lines in 
    this paragraph shall be moved onto the next page to ensure they are 
    displayed together. If this is not possible because all lines in the 
    paragraph would exceed a single page in any case, then lines in this 
    paragraph shall start on a new page, with page breaks as needed afterwards.
    """

    keep_with_next: bool = ct.Boolean('w:pPr/w:keepNext')
    """This element specifies that when rendering this document in a paginated 
    view, the contents of this paragraph are at least partly rendered on the 
    same page as the following paragraph whenever possible.
    
    This means that if the contents of the current paragraph would normally be 
    completely rendered on a different page than the following paragraph 
    (because only one of the two paragraphs would fit on the remaining space on
    the first page), then both paragraphs shall be rendered on a single page. 
    This property can be chained between multiple paragraphs to ensure that all 
    paragraphs are rendered on a single page without any intervening page
    boundaries. If this is not possible the entire set of paragraphs that are 
    grouped together using this property would exceed a single page in any case, 
    then the set of "keep with next" paragraphs shall start on a new page, with 
    page breaks as needed afterwards.
    """

    line_height = ct.LineHeight('w:pPr/w:spacing')
    """Specifies the amount of vertical spacing between lines of text within 
    this paragraph.
    """

    space_after = ct.SpaceAfterParagraph('w:pPr/w:spacing')
    """Specifies the spacing that should be added after the last line in this 
    paragraph in the document in absolute units.
    """

    space_before = ct.SpaceBeforeParagraph('w:pPr/w:spacing')
    """Specifies the spacing that should be added above the first line in this 
    paragraph in the document in absolute units."""

    numbering_instance_id = ct.Integer('w:pPr/w:numPr/w:numId')
    """Numbering definition instance for the numbered paragraph"""

    numbering_instance_level = ct.String('w:pPr/w:numPr/w:ilvl')
    """Numbering level of the numbering definition instance"""

    page_break_before: bool = ct.Boolean('w:pPr/w:pageBreakBefore')
    """This element specifies that when rendering this document in a paginated 
    view, the contents of this paragraph are rendered on the start of a new page 
    in the document.
    
    This means that if the contents of the current paragraph would normally be 
    rendered on the middle of a page in the host document, then the paragraph 
    shall be rendered on a new page as if the paragraph was preceded by a page 
    break in the WordprocessingML contents of the document. This property 
    supersedes any use of the keepNext property, so that if any paragraph wishes 
    to be on the same page as this paragraph, they are still be separated by a 
    page break.
    """

    style: str = ct.String('w:pPr/w:pStyle')
    """This element specifies the style ID of the paragraph style which shall be 
    used to format the contents of this paragraph. 
    """

    text_align = ct.Justification('w:pPr/w:jc')
    """This element specifies the paragraph alignment which shall be applied to 
    text in this paragraph.
    """

    text_indent: CssUnit = ct.TextIndent('w:pPr/w:ind')
    """Specifies the additional indentation which shall be applied to, or 
    removed from, the first line of the parent paragraph.
    """

    widows_control: bool = ct.Boolean('w:pPr/w:widowControl')
    """This element specifies whether a consumer shall prevent a single line of 
    this paragraph from being displayed on a separate page from the remaining 
    content at display time by moving the line onto the following page.
    
    When displaying a paragraph in a page, it is sometimes the case that the 
    first line of that paragraph would display as the last line on one page, 
    and all subsequent lines would display on the following page. This property
    ensures that a consumer shall move the single line to the following page as 
    well to prevent having one line on its own page. As well, if a single line 
    appears at the top of a page, a consumer shall move the preceding line onto
    the following page as well, to prevent a single line from being displayed on 
    a separate page.
    """


class DocxCharacterStyle(DocxStyle, RPrMixin):
    pass


class DocxParagraphStyle(DocxStyle, PPrMixin, RPrMixin):
    pass


class DocxNumberingStyle(DocxStyle, PPrMixin):
    pass


@wordml('docDefaults')
class DocDefaults(RPrMixin, PPrMixin, etree.ElementBase):

    @property
    def styles(self):
        return getattr(self, '_styles', None)

    @styles.setter
    def styles(self, styles):
        setattr(self, '_styles', styles)


class DocxPropertyAdapter(etree.ElementBase):

    def get_boolean_attribute(self, name):
        """Get the boolean value of an attribute or None if the
        attribute doesn't exist.
        """
        attribute_value = self.get(w(name))
        if attribute_value is not None:
            return not attribute_value.lower() in ('false', '0')
        else:
            return None

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


class ColorPropertyAdapter(DocxPropertyAdapter):

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


@wordml('bdr')
class BorderProperty(ColorPropertyAdapter):
    color_attribute = 'color'

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
    pass


@wordml('left')
class BorderLeftProperty(BorderProperty):
    pass


@wordml('right')
class BorderRightProperty(BorderProperty):
    pass


@wordml('top')
class BorderTopProperty(BorderProperty):
    pass


@wordml('insideH')
class BorderInsideHorizontalProperty(BorderProperty):
    pass


@wordml('insideV')
class BorderInsideVerticalProperty(BorderProperty):
    pass


@wordml('color')
class FontColorProperty(ColorPropertyAdapter):
    pass


@wordml('rFonts')
class FontProperty(DocxPropertyAdapter):

    def get_theme_font_or_font_value(self, font_name):
        """Get the theme font associated with the theme or return the
        same value if it's not a theme color
        """
        theme = self.get_theme()
        if font_name in theme.fonts:
            font_name = theme.get_font(font_name)
        return font_name

    def get_font_from_font_table(self, font_name):
        font_table = self.get_font_table()
        font = font_table.get_font(font_name)
        if font is not None:
            return font.css_family
        return font_name,


@wordml('shd')
class ShadingProperty(ColorPropertyAdapter):
    color_attribute = 'fill'
    theme_color_attribute = 'themeFill'
    theme_shade_attribute = 'themeFillShade'
    theme_tint_attribute = 'themeFillTint'


@wordml('u')
class UnderlineProperty(ColorPropertyAdapter):
    color_attribute = 'color'
