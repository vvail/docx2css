from abc import ABC, abstractmethod
from dataclasses import fields
import logging
import re
import warnings

from docx2css import api
from docx2css.api import (
    Counter,
    CounterList,
    PageStyle,
    ParagraphFormatting,
    TextFormatting,
)
from docx2css.ooxml.package import OpcPackage
from docx2css.ooxml.simple_types import ST_NumberFormat, ST_Jc
from docx2css.stylesheet import Stylesheet

logger = logging.getLogger(__name__)


class ParserFactory:

    def __init__(self):
        self.__block_parsers = {}
        self.__property_parsers = {}

    def register(self, property_name, parser_class):
        self.__property_parsers[property_name] = parser_class

    def register_block_parser(self, block_class, block_parser_class):
        self.__block_parsers[block_class] = block_parser_class

    def get_block_parser(self, block):
        block_class = block.__class__
        creator = self.__block_parsers.get(block_class)
        if not creator:
            raise ValueError(f'No converter registered for "{block_class}"')
        return creator(block)

    def get_property_parser(self, property_name):
        creator = self.__property_parsers.get(property_name)
        if not creator:
            msg = f'No parser registered for "{property_name}"'
            warnings.warn(msg)
        return creator


DocxParserFactory = ParserFactory()


class DocxParser:

    def __init__(self, docx_filename, factory=None):
        self.factory = factory or DocxParserFactory
        self.opc_package = OpcPackage(docx_filename)
        self.__stylesheet = Stylesheet()
        self.__counter_definitions = []
        self.__paragraph_counters = {}

    def get_counter_for_paragraph(self, paragraph_id):
        return self.__paragraph_counters.get(paragraph_id, None)

    def parse_partial_table(self, xml_element, style):
        """Parse a table style or a table conditional formatting element"""
        table_properties = tuple(f[0] for f in style.table_properties(False))
        self.parse_xml_style(xml_element, style, table_properties)

    @classmethod
    def normalize_table_id(cls, xml_element_id):
        if xml_element_id == 'TableNormal':
            return ''
        else:
            return xml_element_id

    def parse_docx_table_style(self, xml_element):
        if xml_element.type != 'table':
            return
        logger.debug(f'Parsing table style "{xml_element.name}"')
        style = api.TableStyle(
            name=xml_element.name,
            id=self.normalize_table_id(xml_element.id),
            parent_id=self.normalize_table_id(xml_element.parent_id),
        )

        self.parse_partial_table(xml_element, style)

        formats = style.table_conditional_formatting_properties(False)
        for format_name, _ in formats:
            xml_conditional_format = getattr(xml_element, format_name)
            if xml_conditional_format is None:
                continue
            conditional_format = api.TableConditionalFormatting()
            self.parse_partial_table(xml_conditional_format, conditional_format)
            setattr(style, format_name, conditional_format)

        self.__stylesheet.add_style(style)
        return style

    def parse_abstract_numbering(self, xml_element):
        names = (xml_element.name, xml_element.style_link)
        default_name = f'counter{xml_element.id}'
        name = next((x for x in names if x is not None), default_name)
        name = ''.join(name.split())
        counter_definition = CounterList(
            id=xml_element.id,
            name=name,
        )
        for xml_level in xml_element.levels.values():
            counter = self.parse_level(xml_level, counter_definition)
            counter_definition.counters[counter.name] = counter

            if xml_level.paragraph_style:
                self.__paragraph_counters[xml_level.paragraph_style] = counter
        return counter_definition

    # def parse_level(self, xml_element, counter_list):
    #     counter_format = ST_NumberFormat.css_value(xml_element.number_format)
    #     if xml_element.is_legal_format:
    #         counter_format = 'decimal'
    #
    #     level = LevelDefinition(
    #         counter_list=counter_list,
    #         number=xml_element.level_number,
    #         format=counter_format,
    #         start=xml_element.level_start,
    #         text=xml_element.level_text,
    #         paragraph_style=xml_element.paragraph_style,
    #         justification=ST_Jc.css_value(xml_element.justification),
    #         suffix=xml_element.level_suffix,
    #     )
    #     # Level restart logic: If value is 0 or higher than level, the
    #     # counter never restarts. If there is no value, it restarts at
    #     # previous level
    #     restart = xml_element.level_restart
    #     if restart != 0:
    #         if restart is None:
    #             restart = level.number - 1
    #         else:
    #             # Account for the fact that the xml value is one-based
    #             restart -= 1
    #         previous = counter_list.levels.get(restart, None)
    #         if previous is not None:
    #             previous.restart.add(level.number)
    #
    #     # self.parse_descendants(xml_element, level)
    #     props = tuple(f.name for f in fields(ParagraphFormatting))
    #     self.parse_xml_style(xml_element, level, props)
    #
    #     return level

    def parse_level(self, xml_element, counter_definition):
        number = xml_element.level_number
        name = f'{counter_definition.name}-L{number}'

        counter_format = ST_NumberFormat.css_value(xml_element.number_format)
        if xml_element.is_legal_format:
            counter_format = 'decimal'

        text = ''
        tokens = re.split('(%\\d)', xml_element.level_text)
        for token in (t for t in tokens if t):
            regex = re.match('%(\\d)', token)
            if regex:
                level_number = int(regex.group(1)) - 1
                text += f'{{{counter_definition.name}-L{level_number}}}'
            else:
                text += token

        counter = Counter(
            counter_list=counter_definition,
            name=name,
            style=counter_format,
            start=xml_element.level_start,
            text=text,
            suffix=xml_element.level_suffix,
            justification=ST_Jc.css_value(xml_element.justification) or 'start',
        )
        # Level restart logic: If value is 0 or higher than level, the
        # counter never restarts. If there is no value, it restarts at
        # previous level
        restart = xml_element.level_restart
        if restart != 0:
            if restart is None:
                restart = number - 1
            else:
                # Account for the fact that the xml value is one-based
                restart -= 1
            # previous = counter_list.levels.get(restart, None)
            previous_name = f'{counter_definition.name}-L{restart}'
            previous = counter_definition.counters.get(previous_name)
            if previous is not None:
                restart_add = name
                if counter.start != 1:
                    restart_add += f' {counter.start - 1}'
                previous.restart.add(restart_add)

        props = tuple(f.name for f in fields(ParagraphFormatting))
        self.parse_xml_style(xml_element, counter, props)

        return counter

    def parse_xml_style(self, xml_element, style, properties):
        for prop in properties:
            parser_class = self.factory.get_property_parser(prop)
            if parser_class:
                parser = parser_class(self)
                parser.parse(xml_element, style)

    def parse_docx_doc_defaults(self, doc_defaults):
        style = api.BodyStyle()
        props = tuple(x.name for x in fields(style))
        self.parse_xml_style(doc_defaults, style, props)
        self.__stylesheet.body_style = style
        return style

    def parse_docx_character_style(self, docx_style):
        style = api.SpanStyle(
            name=docx_style.name,
            id=docx_style.id,
            parent_id=docx_style.parent_id,
        )
        props = tuple(f.name for f in fields(TextFormatting))
        self.parse_xml_style(docx_style, style, props)
        self.__stylesheet.add_style(style)
        return style

    def get_or_create_paragraph_style(self, style_id, style_name, parent_id):
        style = self.__stylesheet.paragraph_styles.get(style_id)
        if style is None:
            style = api.ParagraphStyle(
                name=style_name,
                id=style_id,
                parent_id=parent_id,
            )
        else:
            style.name = style_name
            style.id = style_id
            style.parent_id = parent_id

        # Create basic parent style when the parent style has not been
        # parsed yet. This can occur when styles are defined out of
        # order, that is, a child is defined before its parent
        if parent_id and parent_id not in self.__stylesheet.paragraph_styles:
            self.get_or_create_paragraph_style(parent_id, parent_id, None)
        self.__stylesheet.add_style(style)
        return style

    @classmethod
    def normalize_paragraph_id(cls, paragraph_id):
        if paragraph_id == 'Normal':
            return ''
        else:
            return paragraph_id

    def parse_docx_paragraph_style(self, docx_style):
        style_name = self.normalize_paragraph_id(docx_style.name)
        style_id = self.normalize_paragraph_id(docx_style.id)
        parent_id = self.normalize_paragraph_id(docx_style.parent_id)
        style = self.get_or_create_paragraph_style(style_id, style_name, parent_id)
        props = tuple(f.name for f in fields(ParagraphFormatting))
        self.parse_xml_style(docx_style, style, props)

        return style

    def parse_docx_style(self, docx_style):
        if docx_style.type == 'character':
            return self.parse_docx_character_style(docx_style)
        elif docx_style.type == 'paragraph':
            return self.parse_docx_paragraph_style(docx_style)
        elif docx_style.type == 'table':
            return self.parse_docx_table_style(docx_style)

    def parse(self):
        self.__stylesheet.page_style = self.parse_page_style()
        self.parse_numbering()
        doc_defaults = self.opc_package.styles.doc_defaults
        self.__stylesheet.body_style = self.parse_docx_doc_defaults(doc_defaults)
        for docx_style in self.opc_package.styles.values():
            self.parse_docx_style(docx_style)
        return self.__stylesheet

    def parse_numbering(self):
        """Parse all numbering instances of the document"""
        for numbering in self.opc_package.numbering.values():
            counter_definition = self.parse_abstract_numbering(numbering)
            self.__counter_definitions.append(counter_definition)

    def parse_page_style(self):
        section = self.opc_package.sections[-1]
        style = PageStyle(
            margin_bottom=section.margin_bottom,
            margin_left=section.margin_left,
            margin_right=section.margin_right,
            margin_top=section.margin_top,
            page_height=section.page_height,
            page_orientation=section.page_orientation or 'portrait',
            page_width=section.page_width,
        )
        return style


class DocxPropertyParser(ABC):

    def __init__(self, docx_parser: DocxParser):
        self.docx_parser = docx_parser

    @abstractmethod
    def parse(self, xml_element, api_element):
        pass


class SimplePropertyParser(DocxPropertyParser):

    @property
    @abstractmethod
    def property_name(self):
        pass

    def parse(self, xml_element, api_element):
        value = getattr(xml_element, self.property_name)
        setattr(api_element, self.property_name, value)


class AllCapsParser(SimplePropertyParser):
    property_name = 'all_caps'


class BackgroundColorParser(SimplePropertyParser):
    property_name = 'background_color'


class BoldParser(SimplePropertyParser):
    property_name = 'bold'


class BorderParser(SimplePropertyParser):
    property_name = 'border'


class DoubleStrikeParser(SimplePropertyParser):
    property_name = 'double_strike'


class EmbossParser(SimplePropertyParser):
    property_name = 'emboss'


class FontColorParser(SimplePropertyParser):
    property_name = 'font_color'


class FontFamilyParser(SimplePropertyParser):
    property_name = 'font_family'


class FontKerningParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        kerning = xml_element.font_kerning
        if kerning is not None:
            api_element.font_kerning = kerning != 0


class FontSizeParser(SimplePropertyParser):
    property_name = 'font_size'


class HighlightParser(SimplePropertyParser):
    property_name = 'highlight'


class ImprintParser(SimplePropertyParser):
    property_name = 'imprint'


class ItalicsParser(SimplePropertyParser):
    property_name = 'italics'


class LetterSpacingParser(SimplePropertyParser):
    property_name = 'letter_spacing'


class OutlineParser(SimplePropertyParser):
    property_name = 'outline'


class PositionParser(SimplePropertyParser):
    property_name = 'position'


class ShadowParser(SimplePropertyParser):
    property_name = 'shadow'


class SmallCapsParser(SimplePropertyParser):
    property_name = 'small_caps'


class StrikeParser(SimplePropertyParser):
    property_name = 'strike'


class UnderlineParser(SimplePropertyParser):
    property_name = 'underline'


class VerticalAlignParser(SimplePropertyParser):
    property_name = 'vertical_align'


class VisibleParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.visible = xml_element.vanish


DocxParserFactory.register('all_caps',          AllCapsParser)
DocxParserFactory.register('background_color', BackgroundColorParser)
DocxParserFactory.register('bold',              BoldParser)
DocxParserFactory.register('border',            BorderParser)
DocxParserFactory.register('double_strike',     DoubleStrikeParser)
DocxParserFactory.register('emboss',            EmbossParser)
DocxParserFactory.register('font_color',        FontColorParser)
DocxParserFactory.register('font_family',       FontFamilyParser)
DocxParserFactory.register('font_kerning',      FontKerningParser)
DocxParserFactory.register('font_size',         FontSizeParser)
DocxParserFactory.register('highlight',         HighlightParser)
DocxParserFactory.register('imprint',           ImprintParser)
DocxParserFactory.register('italics',           ItalicsParser)
DocxParserFactory.register('letter_spacing',    LetterSpacingParser)
DocxParserFactory.register('outline',           OutlineParser)
DocxParserFactory.register('position',          PositionParser)
DocxParserFactory.register('shadow',            ShadowParser)
DocxParserFactory.register('small_caps',        SmallCapsParser)
DocxParserFactory.register('strike',            StrikeParser)
DocxParserFactory.register('underline',         UnderlineParser)
DocxParserFactory.register('vertical_align',    VerticalAlignParser)
DocxParserFactory.register('visible',           VisibleParser)


########################################################################
#                                                                      #
# Paragraph Formatting Parsers                                         #
#                                                                      #
########################################################################

class BorderBottomParser(SimplePropertyParser):
    property_name = 'border_bottom'


class BorderLeftParser(SimplePropertyParser):
    property_name = 'border_left'


class BorderTopParser(SimplePropertyParser):
    property_name = 'border_top'


class BorderRightParser(SimplePropertyParser):
    property_name = 'border_right'


class CounterParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        instance_id = xml_element.numbering_instance_id
        instance_level = xml_element.numbering_instance_level
        if instance_id is None and instance_level is None:
            return
        counter = self.docx_parser.get_counter_for_paragraph(api_element.id)
        api_element.counter = counter


class KeepTogetherParser(SimplePropertyParser):
    property_name = 'keep_together'


class KeepWithNextParser(SimplePropertyParser):
    property_name = 'keep_with_next'


class LineHeightParser(SimplePropertyParser):
    property_name = 'line_height'


class MarginLeftParser(SimplePropertyParser):
    property_name = 'margin_left'


class MarginRightParser(SimplePropertyParser):
    property_name = 'margin_right'


class MarginBottomParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        value = xml_element.space_after
        api_element.margin_bottom = value


class MarginTopParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        value = xml_element.space_before
        api_element.margin_top = value


class PageBreakBeforeParser(SimplePropertyParser):
    property_name = 'page_break_before'


class TextAlignParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.text_align = ST_Jc.css_value(xml_element.text_align)


class TextIndentParser(SimplePropertyParser):
    property_name = 'text_indent'


class WidowsParser(SimplePropertyParser):
    property_name = 'widows_control'


DocxParserFactory.register('border_bottom',     BorderBottomParser)
DocxParserFactory.register('border_left',       BorderLeftParser)
DocxParserFactory.register('border_top',        BorderTopParser)
DocxParserFactory.register('border_right',      BorderRightParser)
DocxParserFactory.register('counter',           CounterParser)
DocxParserFactory.register('keep_together',     KeepTogetherParser)
DocxParserFactory.register('keep_with_next',    KeepWithNextParser)
DocxParserFactory.register('line_height',       LineHeightParser)
DocxParserFactory.register('margin_bottom',     MarginBottomParser)
DocxParserFactory.register('margin_left',       MarginLeftParser)
DocxParserFactory.register('margin_right',      MarginRightParser)
DocxParserFactory.register('margin_top',      MarginTopParser)
DocxParserFactory.register('page_break_before', PageBreakBeforeParser)
DocxParserFactory.register('text_align',        TextAlignParser)
DocxParserFactory.register('text_indent',       TextIndentParser)
DocxParserFactory.register('widows_control',    WidowsParser)


########################################################################
#                                                                      #
# Table Formatting Parsers                                             #
#                                                                      #
########################################################################

class AlignmentParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.alignment = ST_Jc.css_value(xml_element.justification)


class BorderInsideHorizontalParser(SimplePropertyParser):
    property_name = 'border_inside_horizontal'


class BorderInsideVerticalParser(SimplePropertyParser):
    property_name = 'border_inside_vertical'


class CellPaddingBottomParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.cell_padding_bottom = xml_element.cell_margin_bottom


class CellPaddingLeftParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.cell_padding_left = xml_element.cell_margin_left


class CellPaddingRightParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.cell_padding_right = xml_element.cell_margin_right


class CellPaddingTopParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.cell_padding_top = xml_element.cell_margin_top


class CellSpacingParser(SimplePropertyParser):
    property_name = 'cell_spacing'


class ColBandSizeParser(SimplePropertyParser):
    property_name = 'col_band_size'


class ColSpanParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.colspan = xml_element.grid_span


class DefaultCellParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        default_cell = api.TableCellProperties()
        api_element.default_cell = default_cell
        cell_properties = list(default_cell.table_cell_properties(False))
        for prop_name, _ in cell_properties:
            parser_class = self.docx_parser.factory.get_property_parser(prop_name)
            if parser_class:
                parser = parser_class(self.docx_parser)
                parser.parse(xml_element.cell_properties, default_cell)


class FitTextParser(SimplePropertyParser):
    property_name = 'fit_text'


class IndentParser(SimplePropertyParser):
    property_name = 'indent'


class LayoutParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        layout = xml_element.layout
        if layout == 'autofit':
            layout = 'auto'
        api_element.layout = layout


class PaddingBottomParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.padding_bottom = xml_element.margin_bottom


class PaddingLeftParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.padding_left = xml_element.margin_left


class PaddingRightParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.padding_right = xml_element.margin_right


class PaddingTopParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        api_element.padding_top = xml_element.margin_top


class RowBandSizeParser(SimplePropertyParser):
    property_name = 'row_band_size'


class VAlignParser(SimplePropertyParser):
    property_name = 'valign'


class WidthParser(SimplePropertyParser):
    property_name = 'width'


class WrapTextParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        no_wrap = xml_element.no_wrap
        if no_wrap is not None:
            api_element.wrap_text = not no_wrap


class RowPropertiesParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        default_row = api.TableRowProperties()
        api_element.default_row = default_row
        for prop_name, _ in default_row.table_row_properties(False):
            parser_class = self.docx_parser.factory.get_property_parser(prop_name)
            if parser_class:
                parser = parser_class(self.docx_parser)
                parser.parse(xml_element.row_properties, default_row)


class RowHeightParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        if xml_element.height_type == 'exact':
            api_element.height = xml_element.height


class RowIsHeaderParser(SimplePropertyParser):
    property_name = 'is_header'


class RowMinHeightParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        height_type = xml_element.height_type
        if height_type is None or height_type == 'atLeast':
            api_element.min_height = xml_element.height


class RowSplitParser(DocxPropertyParser):

    def parse(self, xml_element, api_element):
        cant_split = xml_element.cant_split
        if cant_split is not None:
            api_element.split = not cant_split


DocxParserFactory.register('alignment',                 AlignmentParser)
DocxParserFactory.register('border_inside_horizontal',  BorderInsideHorizontalParser)
DocxParserFactory.register('border_inside_vertical',    BorderInsideVerticalParser)
DocxParserFactory.register('cell_padding_bottom',       CellPaddingBottomParser)
DocxParserFactory.register('cell_padding_left',         CellPaddingLeftParser)
DocxParserFactory.register('cell_padding_right',        CellPaddingRightParser)
DocxParserFactory.register('cell_padding_top',          CellPaddingTopParser)
DocxParserFactory.register('cell_spacing',              CellSpacingParser)
DocxParserFactory.register('col_band_size',             ColBandSizeParser)
DocxParserFactory.register('colspan',                   ColSpanParser)
DocxParserFactory.register('default_cell',              DefaultCellParser)
DocxParserFactory.register('default_row',               RowPropertiesParser)
DocxParserFactory.register('fit_text',                  FitTextParser)
DocxParserFactory.register('height',                    RowHeightParser)
DocxParserFactory.register('indent',                    IndentParser)
DocxParserFactory.register('is_header',                 RowIsHeaderParser)
DocxParserFactory.register('layout',                    LayoutParser)
DocxParserFactory.register('min_height',                RowMinHeightParser)
DocxParserFactory.register('padding_bottom',            PaddingBottomParser)
DocxParserFactory.register('padding_left',              PaddingLeftParser)
DocxParserFactory.register('padding_right',             PaddingRightParser)
DocxParserFactory.register('padding_top',               PaddingTopParser)
DocxParserFactory.register('row_band_size',             RowBandSizeParser)
DocxParserFactory.register('split',                     RowSplitParser)
DocxParserFactory.register('valign',                    VAlignParser)
DocxParserFactory.register('width',                     WidthParser)
DocxParserFactory.register('wrap_text',                 WrapTextParser)
