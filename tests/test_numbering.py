from unittest import TestCase

import cssutils

from docx2css.css.serializers import CssStylesheetSerializer, FACTORY
from docx2css.ooxml.numbering import AbstractNumbering
from docx2css.ooxml.package import OpcPackage
from docx2css.ooxml.parsers import DocxParser
from tests.test_styles import TestHarness


cssutils.ser.prefs.indentClosingBrace = False
cssutils.ser.prefs.omitLastSemicolon = False


class CounterApiTestCase(TestCase):

    @classmethod
    def setUpClass(cls):
        cls.parser = DocxParser('test_files/numbering/docx/requete.docx')
        cls.numbering = cls.parser.opc_package.numbering
        cls.styles = cls.parser.opc_package.styles

    def test_abstract_num_id(self):
        xml_element = self.numbering[2]
        allegations = self.parser.parse_abstract_numbering(xml_element)
        self.assertEqual(5, allegations.id)
        self.assertEqual('allegations', allegations.name)
        self.assertEqual(9, len(allegations.counters))

    def test_allegations_l0(self):
        xml_element = self.numbering[2]
        allegations = self.parser.parse_abstract_numbering(xml_element)
        l0 = allegations.counters['allegations-L0']
        self.assertEqual(1, l0.start)
        self.assertEqual('none', l0.style)
        self.assertFalse(l0.bold)
        self.assertFalse(l0.italics)
        self.assertFalse(l0.all_caps)
        self.assertEqual('{allegations-L0}', l0.text)
        self.assertEqual(0, l0.text_indent)
        l1 = allegations.counters['allegations-L1']
        self.assertEqual(-0.5, l1.text_indent.inches)


class NumberingTestCase(TestCase):

    def setUp(self):
        docx_package = OpcPackage('test_files/numbering/docx/requete.docx')
        self.numbering = docx_package.numbering

    def test_abstract_numbering(self):
        self.assertEqual(6, len(self.numbering))

    def test_numbering_instances(self):
        for num in self.numbering.values():
            self.assertIsInstance(num, AbstractNumbering)

    def test_style_link(self):
        resolutions = self.numbering[3]
        self.assertEqual('resolutions', resolutions.style_link)

    def test_style_link_is_none(self):
        num2 = self.numbering[1]
        self.assertIsNone(num2.style_link)


class AbstractNumberingTestCase(TestCase):

    def setUp(self):
        docx_package = OpcPackage('test_files/numbering/docx/requete.docx')
        self.numbering = docx_package.numbering

    def test_id(self):
        resolutions = self.numbering[3]
        self.assertEqual(0, resolutions.id)
        self.assertEqual('resolutions', resolutions.style_link)

    def test_levels(self):
        resolutions = self.numbering[3]
        self.assertEqual(9, len(resolutions.levels))

    def test_resolve_style_links(self):
        self.assertEqual(2, self.numbering[1].id)
        self.assertEqual(5, self.numbering[2].id)
        self.assertEqual(0, self.numbering[3].id)
        self.assertEqual(3, self.numbering[4].id)
        self.assertEqual(5, self.numbering[5].id)
        self.assertEqual(5, self.numbering[6].id)


class LevelTestCase(TestCase):

    def setUp(self):
        parser = DocxParser('test_files/numbering/docx/requete.docx')
        numbering = parser.opc_package.numbering[3]
        self.resolutions = parser.parse_abstract_numbering(numbering)

    def test_start(self):
        level = self.resolutions.counters['resolutions-L0']
        self.assertEqual(1, level.start)

    def test_suffix(self):
        level = self.resolutions.counters['resolutions-L0']
        self.assertEqual('tab', level.suffix)

    def test_level_justification(self):
        level = self.resolutions.counters['resolutions-L0']
        self.assertEqual('start', level.justification)

    def test_counter_restart(self):
        expected = {
            'resolutions-L0': {'resolutions-L1'},
            'resolutions-L1': {'resolutions-L2'},
            'resolutions-L2': {'resolutions-L3'},
            'resolutions-L3': {'resolutions-L4'},
            'resolutions-L4': {'resolutions-L5'},
            'resolutions-L5': {'resolutions-L6'},
            'resolutions-L6': {'resolutions-L7'},
            'resolutions-L7': {'resolutions-L8'},
            'resolutions-L8': set(),
        }
        for i, level in self.resolutions.counters.items():
            self.assertEqual(expected[i], level.restart)

    def test_resolution_l1(self):
        level = self.resolutions.counters['resolutions-L0']
        self.assertEqual('resolutions-L0', level.name)
        self.assertEqual('', level.style)
        self.assertEqual('start', level.justification)
        self.assertEqual({'resolutions-L1'}, level.restart)
        self.assertEqual(1, level.start)
        self.assertEqual('tab', level.suffix)


class RestartNumberingTestCase(TestCase):

    def setUp(self):
        parser = DocxParser('test_files/numbering/docx/restart_numbering.docx')
        numbering = parser.opc_package.numbering[1]
        self.my_list = parser.parse_abstract_numbering(numbering)

    def test_restart_numbering(self):
        expected = {
            0: {'mylist-L1', 'mylist-L2', 'mylist-L3'},
            1: set(),
            2: set(),
            3: {'mylist-L4'},
            4: {'mylist-L5'},
            5: {'mylist-L6'},
            6: {'mylist-L7'},
            7: {'mylist-L8'},
            8: set(),
        }
        # expected = {
        #     0: {1, 2, 3},
        #     1: set(),
        #     2: set(),
        #     3: {4},
        #     4: {5},
        #     5: {6},
        #     6: {7},
        #     7: {8},
        #     8: set(),
        # }
        for i, level in enumerate(self.my_list.counters.values()):
            self.assertEqual(expected[i], level.restart)


class CounterContentTestCase(TestCase):

    def test_css_root_counters_requete(self):
        parser = DocxParser('test_files/numbering/docx/requete.docx')
        stylesheet = parser.parse()
        serializer = CssStylesheetSerializer(stylesheet)
        serializer.initialize_counters_in_body = False
        result = serializer.css_root_counters()
        expected = {'allegations-L0', 'counter3-L0', 'resolutions-L0'}
        self.assertEqual(expected, result)

    def test_css_all_counters_requete(self):
        parser = DocxParser('test_files/numbering/docx/requete.docx')
        stylesheet = parser.parse()
        serializer = CssStylesheetSerializer(stylesheet)
        result = serializer.css_root_counters()
        expected = {'allegations-L0', 'allegations-L1',
                    'allegations-L2', 'allegations-L3', 'counter3-L0',
                    'resolutions-L0', 'resolutions-L1', 'resolutions-L2'
                    }
        self.assertEqual(expected, result)

    def test_css_root_counters_start_at_5(self):
        parser = DocxParser('test_files/numbering/docx/start_at_5.docx')
        stylesheet = parser.parse()
        serializer = CssStylesheetSerializer(stylesheet)
        serializer.initialize_counters_in_body = False
        result = serializer.css_root_counters()
        expected = {'start-at-5-list-L0'}
        self.assertEqual(expected, result)


class ParagraphNumberingParserTestCase(TestCase):

    def test_heading1(self):
        parser = DocxParser('test_files/numbering/docx/requete.docx')
        stylesheet = parser.parse()
        # xml_element = parser.opc_package.styles['Heading1']
        style = stylesheet.paragraph_styles['Heading1']
        self.assertIsNotNone(style)
        counter = style.counter

        self.assertTrue(counter.bold is False)
        self.assertTrue(counter.italics is False)
        self.assertTrue(counter.all_caps is False)
        self.assertEqual(0, counter.margin_left)
        self.assertEqual(0, counter.text_indent)

        self.assertEqual('allegations-L0', counter.name)
        self.assertEqual(1, counter.start)
        self.assertEqual('{allegations-L0}', counter.text)
        self.assertEqual({'allegations-L1'}, counter.restart)
        self.assertEqual('tab', counter.suffix)
        self.assertEqual('start', counter.justification)

    def test_counter_name_from_name(self):
        parser = DocxParser('test_files/numbering/docx/with_names.docx')
        xml_numbering = parser.opc_package.numbering[1]
        definition = parser.parse_abstract_numbering(xml_numbering)
        expected = 'fieldlistname'
        self.assertEqual(expected, definition.name)

    def test_counter_name_from_style_link(self):
        parser = DocxParser('test_files/numbering/docx/requete.docx')
        xml_numbering = parser.opc_package.numbering[3]
        definition = parser.parse_abstract_numbering(xml_numbering)
        expected = 'resolutions'
        self.assertEqual(expected, definition.name)

    def test_counter_name_from_id(self):
        parser = DocxParser('test_files/numbering/docx/legal_numbering.docx')
        xml_numbering = parser.opc_package.numbering[1]
        definition = parser.parse_abstract_numbering(xml_numbering)
        expected = 'counter0'
        self.assertEqual(expected, definition.name)

    def test_counter_format_legal(self):
        parser = DocxParser('test_files/numbering/docx/legal_numbering.docx')
        numbering = parser.opc_package.numbering[1]
        abstract_numbering = parser.parse_abstract_numbering(numbering)
        for i in range(2):
            xml_level = numbering.levels[i]
            counter = parser.parse_level(xml_level, abstract_numbering)
            self.assertTrue(xml_level.is_legal_format)
            self.assertEqual('decimal', counter.style)


class NumberingSerializerTestCase(TestHarness):
    files = ('requete.docx', 'first_line.docx', 'hanging.docx', 'start_at_5.docx')
    css_files_location = 'test_files/numbering/css/'
    docx_files_location = 'test_files/numbering/docx/'
    parse_numbering_instances = True

    def get_counter_serializer(self, docx_filename, style_id):
        style = self.get_style(docx_filename, style_id)
        block_serializer = FACTORY.get_block_serializer(style)
        prop = ('counter', style.counter)
        return FACTORY.get_property_serializer(block_serializer, *prop)

    def get_style(self, docx_filename, style_id):
        parser = DocxParser(f'{self.docx_files_location}{docx_filename}')
        stylesheet = parser.parse()
        style = stylesheet.paragraph_styles[style_id]
        return style

    def test_css_counter_name_resolutions(self):
        serializer = self.get_counter_serializer('requete.docx', 'resolution-L1')
        self.assertEqual('resolutions-L0', serializer.css_counter_name())

    def test_counter_name_allegations(self):
        expected = {
            0: 'allegations-L0',
            1: 'allegations-L1',
            2: 'allegations-L2',
            3: 'allegations-L3',
            4: 'allegations-L4',
            5: 'allegations-L5',
            6: 'allegations-L6',
            7: 'allegations-L7',
            8: 'allegations-L8',
        }
        styles = ('Heading1', 'allegations-L1', 'allegations-L2', 'allegations-L3')
        for i, style in enumerate(styles):
            serializer = self.get_counter_serializer('requete.docx', style)
            self.assertEqual(expected[i], serializer.css_counter_name())

    def test_css_counter_resets(self):
        expected = {
            0: 'resolutions-L1',
            1: 'resolutions-L2',
            2: 'resolutions-L3',
        }
        styles = ('resolution-L1', 'resolution-L2', 'resolution-L3')
        for i, style in enumerate(styles):
            serializer = self.get_counter_serializer('requete.docx', style)
            self.assertEqual(expected[i], serializer.css_counter_resets())

    def test_css_counter_resets_allegations(self):
        expected = {
            0: 'allegations-L1',
            1: 'allegations-L2',
            2: 'allegations-L3',
            3: 'allegations-L4',
        }
        styles = ('Heading1', 'allegations-L1', 'allegations-L2', 'allegations-L3')
        for i, style in enumerate(styles):
            serializer = self.get_counter_serializer('requete.docx', style)
            self.assertEqual(expected[i], serializer.css_counter_resets())

    def test_counter_contents_allegations(self):
        expected = {
            0: None,
            1: 'counter(allegations-L1, decimal) "."',
            2: ('counter(allegations-L1, decimal) "." '
                'counter(allegations-L2, decimal) "."'),
            3: ('counter(allegations-L1, decimal) "." '
                'counter(allegations-L2, decimal) "." '
                'counter(allegations-L3, decimal) "."'),
            4: '"(" counter(allegations-L4, lower-alpha) ")"',
            5: '"(" counter(allegations-L5, lower-roman) ")"',
            6: 'counter(allegations-L6, decimal) "."',
            7: 'counter(allegations-L7, lower-alpha) "."',
            8: 'counter(allegations-L8, lower-roman) "."',
        }
        styles = ('Heading1', 'allegations-L1', 'allegations-L2', 'allegations-L3')
        for i, style in enumerate(styles):
            serializer = self.get_counter_serializer('requete.docx', style)
            self.assertEqual(expected[i], serializer.css_counter_content())

    def test_counter_contents_resolutions(self):
        expected = {
            0: r'"\005C f0b7"',
            1: r'"\005C 006f"',
            2: r'"\005C f0a7"',
            3: r'"\005C f0b7"',
            4: r'"\005C f0a8"',
            5: r'"\005C f0d8"',
            6: r'"\005C f0a7"',
            7: r'"\005C f0b7"',
            8: r'"\005C f0a8"',
        }
        styles = ('resolution-L1', 'resolution-L2', 'resolution-L3')
        for i, style in enumerate(styles):
            serializer = self.get_counter_serializer('requete.docx', style)
            self.assertEqual(expected[i], serializer.css_counter_content())

    def test_paragraph_numbering_allegationsL1(self):
        self.style_test_harness('allegations-L1')

    def test_paragraph_numbering_allegationsL2(self):
        self.style_test_harness('allegations-L2')

    def test_paragraph_numbering_allegationsL3(self):
        self.style_test_harness('allegations-L3')

    def test_paragraph_numbering_resolutionL1(self):
        self.style_test_harness('resolution-L1')

    def test_paragraph_numbering_resolutionL2(self):
        self.style_test_harness('resolution-L2')

    def test_paragraph_numbering_resolutionL3(self):
        self.style_test_harness('resolution-L3')

    def test_paragraph_numbering_first_line(self):
        self.style_test_harness('first_line')

    def test_paragraph_numbering_first_line_space(self):
        self.style_test_harness('first_line_space')

    def test_paragraph_numbering_first_line_nothing(self):
        self.style_test_harness('first_line_nothing')

    def test_paragraph_numbering_hanging_tab(self):
        self.style_test_harness('hanging_tab')

    def test_paragraph_numbering_hanging_space(self):
        self.style_test_harness('hanging_space')

    def test_paragraph_numbering_hanging_nothing(self):
        self.style_test_harness('hanging_nothing')

    def test_list_start_at_5(self):
        self.style_test_harness('list-5-paragraph')

    def test_list_start_at_5_parent(self):
        self.style_test_harness('list-5-parent-paragraph')
