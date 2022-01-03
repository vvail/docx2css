from unittest import TestCase

import cssutils

from docx2css.ooxml.numbering import AbstractNumbering
from docx2css.ooxml.package import OpcPackage
from docx2css.ooxml.styles import NumberingProperty


cssutils.ser.prefs.indentClosingBrace = False
cssutils.ser.prefs.omitLastSemicolon = False


class NumberingTestCase(TestCase):

    def setUp(self):
        parser = OpcPackage('test_files/numbering/docx/requete.docx')
        self.numbering = parser.get_numbering()
        self.styles = parser.styles

    def test_abstract_numbering(self):
        self.assertEqual(6, len(self.numbering.abstract_numbering))

    def test_numbering_instances(self):
        for num in self.numbering.numbering_instances.values():
            self.assertIsInstance(num, AbstractNumbering)

    def test_style_link(self):
        resolutions = self.numbering.abstract_numbering[0]
        self.assertEqual('resolutions', resolutions.style_link)
        num2 = self.numbering.abstract_numbering[2]
        self.assertIsNone(num2.style_link)

    def test_paragraph_style_numbering(self):
        p = self.styles['allegations-L1']
        self.assertIsInstance(p.numbering, NumberingProperty)
        self.assertEqual(1, p.numbering.level)

    def test_levels(self):
        resolutions = self.numbering.abstract_numbering[0]
        self.assertEqual(9, len(resolutions.levels))

    def test_level_start(self):
        resolutions = self.numbering.abstract_numbering[0]
        level = resolutions.levels[0]
        self.assertEqual(1, level.start)

    def test_level_justification(self):
        resolutions = self.numbering.abstract_numbering[0]
        level = resolutions.levels[0]
        self.assertEqual('left', level.justification)

    def test_css_counter_name_resolutions(self):
        resolutions = self.numbering.abstract_numbering[0]
        level = resolutions.levels[0]
        self.assertEqual('resolutions-L0', level.css_counter_name())

    def test_css_counter_name_unnamed(self):
        resolutions = self.numbering.abstract_numbering[2]
        level = resolutions.levels[0]
        self.assertEqual('counter2-L0', level.css_counter_name())

    def test_css_counter_resets(self):
        resolutions = self.numbering.abstract_numbering[0]
        expected = {
            0: 'resolutions-L1',
            1: 'resolutions-L2',
            2: 'resolutions-L3',
            3: 'resolutions-L4',
            4: 'resolutions-L5',
            5: 'resolutions-L6',
            6: 'resolutions-L7',
            7: 'resolutions-L8',
            8: '',
        }
        for i, level in enumerate(resolutions.levels.values()):
            self.assertEqual(expected[i], level.css_counter_resets())

    def test_css_root_counters_requete(self):
        package = OpcPackage('test_files/numbering/docx/requete.docx')
        numbering = package.get_numbering()
        expected = 'allegations-L0 counter2-L0 counter3-L0 resolutions-L0'
        self.assertEqual(expected, numbering.css_root_counters())

    def test_css_root_counters_start_at_5(self):
        package = OpcPackage('test_files/numbering/docx/start_at_5.docx')
        numbering = package.get_numbering()
        expected = 'start-at-5-list-L0 start-at-5-list-L2'
        self.assertEqual(expected, numbering.css_root_counters())


class NumberingFormatTestCase(TestCase):
    files = (
        'test_files/numbering/docx/first_formats.docx',
        'test_files/numbering/docx/second_formats.docx',
        'test_files/numbering/docx/third_formats.docx',
    )

    def setUp(self):
        self.numbering = []
        for file in self.files:
            parser = OpcPackage(file)
            self.numbering.append(parser.get_numbering())

    def test_level_format(self):
        for num in self.numbering:
            for level in num.abstract_numbering[0].levels.values():
                print(f"'{level.format}'")
            print('----')


class RestartNumberingTestCase(TestCase):

    def setUp(self):
        parser = OpcPackage('test_files/numbering/docx/restart_numbering.docx')
        self.numbering = parser.get_numbering().abstract_numbering

    def test_restart_numbering(self):
        num = self.numbering[1]
        expected = {
            0: 'mylist-L1 mylist-L2 mylist-L3',
            1: '',
            2: '',
            3: 'mylist-L4',
            4: 'mylist-L5',
            5: 'mylist-L6',
            6: 'mylist-L7',
            7: 'mylist-L8',
            8: '',
        }
        for level in num.levels.values():
            self.assertEqual(expected[level.level], level.css_counter_resets())


class CounterContentTestCase(TestCase):

    def setUp(self):
        self.parser = OpcPackage('test_files/numbering/docx/requete.docx')

    def test_counter_name_allegations(self):
        style = self.parser.styles['allegations']
        numbering = style.numbering.definition
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
        for level in numbering.levels.values():
            self.assertEqual(expected[level.level],
                             level.css_counter_name(),
                             f'Level {level.level}')

    def test_counter_contents_allegations(self):
        style = self.parser.styles['allegations']
        numbering = style.numbering.definition
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
        for level in numbering.levels.values():
            self.assertEqual(expected[level.level],
                             level.css_counter_content(),
                             f'Level {level.level}')

    def test_counter_contents_resolutions(self):
        style = self.parser.styles['resolutions']
        numbering = style.numbering.definition
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
        for level in numbering.levels.values():
            self.assertEqual(expected[level.level],
                             fr'{level.css_counter_content()}',
                             f'Level {level.level}')

    def test_counter_format_legal(self):
        parser = OpcPackage('test_files/numbering/docx/legal_numbering.docx')
        numbering = parser.get_numbering().abstract_numbering[0]
        for i in range(2):
            self.assertTrue('decimal' in numbering.levels[i].css_counter())

    def test_counter_name_from_name(self):
        parser = OpcPackage('test_files/numbering/docx/with_names.docx')
        numbering = parser.get_numbering().abstract_numbering[0]
        expected = 'fieldlistname-L0'
        actual = numbering.levels[0].css_counter_name()
        self.assertEqual(expected, actual)

    def test_counter_name_from_style_link(self):
        style = self.parser.styles['resolutions']
        numbering = style.numbering.definition
        expected = 'resolutions-L0'
        actual = numbering.levels[0].css_counter_name()
        self.assertEqual(expected, actual)

    def test_counter_name_from_id(self):
        parser = OpcPackage('test_files/numbering/docx/legal_numbering.docx')
        numbering = parser.get_numbering().abstract_numbering[0]
        expected = 'counter0-L0'
        actual = numbering.levels[0].css_counter_name()
        self.assertEqual(expected, actual)

    def compare_style(self, opc_package, style_name):
        style = opc_package.styles[style_name]
        with open(f'test_files/numbering/css/{style_name}.css', 'r') as css_file:
            expected = css_file.read()
            actual = style.css_numbering_style_rule().cssText + '\n'
            actual += style.css_style_rule().cssText
            self.assertEqual(expected, fr'{actual}')

    def test_paragraph_numbering_allegationsL1(self):
        self.compare_style(self.parser, 'allegations-L1')

    def test_paragraph_numbering_allegationsL2(self):
        self.compare_style(self.parser, 'allegations-L2')

    def test_paragraph_numbering_allegationsL3(self):
        self.compare_style(self.parser, 'allegations-L3')

    def test_paragraph_numbering_resolutionL1(self):
        self.compare_style(self.parser, 'resolution-L1')

    def test_paragraph_numbering_resolutionL2(self):
        self.compare_style(self.parser, 'resolution-L2')

    def test_paragraph_numbering_resolutionL3(self):
        self.compare_style(self.parser, 'resolution-L3')

    def test_paragraph_numbering_first_line(self):
        parser = OpcPackage('test_files/numbering/docx/first_line.docx')
        self.compare_style(parser, 'firstline')

    def test_paragraph_numbering_first_line_space(self):
        parser = OpcPackage('test_files/numbering/docx/first_line.docx')
        self.compare_style(parser, 'firstlinespace')

    def test_paragraph_numbering_first_line_nothing(self):
        parser = OpcPackage('test_files/numbering/docx/first_line.docx')
        self.compare_style(parser, 'firstlinenothing')

    def test_paragraph_numbering_hanging_tab(self):
        parser = OpcPackage('test_files/numbering/docx/hanging.docx')
        self.compare_style(parser, 'hangingtab')

    def test_paragraph_numbering_hanging_space(self):
        parser = OpcPackage('test_files/numbering/docx/hanging.docx')
        self.compare_style(parser, 'hangingspace')

    def test_paragraph_numbering_hanging_nothing(self):
        parser = OpcPackage('test_files/numbering/docx/hanging.docx')
        self.compare_style(parser, 'hangingnothing')

    def test_list_start_at_5(self):
        parser = OpcPackage('test_files/numbering/docx/start_at_5.docx')
        self.compare_style(parser, 'list-5-paragraph')

    def test_list_start_at_5_parent(self):
        parser = OpcPackage('test_files/numbering/docx/start_at_5.docx')
        self.compare_style(parser, 'list-5-parent-paragraph')
