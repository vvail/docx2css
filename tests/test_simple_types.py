from unittest import TestCase

from docx2css.ooxml.simple_types import TwoWayDict


class TestTwoWayDict(TestCase):

    class MyTwoWayDict(TwoWayDict):
        docx2css = {'my_docx_value': 'my_css_value'}

    def test_css_value(self):
        self.assertEqual('my_css_value', self.MyTwoWayDict.css_value('my_docx_value'))

    def test_docx_value(self):
        self.assertEqual('my_docx_value', self.MyTwoWayDict.docx_value('my_css_value'))
