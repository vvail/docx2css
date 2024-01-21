from lxml import etree

from docx2css.api import PageStyle
from docx2css.ooxml import w, wordml
from docx2css.utils import CssUnit


class Sections:

    def __init__(self, document_part):
        self.document = document_part
        self._sections = list(
            d for d in document_part.iterdescendants() if isinstance(d, Section)
        )

    def __getitem__(self, item):
        return self._sections[item]


class MarginDescriptor:

    def __set_name__(self, owner, name):
        self.direction = name.partition('_')[2]

    def __get__(self, instance, owner):
        margins = instance.find(w('pgMar'))
        return CssUnit(margins.get(w(self.direction)), 'twip')

    def __set__(self, instance, value):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError


class PageSizeDescriptor:

    def __set_name__(self, owner, name):
        self.property_name = name.partition('_')[2]

    def __get__(self, instance, owner):
        page_size = instance.find(w('pgSz'))
        return {
            'height': CssUnit(page_size.get(w('h')), 'twip'),
            'orientation': page_size.get(w('orient')),
            'width': CssUnit(page_size.get(w('w')), 'twip'),
        }[self.property_name]

    def __set__(self, instance, value):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError


@wordml('sectPr')
class Section(etree.ElementBase, PageStyle):

    margin_bottom = MarginDescriptor()
    margin_left = MarginDescriptor()
    margin_right = MarginDescriptor()
    margin_top = MarginDescriptor()
    page_height = PageSizeDescriptor()
    page_orientation = PageSizeDescriptor()
    page_width = PageSizeDescriptor()
