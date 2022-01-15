from lxml import etree

from docx2css.api import PageStyle
from docx2css.ooxml import w, wordml
from docx2css.ooxml.constants import NAMESPACES as NS
from docx2css.utils import CssUnit

INCLUDE_PAGE_RULE = 'include_page_rule'
SIMULATE_PRINTED_PAGE = 'simulate_printed_page'


class Sections:

    def __init__(self, document_part):
        self.document = document_part
        self._sections = list(
            d for d in document_part.iterdescendants() if isinstance(d, Section)
        )

    def __getitem__(self, item):
        return self._sections[item]


@wordml('sectPr')
class Section(etree.ElementBase, PageStyle):

    def _get_height(self) -> CssUnit:
        return self.page_size.height

    def _get_margin_bottom(self) -> CssUnit:
        return self.margins.bottom

    def _get_margin_left(self) -> CssUnit:
        return self.margins.left

    def _get_margin_right(self) -> CssUnit:
        return self.margins.right

    def _get_margin_top(self) -> CssUnit:
        return self.margins.top

    def _get_width(self) -> CssUnit:
        return self.page_size.width

    @property
    def margins(self):
        return self.find('.//w:pgMar', namespaces=NS)

    @property
    def page_size(self):
        return self.find('.//w:pgSz', namespaces=NS)


@wordml('pgMar')
class PageMargin(etree.ElementBase):

    def _get_margin(self, direction):
        return CssUnit(self.get(w(direction)), 'twip')

    @property
    def bottom(self):
        """Get the bottom margin value in inches"""
        return self._get_margin('bottom')

    @property
    def left(self):
        """Get the left margin value in inches"""
        return self._get_margin('left')

    @property
    def right(self):
        """Get the right margin value in inches"""
        return self._get_margin('right')

    @property
    def top(self):
        """Get the top margin in inches"""
        return self._get_margin('top')


@wordml('pgSz')
class PageSize(etree.ElementBase):

    @property
    def height(self):
        """
        Get the page height. The original value is in 20th of a pt.
        """
        height = self.get(w('h'))
        return CssUnit(height, 'twip')

    @property
    def orientation(self):
        """Get the orientation of the page (portrait or landscape)"""
        return self.get(w('orient')) or 'portrait'

    @property
    def width(self):
        """Get the page width.
        The original value is in 20th of a pt"""
        width = self.get(w('w'))
        return CssUnit(width, 'twip')
