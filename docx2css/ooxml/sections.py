import cssutils
from lxml import etree

from docx2css.ooxml import w, wordml
from docx2css.ooxml.constants import NAMESPACES as NS


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
class Section(etree.ElementBase):

    @property
    def margins(self):
        return self.find('.//w:pgMar', namespaces=NS)

    @property
    def page_size(self):
        return self.find('.//w:pgSz', namespaces=NS)

    def css_style_declaration_print(self):
        css_style = cssutils.css.CSSStyleDeclaration()
        self.page_size.set_css_style_print(css_style)
        self.margins.set_css_style_print(css_style)
        return css_style

    def css_style_rule_print(self):
        if not hasattr(self, '_css_style_rule_print'):
            css_style = self.css_style_declaration_print()
            page = cssutils.css.CSSPageRule(style=css_style)
            setattr(self, '_css_style_rule_print', page)
        return getattr(self, '_css_style_rule_print')

    def css_style_declaration_screen(self):
        css_style = cssutils.css.CSSStyleDeclaration()
        css_style['background-color'] = 'white'
        css_style['border'] = '1px darkgray solid'
        css_style['box-shadow'] = '1rem 0.5rem 1rem rgba(0,0,0,0.15)'
        self.page_size.set_css_style_screen(css_style)
        self.margins.set_css_style_screen(css_style)

        # Adjust max-width to margins
        max_width = self.page_size.width - self.margins.left - self.margins.right
        css_style['max-width'] = f'{max_width}in'

        return css_style

    def css_style_rule_screen(self):
        if not hasattr(self, '_css_style_rule_screen'):
            screen = cssutils.css.CSSMediaRule('screen')

            html_style = cssutils.css.CSSStyleDeclaration()
            html_style['background-color'] = 'gainsboro'
            html_rule = cssutils.css.CSSStyleRule('html', html_style)
            screen.add(html_rule)

            body_style = self.css_style_declaration_screen()
            body_rule = cssutils.css.CSSStyleRule('body', body_style)
            screen.add(body_rule)

            setattr(self, '_css_style_rule_screen', screen)
        return getattr(self, '_css_style_rule_screen')

    def css_style_rules(self, preferences=None):
        preferences = preferences or {}
        rules = []
        if preferences.get(INCLUDE_PAGE_RULE, True):
            rules.append(self.css_style_rule_print())
        if preferences.get(SIMULATE_PRINTED_PAGE, False):
            rules.append(self.css_style_rule_screen())

        return rules


@wordml('pgMar')
class PageMargin(etree.ElementBase):

    def _get_margin(self, direction):
        return int(self.get(w(direction))) / 20 / 72

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

    def set_css_style_print(self, style_rule):
        value = f'{self.top}in {self.right}in {self.bottom}in {self.left}in'
        style_rule['margin'] = value

    def set_css_style_screen(self, style_rule):
        style_rule['margin'] = '1em auto'
        value = f'{self.top}in {self.right}in {self.bottom}in {self.left}in'
        style_rule['padding'] = value


@wordml('pgSz')
class PageSize(etree.ElementBase):

    @property
    def height(self):
        """
        Get the page height in inches. The original value is in 20th of a pt.
        """
        height = self.get(w('h'))
        return int(height) / 20 / 72

    @property
    def orientation(self):
        """Get the orientation of the page (portrait or landscape)"""
        return self.get(w('orient')) or 'portrait'

    @property
    def width(self):
        """Get the page height in inches.
        The original value is in 20th of a pt"""
        width = self.get(w('w'))
        return int(width) / 20 / 72

    def set_css_style_print(self, style_rule):
        style_rule['size'] = f'{self.width}in {self.height}in'

    def set_css_style_screen(self, style_rule):
        style_rule['max-width'] = f'{self.width}in'
