from abc import ABC, abstractmethod

import cssutils

from docx2css.utils import CssUnit


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
