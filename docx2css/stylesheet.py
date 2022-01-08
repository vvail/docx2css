import cssutils

from docx2css.ooxml import NAMESPACES
from docx2css.ooxml.constants import CONTENT_TYPE


class Stylesheet:

    def __init__(self, opc_package, preferences=None):
        self.opc_package = opc_package
        self.preferences = preferences or {}
        self._css_stylesheet = None

    def css_body_style(self):
        style = self.default_css_properties()
        return cssutils.css.CSSStyleRule('body', style=style)

    def default_css_properties(self):
        styles = self.opc_package.parts[CONTENT_TYPE.STYLES]
        defaults = styles.find('./w:docDefaults', namespaces=NAMESPACES)
        defaults.package = self.opc_package
        return defaults.css_style_declaration()

    def merge_doc_defaults(self, style):
        """
        Merge the docDefaults properties found in styles.xml to a style
        :param style: DocxStyle to merge to
        """
        defaults = self.default_css_properties()
        style_properties = style.css_style_rule().style
        for prop in defaults.getProperties():
            if style_properties.getProperty(prop.name) is None:
                style_properties.setProperty(prop)

    def _add_rules(self, rule):
        """Add a set of rules to the CSSStylesheet"""
        for r in rule:
            self._css_stylesheet.add(r)

    @property
    def css_stylesheet(self):
        self._css_stylesheet = cssutils.css.CSSStyleSheet()
        self._serialize_css()
        return self._css_stylesheet

    @property
    def cssText(self):
        return self.css_stylesheet.cssText.decode('utf-8')

    def _serialize_css(self):

        # Serialize body
        section = self.opc_package.sections[-1]
        self._add_rules(section.css_style_rules(self.preferences))
        self._add_rules([self.css_body_style()])

        for style in self.opc_package.styles.values():
            self._add_rules(style.css_style_rules())
