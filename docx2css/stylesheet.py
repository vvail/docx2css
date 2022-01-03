import cssutils

from docx2css.ooxml import NAMESPACES
from docx2css.ooxml.constants import CONTENT_TYPE


class Stylesheet:

    def __init__(self, opc_package):
        self.opc_package = opc_package
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

    def _add_rule(self, rule):
        """Add a rule, or a set of rules to the CSSStylesheet"""
        if isinstance(rule, cssutils.css.CSSRule):
            self._css_stylesheet.add(rule)
        else:
            for r in rule:
                self._css_stylesheet.add(r)

    @property
    def css_preferences(self):
        if not hasattr(self, '_css_preferences'):
            default_preferences = {
                'indentClosingBrace': False,
                'omitLastSemicolon': False,
            }
            preferences = cssutils.serialize.Preferences(**default_preferences)
            setattr(self, '_css_preferences', preferences)
        return getattr(self, '_css_preferences')

    @css_preferences.setter
    def css_preferences(self, preferences):
        setattr(self, '_css_preferences', preferences)

    @property
    def css_stylesheet(self):
        self._css_stylesheet = cssutils.css.CSSStyleSheet()
        self._serialize_css()
        return self._css_stylesheet

    @property
    def cssText(self):
        return self.css_stylesheet.cssText.decode('utf-8')

    def _serialize_css(self):
        cssutils.serialize.CSSSerializer(self.css_preferences)

        # Serialize body
        section = self.opc_package.sections[-1]
        self._add_rule(section.css_style_rule_print())
        self._add_rule(section.css_style_rule_screen())
        self._add_rule(self.css_body_style())

        for style in self.opc_package.styles.values():
            # TODO: Handle numbering styles
            if style.type == 'numbering':
                continue
            if style.numbering is not None:
                self._add_rule(style.css_numbering_style_rule())
            self._add_rule(style.css_style_rule())