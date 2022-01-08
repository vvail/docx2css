import re

from cssutils import css
from lxml import etree

from docx2css.ooxml import w, wordml
from docx2css.ooxml.constants import CONTENT_TYPE, NAMESPACES as NS
from docx2css.ooxml.simple_types import ST_NumberFormat
from docx2css.ooxml.styles import CssPropertyAdapter


class Numbering:

    def __init__(self, opc_package):
        self.abstract_numbering = {}
        self.numbering_instances = {}
        self.opc_package = opc_package
        # Numbering is optional in the package. There might not be one
        # try:
        part = opc_package.parts[CONTENT_TYPE.NUMBERING]
        self.unmarshall_numbering(part, opc_package)
        self.follow_style_links()
        # except KeyError:

    def unmarshall_numbering(self, numbering_part, opc_package):
        for numbering in numbering_part:
            if isinstance(numbering, AbstractNumbering):
                numbering.styles = opc_package.styles
                self.abstract_numbering[numbering.id] = numbering
            elif isinstance(numbering, Num):
                abstract = self.abstract_numbering[numbering.abstract_num_id]
                self.numbering_instances[numbering.id] = abstract

    def follow_style_link(self, style_name):
        styles = self.opc_package.parts[CONTENT_TYPE.STYLES]
        q = f'.//w:style[@w:styleId="{style_name}"]/w:pPr/w:numPr/w:numId'
        num_id = int(styles.find(q, namespaces=NS).get(w('val')))
        instance = self.numbering_instances[num_id]
        return instance

    def follow_style_links(self):
        """Replace the AbstractNumbering that are only defined by pointing
        to a numbering style link
        """
        for key, numbering in self.abstract_numbering.items():
            style_name = numbering.numbering_style_link
            if style_name:
                new_numbering = self.follow_style_link(style_name)
                self.abstract_numbering[key] = new_numbering
                for k, old_numbering in self.numbering_instances.items():
                    if numbering == old_numbering:
                        self.numbering_instances[k] = new_numbering

    def css_root_counters(self):
        """
        Get a space-separated string representing all the counters that
        should be set at the root, or counters that never restart
        :return:
        """
        counters = set()
        for numbering in self.abstract_numbering.values():
            for level in numbering.levels.values():
                if level.level == 0 or level.level_restart == 0:
                    counters.add(level.css_counter_name())
        return ' '.join(sorted(counters))


@wordml('abstractNum')
class AbstractNumbering(etree.ElementBase):

    @property
    def id(self):
        return int(self.get(w('abstractNumId')))

    def get_level_for_paragraph(self, style_id):
        """Get the level where a pStyle element matches the style_id"""
        for level in self.levels.values():
            if level.paragraph_style == style_id:
                return level

    def get_name(self):
        """
        Give a name to this abstract numbering. It will be either:
            * name taken from the child of the same name;
            * style_link it points to; or
            * the value of the _abstractNumId_ attribute
        :return: Name of the numbering
        """
        return self.name or self.style_link or f'counter{self.id}'

    @property
    def levels(self):
        if not hasattr(self, '_levels'):
            levels = {}
            for level in self.findall(w('lvl')):
                level.numbering = self
                levels[level.level] = level
            setattr(self, '_levels', levels)
        return getattr(self, '_levels')

    @property
    def multi_level_type(self):
        element = self.find(w('multiLevelType'))
        return element.get(w('val')) if element is not None else None

    @property
    def name(self):
        element = self.find(w('name'))
        return element.get(w('val')) if element is not None else None

    @property
    def numbering_style_link(self):
        """
        Provide the name of numbering style this abstract definitions
        links to
        :return: Name of a numbering style
        """
        element = self.find(w('numStyleLink'))
        return element.get(w('val')) if element is not None else None

    @property
    def style(self):
        """
        Return the actual numbering style if any is defined with the
        'numStyleLink' element
        :return: Style or None
        """
        link = self.numbering_style_link
        return self.styles.get(link, None)

    @property
    def style_link(self):
        """
        Provide the name of the numbering style this abstract numbering
        refers to
        :return: Name of a numbering style
        """
        element = self.find(w('styleLink'))
        return element.get(w('val')) if element is not None else None


@wordml('num')
class Num(etree.ElementBase):

    @property
    def id(self):
        return int(self.get(w('numId')))

    @property
    def abstract_num_id(self):
        child = self.find(w('abstractNumId'))
        return int(child.get(w('val')))


@wordml('lvl')
class Level(etree.ElementBase):

    # <xsd:element name="pPr" type="CT_PPrGeneral" minOccurs="0"/>
    # <xsd:element name="rPr" type="CT_RPr" minOccurs="0"/>

    def css_counter(self):
        counter_format = ST_NumberFormat.css_value(self.format)
        if self.is_legal:
            counter_format = 'decimal'
        if counter_format != 'none':
            return f'counter({self.css_counter_name()}, {counter_format})'

    def css_counter_content(self):
        if self.format == 'none':
            return
        elif self.format == 'bullet':
            # When the content is a bullet, there should only be one
            # character in the level_text string, and it might not be
            # printable. Therefore, it is best to escape it
            return fr'"\005C {ord(self.level_text):04x}"'
        contents = []
        tokens = re.split('(%\\d)', self.level_text)
        for token in (t for t in tokens if t):
            regex = re.match('%(\\d)', token)
            if regex:
                level = self.numbering.levels[int(regex.group(1)) - 1]
                contents.append(level.css_counter())
            else:
                contents.append(f'"{token}"')
        if self.suffix == 'space':
            contents.append(r'"\005C 00A0"')
        return ' '.join(filter(lambda x: x is not None, contents))

    def css_counter_name(self):
        parent = self.numbering.get_name()
        parent = ''.join(parent.split())
        return f'{parent}-L{self.level}'

    def css_counter_resets(self):
        """
        Get a space-separated list of counters to reset at this level
        :return: String
        """
        levels = set()
        next_levels = set(lvl for lvl in self.numbering.levels.values()
                          if lvl.level > self.level)
        for level in next_levels:
            restart = level.level_restart or level.level
            if restart - 1 == self.level:
                # Handle the fact that counters don't always start from zero
                counter = level.css_counter_name()
                if level.start and level.start != 1:
                    counter += f' {level.start - 1}'

                # Don't reset the counter if no special restart value
                if level.level_restart is None or level.level_restart != 0:
                    levels.add(counter)
        return ' '.join(sorted(levels))

    def css_properties(self):
        for d in self.iterdescendants():
            if isinstance(d, CssPropertyAdapter):
                yield d

    def css_selector(self):
        if self.paragraph_style is not None:
            p_style = self.get_paragraph_style()
            return f'{p_style.css_current_selector()}:before'

    def css_style_declaration(self):
        css_style = css.CSSStyleDeclaration()
        package = self.numbering.styles.opc_package
        for prop in self.css_properties():
            prop.set_css_style(css_style, package)
        css_style['content'] = self.css_counter_content()
        css_style['counter-increment'] = self.css_counter_name()
        css_style['text-align'] = self.justification
        return css_style

    def css_style_rule(self):
        if not hasattr(self, '_css_style_rule'):
            css_style = self.css_style_declaration()
            if self.get_paragraph_style() is not None:
                self.pull_margin_left(css_style)
                self.pull_text_indent(css_style)
                self.adjust_for_indent(css_style)

            rule = css.CSSStyleRule(self.css_selector(), style=css_style)
            setattr(self, '_css_style_rule', rule)
        return getattr(self, '_css_style_rule')

    def pull_margin_left(self, css_style):
        """
        When a paragraph style is attached to this level, the former's
        *margin-left* has priority
        :param css_style:
        :return:
        """
        self.pull_property_from_paragraph_style(css_style, 'margin-left')

    def pull_text_indent(self, css_style):
        """
        When a paragraph style is attached to this level, the former's
        *text-indent* has priority
        :param css_style:
        :return:
        """
        self.pull_property_from_paragraph_style(css_style, 'text-indent')

    def pull_property_from_paragraph_style(self, css_style, prop_name):
        p_style = self.get_paragraph_style().css_style_declaration()
        p_property = p_style.getProperty(prop_name)
        if p_property is not None:
            css_style[prop_name] = p_property.propertyValue.cssText

    def adjust_for_indent(self, css_style):
        """
        Certain properties must be set depending on whether there is a
        hanging indent, or a first line indent.
        """
        text_indent = css_style.getProperty('text-indent')
        text_indent_value = text_indent.propertyValue[0].value
        css_style['margin-left'] = ''
        if text_indent_value < 0:
            self.adjust_for_hanging_indent(css_style)
        else:
            self.adjust_for_first_line_indent(css_style)

    def adjust_for_hanging_indent(self, css_style):
        """
        Adjust the CSS rules in the case of a hanging indent, that is:

        * delete margin-left
        * add display: inline-block
        :return:
        """
        if self.suffix == 'tab':
            css_style['display'] = 'inline-block'
        else:
            css_style['text-indent'] = ''

    def adjust_for_first_line_indent(self, css_style):
        """
        Adjust the CSS rules in the case of a first line indent, that is:

        * add margin-right with the same value as text-indent to fake a tab
        * add a display: inline-block
        :param css_style:
        :return:
        """
        if self.suffix == 'tab':
            text_indent = css_style.getPropertyValue('text-indent')
            css_style['margin-right'] = text_indent
        css_style['display'] = 'inline-block'

    def get_paragraph_style(self):
        """
        Get the paragraph style associated with this level, assuming one
        is defined with the *paragraph_style* property
        :return:
        """
        return self.numbering.styles.get(self.paragraph_style, None)

    @property
    def format(self):
        # This element can be part of alternate content (eg 0001 and such)
        # In this case, we will rely on the fallback, which happens to
        # be the last element
        element = self.findall('.//w:numFmt', namespaces=NS)
        if len(element):
            return element[-1].get(w('val'))

    @property
    def is_legal(self):
        element = self.find(w('isLgl'))
        if element is not None:
            value = element.get(w('val'))
            if value is None:
                return True
            else:
                return value.lower() not in ('false', '0')
        return False

    @property
    def justification(self):
        element = self.find(w('lvlJc'))
        if element is not None:
            return element.get(w('val'))

    @property
    def level(self):
        return int(self.get(w('ilvl')))

    @property
    def level_restart(self):
        element = self.find(w('lvlRestart'))
        if element is not None:
            return int(element.get(w('val')))

    @property
    def level_text(self):
        element = self.find(w('lvlText'))
        if element is not None:
            return element.get(w('val'))

    @property
    def numbering(self):
        """Abstract numbering element this level is part of"""
        return getattr(self, '_numbering', None)

    @numbering.setter
    def numbering(self, numbering):
        setattr(self, '_numbering', numbering)

    @property
    def paragraph_style(self):
        element = self.find(w('pStyle'))
        if element is not None:
            return element.get(w('val'))

    @property
    def start(self):
        element = self.find(w('start'))
        if element is not None:
            return int(element.get(w('val')))

    @property
    def suffix(self):
        element = self.find(w('suff'))
        if element is not None:
            return element.get(w('val'))
        return 'tab'
