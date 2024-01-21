from lxml import etree

from docx2css.ooxml import w, wordml
from docx2css.ooxml.constants import NAMESPACES
from docx2css.ooxml.styles import PPrProxy, RPrProxy


@wordml('abstractNum')
class AbstractNumbering(etree.ElementBase):
    numbering_part = None

    @property
    def id(self):
        return int(self.get(w('abstractNumId')))

    @property
    def levels(self):
        if not hasattr(self, '_levels'):
            levels = {}
            for level in self.findall(w('lvl')):
                level.abstract_numbering = self
                levels[level.level_number] = level
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
    numbering_part = None

    @property
    def id(self):
        return int(self.get(w('numId')))

    @property
    def abstract_num_id(self):
        child = self.find(w('abstractNumId'))
        return int(child.get(w('val')))


@wordml('lvl')
class Level(PPrProxy, RPrProxy):
    abstract_numbering = None

    @property
    def number_format(self):
        # This element can be part of alternate content (eg 0001 and such)
        # In this case, we will rely on the fallback, which happens to
        # be the last element
        element = self.findall('.//w:numFmt', namespaces=NAMESPACES)
        if len(element):
            return element[-1].get(w('val'))

    @property
    def is_legal_format(self):
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
    def level_number(self):
        return int(self.get(w('ilvl')))

    @property
    def level_start(self):
        element = self.find(w('start'))
        if element is not None:
            return int(element.get(w('val'))) or 0

    @property
    def level_restart(self):
        element = self.find(w('lvlRestart'))
        if element is not None:
            return int(element.get(w('val')))

    @property
    def level_suffix(self):
        element = self.find(w('suff'))
        if element is not None:
            return element.get(w('val'))
        else:
            return 'tab'

    @property
    def level_text(self):
        element = self.find(w('lvlText'))
        if element is not None:
            return element.get(w('val'))

    @property
    def paragraph_style(self):
        element = self.find(w('pStyle'))
        if element is not None:
            return element.get(w('val'))
