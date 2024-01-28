from docx2css import api
from docx2css.ooxml import NAMESPACES, normalize_element_name, w
from docx2css.utils import CssUnit


def get_or_create_element(xml_parent, path):
    if '/' in path:
        partition = path.partition('/')
        next_el = get_or_create_element(xml_parent, partition[0])
        return get_or_create_element(next_el, partition[2])
    else:
        element = xml_parent.find(path, namespaces=NAMESPACES)
        element_name = normalize_element_name(path)
        if element is None:
            element = xml_parent.makeelement(element_name)
            xml_parent.append(element)
        return element


class Boolean:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            value = element.get(w('val'))
            return value not in ('false', '0')

    def __set__(self, instance, value: CssUnit):
        element = get_or_create_element(instance, self.path)
        if value is False:
            element.set(w('val'), '0')

    def __delete__(self, instance):
        raise NotImplementedError


class BorderDescriptor:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> api.Border:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            return api.Border(
                color=element.color,
                padding=element.padding,
                shadow=element.shadow,
                style=element.style,
                width=element.width
            )

    def __set__(self, instance, value: api.Border):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError


class Integer:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> int:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            return int(element.get(w('val')))

    def __set__(self, instance, value: int):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError


class HalfPointMeasure:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            sz = int(element.get(w('val')))
            return CssUnit(sz / 2, 'pt')

    def __set__(self, instance, value: CssUnit):
        element = get_or_create_element(instance, self.path)
        element.set(w('val'), str(round(value.pt * 2)))

    def __delete__(self, instance):
        raise NotImplementedError


class TwipMeasure:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            value = int(element.get(w('val')))
            return CssUnit(value, 'twip')

    def __set__(self, instance, value: CssUnit):
        element = get_or_create_element(instance, self.path)
        element.set(w('val'), str(value.twips))

    def __delete__(self, instance):
        raise NotImplementedError


class Shading:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> str:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            return element.get_color()

    def __set__(self, instance, value: str):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError


class String:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> str:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            return element.get(w('val'))

    def __set__(self, instance, value: str):
        element = get_or_create_element(instance, self.path)
        element.set(w('val'), value)

    def __delete__(self, instance):
        raise NotImplementedError


class Justification(String):

    def __delete__(self, instance):
        raise NotImplementedError


class VerticalJustification(String):

    def __delete__(self, instance):
        raise NotImplementedError
