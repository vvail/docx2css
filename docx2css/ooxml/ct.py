from docx2css import api
from docx2css.ooxml import NAMESPACES, normalize_element_name, w
from docx2css.ooxml.simple_types import ST_Underline, ST_FontFamily
from docx2css.utils import AutoLength, CssUnit, Percentage


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
        element = instance.find(f'.//{self.path}', namespaces=NAMESPACES)
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


class FontDescriptor:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> str:
        element = instance.find(f'.//{self.path}', namespaces=NAMESPACES)
        if element is not None:
            # What we want to do here is have a set of the fonts, but at the
            # same time, we want to keep the order so it's easier to use a
            # dict because the order is guaranteed
            fonts = {}
            # Theme values take precedence over explicit values, so we
            # favour the former
            attributes = (
                element.get(w('hAnsiTheme')) or element.get(w('hAnsi')),
                element.get(w('asciiTheme')) or element.get(w('ascii')),
                element.get(w('eastAsiaTheme')) or element.get(w('eastAsia')),
                element.get(w('cstheme')) or element.get(w('cs')),
            )
            for attribute in attributes:
                font_name = element.get_theme_font_or_font_value(attribute)
                if font_name:
                    for f in element.get_font_from_font_table(font_name):
                        fonts[f'"{f}"' if ' ' in f else f] = None
            # Push the generic family at the end. This happens when different
            # fonts are specified, and they are found in the font table
            for generic in ST_FontFamily.docx2css.values():
                if generic in fonts:
                    value = fonts.pop(generic)
                    fonts[generic] = value
            return ', '.join(fonts.keys())

    def __set__(self, instance, value: str):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError


class UnderlineDescriptor:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> api.TextDecoration:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            color = element.get_color()
            style = ST_Underline.css_value(element.get(w('val')))
            value = api.TextDecoration(color=color, style=style)
            if style != 'none':
                value.add_line(api.TextDecoration.UNDERLINE)
            return value

    def __set__(self, instance, value: api.TextDecoration):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError


class LineHeight:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner):
        element = instance.find(f'.//{self.path}', namespaces=NAMESPACES)
        if element is not None:
            height = element.get(w('line'))
            rule = element.get(w('lineRule'))
            if height is not None:
                if rule in ('atLeast', 'exact'):
                    return CssUnit(int(height), 'twip')
                elif rule == 'auto':
                    # Height is 240th of a line
                    return int(height) / 240


class ParagraphIndentLeft:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(f'.//{self.path}', namespaces=NAMESPACES)
        if element is not None:
            right = element.get(w('start')) or element.get(w('left'))
            if right is not None:
                return CssUnit(int(right), 'twip')


class ParagraphIndentRight:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(f'.//{self.path}', namespaces=NAMESPACES)
        if element is not None:
            right = element.get(w('end')) or element.get(w('right'))
            if right is not None:
                return CssUnit(int(right), 'twip')


def parse_boolean(raw_value):
    if raw_value is not None:
        return not raw_value.lower() in ('false', '0')
    else:
        return None


class SpaceAfterParagraph:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(f'.//{self.path}', namespaces=NAMESPACES)
        if element is not None:
            after = element.get(w('after'))
            auto = parse_boolean(element.get(w('afterAutospacing')))
            if auto is not True and after is not None:
                return CssUnit(after, 'twip')


class SpaceBeforeParagraph:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(f'.//{self.path}', namespaces=NAMESPACES)
        if element is not None:
            before = element.get(w('before'))
            auto = parse_boolean(element.get(w('beforeAutospacing')))
            if auto is not True and before is not None:
                return CssUnit(before, 'twip')


class TextIndent:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(f'.//{self.path}', namespaces=NAMESPACES)
        if element is not None:
            # firstLine and hanging attributes are mutually exclusive, if both
            # are specified, then the firstLine value is ignored
            hanging = element.get(w('hanging'))
            if hanging is not None:
                return CssUnit(-1 * int(hanging), 'twip')
            first_line = element.get(w('firstLine'))
            if first_line is not None:
                return CssUnit(int(first_line), 'twip')


class TableLayout:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            return element.get(w('type'))

    def __set__(self, instance, value: CssUnit):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError


class TableMeasure:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> CssUnit:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            unit = element.get(w('type'))
            value = element.get(w('w'))
            if unit == 'auto':
                return AutoLength()
            elif unit == 'dxa':
                return CssUnit(value, 'twip')
            elif unit == 'nil':
                return CssUnit(0)
            elif unit == 'pct':
                return Percentage(int(value) / 50)
            else:
                raise ValueError(f'Unit "{unit}" is invalid!')

    def __set__(self, instance, value: CssUnit):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError


class TableRowHeightType:

    def __init__(self, relative_path):
        self.path = relative_path

    def __get__(self, instance, owner) -> str:
        element = instance.find(self.path, namespaces=NAMESPACES)
        if element is not None:
            return element.get(w('hRule'))

    def __set__(self, instance, value: str):
        raise NotImplementedError

    def __delete__(self, instance):
        raise NotImplementedError
