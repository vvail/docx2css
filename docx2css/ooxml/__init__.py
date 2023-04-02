from lxml import etree

from .constants import NAMESPACES


class DocxStyleLookup(etree.PythonElementClassLookup):

    def lookup(self, doc, element):
        from . import styles, tables
        style_mapping = {
            'character': styles.CharacterStyle,
            'numbering': styles.NumberingStyle,
            'paragraph': styles.ParagraphStyle,
            'table': styles.DocxTableStyle,
        }
        style_tag = f"{{{NAMESPACES['w']}}}style"
        style_type_name = f"{{{NAMESPACES['w']}}}type"
        bottom_tag = f"{{{NAMESPACES['w']}}}bottom"
        left_tag = f"{{{NAMESPACES['w']}}}left"
        right_tag = f"{{{NAMESPACES['w']}}}right"
        top_tag = f"{{{NAMESPACES['w']}}}top"
        if element.tag == style_tag:
            return style_mapping.get(element.get(style_type_name), None)
        # <w:bottom> can be a border or a table cell margin. The latter
        # has a 'type' attribute while the former does not. The presence
        # of this attribute is used to discriminate between borders and
        # cell margins. If the 'type' attribute is not found, the lookup
        # scheme will fallback on the annotations
        elif element.tag == bottom_tag and element.get(w('type')) is not None:
            return tables.TableCellMarginBottom
        elif element.tag == left_tag and element.get(w('type')) is not None:
            return tables.TableCellMarginLeft
        elif element.tag == right_tag and element.get(w('type')) is not None:
            return tables.TableCellMarginRight
        elif element.tag == top_tag and element.get(w('type')) is not None:
            return tables.TableCellMarginTop


lookup = etree.ElementNamespaceClassLookup()
opc_parser = etree.XMLParser()
opc_parser.set_element_class_lookup(DocxStyleLookup(lookup))
drawingml = lookup.get_namespace(NAMESPACES['a'])
wordml = lookup.get_namespace(NAMESPACES['w'])


def a(tag):
    """Shortcut function to build a namespace-qualified element name"""
    return etree.QName(NAMESPACES['a'], tag)


def w(tag):
    """Shortcut function to build a namespace-qualified element name"""
    return etree.QName(NAMESPACES['w'], tag)
