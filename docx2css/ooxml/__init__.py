from lxml import etree

from .constants import NAMESPACES


class DocxStyleLookup(etree.PythonElementClassLookup):

    def lookup(self, doc, element):
        from . import styles, tables
        style_mapping = {
            'character': styles.DocxCharacterStyle,
            'numbering': styles.DocxNumberingStyle,
            'paragraph': styles.DocxParagraphStyle,
            'table': tables.DocxTableStyle,
        }
        style_tag = f"{{{NAMESPACES['w']}}}style"
        style_type_name = f"{{{NAMESPACES['w']}}}type"
        if element.tag == style_tag:
            return style_mapping.get(element.get(style_type_name), None)


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


def normalize_element_name(name):
    """Change an element name from w:name to {w]name"""
    for ns in NAMESPACES:
        name = name.replace(f'{ns}:', f'{{{NAMESPACES[ns]}}}')
    return name
