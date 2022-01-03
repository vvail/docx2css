from lxml import etree

from .constants import NAMESPACES


class DocxStyleLookup(etree.PythonElementClassLookup):

    def lookup(self, doc, element):
        from .styles import STYLE_MAPPING
        tag = f"{{{NAMESPACES['w']}}}style"
        type_name = f"{{{NAMESPACES['w']}}}type"
        if element.tag == tag:
            return STYLE_MAPPING.get(element.get(type_name), None)


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
