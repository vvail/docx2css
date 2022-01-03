from docx2css.ooxml.package import OpcPackage
from docx2css.stylesheet import Stylesheet


def open_docx(filename):
    docx = OpcPackage(filename)
    return Stylesheet(docx)
