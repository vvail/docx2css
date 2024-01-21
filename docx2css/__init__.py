from docx2css.css.serializers import CssStylesheetSerializer
from docx2css.ooxml.parsers import DocxParser


def open_docx(filename):
    parser = DocxParser(filename)
    return parser.parse()


def to_string(stylesheet):
    serializer = CssStylesheetSerializer(stylesheet)
    return serializer.serialize()
