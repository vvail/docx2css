import zipfile

from lxml import etree

from docx2css.ooxml import opc_parser
from docx2css.ooxml.constants import CONTENT_TYPE
from docx2css.ooxml.fonts import FontTable
from docx2css.ooxml.numbering import Numbering
from docx2css.ooxml.sections import Sections
from docx2css.ooxml.styles import Styles
from docx2css.ooxml.theme import Theme


class OpcPackage:

    def __init__(self, filename):
        self.parts = {}
        with zipfile.ZipFile(filename) as file:
            with file.open('[Content_Types].xml') as part_names:
                types = etree.fromstring(part_names.read())
                for override in types.findall('.//{*}Override'):
                    name = override.get('ContentType')
                    location = override.get('PartName')[1:]
                    self.unmarshall_part(file, name, location)

    def unmarshall_part(self, zip_file, content_type, location):
        try:
            with zip_file.open(location) as file:
                element = etree.fromstring(file.read(), opc_parser)
                self.parts[content_type] = element
        except KeyError:
            pass

    @property
    def font_table(self):
        return FontTable(self)

    def get_numbering(self):
        if not hasattr(self, '_numbering'):
            setattr(self, '_numbering', Numbering(self))
        return getattr(self, '_numbering')

    @property
    def styles(self):
        if not hasattr(self, '_styles'):
            setattr(self, '_styles', Styles(self))
        return getattr(self, '_styles')

    @property
    def theme(self):
        return Theme(self)

    @property
    def sections(self):
        if not hasattr(self, '_sections'):
            sections = Sections(self.parts[CONTENT_TYPE.DOCUMENT])
            setattr(self, '_sections', sections)
        return getattr(self, '_sections')
