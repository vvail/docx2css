from abc import ABC, abstractmethod
from collections.abc import Mapping
import logging
import zipfile

from lxml import etree

from docx2css.ooxml import opc_parser
from docx2css.ooxml.constants import CONTENT_TYPE
from docx2css.ooxml.fonts import FontTable
from docx2css.ooxml.numbering import AbstractNumbering, Num
from docx2css.ooxml.sections import Sections
from docx2css.ooxml.styles import Styles
from docx2css.ooxml.theme import Theme


logger = logging.getLogger(__name__)


class OpcPackage:

    def __init__(self, filename):
        self.parts = {}
        self.unmarshall_parts(filename)

    def unmarshall_parts(self, filename):
        # TODO: Remove the try by lazy loading the file
        try:
            with zipfile.ZipFile(filename) as file:
                with file.open('[Content_Types].xml') as part_names:
                    types = etree.fromstring(part_names.read())
                    for override in types.findall('.//{*}Override'):
                        name = override.get('ContentType')
                        location = override.get('PartName')[1:]
                        self.unzip_part(file, name, location)
        except FileNotFoundError:
            logger.error(f'{filename} not found. Docx has NOT been parsed!')

    def unzip_part(self, zip_file, content_type, location):
        try:
            with zip_file.open(location) as file:
                element = etree.fromstring(file.read(), opc_parser)
                self.parts[content_type] = element
        except KeyError:
            pass

    @property
    def font_table(self):
        return FontTable(self)

    @property
    def numbering(self):
        if not hasattr(self, '_numbering'):
            part = NumberingPart(self, self.parts.get(CONTENT_TYPE.NUMBERING))
            setattr(self, '_numbering', part)
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


class PackagePart(ABC):

    def __init__(self, opc_package, xml_content):
        self.opc_package = opc_package
        if xml_content is not None:
            self.unmarshall(xml_content)

    @abstractmethod
    def unmarshall(self, xml_content):
        pass


class NumberingPart(PackagePart, Mapping):

    def __init__(self, opc_package, xml_content):
        self.__abstract_numbering = {}
        self.__numbering_instances = {}
        super().__init__(opc_package, xml_content)

    def unmarshall(self, xml_content):
        for numbering in xml_content:
            numbering.numbering_part = self
            if isinstance(numbering, AbstractNumbering):
                self.__abstract_numbering[numbering.id] = numbering
            if isinstance(numbering, Num):
                abstract_numbering = self.resolve_abstract_numbering(numbering.abstract_num_id)
                self.__numbering_instances[numbering.id] = abstract_numbering

    def __getitem__(self, k):
        return self.__numbering_instances[k]

    def __len__(self) -> int:
        return len(self.__numbering_instances)

    def __iter__(self):
        return iter(self.__numbering_instances)

    def resolve_abstract_numbering(self, abstract_num_id):
        abstract_numbering = self.__abstract_numbering[abstract_num_id]
        style_id = abstract_numbering.numbering_style_link
        if style_id is None:
            return abstract_numbering
        else:
            style = self.opc_package.styles[style_id]
            num_id = style.numbering_instance_id
            return self.__numbering_instances[num_id]
