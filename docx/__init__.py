# encoding: utf-8

from docx.api import Document  # noqa

__version__ = '0.8.6'


# register custom Part classes with opc package reader

from docx.opc.constants import CONTENT_TYPE as CT, RELATIONSHIP_TYPE as RT # type: ignore
from docx.opc.part import PartFactory # type: ignore
from docx.opc.parts.coreprops import CorePropertiesPart # type: ignore

from docx.parts.document import DocumentPart  # type: ignore
from docx.parts.image import ImagePart # type: ignore
from docx.parts.numbering import NumberingPart # type: ignore
from docx.parts.settings import SettingsPart # type: ignore
from docx.parts.styles import StylesPart # type: ignore


def part_class_selector(content_type, reltype):
    if reltype == RT.IMAGE:
        return ImagePart
    return None


PartFactory.part_class_selector = part_class_selector # type: ignore
PartFactory.part_type_for[CT.OPC_CORE_PROPERTIES] = CorePropertiesPart
PartFactory.part_type_for[CT.WML_DOCUMENT_MAIN] = DocumentPart
PartFactory.part_type_for[CT.WML_NUMBERING] = NumberingPart
PartFactory.part_type_for[CT.WML_SETTINGS] = SettingsPart
PartFactory.part_type_for[CT.WML_STYLES] = StylesPart

del (
    CT, CorePropertiesPart, DocumentPart, NumberingPart, PartFactory,
    StylesPart, part_class_selector
)
