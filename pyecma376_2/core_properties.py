import datetime
from typing import Optional, List, Tuple, IO, Callable, Any

import lxml.etree as etree  # type: ignore

XML_NAMESPACE = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}"
XML_NAMESPACE_DC = "{http://purl.org/dc/elements/1.1/}"
XML_NAMESPACE_DCTERMS = "{http://purl.org/dc/terms/}"
XML_NAMESPACE_XSI = "{http://www.w3.org/2001/XMLSchema-instance}"
XML_NAMESPACE_XML = "{http://www.w3.org/XML/1998/namespace}"


def extract_str(el: etree.Element) -> str:
    return ""  # TODO


def to_xml_tag(tag: str, content: Any) -> etree.Element:
    return etree.Element(tag)  # TODO append str(content)


def extract_keywords(el: etree.Element) -> List[Tuple[Optional[str], str]]:
    # TODO
    return []


def keywords_to_xml(tag: str, content: List[Tuple[Optional[str], str]]) -> etree.Element:
    result = etree.Element(tag)
    for lang, keyword in content:
        key_tag = etree.Element(XML_NAMESPACE + "value")
        # TODO append keyword
        if lang:
            key_tag[XML_NAMESPACE_XSI + 'lang'] = lang
        result.appendChild(key_tag)  # TODO
    return result


class OPCCoreProperties:
    """
    Class to represent the Core Properties Part, containing the meta data of an OPC package.

    The Core Properties are stored as an XML part in the package and referenced via a reference of type
    http://schemas.openxmlformats.org/.../core-properties from the package root.
    """
    def __init__(self):
        # TODO check types (see schema and Dublin Core schema)
        self.category: Optional[str] = None  # A categorization of the content of this package. (e.g. "Letter")
        self.contentStatus: Optional[str] = None  # The status of the content. (e.g. "Draft", "Final")
        self.created: Optional[datetime.date] = None  # Date of creation of the resource.
        self.creator: Optional[str] = None  # An entity primarily responsible for making the content of the resource.
        self.description: Optional[str] = None  # An explanation of the content of the resource.
        self.identifier: Optional[str] = None  # An unambiguous reference to the resource within a given context.
        # A list of keywords to support searching and indexing. Each Keyword may or may not have its language specified
        # (using XML lang strings).
        self.keywords: Optional[List[Tuple[Optional[str], str]]] = None
        self.language: Optional[str] = None  # The language of the intellectual content of the resource (see RFC 3066).
        self.lastModifiedBy: Optional[str] = None  # The user who performed the last modification.
        self.lastPrinted: Optional[datetime.date] = None  # The date and time of the last printing.
        self.modified: Optional[datetime.date] = None  # Date on which the resource was changed.
        self.revision: Optional[str] = None  # The revision number.
        self.subject: Optional[str] = None  # The topic of the content of the resource.
        self.title: Optional[str] = None  # The name given to the resource.
        self.version: Optional[str] = None  # The version number. This value is set by the user or by the application.

    # Yay, some dynamic coding. This tuple defines all attributes of the object as well as the corresponding XML tag
    # and the transformation functions for converting the XML string value into the correct python type and vice versa.
    ATTRIBUTES: Tuple[Tuple[str, str, Callable[[etree.Element], Any], Callable[[str, Any], etree.Element]], ...] = (
        ('category', XML_NAMESPACE + 'category', extract_str, to_xml_tag),
        ('contentStatus', XML_NAMESPACE + 'contentStatus', extract_str, to_xml_tag),
        ('created', XML_NAMESPACE_DCTERMS + 'created', extract_str, to_xml_tag),  # TODO date handling
        ('creator', XML_NAMESPACE_DC + 'creator', extract_str, to_xml_tag),
        ('description', XML_NAMESPACE_DC + 'description', extract_str, to_xml_tag),
        ('identifier', XML_NAMESPACE_DC + 'identifier', extract_str, to_xml_tag),
        ('keywords', XML_NAMESPACE + 'keywords', extract_keywords, keywords_to_xml),
        ('language', XML_NAMESPACE_DC + 'language', extract_str, to_xml_tag),
        ('lastModifiedBy', XML_NAMESPACE + 'lastModifiedBy', extract_str, to_xml_tag),
        ('lastPrinted', XML_NAMESPACE + 'lastPrinted', extract_str, to_xml_tag),  # TODO date handling
        ('modified', XML_NAMESPACE_DCTERMS + 'modified', extract_str, to_xml_tag),  # TODO date handling
        ('revision', XML_NAMESPACE + 'revision', extract_str, to_xml_tag),
        ('subject', XML_NAMESPACE_DC + 'subject', extract_str, to_xml_tag),
        ('title', XML_NAMESPACE_DC + 'title', extract_str, to_xml_tag),
        ('version', XML_NAMESPACE + 'version', extract_str, to_xml_tag),
    )
    ATTRIBUTES_BY_TAG = {a[1]: a for a in ATTRIBUTES}
    NSMAP = {
        None: XML_NAMESPACE[1:-1],
        'dc': XML_NAMESPACE_DC[1:-1],
        'dcterms': XML_NAMESPACE_DCTERMS[1:-1],
        'xml': XML_NAMESPACE_XML[1:-1],
        'xsi': XML_NAMESPACE_XSI[1:-1],
    }

    def write_xml(self, file: IO[bytes]) -> None:
        """
        Serialize and write these Core Properties into an XML file or package Part, according to the ECMA standard.
        """
        with etree.xmlfile(file, encoding="UTF-8") as xf:
            with xf.element(XML_NAMESPACE + "coreProperties", nsmap=self.NSMAP):
                for attr, tag, extractor, serializer in self.ATTRIBUTES:
                    value = getattr(self, attr)
                    if value is not None:
                        xf.write(serializer(tag, value))

    @classmethod
    def from_xml(cls, file: IO[bytes]) -> "OPCCoreProperties":
        """
        Read and parse Core Properties from an XML file or package Part, according to the ECMA standard.
        """
        result = cls()
        for _event, elem in etree.iterparse(file):
            if elem.tag in cls.ATTRIBUTES_BY_TAG:
                attr, tag, extractor, serializer = cls.ATTRIBUTES_BY_TAG[elem.tag]
                setattr(result, attr, extractor(elem))
                elem.clear()
        return result
