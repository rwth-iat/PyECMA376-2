import abc
import collections
import enum
import io
import re
import types
import urllib.parse
# We need lxml's ElementTree implementation, as it allows correct handling of default namespaces (xmlns="…") when
# writing XML files. And since we already have it, we also use the iterative writer.
import lxml.etree as etree  # type: ignore
from typing import BinaryIO, Sequence, Dict, Iterable, NamedTuple, Optional, IO, Generator, List, DefaultDict, Tuple

RE_RELS_PARTS = re.compile(r'^(.*/)_rels/([^/]*).rels$', re.IGNORECASE)
RE_FRAGMENT_ITEMS = re.compile(r'^(.*)/\[(\d+)\](.last)?.piece$', re.IGNORECASE)
RELATIONSHIPS_XML_NAMESPACE = "{http://schemas.openxmlformats.org/package/2006/relationships}"


class OPCPackageReader(metaclass=abc.ABCMeta):
    content_types_stream_name: Optional[str] = None

    class _PartDescriptor(types.SimpleNamespace):
        content_type: str
        fragmented: bool
        physical_item_name: str

    def __init__(self):
        # dict mapping all known normalized part names to (content type, fragmented, physical item name)
        self._parts: Dict[str, OPCPackageReader._PartDescriptor] = {}

    def _init_data(self) -> None:
        """ Part of the initializer which should be called after opening the package """
        # First run: Find all parts (including the Content Types Stream)
        for item_name in self.list_items():
            fragment_match = RE_FRAGMENT_ITEMS.match(item_name)
            if fragment_match:
                if fragment_match[2] != "0":
                    continue
                part_name = fragment_match[1]
                self._parts[normalize_part_name(part_name)] = self._PartDescriptor(content_type="", fragmented=True,
                                                                                   physical_item_name=part_name)
            else:
                self._parts[normalize_part_name(item_name)] = self._PartDescriptor(content_type="", fragmented=False,
                                                                                   physical_item_name=item_name)

        # Read ContentTypes data and update parts' data, remove ContentTypesStream part afterwards
        if self.content_types_stream_name is not None:
            with self.open_part(self.content_types_stream_name) as part:
                content_types = ContentTypesData.from_xml(part)
            for part_name, part_record in self._parts.items():
                if part_name == self.content_types_stream_name:
                    continue
                content_type = content_types.get_content_type(part_name)
                if content_type is None:
                    raise ValueError("No content type for part {} given".format(part_name))
                    # TODO add failsafe variant?
                part_record.content_type = content_type
            del self._parts[self.content_types_stream_name]

    def list_parts(self, include_rels_parts=False) -> Iterable[Tuple[str, str]]:
        return ((normalized_name, part_descriptor.content_type)
                for normalized_name, part_descriptor in self._parts.items()
                if include_rels_parts or not RE_RELS_PARTS.match(normalized_name))

    def open_part(self, name: str) -> IO[bytes]:
        try:
            part_descriptor = self._parts[normalize_part_name(name)]
        except KeyError as e:
            raise KeyError("Could not find part {} in package".format(name)) from e
        if part_descriptor.fragmented:
            return FragmentedPartReader(part_descriptor.physical_item_name, self)  # type: ignore
        else:
            return self.open_item(part_descriptor.physical_item_name)

    def get_raw_relationships(self, part_name: str = "/") -> Generator["OPCRelationship", None, None]:
        try:
            rels_part = self.open_part(_rels_part_for(part_name))
        except KeyError:
            return
        yield from self._read_relationships(rels_part)

    def get_related_parts_by_type(self, part_name: str = "/") -> Dict[str, List[str]]:
        result = collections.defaultdict(list)  # type: DefaultDict[str, List[str]]
        for relationship in self.get_raw_relationships(part_name):
            if relationship.target_mode == OPCTargetMode.INTERNAL:
                result[relationship.type].append(normalize_part_name(part_realpath(relationship.target, part_name)))
        return dict(result)

    def close(self) -> None:
        pass

    def __enter__(self):
        return self

    def __exit__(self, _exc_type, _exc_val, _exc_tb):
        self.close()

    @staticmethod
    def _read_relationships(rels_part: IO[bytes]) -> Generator["OPCRelationship", None, None]:
        for _event, elem in etree.iterparse(rels_part):
            if elem.tag == RELATIONSHIPS_XML_NAMESPACE + "Relationship":
                yield OPCRelationship(
                    elem.attrib["Id"],
                    elem.attrib["Type"],
                    elem.attrib["Target"],
                    OPCTargetMode.from_serialization(elem.attrib.get('TargetMode', 'Internal')))

    @abc.abstractmethod
    def list_items(self) -> Iterable[str]:
        pass

    @abc.abstractmethod
    def open_item(self, name: str) -> IO[bytes]:
        pass


class FragmentedPartReader(io.RawIOBase):
    def __init__(self, name: str, reader: OPCPackageReader):
        self._name: str = name
        self._reader: OPCPackageReader = reader
        self._fragment_number: int = 0
        self._finished = False
        self._current_item_handle: IO[bytes]
        self._open_next_item()

    def _open_next_item(self) -> None:
        try:
            self. _current_item_handle = self._reader.open_item("{}/[{}].piece"
                                                                .format(self._name, self._fragment_number))
            self._fragment_number += 1
        except KeyError:
            self._finished = True
            try:
                self._current_item_handle = self._reader.open_item("{}/[{}].last.piece"
                                                                   .format(self._name, self._fragment_number))
                self._fragment_number += 1
            except KeyError as e:
                raise KeyError("Fragment {} of part {} is missing in package"
                               .format(self._fragment_number, self._name)) from e

    def seekable(self) -> bool:
        return False

    def read(self, size: int = -1) -> Optional[bytes]:
        result = self._current_item_handle.read(size)
        while result is not None and (size == -1 or 0 == len(result) < size) and not self._finished:
            self._current_item_handle.close()
            self._open_next_item()
            result += self._current_item_handle.read(size)
        return result

    def close(self) -> None:
        self._current_item_handle.close()


class OPCPackageWriter(metaclass=abc.ABCMeta):
    content_types_stream_name: Optional[str] = None

    def __init__(self):
        if self.content_types_stream_name is not None:
            self.content_types = ContentTypesData()
            self.content_types_written = False

    def open_part(self, name: str, content_type: str) -> IO[bytes]:
        name = normalize_part_name(name)
        check_part_name(name)
        if self.content_types_stream_name is not None:
            if self.content_types.get_content_type(name) != content_type:
                if self.content_types_written:
                    raise RuntimeError("Content Type of part {} is not set correctly but ContentTypeStream has been "
                                       "written already.".format(name))
                else:
                    self.content_types.overrides[name] = content_type
        return self.create_item(name, content_type)

    def write_relationships(self, relationships: Iterable["OPCRelationship"], part_name: str = "/") -> None:
        # We do currently not support fragmented relationships parts
        part_name = normalize_part_name(part_name)
        if part_name != "/":
            # "/" is a special case, as it is allowed to end on a "/"
            check_part_name(part_name)
        with self.open_part(_rels_part_for(part_name), "application/vnd.openxmlformats-package.relationships+xml") as i:
            self._write_relationships(i, relationships)

    def create_fragmented_part(self, name: str, content_type: str) -> "FragmentedPartWriterHandle":
        part_name = normalize_part_name(name)
        check_part_name(part_name)
        if self.content_types_stream_name is not None:
            if self.content_types.get_content_type(name) != content_type:
                if self.content_types_written:
                    raise RuntimeError("Content Type of part {} is not set correctly but ContentTypeStream has been "
                                       "written already.".format(name))
                else:
                    self.content_types.overrides[name] = content_type
        return FragmentedPartWriterHandle(name, content_type, self)

    def close(self) -> None:
        self.write_content_types_stream()

    def __enter__(self):
        return self

    def __exit__(self, _exc_type, _exc_val, _exc_tb):
        self.close()

    def write_content_types_stream(self) -> None:
        # We do currently not support interleaved Content Types Streams yet
        if self.content_types_stream_name is None:
            raise RuntimeError("Physical Package Format uses native content types. No Content Types Stream is required")
        if self.content_types_written:
            return
        with self.create_item(self.content_types_stream_name, "application/xml") as i:
            self.content_types.write_xml(i)
        self.content_types_written = True

    @abc.abstractmethod
    def create_item(self, name: str, content_type: str) -> IO[bytes]:
        pass

    @staticmethod
    def _write_relationships(rels_part: IO[bytes], relationships: Iterable["OPCRelationship"]) -> None:
        with etree.xmlfile(rels_part, encoding="UTF-8") as xf:
            with xf.element(RELATIONSHIPS_XML_NAMESPACE + "Relationships",
                            nsmap={None: RELATIONSHIPS_XML_NAMESPACE[1:-1]}):
                for relationship in relationships:
                    xf.write(etree.Element(RELATIONSHIPS_XML_NAMESPACE + 'Relationship', {
                        'Target': relationship.target,
                        'Id': relationship.id,
                        'Type': relationship.type,
                        'TargetMode': relationship.target_mode.serialize()}))


class FragmentedPartWriterHandle:
    def __init__(self, name: str, content_type: str, writer: OPCPackageWriter):
        self.name: str = name
        self.content_type = content_type
        self.writer: OPCPackageWriter = writer
        self.fragment_number: int = 0
        self.finished = False

    def open(self, last: bool = False) -> IO[bytes]:
        if self.finished:
            raise RuntimeError("Fragmented Part {} has already been finished".format(self.name))
        f = self.writer.create_item("{}/[{}]{}.piece".format(self.name, self.fragment_number, ".last" if last else ""),
                                    self.content_type)
        self.fragment_number += 1
        self.finished = last
        return f


class OPCTargetMode(enum.Enum):
    INTERNAL = 1
    EXTERNAL = 2

    @classmethod
    def from_serialization(cls, serialization: str) -> "OPCTargetMode":
        return cls[serialization.upper()]

    def serialize(self) -> str:
        return self.name.capitalize()


class OPCRelationship(NamedTuple):
    id: str
    type: str
    target: str
    target_mode: OPCTargetMode


class ContentTypesData:
    XML_NAMESPACE = "{http://schemas.openxmlformats.org/package/2006/content-types}"

    def __init__(self):
        self.default_types: Dict[str, str] = {}   # dict mapping file extensions to mime types
        self.overrides: Dict[str, str] = {}   # dict mapping part names to mime types

    def get_content_type(self, part_name: str) -> Optional[str]:
        part_name = normalize_part_name(part_name)
        if part_name in self.overrides:
            return self.overrides[part_name]
        extension = part_name.split("/")[-1].split(".")[-1]
        if extension in self.default_types:
            return self.default_types[extension]
        return None

    @classmethod
    def from_xml(cls, content_types_file: IO[bytes]) -> "ContentTypesData":
        result = cls()
        for _event, elem in etree.iterparse(content_types_file):
            if elem.tag == cls.XML_NAMESPACE + "Default":
                result.default_types[elem.attrib["Extension"].lower()] = elem.attrib["ContentType"]
            elif elem.tag == cls.XML_NAMESPACE + "Override":
                result.overrides[normalize_part_name(elem.attrib["PartName"])] = elem.attrib["ContentType"]
        return result

    def write_xml(self, file: IO[bytes]) -> None:
        with etree.xmlfile(file, encoding="UTF-8") as xf:
            with xf.element(self.XML_NAMESPACE + "Types",
                            nsmap={None: self.XML_NAMESPACE[1:-1]}):
                for extension, content_type in self.default_types.items():
                    xf.write(etree.Element(self.XML_NAMESPACE + 'Default',
                                           {'Extension': extension, 'ContentType': content_type}))
                for part_name, content_type in self.overrides.items():
                    xf.write(etree.Element(self.XML_NAMESPACE + 'Override',
                                           {'PartName': part_name, 'ContentType': content_type}))


def _rels_part_for(part_name: str) -> str:
    name_parts = part_name.split("/")
    return "/".join(name_parts[:-1] + ["_rels", name_parts[-1]]) + ".rels"


def normalize_part_name(part_name: str) -> str:
    """ Converts a part name to the normalized URI representation (i.e. uschars are %-encoded) and to lowercase """
    part_name = urllib.parse.quote(part_name, safe='/#%[]=:;$&()+,!?*@\'~')
    return part_name.lower()


RE_PART_NAME = re.compile(r'^(/[A-Za-z0-9\-\._~%:@!$&\'()*+,;=]*[A-Za-z0-9\-_~%:@!$&\'()*+,;=])+$')
RE_PART_NAME_FORBIDDEN = re.compile(r'%5c|%2f', re.IGNORECASE)


def check_part_name(part_name: str) -> None:
    if not RE_PART_NAME.match(part_name):
        raise ValueError("{} is not an URI path with multiple segments (each not empty and not starting with '.') "
                         "or not starting with '/' or ending wit '/'".format(repr(part_name)))
    if RE_PART_NAME_FORBIDDEN.match(part_name):
        raise ValueError("{} contains URI encoded '/' or '\\'.".format(repr(part_name)))


def part_realpath(part_name: str, source_part_name: str) -> str:
    """ Get an absolute part name from a relative part name (e.g. from a relationship) """
    path_segments = part_name.split("/")
    result = source_part_name.split("/")[:-1]
    for segment in path_segments:
        if segment == '.':
            pass
        elif segment == '..':
            result.pop()
        else:
            result.append(segment)
    return "/".join(result)
