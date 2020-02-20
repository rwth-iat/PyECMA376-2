import abc
import enum
import io
import re
import urllib.parse
import xml.etree.ElementTree as ETree
from typing import BinaryIO, Sequence, Dict, Iterable, NamedTuple, Optional, IO, Generator

RE_RELS_PARTS = re.compile(r'^(.*/)_rels/([^/]*).rels$', re.IGNORECASE)
RE_FRAGMENT_ITEMS = re.compile(r'^(.*)/\[(\d+)\](.last)?.piece$', re.IGNORECASE)
RELATIONSHIPS_XML_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"


class OPCPackageReader(metaclass=abc.ABCMeta):
    content_types_stream_name: Optional[str] = None

    def __init__(self):
        # TODO provide single mechanism for getting content types (including physical mappings with native content type)
        if self.content_types_stream_name is not None:
            self.content_types = ContentTypesData()

    def _init_data(self):
        """ Part of the initializer which should be called after opening the package """
        # TODO initialize dict of part names
        if self.content_types_stream_name is not None:
            with self.open_part(self.content_types_stream_name) as part:
                self.content_types = ContentTypesData.from_xml(part)

    def list_parts(self, include_rels_parts=False) -> Iterable[str]:
        for item_name in self.list_items():
            if not include_rels_parts and RE_RELS_PARTS.match(item_name):
                continue
            fragment_match = RE_FRAGMENT_ITEMS.match(item_name)
            if fragment_match:
                if fragment_match[2] != 0:
                    continue
                part_name = fragment_match[1]
            else:
                part_name = item_name
            part_name = normalize_part_name(part_name)
            if part_name == self.content_types_stream_name:
                continue
            yield part_name

    def open_part(self, name: str) -> IO[bytes]:
        # TODO create dict of normalized part names and parts in the package to lookup correct spelling
        try:
            return self.open_item(name)
        except KeyError:
            try:
                return FragmentedPartReader(name, self)  # type: ignore
            except KeyError as e:
                raise KeyError("Could not find part {} in package".format(name))

    def get_relationships(self) -> Generator["OPCRelationship"]:
        return self._read_relationships(self.open_part("/_rels/.rels"))

    def get_part_relationships(self, part_name: str) -> Generator["OPCRelationship", None, None]:
        return self._read_relationships(self.open_part(_rels_part_for(part_name)))

    def close(self) -> None:
        pass

    def __enter__(self):
        return self

    def __exit__(self, _exc_type, _exc_val, _exc_tb):
        self.close()

    @staticmethod
    def _read_relationships(rels_part: IO[bytes]) -> Generator["OPCRelationship", None, None]:
        for _event, elem in ETree.iterparse(rels_part):
            if elem.tag == "Relationship":
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
        self._current_item_handle: IO[bytes] = self._open_next_item()

    def _open_next_item(self) -> IO[bytes]:
        try:
            return self._reader.open_item("{}/[{}].piece".format(self._name, self._fragment_number))
        except KeyError:
            self.finished = True
            try:
                return self._reader.open_item("{}/[{}].last.piece".format(self._name, self._fragment_number))
            except KeyError as e:
                raise KeyError("Fragment {} of part {} is missing in package".format(self._fragment_number, self._name))

    def seekable(self) -> bool:
        return False

    def read(self, size: int = -1) -> Optional[bytes]:
        result = self._current_item_handle.read(size)
        while result is not None and (size == -1 or 0 == len(result) < size) and not self.finished:
            self._current_item_handle.close()
            self._current_item_handle = self._open_next_item()
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

    def write_package_relationships(self, relationships: Iterable["OPCRelationship"]) -> None:
        with self.create_item("/_rels/.rels", "application/xml") as i:
            self._write_relationships(i, relationships)

    def open_part(self, name: str, content_type: str) -> IO[bytes]:
        part_name = normalize_part_name(name)
        check_part_name(part_name)
        if self.content_types_stream_name is not None:
            if self.content_types.get_content_type(name) != content_type:
                if self.content_types_written:
                    raise RuntimeError("Content Type of part {} is not set correctly but ContentTypeStream has been "
                                       "written already.".format(name))
                else:
                    self.content_types.overrides[name] = content_type
        return self.create_item(name, content_type)

    def write_part_relationships(self, part_name: str, relationships: Iterable["OPCRelationship"]) -> None:
        part_name = normalize_part_name(part_name)
        check_part_name(part_name)
        with self.open_part(_rels_part_for(part_name), "application/xml") as i:
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
        # TODO this does not support interleaved Content Types Streams yet
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
        root = ETree.Element("Relationships")
        for relationship in relationships:
            ETree.SubElement(root, 'Relationship', {
                'Target': relationship.target,
                'Id': relationship.id,
                'Type': relationship.type,
                'TargetMode': relationship.target_mode.serialize()})
        ETree.ElementTree(root).write(rels_part, default_namespace=RELATIONSHIPS_XML_NAMESPACE)


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
    XML_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types"

    def __init__(self):
        self.default_types: Dict[str, str] = {}   # dict mapping file extensions to mime types
        self.overrides: Dict[str, str] = {}   # dict mapping part names to mime types

    def get_content_type(self, part_name: str) -> Optional[str]:
        part_name = normalize_part_name(part_name)
        if part_name in self.overrides:
            return self.overrides[part_name]
        extension = part_name.split(".")[-1]
        if extension in self.default_types:
            return self.default_types[extension]
        return None

    @classmethod
    def from_xml(cls, content_types_file: IO[bytes]) -> "ContentTypesData":
        result = cls()
        for _event, elem in ETree.iterparse(content_types_file):
            if elem.tag == "Default":
                result.default_types[elem.attrib["Extension"].lower()] = elem.attrib["ContentType"]
            elif elem.tag == "Override":
                result.default_types[normalize_part_name(elem.attrib["PartName"])] = elem.attrib["ContentType"]
        return result

    def write_xml(self, file: IO[bytes]) -> None:
        root = ETree.Element("Types")
        for extension, content_type in self.default_types:
            ETree.SubElement(root, 'Default', {'Extension': extension, 'ContentType': content_type})
        for part_name, content_type in self.overrides:
            ETree.SubElement(root, 'Override', {'PartName': part_name, 'ContentType': content_type})
        ETree.ElementTree(root).write(file, default_namespace=self.XML_NAMESPACE)


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
