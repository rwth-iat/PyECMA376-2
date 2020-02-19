import abc
import enum
import io
import re
from typing import BinaryIO, Sequence, Dict, Iterable, NamedTuple, Optional, IO

RE_RELS_PARTS = re.compile(r'^(.*/)_rels/([^/]*).rels$', re.IGNORECASE)
RE_FRAGMENT_ITEMS = re.compile(r'^(.*)/\[(\d+)\](.last)?.piece$', re.IGNORECASE)


class OPCPackageReader(metaclass=abc.ABCMeta):
    content_types_stream_name: Optional[str] = None

    def __init__(self):
        if self.content_types_stream_name is not None:
            self.content_types = ContentTypesData()  # TODO should be initialized by inheriting class after opening data

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
        # TODO handle fragmented parts
        try:
            return self.open_item(name)
        except KeyError:
            try:
                return FragmentedPartReader(name, self)  # type: ignore
            except KeyError as e:
                raise KeyError("Could not find part {} in package".format(name))

    def get_relationships(self) -> Sequence["OPCRelationship"]:
        return self._read_relationships(self.open_part("/_rels/.rels"))

    def get_part_relationships(self, part_name: str) -> Sequence["OPCRelationship"]:
        return self._read_relationships(self.open_part(_rels_part_for(part_name)))

    def close(self) -> None:
        pass

    def __enter__(self):
        return self

    def __exit__(self, _exc_type, _exc_val, _exc_tb):
        self.close()

    @staticmethod
    def _read_relationships(rels_part: IO[bytes]) -> Sequence["OPCRelationship"]:
        # TODO
        pass

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
        if self.content_types_stream_name is not None:
            if self.content_types.get_content_type(name) != content_type:
                if self.content_types_written:
                    raise RuntimeError("Content Type of part {} is not set correctly but ContentTypeStream has been "
                                       "written already.".format(name))
                else:
                    self.content_types.overrides[name] = content_type
        return self.create_item(name, content_type)

    def write_part_relationships(self, part_name: str, relationships: Iterable["OPCRelationship"]) -> None:
        with self.open_part(_rels_part_for(part_name), "application/xml") as i:
            self._write_relationships(i, relationships)

    def create_fragmented_part(self, name: str, content_type: str) -> "FragmentedPartWriterHandle":
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
    def _write_relationships(rels_part: IO[bytes], relationships: Iterable["OPCRelationship"]):
        # TODO
        pass


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


class OPCRelationship(NamedTuple):
    id: str
    type: str
    target: str
    target_mode: OPCTargetMode


class ContentTypesData:
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
    def from_xml(cls, content_types_file: BinaryIO) -> "ContentTypesData":
        # TODO
        pass

    def write_xml(self, file: IO[bytes]) -> None:
        # TODO
        pass


def _rels_part_for(part_name: str) -> str:
    name_parts = part_name.split("/")
    return "/".join(name_parts[:-1] + ["_rels", name_parts[-1]]) + ".rels"


def normalize_part_name(part_name: str) -> str:
    # TODO IRI → URI conversion
    return part_name.lower()
