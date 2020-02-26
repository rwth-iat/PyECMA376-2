# Copyright 2020 Michael Thies
#
# Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with
# the License. You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the
# specific language governing permissions and limitations under the License.

import zipfile
from typing import Iterable, BinaryIO, IO

from . import package_model

CONTENT_TYPES_STREAM_NAME = "/[Content_Types].xml"


class ZipPackageReader(package_model.OPCPackageReader, zipfile.ZipFile):
    content_types_stream_name = package_model.normalize_part_name(CONTENT_TYPES_STREAM_NAME)

    def __init__(self, file):
        package_model.OPCPackageReader.__init__(self)
        zipfile.ZipFile.__init__(self, file)
        self._init_data()

    def list_items(self) -> Iterable[str]:
        return ["/" + name for name in self.namelist()
                if name[-1] != '/']

    def open_item(self, name: str) -> IO[bytes]:
        return self.open(name[1:])

    def close(self) -> None:
        zipfile.ZipFile.close(self)


class ZipPackageWriter(package_model.OPCPackageWriter, zipfile.ZipFile):
    content_types_stream_name = CONTENT_TYPES_STREAM_NAME

    def __init__(self, file):
        package_model.OPCPackageWriter.__init__(self)
        zipfile.ZipFile.__init__(self, file, mode='w')

    def close(self) -> None:
        package_model.OPCPackageWriter.close(self)
        zipfile.ZipFile.close(self)

    def create_item(self, name: str, content_type: str) -> IO[bytes]:
        return self.open(name[1:], mode='w')
