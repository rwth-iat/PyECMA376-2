# Copyright 2020 by Michael Thies
#
# Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with
# the License. You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the
# specific language governing permissions and limitations under the License.
""" Tests for the main Reader/Writer functionality of the PyECMA376_2 package"""
import datetime
import os.path
import tempfile
import unittest

import lxml.etree as etree  # type: ignore
import pyecma376_2


class TestZipReader(unittest.TestCase):
    def test_reading_empty_docx(self) -> None:
        file_name = os.path.join(os.path.dirname(__file__), "empty_document.docx")
        reader = pyecma376_2.ZipPackageReader(file_name)

        parts = list(reader.list_parts(True))
        self.assertGreater(len(parts), 0)
        package_rels = list(reader.get_raw_relationships())
        self.assertGreater(len(package_rels), 0)
        document_part = reader.get_related_parts_by_type()[
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'][0]
        self.assertEqual("/word/document.xml", document_part)
        core_properties_part = reader.get_related_parts_by_type()[
            pyecma376_2.RELATIONSHIP_TYPE_CORE_PROPERTIES][0]
        self.assertEqual(pyecma376_2.DEFAULT_CORE_PROPERTIES_NAME, core_properties_part)
        document_rels = list(reader.get_raw_relationships("/word/document.xml"))
        self.assertGreater(len(document_rels), 0)

        reader.close()

    def test_reading_fragmented(self) -> None:
        file_name = os.path.join(os.path.dirname(__file__), "fragmented.docx")
        reader = pyecma376_2.ZipPackageReader(file_name)

        self.assertIn("/word/document.xml", (n for n, ct in reader.list_parts()))
        with reader.open_part("/word/document.xml") as doc:
            etree.parse(doc)

        reader.close()


class TestZipWriter(unittest.TestCase):
    def test_rewrite_docx(self) -> None:
        fd, new_filename = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        file_name = os.path.join(os.path.dirname(__file__), "empty_document.docx")
        reader = pyecma376_2.ZipPackageReader(file_name)
        writer = pyecma376_2.ZipPackageWriter(new_filename)

        writer.write_relationships(reader.get_raw_relationships())
        for name, content_type in reader.list_parts():
            with writer.open_part(name, content_type) as w:
                with reader.open_part(name) as r:
                    w.write(r.read())
            relationships = list(reader.get_raw_relationships(name))
            if relationships:
                writer.write_relationships(relationships, name)
        writer.close()
        reader.close()

        with pyecma376_2.ZipPackageReader(new_filename) as reader:
            parts = list(reader.list_parts(True))
            self.assertGreater(len(parts), 0)
            package_rels = list(reader.get_raw_relationships())
            self.assertGreater(len(package_rels), 0)

        os.unlink(new_filename)

    def test_write_example(self) -> None:
        fd, filename = tempfile.mkstemp(suffix=".myx")
        os.close(fd)
        with pyecma376_2.ZipPackageWriter(filename) as writer:
            # Add a part
            with writer.open_part("/example/document.txt", "text/plain") as part:
                part.write("Lorem ipsum dolor sit amet.".encode())

            # Write core properties (meta data)
            # To make those work, we need to add the RELATIONSHIP_TYPE_CORE_PROPERTIES relationship below.
            cp = pyecma376_2.OPCCoreProperties()
            cp.created = datetime.datetime.now()
            with writer.open_part(pyecma376_2.DEFAULT_CORE_PROPERTIES_NAME, "application/xml") as part:
                cp.write_xml(part)

            # Write the packages root relationships
            writer.write_relationships([
                pyecma376_2.OPCRelationship("r1", "http://example.com/my-package-relationship-id", "http://example.com",
                                            pyecma376_2.OPCTargetMode.EXTERNAL),
                pyecma376_2.OPCRelationship("r2", "http://example.com/my-document-rel", "example/document.txt",
                                            pyecma376_2.OPCTargetMode.INTERNAL),
                pyecma376_2.OPCRelationship("r3", pyecma376_2.RELATIONSHIP_TYPE_CORE_PROPERTIES,
                                            pyecma376_2.DEFAULT_CORE_PROPERTIES_NAME,
                                            pyecma376_2.OPCTargetMode.INTERNAL),
            ])
        os.unlink(filename)

    def test_write_fragmented(self) -> None:
        """ Test the writing of interleaved parts. Also test premature writing of ContentTypes (using defaults)"""
        fd, filename = tempfile.mkstemp(suffix=".myx")
        os.close(fd)

        with pyecma376_2.ZipPackageWriter(filename) as writer:
            writer.content_types.default_types['txt'] = "text/plain"
            writer.write_content_types_stream()

            handle = writer.create_fragmented_part("/foo.txt", "text/plain")
            with handle.open() as f:
                f.write(b"Hello, ")
            with writer.open_part("/bar.txt", "text/plain") as f:
                f.write(b"Other part's contents")
            with handle.open(last=True) as f:
                f.write(b"World!")

        with pyecma376_2.ZipPackageReader(filename) as reader:
            with reader.open_part("/foo.txt") as f:
                foo_content = f.read()
            with reader.open_part("/bar.txt") as f:
                bar_content = f.read()
        self.assertEqual(b"Hello, World!", foo_content)
        self.assertEqual(b"Other part's contents", bar_content)

        os.unlink(filename)

    def test_empty_file(self) -> None:
        fd, filename = tempfile.mkstemp(suffix=".myx")
        os.close(fd)

        writer = pyecma376_2.ZipPackageWriter(filename)
        writer.close()

        with pyecma376_2.ZipPackageReader(filename) as reader:
            self.assertEqual([], list(reader.list_parts()))
            self.assertEqual([], list(reader.get_related_parts_by_type()))

            core_properties = reader.get_core_properties()
            self.assertIsNone(core_properties.created)
