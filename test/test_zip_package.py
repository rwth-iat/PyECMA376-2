import os.path
import unittest

import lxml.etree as etree  # type: ignore
from pyecma376_2 import zip_package


class TestZipReader(unittest.TestCase):
    def test_reading_empty_docx(self) -> None:
        file_name = os.path.join(os.path.dirname(__file__), "empty_document.docx")
        reader = zip_package.ZipPackageReader(file_name)

        parts = list(reader.list_parts(True))
        self.assertGreater(len(parts), 0)
        package_rels = list(reader.get_raw_relationships())
        self.assertGreater(len(package_rels), 0)
        document_part = reader.get_related_parts_by_type()[
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'][0]
        self.assertEqual("/word/document.xml", document_part)
        document_rels = list(reader.get_raw_relationships("/word/document.xml"))
        self.assertGreater(len(document_rels), 0)

        reader.close()

    def test_reading_fragmented(self) -> None:
        file_name = os.path.join(os.path.dirname(__file__), "fragmented.docx")
        reader = zip_package.ZipPackageReader(file_name)

        self.assertIn("/word/document.xml", (n for n, ct in reader.list_parts()))
        with reader.open_part("/word/document.xml") as doc:
            etree.parse(doc)

        reader.close()


class TestZipWriter(unittest.TestCase):
    def test_rewrite_docx(self):
        file_name = os.path.join(os.path.dirname(__file__), "empty_document.docx")
        file_name_new = os.path.join(os.path.dirname(__file__), "empty_document_new.docx")
        reader = zip_package.ZipPackageReader(file_name)
        writer = zip_package.ZipPackageWriter(file_name_new)

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

        with zip_package.ZipPackageReader(file_name_new) as reader:
            parts = list(reader.list_parts(True))
            self.assertGreater(len(parts), 0)
            package_rels = list(reader.get_raw_relationships())
            self.assertGreater(len(package_rels), 0)
