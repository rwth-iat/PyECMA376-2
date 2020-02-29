import datetime
import io
import os.path
import unittest

import pyecma376_2


class CorePropertiesTest(unittest.TestCase):
    def test_parsing_core_properties(self) -> None:
        file_name = os.path.join(os.path.dirname(__file__), "CorePropertiesExample.xml")
        with open(file_name, 'rb') as f:
            cp = pyecma376_2.OPCCoreProperties.from_xml(f)

        self.assertEqual("Reviewed", cp.contentStatus)
        self.assertEqual("OPC Core Properties", cp.title)
        self.assertEqual(datetime.date(2005, 6, 12), cp.created)
        self.assertListEqual([('en-US', "color"), ('en-CA', "colour"), ('fr-FR', "couleur")],
                             cp.keywords)  # type: ignore

    def test_serializing_core_properties(self) -> None:
        cp = pyecma376_2.OPCCoreProperties()
        cp.contentStatus = "Reviewed"
        cp.title = "OPC Core Properties"
        cp.created = datetime.date(2005, 6, 12)
        cp.keywords = [('en-US', "color"), ('en-CA', "colour"), ('fr-FR', "couleur")]

        f = io.BytesIO()
        cp.write_xml(f)
        f.seek(0)

        cp2 = pyecma376_2.OPCCoreProperties.from_xml(f)

        self.assertEqual("Reviewed", cp2.contentStatus)
        self.assertEqual("OPC Core Properties", cp2.title)
        self.assertEqual(datetime.date(2005, 6, 12), cp2.created)
        self.assertListEqual([('en-US', "color"), ('en-CA', "colour"), ('fr-FR', "couleur")],
                             cp2.keywords)  # type: ignore

    def test_read_empty_doc_properties(self) -> None:
        file_name = os.path.join(os.path.dirname(__file__), "empty_document.docx")
        with pyecma376_2.ZipPackageReader(file_name) as reader:
            core_properties = reader.get_core_properties()
        self.assertEqual("Michael Thies", core_properties.creator)
        self.assertEqual(2020, core_properties.created.year)
