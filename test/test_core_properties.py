import datetime
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
