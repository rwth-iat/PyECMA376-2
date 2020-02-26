import unittest

from pyecma376_2.package_model import part_realpath


class TestUtilFunctions(unittest.TestCase):
    def test_part_realpath(self) -> None:
        self.assertEqual("/word/document.xml", part_realpath("word/document.xml", "/"))
        self.assertEqual("/word/document.xml", part_realpath("./document.xml", "/word/anotherPart.xml"))
        self.assertEqual("/document.xml", part_realpath("../document.xml", "/word/anotherPart.xml"))
