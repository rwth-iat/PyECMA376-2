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
""" Tests for the utility functions of the PyECMA376_2 package"""

import unittest

from pyecma376_2.package_model import part_realpath, check_part_name


class TestUtilFunctions(unittest.TestCase):
    def test_part_realpath(self) -> None:
        self.assertEqual("/word/document.xml", part_realpath("word/document.xml", "/"))
        self.assertEqual("/word/document.xml", part_realpath("./document.xml", "/word/anotherPart.xml"))
        self.assertEqual("/document.xml", part_realpath("../document.xml", "/word/anotherPart.xml"))
        self.assertEqual("/document.xml", part_realpath("/document.xml", "/word/anotherPart.xml"))

    def test_check_part_name(self) -> None:
        check_part_name("/word/document.xml")
        # Part names must start with '/'
        with self.assertRaises(ValueError):
            check_part_name("word/document.xml")
        # Part names must not end with '/'
        with self.assertRaises(ValueError):
            check_part_name("/word/document/")
        # Unicode characters (IRIs) are not allowed, must be URI-encoded
        with self.assertRaises(ValueError):
            check_part_name("/schönes/Dokument")
        # Part names must not contain encoded '/'
        with self.assertRaises(ValueError):
            check_part_name("/ein%2fDokument.xml")
