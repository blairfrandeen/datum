import datetime
import unittest
from unittest.mock import patch

import xlwings as xw

import xl_populate_named_ranges as xlpnr

TEST_JSON_FILE = "tests/nx_measurements_test.json"
TEST_EXCEL_WB = "tests/datum_excel_tests.xlsx"


class TestXL(unittest.TestCase):
    def setUp(self):
        self._load_json_test()
        self._load_excel_test()

    def _load_json_test(self):
        self.json_file = TEST_JSON_FILE

    def _load_excel_test(self):
        # visible=False tag will run tests in background
        # without opening Excel window
        self.app = xw.App(visible=False)
        self.workbook = self.app.books.open(TEST_EXCEL_WB)

        # starting self.app will open a blank workbook that
        # isn't needed. Close it prior to tests
        self.app.books[0].close()

    def test_get_named_range(self):
        print("TEST GET")
        self.assertIsNone(xlpnr.xw_get_named_range(self.workbook, "non-existant range"))
        self.assertIsInstance(
            xlpnr.xw_get_named_range(self.workbook, "DIPSTICK"),xw.Range
            )
        self.assertIsNone(xlpnr.xw_get_named_range(self.workbook, "missing_ref"))

    def test_write_named_range(self):
        print("WRITE_NORMAL")
        testvalue = 700_000
        xlpnr.write_named_range(self.workbook, "SURFACE_PAINTED.area", testvalue)
        result = xlpnr.read_named_range(self.workbook, "SURFACE_PAINTED.area")
        self.assertEqual(testvalue, result)
    
    def test_write_empty_range(self):
        print("WRITE EMPTY")
        xlpnr.write_named_range(self.workbook, "SURFACE_PAINTED.area", None)
        result = xlpnr.read_named_range(self.workbook, "SURFACE_PAINTED.area")
        self.assertIsNone(result)

    def test_read_named_range(self):
        print("READ NAMED RANGES")
        test_ranges = {
            "Test_Int": 4,
            "Test_Float": 3.141519,
            "Test_Str": "Kivo is a dork",
            "Test_Date": datetime.datetime(2022, 5, 1, 0, 0),
            "Invalid_Range_Name": None,
        }
        for named_range in test_ranges.keys():
            expected_value = test_ranges[named_range]
            read_value = xlpnr.read_named_range(self.workbook, named_range)
            self.assertEqual(read_value, expected_value)

    def test_get_workbook(self):
        non_existant_workbook = xlpnr.xw_get_workbook("DNE.xlsx")
        self.assertIsNone(non_existant_workbook)
        valid_workbook = xlpnr.xw_get_workbook("datum_excel_tests.xlsx")
        self.assertIsInstance(valid_workbook, xw.Book)

    def test_update(self):
        xlpnr.write_named_range(self.workbook, "DIPSTICK", 95)  # should be: 90
        xlpnr.write_named_range(self.workbook, "GEARS.mass", 45)  # should be: 55.456

        # confirm overwrite automatically
        with patch("builtins.input", return_value="y"):
            xlpnr.update_named_ranges(self.json_file, self.workbook, backup=False)
        updated_value = xlpnr.read_named_range(self.workbook, "DIPSTICK")
        self.assertEqual(updated_value, 90)
        updated_value = xlpnr.read_named_range(self.workbook, "GEARS.mass")
        self.assertAlmostEqual(updated_value, 54.456, places=2)

    def tearDown(self):
        for book in self.app.books:
            book.save()
            book.close()
        self.app.quit()


if __name__ == "__main__":
    unittest.main()
