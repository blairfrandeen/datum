import datetime
import unittest
import logging
from unittest.mock import patch

import xlwings as xw

import datum.xl_populate_named_ranges as xlpnr

xlpnr.logger = logging.getLogger('testLogger')

# TODO: Restructure so we don't need to call out tests
# folder for test data files
TEST_JSON_FILE = "tests/json/nx_measurements_test.json"
JSON_WITHOUT_USEFUL_DATA = "tests/json/useless.json"
TEST_EXCEL_WB = "tests/xl/datum_excel_tests.xlsx"


class TestJSON(unittest.TestCase):
    def test_get_json_measurement_names(self):
        valid_names = xlpnr.get_json_measurement_names(TEST_JSON_FILE)
        self.assertIsInstance(valid_names, dict)

        no_names = xlpnr.get_json_measurement_names(JSON_WITHOUT_USEFUL_DATA)
        self.assertIsNone(no_names)

        non_existant_json = xlpnr.get_json_measurement_names("DNE.json")
        self.assertIsNone(non_existant_json)

        broken_json = xlpnr.get_json_measurement_names("tests/xl/broken.json")
        self.assertIsNone(broken_json)

        no_measurements = xlpnr.get_json_measurement_names("tests/xl/no_measurements.json")
        self.assertIsNone(no_measurements)


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
        self.assertIsNone(xlpnr.xw_get_named_range(self.workbook, "non-existant range"))
        self.assertIsInstance(
            xlpnr.xw_get_named_range(self.workbook, "DIPSTICK.angle"), xw.Range
        )
        self.assertIsNone(xlpnr.xw_get_named_range(self.workbook, "missing_ref"))

    def test_get_workbook_range_names(self):
        valid_workbook = xlpnr.get_workbook_range_names(self.workbook)
        self.assertIsInstance(valid_workbook, dict)

        blank_wb = self.app.books.add()
        empty_wb = xlpnr.get_workbook_range_names(blank_wb)
        self.assertIsNone(empty_wb)
        blank_wb.close()

    def test_write_named_range(self):
        testvalue = 700_000
        xlpnr.write_named_range(self.workbook, "SURFACE_PAINTED.area", testvalue)
        result = xlpnr.read_named_range(self.workbook, "SURFACE_PAINTED.area")
        self.assertEqual(testvalue, result)

    def test_write_empty_range(self):
        xlpnr.write_named_range(self.workbook, "SURFACE_PAINTED.area", None)
        result = xlpnr.read_named_range(self.workbook, "SURFACE_PAINTED.area")
        self.assertIsNone(result)

    def test_read_named_range(self):
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
    
    def test_read_named_vector_range(self):
        horizontal_range = "AIR_NUT.point"
        vertical_range = "HOUSING.moments_of_inertia_centroidal"
        vertical_range_value = xlpnr.read_named_range(self.workbook, vertical_range)
        horizontal_range_value = xlpnr.read_named_range(self.workbook, horizontal_range)
        self.assertIsInstance(vertical_range_value, list)
        self.assertIsInstance(horizontal_range_value, list)

    def test_write_named_vector_range(self):
        horizontal_range = "AIR_NUT.point"
        vertical_range = "HOUSING.moments_of_inertia_centroidal"
        xlpnr.write_named_range(self.workbook, horizontal_range, [3.0, 2.11, 9.99])
        xlpnr.write_named_range(self.workbook, vertical_range, [0.012, 0.11, 0.99])
        vertical_result = xlpnr.read_named_range(self.workbook, vertical_range)
        horizontal_result = xlpnr.read_named_range(self.workbook, horizontal_range)
        self.assertEqual(vertical_result[1], 0.11)
        self.assertEqual(horizontal_result[2], 9.99)

    def test_get_workbook(self):
        non_existant_workbook = xlpnr.xw_get_workbook("DNE.xlsx")
        self.assertIsNone(non_existant_workbook)
        valid_workbook = xlpnr.xw_get_workbook("datum_excel_tests.xlsx")
        self.assertIsInstance(valid_workbook, xw.Book)

    def test_update(self):
        xlpnr.write_named_range(self.workbook, "DIPSTICK.angle", 95)  # should be: 90
        xlpnr.write_named_range(self.workbook, "GEARS.mass", 45)  # should be: 55.456

        # confirm overwrite automatically
        with patch("builtins.input", return_value="y"):
            xlpnr.update_named_ranges(self.json_file, self.workbook, backup=False)
        updated_value = xlpnr.read_named_range(self.workbook, "DIPSTICK.angle")
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
