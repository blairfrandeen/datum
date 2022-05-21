import datetime
import logging
import os
import unittest
from unittest.mock import patch

import pytest
import xlwings as xw

import datum.xl_populate_named_ranges as xlpnr

xlpnr.logger = logging.getLogger("testLogger")

TEST_JSON_FILE = "tests/json/nx_measurements_test.json"
JSON_WITHOUT_USEFUL_DATA = "tests/json/useless.json"
TEST_EXCEL_WB = "tests/xl/datum_excel_tests.xlsx"


def test_get_json_measurement_names():
    valid_names = xlpnr.get_json_measurement_names(TEST_JSON_FILE)
    assert isinstance(valid_names, dict)

    no_names = xlpnr.get_json_measurement_names(JSON_WITHOUT_USEFUL_DATA)
    assert no_names is None

    non_existant_json = xlpnr.get_json_measurement_names("DNE.json")
    assert non_existant_json is None

    broken_json = xlpnr.get_json_measurement_names("tests/json/broken.json")
    assert broken_json is None

    no_measurements = xlpnr.get_json_measurement_names(
        "tests/json/no_measurements.json"
    )
    assert no_measurements is None

class TestUtilities(unittest.TestCase):
    def test_check_dict_keys(self):
        test_dict = {
            "test_key1": ["item 1", "item 2"],
            "test_key2": [],
            "test_key4": [1, 2, 3],
            "test_non_list": 5,
        }
        assert xlpnr.check_dict_keys(test_dict, ["test_key1"]) is True
        assert xlpnr.check_dict_keys(test_dict, ["test_key2"]) is False
        assert xlpnr.check_dict_keys(test_dict, ["test_key3"]) is False
        assert xlpnr.check_dict_keys(test_dict, ["test_key1", "test_key2"]) is False
        assert xlpnr.check_dict_keys(test_dict, ["test_key1", "test_key3"]) is False
        assert xlpnr.check_dict_keys(test_dict, ["test_key1", "test_key4"]) is True
        assert xlpnr.check_dict_keys(test_dict, ["test_non_list"]) is True


    def test_report_difference(self):
        test_date_1 = datetime.datetime(1984, 6, 17)
        test_date_2 = datetime.datetime(1982, 6, 25)
        # test comparison between any combination of
        # int, float, string, date, and None
        int_str = (7, "seven")
        int_int = (7, 8)
        int_flt = (7, 7.5)
        flt_flt = (5.5, 7.5)
        flt_str = (7.5, "seven and a half")
        str_str = ("seven", "eight")
        non_non = (None, None)
        int_non = (7, None)
        flt_non = (7.5, None)
        str_non = ("seven", None)
        dat_str = (test_date_1, "Blair's Birthday")
        dat_non = (test_date_1, None)
        dat_int = (test_date_1, 7)
        dat_flt = (test_date_1, 7.5)
        dat_dat = (test_date_1, test_date_2)
        zer_int = (0, 7)
        zer_flt = (0, 7.5)

        assert xlpnr.report_difference(*int_int) == 1 / 7
        assert xlpnr.report_difference(*int_flt) == (7.5 - 7) / 7
        assert xlpnr.report_difference(*flt_flt) == 2 / 5.5
        assert xlpnr.report_difference(*dat_dat) == datetime.timedelta(days=-723)

        assert xlpnr.report_difference(*zer_int) is None
        assert xlpnr.report_difference(*zer_flt) is None
        assert xlpnr.report_difference(*non_non) is None
        assert xlpnr.report_difference(*int_non) is None
        assert xlpnr.report_difference(*flt_non) is None
        assert xlpnr.report_difference(*str_non) is None
        assert xlpnr.report_difference(*dat_non) is None
        assert xlpnr.report_difference(*int_str) is None
        assert xlpnr.report_difference(*flt_str) is None
        assert xlpnr.report_difference(*str_str) is None
        assert xlpnr.report_difference(*dat_str) is None
        assert xlpnr.report_difference(*dat_int) is None
        assert xlpnr.report_difference(*dat_flt) is None

    def test_print_columns(self):
        column_widths = [42, 15, 15, 15]
        column_headings = ["PARAMETER", "OLD VALUE", "NEW VALUE", "PERCENT CHANGE"]
        underlines = ["-" * 20, "-" * 12, "-" * 12, "-" * 15]
        floats = ["Your mom lol", 42.5, 95.2, 0.738]
        no_change = ["Your mom lol", "she sits", "around the house", None]
        mostly_none = ["Nothing", None, None, None]
        too_many_values = list(range(10))
        print()  # newline
        xlpnr.print_columns(column_widths, column_headings)
        xlpnr.print_columns(column_widths, underlines)
        xlpnr.print_columns(column_widths, floats)
        xlpnr.print_columns(column_widths, no_change)
        xlpnr.print_columns(column_widths, mostly_none)
        test_date_1 = datetime.datetime(1984, 6, 17)
        test_date_2 = datetime.datetime(1982, 6, 25)
        xlpnr.print_columns(
            column_widths,
            ["Birthdays", test_date_1, test_date_2, test_date_2 - test_date_1],
        )
        with pytest.raises(IndexError):
            xlpnr.print_columns(column_widths, too_many_values)

        with pytest.raises(TypeError):
            xlpnr.print_columns(["1", 2, 3.4, "five"], floats)
        assert 1

    def test_flattened_list(self):
        list_1d = [1.3, 2, "three", 4.2, 5]
        list_2d = [list_1d, list_1d]
        list_mixed = list_1d + [list_1d]
        assert list(xlpnr.flatten_list(list_1d)) == list_1d
        assert list(xlpnr.flatten_list(list_2d)) == list_1d + list_1d
        assert list(xlpnr.flatten_list(list_mixed)) == list_1d + list_1d


class TestXL(unittest.TestCase):
    def setUp(self):
        self._load_json_test()
        self._load_excel_test()

    def _load_json_test(self):
        self.json_file = TEST_JSON_FILE

    def _load_excel_test(self):
        """Open an invisible Excel workbook to execute tests."""
        # visible=False tag will run tests in background
        # without opening Excel window
        self.app = xw.App(visible=False)
        self.workbook = self.app.books.open(TEST_EXCEL_WB)

        # starting self.app will open a blank workbook that
        # isn't needed. Close it prior to tests
        self.app.books[0].close()

    def test_get_workbook_kvp(self):
        valid_workbook = xlpnr.get_workbook_key_value_pairs(self.workbook)
        self.assertIsInstance(valid_workbook, dict)

        blank_wb = self.app.books.add()
        empty_wb = xlpnr.get_workbook_key_value_pairs(blank_wb)
        self.assertIsNone(empty_wb)
        blank_wb.close()

    def test_write_named_range(self):
        testvalue = 700_000
        assert (
            xlpnr.write_named_range(self.workbook, "SURFACE_PAINTED.area", testvalue)
            == testvalue
        )

        with pytest.raises(KeyError):
            xlpnr.write_named_range(self.workbook, "NON-EXISTANT-RANGE", testvalue)
        with pytest.raises(TypeError):
            xlpnr.write_named_range(self.workbook, "missing_ref", testvalue)

        illegal_dict = {"kivo": "stinker", "layla": "earflops"}
        self.assertIsNone(
            xlpnr.write_named_range(self.workbook, "SURFACE_PAINTED.area", illegal_dict)
        )

        illegal_tuple = (1, 2, 3)
        self.assertIsNone(
            xlpnr.write_named_range(
                self.workbook, "SURFACE_PAINTED.area", illegal_tuple
            )
        )

    def test_write_empty_range(self):
        assert (
            xlpnr.write_named_range(self.workbook, "SURFACE_PAINTED.area", None) is None
        )

    def test_write_named_vector_range(self):
        horizontal_range = "AIR_NUT.point_1"
        vertical_range = "HOUSING.moments_of_inertia_centroidal"
        horizontal_result = xlpnr.write_named_range(
            self.workbook, horizontal_range, [3.0, 2.11, 9.99]
        )
        vertical_result = xlpnr.write_named_range(
            self.workbook, vertical_range, [0.012, 0.11, 0.99]
        )
        self.assertEqual(vertical_result[1], 0.11)
        self.assertEqual(horizontal_result[2], 9.99)

    def test_wrong_size_range(self):
        small_range = "too_small_range"
        large_range = "Range_that_s_too_large"
        wrong_size_list = [[1, 2, 3, 4, 5], [6, 7, 8, 9, 10]]
        self.assertEqual(
            xlpnr.write_named_range(self.workbook, small_range, wrong_size_list), [1, 2]
        )

        self.assertEqual(
            xlpnr.write_named_range(self.workbook, large_range, wrong_size_list),
            list(xlpnr.flatten_list(wrong_size_list)),
        )
        self.assertIsNone(self.workbook.names[large_range].refers_to_range.value[2][0])
        # self.assertIsNone(xlpnr.read_named_range(self.workbook, large_range)[2][0])

    def test_update(self):
        xlpnr.write_named_range(self.workbook, "DIPSTICK.angle", 95)  # should be: 90
        xlpnr.write_named_range(self.workbook, "GEARS.mass", 45)  # should be: 55.456

        # confirm overwrite automatically
        with patch("builtins.input", return_value="y"):
            range_undo_buffer = xlpnr.update_named_ranges(
                self.json_file, self.workbook, backup=False
            )
            assert range_undo_buffer["DIPSTICK.angle"] == 95
            assert range_undo_buffer["GEARS.mass"] == 45

        blank_wb = self.app.books.add()
        assert xlpnr.update_named_ranges(self.json_file, blank_wb) is None
        blank_wb.close()

        assert (
            xlpnr.update_named_ranges(JSON_WITHOUT_USEFUL_DATA, self.workbook) is None
        )

    def test_backup(self):
        backup_wb = xlpnr.backup_workbook(self.workbook, backup_dir="tests\\xl")
        assert os.path.isfile(backup_wb)

    def tearDown(self):
        """Close all open workbooks, and quit Excel."""
        for book in self.app.books:
            book.close()
        self.app.quit()


if __name__ == "__main__":
    unittest.main()
