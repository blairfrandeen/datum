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


class TestJson():
    def _empty_measurements(self, fh):
        return {
            "measurements": [
                {
                    "goose": "Frank",
                    "feathers": True
                }
            ]
        }

    def _empty_expressions(self, fh):
        return {
            "measurements": [
                {
                    "name": "Frank",
                    "expressions": [
                        {
                            "bbq": False,
                            "answer": 42,
                            "networth": None
                        }
                    ]
                }
            ]
        }

    def test_get_json_key_value_pairs(self, caplog, monkeypatch):
        # test valid JSON file
        valid_names = xlpnr.get_json_key_value_pairs(TEST_JSON_FILE)
        assert isinstance(valid_names, dict)
        from math import isclose
        # Test for a single value
        assert isclose
        (
            valid_names['GEARS.mass'],
            54.45,
            0.001
        )
        # Test for value with a named component
        assert isclose
        (
            valid_names['BOUNDING.length.x'],
            547.0,
            0.001
        )
        # Test for value from a list
        assert isclose
        (
            valid_names['HOUSING.first_moments_of_inertia.1'],
            120320.589,
            0.001
        )
        assert isinstance(valid_names['HOUSING.first_moments_of_inertia'], list)
                
        # test JSON file that doesn't exist
        non_existant_json = xlpnr.get_json_key_value_pairs("DNE.json")
        assert "Unable to open" in caplog.text
        assert non_existant_json is None

        # test JSON file that is corrupt
        broken_json = xlpnr.get_json_key_value_pairs("tests/json/broken.json")
        assert "is corrupt." in caplog.text
        assert broken_json is None

        # test JSON file without measurement key
        assert xlpnr.get_json_key_value_pairs(JSON_WITHOUT_USEFUL_DATA) is None
        assert 'Key "measurements" not found.' in caplog.text

        # test JSON file with measurement key that is empty
        assert xlpnr.get_json_key_value_pairs("tests/json/no_measurements.json") is None
        assert 'No "measurement" field' in caplog.text

        # test JSON that has measurement with no name or expression
        monkeypatch.setattr("json.load", self._empty_measurements)
        assert len(xlpnr.get_json_key_value_pairs(TEST_JSON_FILE)) == 0
        assert 'is missing name and/or expressions' in caplog.text

        # test JSON that has expression with no name, type, or value
        monkeypatch.setattr("json.load", self._empty_expressions)
        assert len(xlpnr.get_json_key_value_pairs(TEST_JSON_FILE)) == 0
        assert 'missing name/type/value fields in' in caplog.text


class TestUtilities():
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
        assert xlpnr.check_dict_keys('not a dict', ["test_non_list"]) is False

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

class MockXLName():
    def __init__(self, name):
        self.name = name
        self.refers_to = 'none'
        self.refers_to_range = MockXLRange(name, 5)

class MockXLRange():
    def __init__(self, name, value):
        self.name = name
        self.value = value

class MockWorkbook():
    def __init__(self, name, names):
        self.name = name
        self.names = [MockXLName(n) for n in names]

def test_get_workbook_kvp(caplog):
    mockwb = MockWorkbook('mock', ['_xlfn.mockfun', 'test'])
    assert '_xlfn.mockfun' not in xlpnr.get_workbook_key_value_pairs(mockwb)
    assert 'Skipping range _xlfn.mockfun' in caplog.text

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

        # mockwb = MockWorkbook('mock', ['_xlfn.mockfun', 'test'])
        # assert '_xlfn.mockfun' not in xlpnr.get_workbook_key_value_pairs(mockwb)
        # assert 'Skipping range _xlfn.mockfun' in caplog.text

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
