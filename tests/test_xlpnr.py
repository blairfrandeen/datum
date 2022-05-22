import datetime
import logging
import os
from pathlib import Path

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
        with pytest.raises(TypeError):
            assert list(xlpnr.flatten_list({1: 'one', 2: 'two'})) == [1]

class MockXLName():
    def __init__(self, name):
        self.name = name
        self.refers_to = 'Sheet 1 A1 or something'
        self.refers_to_range = MockXLRange(name, 5)

class MockXLRange():
    def __init__(self, name, value):
        self.name = name
        self.value = value

class MockWorkbook():
    def __init__(self, name, names):
        self.name = name
        self.names = [MockXLName(n) for n in names]

class TestXL():
    mock_source_dict = {
        "k1": 12,
        "k2": 3,
        "k4": [4, 3, 2]
    }
    def _mock_target_dict(self, arg):
        return {
            "k1": 15,
            "k3": 9,
            "k4": [1, 2, 3]
        }
    @classmethod
    def setup_class(cls):
        cls._load_json_test(cls)
        cls._load_excel_test(cls)

    @classmethod
    def teardown_class(cls):
        """Close all open workbooks, and quit Excel."""
        for book in cls.app.books:
            book.close()
        cls.app.quit()

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

    def test_get_workbook_kvp(self, caplog):
        # test empty workbook with no names
        blank_wb = self.app.books.add()
        assert xlpnr.get_workbook_key_value_pairs(blank_wb) is None
        blank_wb.close()

        # test mockworkbook to ensure no _xlfn functions included
        mockwb = MockWorkbook('mock', ['_xlfn.mockfun', 'test'])
        assert '_xlfn.mockfun' not in xlpnr.get_workbook_key_value_pairs(mockwb)
        assert 'Skipping range _xlfn.mockfun' in caplog.text

        # add a name with a !#REF error to mock wb, verify it's skipped
        mock_ref_rng = MockXLName('bad_ref')
        mock_ref_rng.refers_to = "!#REF"
        del mock_ref_rng.refers_to_range
        mockwb.names.append(mock_ref_rng)

        xlpnr.get_workbook_key_value_pairs(mockwb)
        assert '#REF! error' in caplog.text

        valid_workbook = xlpnr.get_workbook_key_value_pairs(self.workbook)
        assert valid_workbook['Test_Int'] == 4
        assert valid_workbook['Test_Str'] == 'Kivo is a dork'
        assert valid_workbook['Test_Float'] == 3.141519
        assert valid_workbook['Test_Date'] == datetime.datetime(2022,5,1)
        assert valid_workbook['Empty_Range'] is None
        assert valid_workbook['Test_List'] == [1, 2, 3]
        assert valid_workbook['Test_Vector'] == [1, 2, 3]
        assert valid_workbook['Test_Empty_List'] == [None, None, None]
        assert valid_workbook['Test_Matrix'] == [
            [1, 2, 3],
            [4, 5, 6],
            [7, 8, 9]
        ]
        assert valid_workbook['Test_Empty_Matrix'] == [
            [None, None, None],
            [None, None, None],
            [None, None, None]
        ]

    def test_backup(self):
        backup_path = Path(TEST_EXCEL_WB.replace('.','_BACKUP.'))
        if os.path.isfile(backup_path):
            os.remove(backup_path)
        backup_wb = xlpnr.backup_workbook(self.workbook, backup_dir="tests\\xl")
        assert os.path.isfile(backup_wb)
        assert backup_path == backup_wb
        os.remove(backup_path)

    def test_update_named_ranges(self, monkeypatch, capsys):

        # Test for source dict as argument
        monkeypatch.setattr(xlpnr, 'get_workbook_key_value_pairs',
            self._mock_target_dict)
        monkeypatch.setattr(xlpnr, 'write_named_ranges', lambda *_: None)
        unr_ret = xlpnr.update_named_ranges(self.mock_source_dict, self.workbook)
        assert sorted(list(unr_ret.keys())) == ['k1', 'k4']
        assert unr_ret['k1'] == 15

        # Test for get_json_key_value_pairs to return None when given a string
        monkeypatch.setattr(xlpnr, 'get_json_key_value_pairs', lambda _: None)
        assert xlpnr.update_named_ranges(self.json_file, self.workbook) is None
        captured = capsys.readouterr()
        assert "No measurement data found in JSON file." in captured.out

        # Test for get_workbook_key_value pairs returns none
        monkeypatch.setattr(xlpnr, 'get_workbook_key_value_pairs', lambda _: None)
        assert xlpnr.update_named_ranges(self.json_file, self.workbook) is None
        captured = capsys.readouterr()
        assert "No named ranges in Excel file." in captured.out
        
    def test_write_named_ranges(self, monkeypatch, capsys, caplog):
        # disable functions
        for function in [
            'preview_named_range_update',
            'backup_workbook',
            'write_named_range'
        ]:
            monkeypatch.setattr(xlpnr, function, lambda *_: None)
        
        # Test with user abort
        monkeypatch.setattr('builtins.input', lambda _: 'n')
        xlpnr.write_named_ranges(
            self._mock_target_dict,
            self.mock_source_dict,
            self.workbook, "test"
        )
        captured = capsys.readouterr()
        assert ("Aborted.") in captured.out

        # Test with user confirm
        monkeypatch.setattr('builtins.input', lambda _: 'y')
        xlpnr.write_named_ranges(
            self._mock_target_dict,
            self.mock_source_dict,
            self.workbook, "test",
            backup = True
        )
        assert "Backed up to" in caplog.text
        assert "Source: test" in caplog.text

    def test_preview_named_range_update(self, monkeypatch, capsys):
        def _mock_print_cols(widths, values):
            print(f"{[v for v in values]}")
        def _mock_rep_diff(v1, v2):
            try:
                return v1 - v2
            except TypeError:
                return 'none'
        monkeypatch.setattr(xlpnr, 'print_columns', _mock_print_cols)
        monkeypatch.setattr(xlpnr, 'report_difference', _mock_rep_diff)
        mock_target_dict = self._mock_target_dict(None)
        mock_target_dict["k2"] = 3.00001
        self.mock_source_dict["k5"] = {"x": 1.1, "y": 2.2, "z": 3.3}
        mock_target_dict["k5"] =  1.2
        xlpnr.preview_named_range_update(
            mock_target_dict,
            self.mock_source_dict
        )
        captured = capsys.readouterr()
        # ensure items are being hidden
        assert "k2" not in captured.out

        # test that valid value is being printed
        assert "k1" in captured.out

        # test that lists are being printed
        assert "k4[0]" in captured.out

        # test that dicts are being printed
        assert "'k5[x]', 1.2" in captured.out
        
        # test that excel values aren't repeated if not list
        assert "'k5[y]', None" in captured.out

    def test_write_named_range(self, monkeypatch, caplog):
        # test writing to non-existant range
        with pytest.raises(KeyError):
            xlpnr.write_named_range(self.workbook, "NON-EXISTANT-RANGE", None)

        # test writing to range with missing reference
        with pytest.raises(TypeError):
            xlpnr.write_named_range(self.workbook, "missing_ref", None)

        # test writing to range with invalid type
        with pytest.raises(TypeError):
            xlpnr.write_named_range(self.workbook, "Test_Single_Value", MockXLName('test'))
        
        # test single values of all valid types
        single_test_values = [ None, 42, None, 3.14, None, 'good morning!', None,
            datetime.datetime(1984,6,17), None]
        for value in single_test_values:
            # test that function executes
            assert xlpnr.write_named_range(self.workbook, 'Test_Single_Value', value) == value
            # independent test to show it wrote correctly
            assert self.workbook.names['Test_Single_Value'].refers_to_range.value == value

        # test writing of lists
        list_test_values = [
            [ 1, 2, 3 ],
            [ None, None, None],
            [ 2.1, 4.2, 99.99 ],
            [ 'blair had', 'too much', 'wine tonight' ],
            [ datetime.datetime(2021,4,23,4,23,42), None, 42 ],
            [ None, None, None]
        ]
        def _mock_flatten(list):
            return list # this test method only uses flat lists to begin with.
        monkeypatch.setattr(xlpnr, 'flatten_list', _mock_flatten)
        for list in list_test_values:
            # test that function executes
            assert xlpnr.write_named_range(self.workbook, 'List_Test', list) == list
            # independent test to show it wrote correctly
            assert self.workbook.names['List_Test'].refers_to_range.value == list

        # test a list of things that won't work
        badlist = [ {1: 'one', 2: 'two'}, MockXLName('test') ]
        with pytest.raises(TypeError):
            xlpnr.write_named_range(self.workbook, 'List_Test', badlist)

        # test writing of matrices
        test_matrices = [ # assume already flattened
            [1,2,3,4,5,6,7,8,0],
            [1.1,2.2,3.3,4.4,5.3,0.93,1,None,.98],
            ['blair','needs','to','shave',None,'at','all','!!!', datetime.datetime(2022,5,21)],
            [None, None, None,None, None, None,None, None, None]
        ]
        for matrix in test_matrices:
            # test that function executes
            assert xlpnr.write_named_range(self.workbook, 'Matrix_Test', matrix) == matrix
            # independent test to show it wrote correctly
            # requires manual flattening, only check one value
            assert self.workbook.names['Matrix_Test'].refers_to_range.value[1][1] == matrix[4]

        # test writing to a range that's too small
        assert xlpnr.write_named_range(self.workbook, 'too_small_range', [1, 2, 3]) == [1, 2]
        assert xlpnr.write_named_range(self.workbook, 'too_small_range', [None, None]) == [None, None]
        target_range = self.workbook.names['too_small_range'].refers_to_range
        overflow_range = self.workbook.sheets[target_range.sheet].cells[
            target_range.row - 1, target_range.column + 1
        ]
        assert overflow_range.value is None
        assert 'Vector of length 3 will be truncated.' in caplog.text

        # test writing to a range that's too large
        assert xlpnr.write_named_range(self.workbook, 'too_large_range', [1, 2, 3]) == [1, 2, 3]
        assert xlpnr.write_named_range(self.workbook, 'too_large_range', [None, None, None]) == [None, None, None]
        target_range = self.workbook.names['too_large_range'].refers_to_range
        overflow_range = self.workbook.sheets[target_range.sheet].cells[
            target_range.row, target_range.column + 1
        ]
        assert overflow_range.value is None
        assert 'larger than required.' in caplog.text
