import unittest
import xlwings as xw
import datetime

from unittest.mock import patch

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
        self.assertIsNone(xlpnr.xw_get_named_range(self.workbook,\
            "non-existant range"))
        self.assertIsNotNone(xlpnr.xw_get_named_range(self.workbook,\
            "DIPSTICK"))
        self.assertIsNone(xlpnr.xw_get_named_range(self.workbook,\
            "missing_ref"))


    def test_read_named_range(self):
        test_ranges = {
            'Test_Int': 4,
            'Test_Float': 3.141519,
            'Test_Str': 'Kivo is a dork',
            'Test_Date': datetime.datetime(2022, 5, 1, 0, 0),
            'Invalid_Range_Name': None
            }
        for named_range in test_ranges.keys():
            expected_value = test_ranges[named_range]
            read_value = xlpnr.read_named_range(self.workbook, named_range)
            self.assertEqual(read_value, expected_value)


    def test_get_workbook(self):
        non_existant_workbook = xlpnr.xw_get_workbook('DNE.xlsx')
        self.assertIsNone(non_existant_workbook)
        print([b for b in xw.books])
        valid_workbook = xlpnr.xw_get_workbook('datum_excel_tests.xlsx')
        self.assertIsInstance(valid_workbook, xw.Book)


    def test_update(self):
        # value from JSON measurement = 90
        xlpnr.write_named_range(self.workbook, "DIPSTICK", 95)

        # confirm overwrite automatically
        with patch('builtins.input', return_value='y'):
            xlpnr.update_named_ranges(self.json_file, self.workbook, backup=False)
        updated_value = xlpnr.read_named_range(self.workbook, "DIPSTICK")
        self.assertEqual(updated_value, 90)


    def tearDown(self):
        for book in self.app.books:
            book.save()
            book.close()
        self.app.quit()
        

if __name__ == '__main__':
    unittest.main()