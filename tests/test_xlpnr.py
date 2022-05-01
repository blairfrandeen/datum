import unittest
import xlwings as xw
import json

from xl_populate_named_ranges import xw_get_named_range

class TestXL(unittest.TestCase):
    def setUp(self):
        self._load_json_test()
        self._load_excel_test()

    def _load_json_test(self):
        self.json_file = "nx_measurements.test.json"

    def _load_excel_test(self):
        self.app = xw.App(visible=False)
        self.workbook = self.app.books.open('tests/datum_excel_tests.xlsx')
        self.app.books[0].close()

    def test_xl(self):
        for name in self.workbook.names:
            print(name.name, end='')
            print(f"\t{name.refers_to_range.value}")
        result = 6
        self.assertEqual(result, 6)

    def test_get_named_range(self):
        self.assertIsNone(xw_get_named_range(self.workbook,\
            "non-existant range"))
        self.assertIsNotNone(xw_get_named_range(self.workbook,\
            "DIPSTICK"))
        self.assertIsNone(xw_get_named_range(self.workbook,\
            "missing_ref"))

    def tearDown(self):
        for book in self.app.books:
            book.close()
        self.app.quit()
        

if __name__ == '__main__':
    unittest.main()