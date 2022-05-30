import datetime
import pytest
import sqlite3

import datum.xl_populate_named_ranges as xlpnr

class MockWorkbook:
    def __init__(self, name):
        self.name = name

PARAM_TABLE = """--sql
    CREATE TABLE parameters (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        param_key TEXT,
        param_value NUMERIC,
        gen_time DATETIME
    )"""

PARAM_VALUES = """--sql
    INSERT INTO parameters (param_key, param_value)
    VALUES (?, ?)
    """

METADATA_DICT = {
    "part_name": "datum_nx_test_measurements",
    "part_path": "C:\\Users\\frandeen\\Documents\\datum\\tests\\nx\\datum_nx_test_measurements.prt",
    "part_rev": None,
    "part_units": "Millimeters",
    "retrieval_date": "2022-05-08 09:27:57",
    "user": "frandeen",
    "computer": "MT-205700"
}

PARAMETER_DICT = {
    "test_float": 34.7,
    "test_int": 3823,
    "test_str": "Banana",
}

BAD_DICT = {
    "test_float": [12.3, 4, 53.2],
    "test_int": {1: "one", 3823: "a lot"},
    "test_str": MockWorkbook("Banana"),
}

@pytest.fixture
def db_connection(monkeypatch):
    connection = sqlite3.connect(':memory:', detect_types=sqlite3.PARSE_DECLTYPES)
    print("Connected to Test DB")
    # cursor = connection.cursor()
    yield connection
    connection.commit()
    print("Closing Test DB...")
    connection.close()
    
def test_update_parameters(db_connection, monkeypatch, caplog):
    GET_KVP = """--sql
        SELECT param_key, param_value, generation_time FROM parameters
        WHERE param_key=?"""

    def _mock_db_con(*args, **kwargs):
        return db_connection
    monkeypatch.setattr("sqlite3.connect", _mock_db_con)
    xlpnr.write_database_parameters(PARAMETER_DICT, METADATA_DICT, test_flag=True)
    cursor = db_connection.cursor()
    for key, value in PARAMETER_DICT.items():
        cursor.execute(GET_KVP, [key])
        test_key, test_val, timestamp = cursor.fetchone()
        assert test_key == key
        assert test_val == value
        assert timestamp == datetime.datetime(2022, 5, 8, 9, 27, 57)
    
    xlpnr.write_database_parameters(BAD_DICT, METADATA_DICT, test_flag=True)
    for bad_type in ['list', 'dict', 'tests.test_db.MockWorkbook']:
        assert f"Dict with type <class '{bad_type}'>" in caplog.text
    # monkeypatch.setitem(METADATA_DICT, 'retrieval_date', '2021-04-23 04:23:23')
    