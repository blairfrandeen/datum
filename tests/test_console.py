import os

import pytest

from datum import datum_console as dc


class MockWorkbook:
    def __init__(self, name):
        self.name = name


@pytest.fixture
def console_test_session():
    return dc.ConsoleSession()


@pytest.fixture
def console_command_list(console_test_session):
    cs = console_test_session
    command_list = [
        (["b"], cs.backup),
        (["cd"], cs.chdir),
        (["d", "dump"], cs.dump_json),
        (["lm"], cs.load_measurement),
        (["lw"], cs.load_workbook),
        (["pwd"], cs.pwd),
        (["s"], cs.status),
        (["u"], cs.update_named_ranges),
        (["z", "undo"], cs.undo_last_update),
    ]
    return command_list


class TestConsoleSession:
    def test_chdir(self, capsys, console_test_session):
        # cd with no arguments
        console_test_session.chdir()
        captured = capsys.readouterr()
        assert "Change directory: cd <directory>" in captured.out

        # cd to bad directory
        console_test_session.chdir("bad_directory")
        captured = capsys.readouterr()
        assert "Directory not found" in captured.out

        # cd with wrong argument type:
        for wrong_arg in [5, 3.2, (1, 2, "three"), [1, 2, 3]]:
            with pytest.raises(TypeError):
                console_test_session.chdir(wrong_arg)

        # cd to temporary directory directory
        os.makedirs("./temp_test")
        console_test_session.chdir("temp_test")
        assert "temp_test" in os.getcwd()
        console_test_session.chdir("..")
        os.rmdir("./temp_test")

    def test_backup(self, monkeypatch, console_test_session, capsys):
        def _mock_backup(*args):
            print(f"backup_test_success")

        monkeypatch.setattr(dc, "backup_workbook", _mock_backup)

        # test backup with no workbook loaded
        console_test_session.backup()
        captured = capsys.readouterr()
        assert "No Excel workbook is loaded." in captured.out

        # test backup with workbook loaded
        console_test_session.excel_workbook = MockWorkbook("test")
        console_test_session.backup()
        captured = capsys.readouterr()
        assert "backup_test_success" in captured.out

    def test_dump(self, monkeypatch, console_test_session, capsys):
        def _mock_dump(*args):
            print(f"dump_test_success")

        monkeypatch.setattr(dc, "dump", _mock_dump)

        # test dump with no JSON loaded
        console_test_session.dump_json()
        captured = capsys.readouterr()
        assert "No JSON data is loaded." in captured.out

        # test dump with no Workbook loaded
        console_test_session.json_file = "test.json"
        console_test_session.dump_json()
        captured = capsys.readouterr()
        assert "No Excel workbook is loaded." in captured.out

        # test dump with JSON and Excel loaded
        console_test_session.excel_workbook = MockWorkbook("test")
        console_test_session.dump_json()
        captured = capsys.readouterr()
        assert "dump_test_success" in captured.out

    def test_load_measurement(self, monkeypatch, console_test_session):
        def _mock_select_json():
            return "select_json"

        monkeypatch.setattr(dc, "user_select_json_file", _mock_select_json)
        console_test_session.load_measurement()
        assert console_test_session.json_file == "select_json"

    def test_load_workbook(self, monkeypatch, console_test_session):
        def _mock_select_wb():
            return "select_wb"

        monkeypatch.setattr(dc, "user_select_open_workbook", _mock_select_wb)
        console_test_session.load_workbook()
        assert console_test_session.excel_workbook == "select_wb"

    def test_pwd(self, capsys, console_test_session):
        console_test_session.pwd()
        captured = capsys.readouterr()
        assert os.getcwd() in captured.out

    def test_status(self, capsys, console_test_session):
        console_test_session.status()
        captured = capsys.readouterr()
        assert "Loaded Measurement:" in captured.out

    def test_undo(self, console_test_session, monkeypatch, capsys):
        def _mock_xlpnr_update(arg1, arg2, backup=False):
            return {"update_success": True}

        monkeypatch.setattr(dc, "update_named_ranges", _mock_xlpnr_update)
        cts = console_test_session
        cts.undo_last_update()
        captured = capsys.readouterr()
        assert "No undo history available." in captured.out
        cts.undo_buffer = {"update_success": False}
        cts.undo_last_update()
        assert cts.undo_buffer["update_success"] is True

    def test_update_named_ranges(self, console_test_session, monkeypatch):
        cts = console_test_session

        def _mock_select_json():
            cts.json_file = "json_file"

        monkeypatch.setattr(cts, "load_measurement", _mock_select_json)

        def _mock_select_workbook():
            cts.excel_workbook = "excel_workbook"

        monkeypatch.setattr(cts, "load_workbook", _mock_select_workbook)

        def _mock_xlpnr_update(arg1, arg2, arg3):
            return {"update_success": True}

        monkeypatch.setattr(dc, "update_named_ranges", _mock_xlpnr_update)

        cts.update_named_ranges()
        assert cts.json_file == "json_file"
        assert cts.excel_workbook == "excel_workbook"
        assert cts.undo_buffer["update_success"] is True


def test_console(monkeypatch, capsys, console_command_list):
    # verify bad commands are handled
    bad_commands = ["5", "gettrdun"]
    for cmd in bad_commands:
        monkeypatch.setattr("builtins.input", lambda _: cmd)
        dc.console(console_command_list, test_flag=True)
        captured = capsys.readouterr()
        assert "Unknown command." in captured.out

    # verify help command is called
    for help_cmd in ["h", "help"]:
        monkeypatch.setattr("builtins.input", lambda _: help_cmd)
        dc.console(console_command_list, test_flag=True)
        captured = capsys.readouterr()
        assert "Available commands" in captured.out

    # Verify no response for empty input
    monkeypatch.setattr("builtins.input", lambda _: "")
    assert dc.console(console_command_list, test_flag=True) is None


def test_user_select_item(monkeypatch, capsys):
    # test empty list returns None
    empty_list = []
    assert dc.user_select_item(empty_list, "nothing") is None

    # valid list of optoins
    valid_list = ["muffins", "cupcakes", "cookies", "more cookies"]

    # test that string input is invalid
    monkeypatch.setattr("builtins.input", lambda _: "some string")
    dc.user_select_item(valid_list, "treats", test_flag=True)
    captured = capsys.readouterr()
    assert "Invalid input." in captured.out

    # test idnex out of bounds
    bad_indices = [-1, 434]
    for cmd in bad_indices:
        monkeypatch.setattr("builtins.input", lambda _: cmd)
        dc.user_select_item(valid_list, "treats", test_flag=True)
        captured = capsys.readouterr()
        assert "Index out of bounds." in captured.out

    # test valid choices
    good_choices = list(range(4))
    for choice in good_choices:
        monkeypatch.setattr("builtins.input", lambda _: choice)
        assert dc.user_select_item(valid_list, "treats", test_flag=True) == choice

    # test quit command
    monkeypatch.setattr("builtins.input", lambda _: "q")
    assert dc.user_select_item(valid_list, "treats", test_flag=True) is None


def test_user_select_json_file(monkeypatch):
    # mock response from os.listdir()
    file_list = ["test1.json", "something.txt", "test2.json"]
    monkeypatch.setattr("os.listdir", lambda: file_list)

    # mock response from os.getcwd()
    monkeypatch.setattr("os.getcwd", lambda: "")

    # mock responses from dc.user_select_item
    item_selections = iter([0, 1, None])

    def _mock_select_item(arg1, arg2):
        return next(item_selections)

    monkeypatch.setattr(dc, "user_select_item", _mock_select_item)

    assert dc.user_select_json_file().endswith(file_list[0])
    assert dc.user_select_json_file().endswith(file_list[2])
    assert dc.user_select_json_file() is None


def test_user_select_workbook(monkeypatch):
    import xlwings as xw

    monkeypatch.setattr(xw, "apps", [])
    assert dc.user_select_open_workbook() is None
    monkeypatch.setattr(xw, "books", [])
    monkeypatch.setattr(xw, "apps", ["fake app"])
    assert dc.user_select_open_workbook() is None

    # mock workbook list
    wblist = [MockWorkbook(wbname) for wbname in ["wb1", "wb2", "wb3"]]
    monkeypatch.setattr(xw, "books", wblist)

    # mock responses from dc.user_select_item
    item_selections = iter([0, 1, None])

    def _mock_select_item(arg1, arg2):
        return next(item_selections)

    monkeypatch.setattr(dc, "user_select_item", _mock_select_item)

    assert dc.user_select_open_workbook() == wblist[0]
    assert dc.user_select_open_workbook() == wblist[1]
    assert dc.user_select_open_workbook() is None


def test_main(monkeypatch):
    quit_commands = ["q", "quit"]
    for cmd in quit_commands:
        with pytest.raises(SystemExit):
            monkeypatch.setattr("builtins.input", lambda _: cmd)
            dc.main()
