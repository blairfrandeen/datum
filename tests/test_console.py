import os
import pytest
from datum import datum_console as dc

@pytest.fixture
def console_test_session():
    return dc.ConsoleSession()

@pytest.fixture
def console_command_list(console_test_session):
    cs = console_test_session
    command_list = [
        (["cd"], cs.chdir),
        (["lm"], cs.load_measurement),
        (["lw"], cs.load_workbook),
        (["pwd"], cs.pwd),
        (["s"], cs.status),
        (["u"], cs.update_named_ranges),
        (["z", "undo"], cs.undo_last_update),
    ]
    return command_list

def test_main(monkeypatch, capsys):
    monkeypatch.setattr('sys.argv', lambda: ['--test'])
    assert 'tests' in os.getpwd()
    # captured = capsys.readouterr()
    # assert "DATUM - Version" in captured.out


def test_chdir(capsys, console_test_session):
    # cd with no arguments
    console_test_session.chdir()
    captured = capsys.readouterr()
    assert "Change directory: cd <directory>" in captured.out

    # cd to bad directory
    console_test_session.chdir('bad_directory')
    captured = capsys.readouterr()
    assert "Directory not found" in captured.out

    # cd to bad directory
    os.makedirs("./temp_test")
    console_test_session.chdir('temp_test')
    console_test_session.chdir('..')
    os.rmdir("./temp_test")
    captured = capsys.readouterr()
    assert "temp_test" in captured.out

def test_pwd(capsys, console_test_session):
    console_test_session.pwd()
    captured = capsys.readouterr()
    assert os.getcwd() in captured.out

def test_status(capsys, console_test_session):
    console_test_session.status()
    captured = capsys.readouterr()
    assert "Loaded Measurement:" in captured.out

def test_load_measurement(monkeypatch, console_test_session):
    monkeypatch.setattr('builtins.input', lambda _: 'q')
    console_test_session.load_measurement()
    assert console_test_session.json_file is None

def test_load_workbook(monkeypatch, console_test_session):
    monkeypatch.setattr('builtins.input', lambda _: 'q')
    console_test_session.load_workbook()
    assert console_test_session.excel_workbook is None

def test_console(monkeypatch, capsys, console_command_list):
    bad_commands = ['5', 'gettrdun']
    for cmd in bad_commands:
        monkeypatch.setattr('builtins.input', lambda _: cmd)
        dc.console(console_command_list, test_flag=True)
        captured = capsys.readouterr()
        assert "Unknown command." in captured.out
        # with pytest.raises(SystemExit):
        #     monkeypatch.setattr('builtins.input', lambda _: 'q')
    
    # verify help command is called
    for help_cmd in ['h', 'help']:
        monkeypatch.setattr('builtins.input', lambda _: help_cmd)
        dc.console(console_command_list, test_flag=True)
        captured = capsys.readouterr()
        assert "Available commands" in captured.out

    # Verify no response for empty input
    monkeypatch.setattr('builtins.input', lambda _: '')
    assert dc.console(console_command_list, test_flag=True) is None
    # captured = capsys.readouterr()
    # assert "Available commands" in captured.out


def test_user_select_item(monkeypatch, capsys):
    empty_list = []
    assert(dc.user_select_item(empty_list, "nothing") is None)
    valid_list = ['muffins', 'cupcakes', 'cookies', 'more cookies']
        
    monkeypatch.setattr('builtins.input', lambda _: 'some string')
    dc.user_select_item(valid_list, 'treats', test_flag=True)
    captured = capsys.readouterr()
    assert "Invalid input." in captured.out

    bad_indices = [-1, 434]
    for cmd in bad_indices:
        monkeypatch.setattr('builtins.input', lambda _: cmd)
        dc.user_select_item(valid_list, 'treats', test_flag=True)
        captured = capsys.readouterr()
        assert "Index out of bounds." in captured.out

    good_choices = list(range(4))
    for choice in good_choices:
        monkeypatch.setattr('builtins.input', lambda _: choice)
        assert dc.user_select_item(valid_list, 'treats', test_flag=True) == choice

    monkeypatch.setattr('builtins.input', lambda _: 'q')
    assert dc.user_select_item(valid_list, 'treats', test_flag=True) is None

def test_user_select_json_file(monkeypatch):
    file_list = [ 'test1.json', 'something.txt', 'test2.json' ]
    monkeypatch.setattr('os.listdir', lambda: file_list)
    monkeypatch.setattr('os.getcwd', lambda: "")
    monkeypatch.setattr('builtins.input', lambda _: '0')
    assert dc.user_select_json_file().endswith(file_list[0])

    monkeypatch.setattr('builtins.input', lambda _: '1')
    assert dc.user_select_json_file().endswith(file_list[2])

    monkeypatch.setattr('builtins.input', lambda _: 'q')
    assert dc.user_select_json_file() is None

def test_main(monkeypatch):
    quit_commands = ['q', 'quit']
    for cmd in quit_commands:
        with pytest.raises(SystemExit):
            monkeypatch.setattr('builtins.input', lambda _: cmd)
            dc.main()

    