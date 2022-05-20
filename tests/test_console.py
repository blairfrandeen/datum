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
    monkeypatch.setattr('__name__', lambda: "__main__")
    dc()
    captured = capsys.readouterr()
    assert "DATUM - Version" in captured.out

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

    