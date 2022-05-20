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

def test_console(monkeypatch, capsys, console_command_list):
    bad_commands = ['5', 'gettrdun']
    for cmd in bad_commands:
        monkeypatch.setattr('builtins.input', lambda _: cmd)
        dc.console(console_command_list, test_flag=True)
        captured = capsys.readouterr()
        assert "Unknown command." in captured.out
        # with pytest.raises(SystemExit):
        #     monkeypatch.setattr('builtins.input', lambda _: 'q')

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

def test_main(monkeypatch):
    quit_commands = ['q', 'quit']
    for cmd in quit_commands:
        with pytest.raises(SystemExit):
            monkeypatch.setattr('builtins.input', lambda _: cmd)
            dc.main()

    