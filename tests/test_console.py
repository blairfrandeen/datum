from unittest.mock import patch
from datum import datum_console as dc

class TestConsole:
    def test_user_select_item(self):
        empty_list = []
        assert(dc.user_select_item(empty_list, "nothing") is None)

        test_list = 'my name is prince'.split(' ')
        bad_inputs = [0, -1, 42, 'Q', 'z', 9.867]
        # for bad_in in bad_inputs:
        #     with patch("builtins.input", return_value=bad_in):
        #        dc.user_select_item(test_list, item_type="words")
        # with patch("builtins.input", return_value='q'):
        #     assert(dc.user_select_item(test_list) is None)
