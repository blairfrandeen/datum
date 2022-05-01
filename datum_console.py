import os
import sys

from collections import namedtuple

from xl_populate_named_ranges import update_named_ranges, user_select_json_file, user_select_open_workbook


class ConsoleSession:
    def __init__(self):
        self.json_file = None
        self.excel_workbook = None

    def load_measurement(self):
        self.json_file = user_select_json_file()

    def load_workbook(self):
        self.excel_workbook = user_select_open_workbook()

    def update_named_ranges(self, backup=False):
        if not self.json_file:
            self.load_measurement()
        if not self.excel_workbook:
            self.load_workbook()
        update_named_ranges(self.json_file, self.excel_workbook, backup)

    def status(self):
        print(f"Loaded Measurement:\t{self.json_file}")
        print(f"Loaded Workbook:\t{self.excel_workbook}")


def console(console_session, command_list, config_file=None):
    """
    Run a console within your python program.
    Some configuration options in JSON file.
    """

    def _help_msg():
        """Display this message and list all available commands."""
        print("Available commands and descriptions:")
        # TODO: Cleaner formatting for this function.
        for cmd in command_list:
            # print docstring for each command if available
            # otherwise print the function name
            if cmd.function.__doc__:
                docstr = cmd.function.__doc__
            else:
                docstr = cmd.function.__name__
            print(f"\t{cmd.id}\t\t{docstr}")

    # Default commands appear at the end
    command_list.append((["h", "help"], _help_msg))
    command_list.append((["q", "quit"], exit))

    # Create a docstring for exit function
    exit.__doc__ = "Quit"

    # Keep formatting neat for command list above
    # while still leveraging named tuples
    Command = namedtuple("Command", "id function")
    command_list = [Command._make(cmd) for cmd in command_list]

    user_input = None
    while user_input != "q":
        print("> ", end="")
        user_input = input()
        user_command = user_input.split(" ")[0]
        if user_command == "":
            continue
        user_args = user_input.split(" ")[1:]
        valid_command = False
        for cmd in command_list:
            if user_command in cmd.id:
                valid_command = True
                cmd.function(*user_args)

        if not valid_command:
            print("Unknown command. Type 'h' for help, 'q' to quit.")


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--test":
        os.chdir("tests")
    cs = ConsoleSession()
    command_list = [
        (["lm"], cs.load_measurement),
        (["lw"], cs.load_workbook),
        (["s"], cs.status),
        (["u"], cs.update_named_ranges)
    ]
    console(cs, command_list)


if __name__ == "__main__":
    main()
