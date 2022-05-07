import os
import sys
from collections import namedtuple

import xlwings as xw

from xl_populate_named_ranges import logger, update_named_ranges


# TODO: Set up logging for tests
class ConsoleSession:
    def __init__(self):
        self.json_file = None
        self.excel_workbook = None
        self.undo_buffer = None

    def load_measurement(self):
        """Load measurement data from a JSON file"""
        self.json_file = user_select_json_file()

    def load_workbook(self):
        """Select an open Excel workbook to write to"""
        self.excel_workbook = user_select_open_workbook()

    def update_named_ranges(self, backup=False):
        """Update named ranges in the Excel file with matching
        data from the JSON measurement file"""
        if not self.json_file:
            self.load_measurement()
        if not self.excel_workbook:
            self.load_workbook()
        if self.json_file and self.excel_workbook:
            undo_buffer = update_named_ranges(
                self.json_file, self.excel_workbook, backup
            )
            # Do not clear undo buffer to None on abort
            if undo_buffer:
                self.undo_buffer = undo_buffer

    def undo_last_update(self):
        if self.undo_buffer:
            self.undo_buffer = update_named_ranges(
                self.undo_buffer, self.excel_workbook, backup=False
            )
        else:
            print("No undo history available.")

    def status(self):
        """Display loaded measurement & loaded workbook"""
        print(f"Loaded Measurement:\t{self.json_file}")
        print(f"Loaded Workbook:\t{self.excel_workbook}")


def user_select_item(item_list, item_type="choice"):
    """Given a list of files or workbooks, enumerate them and
    ask the user to select one item.

    Return the index of the selected item."""
    if len(item_list) < 1:
        logger.error("user_select_item passed an empty list.")
        return None

    # list the items
    print(f"Available {item_type}s:")
    for index, element in enumerate(item_list):
        print(f"[{index}] - {element}")

    # keep asking for input until a valid input or quit command received
    while True:
        print(f"Select {item_type} index (q to quit): ", end="")
        selection_index = input()
        if selection_index == "q":
            return None
        try:
            selection_index = int(selection_index)
        except ValueError:  # if selection is non-integer
            print("Invalid input.")
            continue
        if selection_index >= len(item_list) or selection_index < 0:
            print("Index out of bounds.")
            continue
        else:
            return selection_index


def user_select_open_workbook():
    if len(xw.apps) == 0:
        logger.error("Excel app not open.")
        return None
    if len(xw.books) == 0:
        logger.error("No Excel workbooks are open.")
        return None

    workbook_list = [book.name for book in xw.books]

    workbook_index = user_select_item(workbook_list, "Excel Workbook")
    if workbook_index is None:
        return None

    return xw.books[workbook_index]


def user_select_json_file():
    """Select a JSON file"""
    json_file_list = []
    for file in os.listdir():
        if file.endswith(".json"):
            json_file_list.append(file)
    json_index = user_select_item(json_file_list, "JSON file")
    if json_index is None:
        return None

    # TODO: Make this compatible with linux/macOS - use pathlib?
    json_file_path = f"{os.getcwd()}\\{json_file_list[json_index]}"
    return json_file_path


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
        (["u"], cs.update_named_ranges),
        (["z", "undo"], cs.undo_last_update),
    ]
    console(cs, command_list)


if __name__ == "__main__":
    main()
