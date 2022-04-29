import json
from collections import namedtuple


def print_name(name=None):
    """Prints your name in ALL CAPS (test function)"""
    if name:
        print(f"Your name is {name.upper()}!!")
    else:
        print("Name argument required")


def console(command_list, config_file=None):
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
            print(f"\t{cmd.id}\t{docstr}")

    # Default commands appear at the end
    command_list.append((["h", "help"], _help_msg))
    command_list.append((["q", "quit"], exit))

    # Create a docstring for exit function
    exit.__doc__ = "Quit"


    # Keep formatting neat for command list above
    # while still leveraging named tuples
    Command = namedtuple("Command", "id function")
    command_list = [ Command._make(cmd) for cmd in command_list ]

    user_input = None
    while user_input != "q":
        print('> ', end='')
        user_input = input()
        user_command = user_input.split(' ')[0]
        if user_command == "":
            continue
        user_args = user_input.split(' ')[1:]
        num_args = len(user_args)
        valid_command = False
        for cmd in command_list:
            if user_command in cmd.id:
                valid_command = True
                cmd.function(*user_args)

        if not valid_command:
            print("Unknown command. Type 'h' for help, 'q' to quit.")

def main():
    sample_command_list = [
            (["n", "name"], print_name),
        ]
    console(sample_command_list)

if __name__ == "__main__":
    main()
