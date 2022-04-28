import json
from collections import namedtuple

def load_console_config(console_config_file):
    if not console_config_file:
        return None
    try:
        config_file = open(console_config_file, "r")
    except FileNotFoundError as err:
        print(err)
        return None
    except IsADirectoryError as err:
        print(err)
        return None
    except UnicodeDecodeError as err:
        print(err)
        return None
    try:
        console_configuration = json.load(config_file)
    except json.decoder.JSONDecodeError as err:
        print(err)
        return None
    config_file.close()
    return console_configuration

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

    console_config = load_console_config(config_file)
    if not console_config:
        console_config = {
            "quit_command": "q",
            "help_intro": "Available commands and descriptions:",
            "prompt": "> ",
            "invalid_command": "Invalid command. Type 'h' for help, 'q' to quit."
        }

    def _help_msg():
        """Display this message and list all available commands."""
        print(console_config["help_intro"])
        # TODO: Cleaner formatting for this function.
        for cmd in command_list:
            # print docstring for each command if available
            # otherwise print the function name
            if cmd.function.__doc__:
                docstr = cmd.function.__doc__
            else:
                docstr = cmd.function.__name__
            print(f"\t{cmd.id}\t{docstr}")

    command_list.append((["h", "help"], _help_msg))

    # Create a docstring for exit function
    exit.__doc__ = "Quit"


    # Keep formatting neat for command list above
    # while still leveraging named tuples
    Command = namedtuple("Command", "id function")
    command_list = [ Command._make(cmd) for cmd in command_list ]

    user_input = None
    while user_input != console_config["quit_command"]:
        print(console_config["prompt"], end='')
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
            print(console_config["invalid_command"])

def main():
    command_list = [
            (["n", "name"], print_name),
            (["q", "quit"], exit)
        ]
    console(command_list)

if __name__ == "__main__":
    main()
