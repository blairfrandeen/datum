import json
from collections import namedtuple

def load_console_config(console_config_file):
    try:
        config_file = open(console_config_file, "r")
    except FileNotFoundError as err:
        print(err)
        print(f"Console configuration file {console_config_file} not found.")
        exit(-1)
    except IsADirectoryError as err:
        print(err)
        print(f"Console configuration file {console_config_file} is a directory.")
        exit(-1)
    except UnicodeDecodeError as err:
        print(err)
        print(f"{console_config_file} could not be read.")
        exit(-1)
    try:
        console_configuration = json.load(config_file)
    except json.decoder.JSONDecodeError as err:
        print(f"Error in {console_config_file}:")
        print(err)
        exit(-1)
    config_file.close()
    return console_configuration

def print_name(name=None):
    """Prints your name in ALL CAPS (test function)"""
    if name:
        print(f"Your name is {name.upper()}!!")
    else:
        print("Name argument required")


def console(config_file):
    """
    Run a console within your python program.
    Some configuration options in JSON file.
    """
    console_config = load_console_config(config_file)

    def _help_msg():
        """Display this message and list all available commands."""
        print(console_config["help_intro"])
        # TODO: Cleaner formatting for this function.
        for cmd in COMMANDS:
            # print docstring for each command if available
            # otherwise print the function name
            if cmd.function.__doc__:
                docstr = cmd.function.__doc__
            else:
                docstr = cmd.function.__name__
            print(f"\t{cmd.id}\t{docstr}")


    # Create a docstring for exit function
    exit.__doc__ = "Quit"

    COMMANDS = [
            (["n", "name"], print_name),
            (["h", "help"], _help_msg),
            (["q", "quit"], exit)
        ]

    # Keep formatting neat for command list above
    # while still leveraging named tuples
    Command = namedtuple("Command", "id function")
    COMMANDS = [ Command._make(cmd) for cmd in COMMANDS ]

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
        for cmd in COMMANDS:
            if user_command in cmd.id:
                valid_command = True
                cmd.function(*user_args)

        if not valid_command:
            print(console_config["invalid_command"])

def main():
    console("console_config.json")

if __name__ == "__main__":
    main()
