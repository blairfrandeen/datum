import json

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

def print_name(name):
    print(f"Your name is {name.upper()}!!")

def help_msg(console_config):
    """Display a help message and list all available commands."""
    print(console_config["help_intro"])
    for cmd in console_config["commands"]:
        print("\t", end="")
        for index, cmd_id in enumerate(cmd["id"]):
            print(f"{cmd_id}", end="")
            if index < len(cmd["id"]) - 1:
                print(", ", end="")
        print(f"\t\t{cmd['description']}")


def console(config_file):
    """
    Run a console within your python program.
    Configure commands and help message in JSON file.
    """
    console_config = load_console_config(config_file)
    while True:
        print(console_config["prompt"], end='')
        user_input = input()
        user_command = user_input.split(' ')[0]
        user_args = user_input.split(' ')[1:]
        num_args = len(user_args)
        if num_args > 0:
            print("Warning: Argument handling not implemented.")
        valid_command = False
        for cmd in console_config["commands"]:
            if user_command in cmd["id"] and user_command != "":
                valid_command = True
                eval(cmd["function"])
        if not valid_command:
            print(console_config["invalid_command"])


def main():
    console("console_config.json")

if __name__ == "__main__":
    main()
