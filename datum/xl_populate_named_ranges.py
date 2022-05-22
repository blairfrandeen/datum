"""
Purpose of this package is to populate named ranges
in Excel with NX measurement data from a JSON file.

Prior to use, the Excel document needs to have ranges named
and ready to populate, which takes some time to setup
the first time this is used.

xlwings requires that Excel be open in order to run this code.
"""
import datetime
import json
import logging
import logging.config
from pathlib import Path
from typing import List, Optional, Union

import xlwings as xw

# logging set-up
logging.config.fileConfig("logging.conf")
logger: logging.Logger = logging.getLogger(__name__)

########################
## XLWINGS INTERFACES ##
########################


def backup_workbook(workbook: xw.main.Book, backup_dir: str = ".") -> Path:
    """Create a backup copy of the workbook.
    Returns the path of the backup copy."""
    # TODO: Make more robust naming convention
    # TODO: Verify M365 files are backing up correctly
    # TODO: Implement unit tests
    print(f"Backuping up {workbook.name}...")
    wb_name: str = workbook.name.split(".xlsx")[0]

    backup_path = Path(f"{backup_dir}\\{wb_name}_BACKUP.xlsx")

    # Open a new blank workbook
    backup_wb: xw.main.Book = xw.Book()

    # Copy sheets individually
    for sheet in workbook.sheets:
        sheet.copy(after=backup_wb.sheets[0])

    # Delete the first blank sheet
    backup_wb.sheets[0].delete()

    # Save & close
    backup_wb.save(path=backup_path)
    backup_wb.close()

    return backup_path


def dump(workbook: xw.main.Book, json_file: dict) -> None:
    """Take data frome a dictionary of key-value pairs
    that originated from a JSON file, and place it in Excel
    in a new worksheet for easy access."""
    sheet_name: str = 'DATUM'
    
    # get the data from the json_file
    data = get_json_key_value_pairs(json_file)
    if data is None:
        logger.error("No key-value pairs in JSON file to dump.")
    else:
        # create a new worksheet
        workbook.sheets.add(sheet_name)
        
        # create header row
        workbook.sheets[sheet_name].range('A1:B1').value = ['PARAMETER', 'VALUE']

        # add each key-value pair from the data
        starting_row: int = 2
        for index, key in enumerate(data):
            target_range: str = f"A{starting_row + index}:B{starting_row + index}"
            workbook.sheets[sheet_name].range(target_range).value = list(
                flatten_list([key, data[key]])
            )

def get_workbook_key_value_pairs(workbook: xw.main.Book) -> Optional[dict]:
    """Find all named ranges in a workbook and return
    a dictionary of name-value pairs."""
    if len(workbook.names) == 0:
        logger.error(f"workbook{workbook.name} has no named ranges.")
        return None
    # make a dict of named ranges, measurement names, and measurement types
    workbook_named_ranges = dict()
    for named_range in workbook.names:
        # Sometimes Excel puts in hidden names that start
        # with _xlfn. -- skip these
        if named_range.name.startswith("_xlfn."):
            logger.debug(f"Skipping range {named_range.name}")
            continue
        if "!#REF" in named_range.refers_to:
            logger.error(f"Name {named_range.name} has a #REF! error.")
            continue
        workbook_named_ranges[named_range.name] = named_range.refers_to_range.value

    return workbook_named_ranges


def load_metadata_from_json(json_file: str) -> Optional[dict]:
    try:
        with open(json_file, "r") as json_handle:
            json_metadata: dict = json.load(json_handle)
            if check_dict_keys(json_metadata, ['METADATA']):
                return json_metadata['METADATA']
            else:
                logger.debug(f'No "METADATA" in {json_file}')

    except FileNotFoundError:
        logger.debug(f'{json_file} not found.')
        return None

def write_named_range(
    workbook: xw.main.Book,
    range_name: str,
    new_value: Optional[Union[list, float, int, str, datetime.datetime]],
) -> Optional[Union[list, int, str, float, datetime.datetime]]:
    """Write a value or list to a named range.

    Keyword arguments:
    workbook -- xlwings Book object
    range_name -- string with range name
    new_value -- new value or list of values to write
    """
    if range_name not in workbook.names:
        raise KeyError(f"Name {range_name} not in {workbook.name}")
    if "!#REF" in workbook.names[range_name].refers_to:
        raise TypeError(f"Name {range_name} has a #REF! error.")

    target_range: xw.main.Range = workbook.names[range_name].refers_to_range

    if isinstance(new_value, list):
        # Flatten any arbitrary list
        new_value = list(flatten_list(new_value))
        new_value_len: int = len(new_value)
        # new_value_len: int = list_len(new_value)
        # TODO: Unit test for this case
        if new_value_len > target_range.size:
            # Truncate input if range size is too small
            new_value = new_value[: target_range.size]
            logger.warning(f"range {target_range.name} has size {target_range.size}.")
            logger.warning(f"Vector of length {new_value_len} will be truncated.")
        elif new_value_len < target_range.size:
            # If range is too big, warn that not all cells will be populated
            logger.warning(
                f"Range {range_name} of size\
                {target_range.size} is larger than required."
            )
        for index in range(min(target_range.size, new_value_len)):
            if (
                not isinstance(
                    new_value[index], (int, str, float, list, datetime.datetime)
                )
                and new_value[index] is not None
            ):
                raise TypeError(f"Write {type(new_value[index])} not allowed.")
            target_range[index].value = new_value[index]
        return new_value
    elif (
        isinstance(new_value, (int, str, float, datetime.datetime)) or new_value is None
    ):
        target_range.value = new_value
        return new_value
    else:
        raise TypeError(f"Cannot write value of type {type(new_value)}")


#######################
###### UTILITIES ######
#######################
def check_dict_keys(target_dict: dict, keys_to_check: list) -> bool:
    """Check that a dictionary has the required keys.

    Ensure keys are not empty lists. Warn if keys not found."""
    if not isinstance(target_dict, dict):
        logger.warning("Not a dictionary.")
        return False
    for key in keys_to_check:
        if key not in target_dict.keys():
            logger.warning(f'Key "{key}" not found.')
            return False
        if isinstance(target_dict[key], (list, dict)) and len(target_dict[key]) == 0:
            logger.warning(f'Key "{key}" has no entires.')
            return False

    return True


def flatten_list(target_list: list):
    """Flatten any nested list."""
    if not isinstance(target_list, list):
        raise TypeError("Cannot flatten a non-list")
    for element in target_list:
        if isinstance(element, list):
            for sub_element in flatten_list(element):
                yield sub_element
        else:
            yield element


def report_difference(
    old_value: Optional[Union[int, float, str, datetime.datetime]],
    new_value: Optional[Union[int, float, str, datetime.datetime]],
) -> Optional[Union[float, datetime.timedelta]]:
    """Report the difference between any two instances of
    int, float, str, date or None. Return None if no numerical
    comparison can be made. Return percent difference for ints and floats.
    Return timedelta for comparison of dates."""
    if old_value == 0:
        return None  # otherwise divide by zero error
    if isinstance(old_value, datetime.datetime) and isinstance(
        new_value, datetime.datetime
    ):
        return new_value - old_value
    if isinstance(old_value, (int, float)) and isinstance(new_value, (int, float)):
        return (new_value - old_value) / old_value

    return None


########################
#### CORE FUNCTIONS ####
########################


def get_json_key_value_pairs(json_file: str) -> Optional[dict]:
    """Load JSON measurement dict from a JSON file."""
    try:
        with open(json_file, "r") as json_file_handle:
            json_data: dict = json.load(json_file_handle)
    except FileNotFoundError:
        logger.error(f"Unable to open {json_file}")
        return None
    except json.decoder.JSONDecodeError:
        logger.error(f"JSON file {json_file} is corrupt.")
        return None

    if not check_dict_keys(json_data, ["measurements"]):
        logger.warning(f'No "measurement" field in {json_file}')
        return None

    json_named_measurements: dict = dict()
    for measurement in json_data["measurements"]:
        if not check_dict_keys(measurement, ["name", "expressions"]):
            logger.warning(f"{measurement} is missing name and/or expressions")
            continue

        # replace spaces with underscores - no spaces allowed in excel range names
        measurement_name: str = measurement["name"].replace(" ", "_")
        for expr in measurement["expressions"]:
            if not check_dict_keys(expr, ["name", "type", "value"]):
                logger.warning(f"missing name/type/value fields in {expr}")
                continue
            range_name: str = f"{measurement_name}.{expr['name']}"
            if expr["type"] == "Point" or expr["type"] == "Vector":
                vector: List[float] = []
                for coordinate in ["x", "y", "z"]:
                    coordinate_name: str = f"{range_name}.{coordinate}"
                    json_named_measurements[coordinate_name] = expr["value"][coordinate]
                    vector.append(expr["value"][coordinate])
                # TODO: Keep dicts as dicts, don't conver them to vectors.
                # Requires update of write_named_range to handle dicts.
                # json_named_measurements[range_name] = expr["value"]
                json_named_measurements[range_name] = vector
            elif expr["type"] == "List":
                json_named_measurements[range_name] = expr["value"]
                for index in range(3):
                    range_index = f"{range_name}.{index}"
                    json_named_measurements[range_index] = expr["value"][index]
            else:
                json_named_measurements[range_name] = expr["value"]

    return json_named_measurements


def preview_named_range_update(
    existing_values: dict, new_values: dict, min_diff: float = 0.001
) -> None:
    """Print out list of values that will be overwritten."""
    column_widths = [36, 17, 17, 17]
    column_headings = ["PARAMETER", "OLD VALUE", "NEW VALUE", "PERCENT CHANGE"]
    underlines = ["-" * 20, "-" * 12, "-" * 12, "-" * 15]
    print()  # newline
    print_columns(column_widths, column_headings)
    print_columns(column_widths, underlines)

    for range_name in new_values.keys():
        json_value = new_values[range_name]
        excel_value = existing_values[range_name]
        if isinstance(json_value, (list, dict)):
            for index, json_item in enumerate(json_value):
                if isinstance(json_value, dict):
                    item_name = f"{range_name}[{json_item}]"
                    json_item = json_value[json_item]
                else:
                    item_name = f"{range_name}[{index}]"
                if isinstance(excel_value, list):
                    excel_item = excel_value[index]
                elif index == 0:
                    excel_item = excel_value
                else:
                    excel_item = None

                difference = report_difference(excel_item, json_item)
                if isinstance(difference, float) and abs(difference) < min_diff:
                    continue  # @pytest-pass
                print_columns(
                    column_widths, [item_name, excel_item, json_item, difference]
                )
        else:
            difference = report_difference(excel_value, json_value)
            if isinstance(difference, float) and abs(difference) < min_diff:
                continue
            print_columns(
                column_widths, [range_name, excel_value, json_value, difference]
            )


def print_columns(
    widths: list, values: list, decimals: int = 3, na_string: str = "-"
) -> None:
    if len(widths) != len(values):
        raise IndexError("Mismatch of columns & values.")
    if any([not isinstance(item, int) for item in widths]):
        raise TypeError("Column widths must be integers")
    alignments = ["<", ">", ">", ">"]  # align left for first column
    for column, value in enumerate(values):
        if isinstance(value, float):
            # use percentage on last column
            if column == 3:
                fspec = "%"
            else:
                fspec = "g"
            print(
                "{val:{al}{wid}.{prec}{fspec}}".format(
                    val=value,
                    al=alignments[column],
                    wid=widths[column],
                    prec=decimals,
                    fspec=fspec,
                ),
                end="",
            )
        elif isinstance(value, str):
            # truncate extra long strings
            if len(value) > widths[column]:
                num_chars = int(widths[column] / 2) - 3
                value = value[:num_chars] + "..." + value[-num_chars:]
            print(
                "{val:{al}{wid}}".format(
                    val=value, al=alignments[column], wid=widths[column]
                ),
                end="",
            )
        elif isinstance(value, datetime.datetime):
            datestr = value.strftime("%Y-%m-%d")
            print(
                "{val:{al}{wid}}".format(val=datestr, al=">", wid=widths[column]),
                end="",
            )
        elif isinstance(value, datetime.timedelta):
            date_delta = f"{value.days} days"
            print(
                "{val:{al}{wid}}".format(val=date_delta, al=">", wid=widths[column]),
                end="",
            )
        elif value is None:
            print(
                "{val:{al}{wid}}".format(val=na_string, al="^", wid=widths[column]),
                end="",
            )
    print()  # newline


def update_named_ranges(
    source: Union[str, dict], target: xw.main.Book, backup: bool = False
) -> Optional[dict]:
    """
    Open a JSON file and an excel file. Update the named
    ranges in the excel file with the corresponding JSON values.

    Named ranges may correspond to a particular expression type.
    For example, the measurement SURFACE_SPHERICAL has an expression
    of type "area", along with other expressions. To populate this in
    Excel, we need to name the range "SURFACE_SPHERICAL.area"
    """
    # Assume target is open excel worksheet
    # TODO: Implement ability to take .xlsx file path as argument
    target_data: Optional[dict] = get_workbook_key_value_pairs(target)
    if not target_data:
        print("No named ranges in Excel file.")
        return None

    # Check if source is json file
    if isinstance(source, str) and source.lower().endswith(".json"):
        source_data: Optional[dict] = get_json_key_value_pairs(source)
        source_str: str = source
        if not source_data:
            print("No measurement data found in JSON file.")
            return None

    # TODO: Unit test
    elif isinstance(source, dict):
        source_data = source
        source_str = "UNDO BUFFER"

    # find range names that occur both in Excel and JSON
    ranges_to_update: list = list(source_data.keys() & target_data.keys())

    range_update_buffer: dict = dict()
    range_undo_buffer: dict = dict()

    for range in ranges_to_update:
        range_update_buffer[range] = source_data[range]
        range_undo_buffer[range] = target_data[range]

    write_named_ranges(
        range_undo_buffer, range_update_buffer, target, source_str, backup
    )

    return range_undo_buffer


def write_named_ranges(
    exiting_values: dict,
    new_values: dict,
    workbook: xw.main.Book,
    source_str: str,
    backup: bool = False,
) -> None:
    """Update named ranges in a workbook from a dictionary."""

    preview_named_range_update(exiting_values, new_values)

    print("The values listed above will be overwritten.")
    # TODO: Add argument to function to skip confirmation
    # print("Enter 'y' to continue: ", end="")
    overwrite_confirm: str = input("Enter 'y' to continue: ")
    if overwrite_confirm == "y":
        if backup:
            backup_path: Path = backup_workbook(workbook)
            logger.debug(f"Backed up to {backup_path}")
        logger.debug(
            f"Updating named ranges.\n\
            Source: {source_str}\n\
            Target: {workbook.fullname}"
        )
        for range in new_values.keys():
            write_named_range(workbook, range, new_values[range])
    else:
        print("Aborted.")
