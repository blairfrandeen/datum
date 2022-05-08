"""
Purpose of this package is to populate named ranges
in Excel with NX measurement data from a JSON file.

Prior to use, the Excel document needs to have ranges named
and ready to populate, which takes some time to setup
the first time this is used.

xlwings requires that Excel be open in order to run this code.

This only works with named ranges that are a single cell.
Behavior with multiple-cell ranges has not been tested.

"""
import json
import logging
import logging.config
import os

import xlwings as xw

# logging set-up
logging.config.fileConfig("logging.conf")
logger = logging.getLogger("root")


def xw_get_workbook(target_workbook):
    """Searches the workbooks that Excel has open
    Returns a workbook object if found, throws an error otherwise"""
    for book in xw.books:
        if book.name == target_workbook:
            return book
    logger.error(f"Target workbook {target_workbook} is not open.")
    return None


def xw_get_named_range(workbook, range_name):
    """Find a named range in an open workbook
    Returns the range object if found, returns None if not found"""
    if range_name in workbook.names:
        # print(f'Found {range_name}')
        if "!#REF!" in workbook.names[range_name].refers_to:
            logger.error(
                f"Name {range_name} has a #REF! error. Please fix prior to continuing."
            )
            logger.error("Use the name manager to remove or fix any names with errors.")
            return None
        rng = workbook.names[range_name].refers_to_range
        return rng
    else:
        logger.debug(f"Name {range_name} not found in {workbook}")
        return None


def read_named_range(workbook, range_name):
    """Read a value from a named range in the work book
    Note: This function will return 0 if an empty cell is read"""
    # TODO: Make this work with ranges of more than one cell
    rng = xw_get_named_range(workbook, range_name)
    if rng is None:
        return None

    return rng.value


def write_named_range(workbook, range_name, new_value):
    rng = xw_get_named_range(workbook, range_name)
    if rng.size > 1:
        logger.warning(f"range {rng.name} has size {rng.size}")
        logger.warning("Sizes larger than 1 not supported.")
    if rng is None:
        return None

    rng.value = new_value


def backup_workbook(workbook):
    """Create a backup copy of the workbook. Returns
    a new workbook object for the current workbook."""
    # TODO: Implement optional backup_dir argument to specify backup directory
    # TODO: Make more robust naming convention
    # TODO: Troubleshoot backup w/ M365 files
    if not workbook.name.endswith(".xlsx"):
        logger.warning(f'Warning: workbook "{workbook.name}" not a .xlsx file')

    wb_name = workbook.name.split(".")[-2]
    wb_full_path = workbook.fullname
    backup_path = f"{os.getcwd()}\\{wb_name}_BACKUP.xlsx"
    try:
        new_workbook = xw.books.open(fullname=wb_full_path)
    except FileNotFoundError:
        logger.exception(
            f"Cannot open workbook at {wb_full_path}.\
            Workbook will NOT be backed up prior to write!"
        )
    else:
        workbook.save(path=backup_path)
        new_workbook = xw.books.open(fullname=wb_full_path)
        workbook.close()
    logger.info(f"Successfully backed up to {backup_path}")
    return new_workbook


def preview_named_range_update(range_update_buffer, workbook):
    """Print out list of values that will be overwritten."""

    print(
        "{0:<32} {1:>12} {2:>12} {3:>15}".format(
            "NAME", "OLD VALUE", "NEW VALUE", "PERCENT CHANGE"
        )
    )
    print(
        "{0:<32} {1:>12} {2:>12} {3:>15}".format(
            "------------", "------------", "------------", "------------"
        )
    )

    for range_name in range_update_buffer.keys():
        json_value = range_update_buffer[range_name]
        # TODO: Better handling of non-float values.
        try:
            excel_value = float(read_named_range(workbook, range_name))
        except TypeError:
            excel_value = 0.0

        try:
            percent_change = (json_value - excel_value) / excel_value * 100
        except ZeroDivisionError:
            percent_change = 100.0
        # TODO: Better handling of empty values or zero values in Excel
        print(
            "{0:<32} {1:>12.5} {2:>12.5} {3:>14.3}%".format(
                range_name, excel_value, json_value, percent_change
            )
        )


def write_named_ranges(workbook, range_update_buffer, source_str, backup=False):

    preview_named_range_update(range_update_buffer, workbook)

    print("The values listed above will be overwritten.")
    # TODO: Add argument to function to skip confirmation
    print("Enter 'y' to continue: ", end="")
    overwrite_confirm = input()
    if overwrite_confirm == "y":
        if backup:
            workbook = backup_workbook(workbook)
        logger.debug(
            f"Updating named ranges.\n\
            Source: {source_str}\n\
            Target: {workbook.fullname}"
        )
        for range in range_update_buffer.keys():
            write_named_range(workbook, range, range_update_buffer[range])
    else:
        print("Aborted.")


def get_json_measurement_names(json_file):
    try:
        with open(json_file, "r") as json_file_handle:
            json_data = json.load(json_file_handle)
    except FileNotFoundError:
        logger.error(f"Unable to open {json_file}")
        return None

    json_named_measurements = dict()
    if "measurements" not in json_data.keys():
        logger.error(f"JSON file {json_file} has no measurement keys.")
        return None
    if len(json_data["measurements"]) == 0:
        logger.error(f"JSON file {json_file} has no measurement objects.")
        return None

    for measurement in json_data["measurements"]:
        measurement_name = measurement["name"]
        # replace spaces with underscores - no spaces allowed in excel range names
        measurement_name = measurement_name.replace(' ','_')
        for expr in measurement["expressions"]:
            expression_name = expr["name"]

            if expr["type"] == "Point" or expr["type"] == "Vector":
                for coordinate in ['x', 'y', 'z']:
                    range_name = f"{measurement_name}.{expression_name}.{coordinate}"
                    json_named_measurements[range_name] = expr["value"][coordinate]
            elif expr["type"] == "List":
                pass
            else:
                range_name = f"{measurement_name}.{expression_name}"
                json_named_measurements[range_name] = expr["value"]

    return json_named_measurements


def get_workbook_range_names(workbook):
    # make a dict of named ranges, measurement names, and measurement types
    workbook_named_ranges = dict()
    if len(workbook.names) == 0:
        logger.error(f"workbook{workbook.name} has no named ranges.")
        return None
    for named_range in workbook.names:
        range_value = read_named_range(workbook, named_range.name)
        workbook_named_ranges[named_range.name] = range_value

    return workbook_named_ranges


def update_named_ranges(source, target, backup=False):
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
    target_data = get_workbook_range_names(target)
    if not target_data:
        print("No named ranges in Excel file.")
        return None

    # Check if source is json file
    if isinstance(source, str) and source.lower().endswith(".json"):
        source_data = get_json_measurement_names(source)
        if not source_data:
            print("No measurement data found in JSON file.")
            return None
        source_str = source

    elif isinstance(source, dict):
        source_data = source
        source_str = "UNDO BUFFER"

    # find range names that occur both in Excel and JSON
    ranges_to_update = list(source_data.keys() & target_data.keys())

    range_update_buffer = dict()
    range_undo_buffer = dict()

    for range in ranges_to_update:
        range_update_buffer[range] = source_data[range]
        range_undo_buffer[range] = target_data[range]

    write_named_ranges(target, range_update_buffer, source_str, backup)

    return range_undo_buffer
