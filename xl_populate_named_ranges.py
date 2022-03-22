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
import xlwings as xw
import os
import json
import re
import logging

# logging set-up
# TODO: Move to external config file
logger = logging.getLogger('xl_pnr')
fh = logging.FileHandler('xl_pnr.log')
fh.setLevel(logging.DEBUG)
fh_format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s',\
    datefmt='%Y-%m-%d %H:%M:%S')
fh.setFormatter(fh_format)
logger.addHandler(fh)

sh = logging.StreamHandler()
sh.setLevel(logging.INFO)
logger.addHandler(sh)

# regex to split ranges into sheet and cells
# illegal characters in excel sheet names: ?*[]\/:
RE_RANGE_SPLIT = re.compile(r"(?<==)('?)([^\[\]\?*\\\/]+)\1!(\$[A-Z]+\$\d+)")

def xw_get_workbook(target_workbook):
    """Searches the workbooks that Excel has open
    Returns a workbook object if found, throws an error otherwise"""
    for book in xw.apps[0].books:
        if book.name == target_workbook:
            return book
    print(f'Error: Target workbook {target_workbook} is not open.')
    return None

def xw_get_named_range(workbook, range_name):
    """Find a named range in an open workbook
    Returns the range object if found, returns None if not found"""
    if range_name in workbook.names:
        # print(f'Found {range_name}')
        range_split = re.search(RE_RANGE_SPLIT,\
            workbook.names[range_name].refers_to).groups(0)
        worksheet = range_split[1]
        cell = range_split[2]
        # print(f'Refers to {worksheet} {cell}')
        rng = workbook.sheets[worksheet].range(cell)
        return rng
    else:
        logger.debug(f'Name {range_name} not found in {workbook}')
        return None

def read_named_range(workbook, range_name):
    """Read a value from a named range in the work book"""
    rng = xw_get_named_range(workbook, range_name)
    if rng is None: return None

    return rng.value

def write_named_range(workbook, range_name, new_value):
    """Write a value to a named range in the workbook"""
    rng = xw_get_named_range(workbook, range_name)
    if rng is None: return None

    rng.value = new_value

def backup_workbook(workbook):
    """Create a backup copy of the workbook. Returns
    a new workbook object for the current workbook."""
    # TODO: Implement optional backup_dir argument to specify backup directory
    # TODO: Make more robust naming convention
    if not workbook.name.endswith('.xlsx'):
        logger.warning(f'Warning: workbook "{workbook.name}" not a .xlsx file')

    wb_name = workbook.name.split('.')[-2]
    wb_full_path = workbook.fullname
    backup_path = f'{os.getcwd()}\\{wb_name}_BACKUP.xlsx'
    try:
        new_workbook = xw.books.open(fullname=wb_full_path)
    except FileNotFoundError:
        logger.exception(f'Cannot open workbook at {wb_full_path}. Workbook will NOT be backed up prior to write!')
    else:
        workbook.save(path=backup_path)
        new_workbook = xw.books.open(fullname=wb_full_path)
        workbook.close()
    logger.info(f'Successfully backed up to {backup_path}')
    return new_workbook


def update_named_ranges(json_file, workbook, backup=True):
    """
    Open a JSON file and an excel file. Update the named
    ranges in the excel file with the corresponding JSON values.
    
    Named ranges may correspond to a particular expression type.
    For example, the measurement SURFACE_SPHERICAL has an expression
    of type "area", along with other expressions. To populate this in 
    Excel, we need to name the range "SURFACE_SPHERICAL.area"
    """

    # make a dict of named ranges, measurement names, and measurement types
    workbook_named_ranges = dict()
    for named_range in workbook.names:
        measurement_name = named_range.name.split('.')[0]
        try:
            # example: MEASUREMENT.mass will take "mass" as the measurement type
            measurement_type = named_range.name.split('.')[1]
        except IndexError:
            # if out of bounds, there is no "." in range name; use default measurement
            measurement_type = None
        workbook_named_ranges[measurement_name] = measurement_type
    # print(workbook_named_ranges)
    write_list = dict()
    with open(json_file, "r") as json_file_handle:
        json_data = json.load(json_file_handle)
        #TODO: Add argument to function to skip confirmation
        print('{0:<32} {1:>12} {2:>12} {3:>15}'.format('NAME', 'OLD VALUE', 'NEW VALUE', 'PERCENT CHANGE'))
        print('{0:<32} {1:>12} {2:>12} {3:>15}'.format('------------', '------------', '------------', '------------'))
        for measurement in json_data['measurements']:
            if measurement['name'] in workbook_named_ranges.keys():
                range_name = measurement['name']
                range_type = workbook_named_ranges[measurement['name']]
                if range_type:
                    range_name = f'{range_name}.{range_type}'
                    json_value = None
                    for expr in measurement['expressions']:
                        if expr['type'] == range_type:
                            json_value = expr['value']
                            break
                else:
                    json_value = measurement['expressions'][0]['value']
                if workbook_named_ranges[measurement['name']] is not None:
                    measurement_type = workbook_named_ranges[measurement['name']]
                # print the value currently in Excel
                excel_value = read_named_range(workbook, range_name)
                # print(f'Excel value of {range_name}: {excel_value}')
                # print the value currently in JSON
                # print(f'JSON Value: {json_value}')
                percent_change = (json_value - excel_value) / excel_value * 100
                print('{0:<32} {1:>12.5} {2:>12.5} {3:>14.3}%'.format(range_name, excel_value, json_value, percent_change))
                write_list[range_name] = json_value

    print("The values listed above will be overwritten. Enter 'y' to continue: ", end='')
    overwrite_confirm = input()
    if overwrite_confirm == 'y':
        if backup:
            workbook = backup_workbook(workbook)
        logger.debug(f'Updating named ranges.\n\
            Source: {json_file}\n\
            Target: {workbook.fullname}')
        for range in write_list.keys():
            write_named_range(workbook, range, write_list[range])
    else:
        print("Aborted.")

def user_select_item(item_list, item_type='choice'):
    """Given a list of files or workbooks, enumerate them and
    ask the user to select one item.
    
    Return the index of the selected item."""
    # list the items
    print(f'Available {item_type}s:')
    for index, element in enumerate(item_list):
        print(f'[{index}] - {element}')
    while True: # keep asking for input until a valid input or quit command received
        print(f'Select {item_type} index (q to quit): ', end='')
        selection_index = input()
        if selection_index == 'q':
            return None
        try:
            selection_index = int(selection_index)
        except ValueError: # if selection is non-integer
            print('Invalid input.')
            continue
        if selection_index >= len(item_list) or selection_index < 0:
            print('Index out of bounds.')
            continue
        else:
            return selection_index

def user_select_open_workbook():
    workbook_list = [ book.name for book in xw.apps[0].books ]
    # TODO: Incorporate this test into user_select_item, throw an
    #   error if empty list received
    if len(workbook_list) > 0:
        workbook_index = user_select_item(workbook_list, 'Excel Workbook')
        if workbook_index is not None:
            return xw.apps[0].books[workbook_index]
        else:
            return None
    else:
        print("No open workbooks detected.")
        return None

def user_select_json_file():
    json_file_list = []
    for file in os.listdir():
        if file.endswith('.json'):
            json_file_list.append(file)
    # TODO: Incorporate this test into user_select_item, throw an
    #   error if empty list received
    if len(json_file_list) > 0:
        json_index = user_select_item(json_file_list, 'JSON file')
        if json_index is not None:
            #TODO: Make this compatible with linux/macOS - use pathlib?
            json_file_path = f'{os.getcwd()}\\{json_file_list[json_index]}'
            return json_file_path
        else:
            return None
    else:
        print("No JSON files detected in working directory.")
        return None

def main():
    json_file = user_select_json_file()
    excel_workbook = user_select_open_workbook()
    update_named_ranges(json_file, excel_workbook)

if __name__ == '__main__':
    main()