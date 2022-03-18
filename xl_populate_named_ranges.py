"""
Purpose of this package is to populate named ranges
in Excel with NX measurement data from a JSON file.

Prior to use, the Excel document needs to have ranges named
and ready to populate, which takes some time to setup
the first time this is used.

xlwings requires that Excel be open in order to run this code.

This only works with named ranges that are a single cell.
Behavior with multiple-cell ranges has not been tested.

TODO: Add proper logging setup (remove many print statements)
"""
import xlwings as xw
import os
import json
import re

# regex to split ranges into sheet and cells
# illegal characters in excel sheet names: ?*[]\/:
RE_RANGE_SPLIT = re.compile(r"(?<==)('?)([^\[\]\?*\\\/]+)\1!(\$[A-Z]+\$\d+)")

def xw_get_workbook(target_workbook):
    """Searches the workbooks that Excel has open
    Returns a workbook object if found, throws an error otherwise"""
    for book in xw.apps[0].books:
        if book.name == target_workbook:
            return book
    print(f'Error: Target workbook {target_workbook} not found in {os.getcwd()}')
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
        print(f'Name {range_name} not found in {workbook}')
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

    rng.value = new_value * 1000

def backup_workbook(workbook):
    """Create a backup copy of the workbook. Returns
    a new workbook object for the current workbook."""
    # TODO: Implement optional backup_dir argument to specify backup directory
    # TODO: Make more robust naming convention
    if not workbook.name.endswith('.xlsx'):
        print(f'Warning: workbook "{workbook.name}" not a .xlsx file')

    wb_name = workbook.name.split('.')[-2]
    wb_full_path = workbook.fullname
    backup_path = f'{os.getcwd()}\\{wb_name}_BACKUP.xlsx'
    try:
        workbook.save(path=backup_path)
        new_workbook = xw.books.open(fullname=wb_full_path)
        workbook.close()
    except:
        raise(f'Error saving {workbook.name} to {backup_path}')
    print(f'Successfully backed up to {backup_path}')
    return new_workbook


def update_named_ranges(json_file, excel_file, backup=True):
    """
    Open a JSON file and an excel file. Update the named
    ranges in the excel file with the corresponding JSON values.
    
    Named ranges may correspond to a particular expression type.
    For example, the measurement SURFACE_SPHERICAL has an expression
    of type "area", along with other expressions. To populate this in 
    Excel, we need to name the range "SURFACE_SPHERICAL.area"
    """
    workbook = xw_get_workbook(excel_file)
    if backup:
        workbook = backup_workbook(workbook)

    # make a dict of named ranges, measurement names, and measurement types
    workbook_named_ranges = dict()
    for named_range in workbook.names:
        measurement_name = named_range.name.split('.')[0]
        try:
            measurement_type = named_range.name.split('.')[1]
        except IndexError:
            measurement_type = None
        workbook_named_ranges[measurement_name] = measurement_type
    # print(workbook_named_ranges)
    with open(json_file, "r") as json_file_handle:
        json_data = json.load(json_file_handle)
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
                print(f'Excel value of {range_name}: {excel_value}')
                # print the value currently in JSON
                print(f'JSON Value: {json_value}')
                write_named_range(workbook, range_name, json_value)

def main():
    # wb = xw_get_workbook("named_range_test.xlsx")
    json_file = "C:\\Users\\frandeen\\Documents\\datum\\json_test.json"
    # update_named_ranges(json_file, "Surface Coating Usable Area.xlsx")
    # update_named_ranges(json_file, "OS Mass Threats & Opportunities.xlsx")
    update_named_ranges(json_file, "named_range_test.xlsx")
    # for named_range in wb.names:
    #     print(f'{named_range.name} = {read_named_range(wb, named_range.name)}')

if __name__ == '__main__':
    main()