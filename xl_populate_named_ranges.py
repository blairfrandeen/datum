import xlwings as xw
import os
import json

def xw_get_workbook(target_workbook):
    for book in xw.apps[0].books:
        if book.name == target_workbook:
            return book
    print(f'Error: Target workbook {target_workbook} not found in {os.getcwd()}')
    return None

def xw_get_named_range(workbook, range_name):
    if range_name in workbook.names:
        # print(f'Found {range_name}')
        # TODO: Write better regex for this. This will fail for worksheet
        # names with '!' included, and has not been tested.
        worksheet = workbook.names[range_name].refers_to.split('!')[0][1:]
        cell = workbook.names[range_name].refers_to.split('!')[1]
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

    rng.value = new_value

def update_named_ranges(json_file, excel_file):
    """
    Open a JSON file and an excel file. Update the named
    ranges in the excel file with the corresponding JSON values.
    
    Named ranges may correspond to a particular expression type.
    For example, the measurement SURFACE_SPHERICAL has an expression
    of type "area", along with other expressions. To populate this in 
    Excel, we need to name the range "SURFACE_SPHERICAL.area"
    """
    workbook = xw_get_workbook(excel_file)
    # make a dict of named ranges, measurement names, and measurement types
    workbook_named_ranges = dict()
    for named_range in workbook.names:
        measurement_name = named_range.name.split('.')[0]
        try:
            measurement_type = named_range.name.split('.')[1]
        except IndexError:
            measurement_type = None
        workbook_named_ranges[measurement_name] = measurement_type
    print(workbook_named_ranges)
    with open(json_file, "r") as json_file_handle:
        json_data = json.load(json_file_handle)
        for measurement in json_data['measurements']:
            if measurement['name'] in workbook_named_ranges.keys():
                range_name = measurement['name']
                range_type = workbook_named_ranges[measurement['name']]
                if range_type:
                    range_name = f'{range_name}.{range_type}'
                if workbook_named_ranges[measurement['name']] is not None:
                    measurement_type = workbook_named_ranges[measurement['name']]
                # print the value currently in JSON
                # print the value currently in Excel
                excel_value = read_named_range(workbook, range_name)
                print(f'Excel value of {range_name}: {excel_value}')

def main():
    wb = xw_get_workbook("named_range_test.xlsx")
    json_file = "C:\\Users\\frandeen\\Documents\\datum\\json_test.json"
    update_named_ranges(json_file, "named_range_test.xlsx")
    # update_named_ranges(json_file, "OS Mass Threats & Opportunities.xlsx")
    # for named_range in wb.names:
    #     print(f'{named_range.name} = {read_named_range(wb, named_range.name)}')

if __name__ == '__main__':
    main()