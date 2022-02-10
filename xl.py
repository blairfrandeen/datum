# Some functions for working with named ranges in excel
from openpyxl import load_workbook

def get_named_range_cell(range_name, workbook):
    """
    Return a tuple of the sheet and cell for a named range.
    Only reliable for ranges with a single cell.
    """
    try:
        rng = workbook.defined_names[range_name]
    except:
        print(f'Range {range_name} not found.')
        return (None, None)
    sheet, cell = next(rng.destinations)
    if len(cell.split(':')) > 1:
        print(f'Warning: Range {range_name} is more than one cell; returning only first cell in range.')
        cell = cell.split(':')[0]
    return (sheet, cell)

def read_named_range(range_name, workbook):
    """Read a value from a named range in the work book"""
    sheet, cell = get_named_range_cell(range_name, workbook)
    if sheet is None: return None

    value = workbook[sheet][cell].value
    return value

def write_named_range(range_name, value, workbook):
    """Write a value to a named range in the workbook"""
    sheet, cell = get_named_range_cell(range_name, workbook)
    if sheet is None: return None

    workbook[sheet][cell].value = value


def main():
    filename = "test.xlsx"
    wkbk = load_workbook(filename, data_only=False)
    print(f'Radius: {read_named_range("Radius", wkbk)}')
    area = read_named_range("Area", wkbk)
    print(f'Area: {area}')
    write_named_range("Radius", 5, wkbk)
    print(f'Radius: {read_named_range("Radius", wkbk)}')
    print(f'Big Range: {read_named_range("bigrange", wkbk)}')
    try:
        wkbk.save(filename)
    except:
        print(f'Warning: could not access {filename}. Saved a copy.')
        wkbk.save(f'{filename.split(".")[0]}-1.xlsx')

if __name__ == '__main__':
    main()
