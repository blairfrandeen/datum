# DATUM
Tools to pull engineering data from dispirate sources and understand performance metrics.

## Transfer Measurement Data from NX to Excel
Currently this code includes two modules that allow the user to transfer limited data from measurement features in NX into named ranges in Excel.

### Requirements
- Python 3.6 or above
- NX running on Windows. This code works in NX 1953, may work in earlier verisons in why NXOpen supports Python
- Microsoft Excel
- Python `xlwings` module

### Pulling measurement data from NX with `nx_get_measurements.py`
Open a work part or assembly that has saved measurement features. Use the `Measure` tool and ensure that `associative` is checked under settings - the measurement should show up in the feature navigator. Give your measurement a meaningful name such as `MASS`, `SURFACE_AREA` or `DIAMETER`.

> NOTE: Only numerical measurements are supported. Points, vectors, and lists will be skipped.

Ensure the developer tab is activated in NX. See the [NXOpen_Python_tutorials](https://github.com/Foadsf/NXOpen_Python_tutorials) repository for instructions.

Press ALT+F8 to open the journal interface. Navigate to where you saved `nx_get_measurements.py` and run it. If you manage to get `tkinter` to work with the NX python installation, you should get a popup window asking where to save the measurement data. Otherwise it will save your data to `C:\Users\<your_username>\Documents\datum\nx_measurements.json`.

#### Adding `tkinter` to the NX installation
Some additional functionality (file picker) can be added to the NX Python installation. This is hacky af, and at your own risk.
- In your Python installation, find the `\tcl\tcl8.6` and `\tcl\tk8.6` folders. These were under `C:\ProgramData\Anaconda3` for me
- Copy the two folders above to the `\Siemens\NX1953\NXBIN\lib\tcl8.6` folder
- If still not working, edit `tk.tcl` and `init.tcl` to specify verison 8.6.9 rather than 8.6.6


### Populating named Excel ranges with measurement data
In your Excel file, you'll need to name the cells that you want to autopopulate. Select each cell and choose Formulas -> Define Name. Note that named ranges of multiple cells are not supported. These range names must correspond to the names you gave your measurement features in NX:
- If the measurement result is a **single number**, simply match the name of the NX feature. An example of this is a measurement of the distance between two points.
- If the measurement result has **more than one expression**, you'll need to name the range `<FEATURE_NAME>.expression_name` where:
    - `<FEATURE_NAME>` is the name of the feature in NX
    - `expression_name` is the type of expression you want. Examples are `mass`, `area`, or `radius`.

Once the Excel sheet is set up, run `xl_populate_named_ranges.py` from the directory where the JSON file was saved. The script will prompt you to choose the JSON file to read from (searches working directory only), and the Excel file to write to (lists open workbooks detected by xlwings). The script will also give you a preview of values to be overwritten, and prompts you prior to doing so. A backup copy of your file will be saved as `<filename>_BACKUP.xlsx` in the working directory in case you find running this code regrettable.
