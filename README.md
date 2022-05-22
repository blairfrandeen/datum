# DATUM
Tools to pull engineering data from dispirate sources and understand performance metrics.

## Transfer Measurement Data from NX to Excel
Currently this code includes two modules that allow the user to transfer data from measurement features in NX into named ranges in Excel.

**BASIC USAGE**:
- In NX part with saved measurements, run `nx_get_measurements.py` as an NX journal.
- With Excel open, run `python datum/datum_console.py` and use the `dump` command to transfer values to Excel.

### Requirements
- Python 3.8.3 or above
- NX running on Windows. This code works in NX 1953, may work in earlier verisons in which NXOpen supports Python
- Microsoft Excel
- See `requirements.txt` for full list of modules.

### Pulling measurement data from NX with `nx_get_measurements.py`
Open a work part or assembly that has saved measurement features. Use the `Measure` tool and ensure that `associative` is checked under settings - the measurement should show up in the feature navigator. Give your measurement a meaningful name, ideally based on the component. All valid measurements such as points, vectors, surface area, distance, mass, volume, MOI, POI, etc. will be saved.

Ensure the developer tab is activated in NX. See the [NXOpen_Python_tutorials](https://github.com/Foadsf/NXOpen_Python_tutorials) repository for instructions.

Press ALT+F8 to open the journal interface. Navigate to where you saved `nx_journals/nx_get_measurements.py` and run it. If you manage to get `tkinter` to work with the NX python installation, you should get a popup window asking where to save the measurement data. Otherwise it will save your data to `C:\Users\<your_username>\Documents\datum\nx_measurements.json`.

#### Adding `tkinter` to the NX installation
Some additional functionality (file picker) can be added to the NX Python installation. This is hacky af, and at your own risk.
- In your Python installation, find the `\tcl\tcl8.6` and `\tcl\tk8.6` folders. These were under `C:\ProgramData\Anaconda3` for me
- Copy the two folders above to the `\Siemens\NX1953\NXBIN\lib\tcl8.6` folder
- If still not working, edit `tk.tcl` and `init.tcl` to specify verison 8.6.9 rather than 8.6.6


### Populating named Excel ranges with measurement data
In your Excel file, you'll need to name the cells that you want to autopopulate. Select each cell and choose Formulas -> Define Name. You can select single cells, or multiple cells to define vectors, points, or lists. Named ranges of multiple cells that accept lists of values. can be oriented horizontally or vertically. 

The range names must be in the format `<FEATURE_NAME>.<expression_name>`. `<FEATURE_NAME>` is the name of your NX measurement feature; `<expression_name>` can be one of the following:

**Single Value Expressions**
- `surface_area`
- `volume`
- `mass`
- `weight`
- `density`
- `distance`
- `angle`

**Multi-Value Expressions**
- `center_of_mass`
- `moments_of_inertia`
- `first_moments_of_inertia`
- ... and many more

For example, if I wanted to know the mass and center of gravity of my chassis, I would follow these steps:
1. Make the associative measurement in NX and name it `CHASSIS`
2. Name a range with a single cell in Excel as `CHASSIS.mass`
3. Name a range of three cells in Excel `CHASSIS.center_of_mass`

Once the Excel sheet is set up, run `datum/datum_console.py` from the directory where the JSON file was saved. The script will prompt you to choose the JSON file to read from (searches working directory only), and the Excel file to write to (lists open workbooks detected by xlwings). The script will also give you a preview of values to be overwritten, and prompts you prior to doing so. Basic undo functionality is now built in.

Code exists to save a backup copy of your file as `<filename>_BACKUP.xlsx` in the working directory in case you find running this code regrettable.
