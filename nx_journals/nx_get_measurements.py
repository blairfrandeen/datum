"""
NX Journal for quickly updating assembly level measurements
that are specified in a CSV file.

NX model needs to have named measurements, e.g. "SURFACE_AREA"
that are saved (associative)

TODO: Make this also work for PMI
TODO: Be able to take arguments passed through NX interface
TODO: Implement proper logging
"""
import datetime
import json
import os
import re
import sys
import sqlite3
from tkinter import TclError

try:
    import NXOpen
except ModuleNotFoundError as err:  # pragma: no cover
    print(err)
    print("Please run this module from NX.")

try:
    from nxmods import nxprint
except ModuleNotFoundError:
    from nx_journals.nxmods import nxprint

try:
    from datum import __version__ as datum_version
except ModuleNotFoundError:  # pragma: no cover
    nxprint("datum module not found.")
    datum_version = "UNKNOWN"

# user settable defaults for where to save JSON file
DATUM_DIR = f"C:\\Users\\{os.getlogin()}\\Documents\\datum"
DATUM_DB_FILE = f"C:\\Users\\{os.getlogin()}\\Documents\\datum\\datum.db"
JSON_DEFAULT_FILE = "nx_measurements.json"
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"

sys.path.insert(0, DATUM_DIR)


# TODO: Test case in test_db
def write_metadata_db(metadata_dict: dict) -> None:
    """Write metadata to an SQLite DB"""
    db_connection = sqlite3.connect(DATUM_DB_FILE, detect_types=sqlite3.PARSE_DECLTYPES)
    cur = db_connection.cursor()
    # NOTE: Keys in metadata_dict must match keys in table
    metadata_table_create = """--sql
        CREATE TABLE IF NOT EXISTS source_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            part_name TEXT,
            part_path TEXT,
            part_rev TEXT,
            part_units TEXT,
            user TEXT,
            computer TEXT,
            datum_version TEXT,
            source_type TEXT,
            source_version TEXT,
            retrieval_ts TIMESTAMP DEFAULT CURRENT_TIMESTAMP /* timestamp for source access */
        )
    """

    cur.execute(metadata_table_create)

    key_str = ", ".join([key for key in metadata_dict.keys()])
    value_str = '"' + '", "'.join(value for value in metadata_dict.values()) + '"'
    insert_command = (
        "INSERT INTO source_history (" + key_str + ") VALUES (" + value_str + ")"
    )
    nxprint(insert_command)
    cur.execute(insert_command)
    get_last_key = """--sql
        SELECT MAX(id) FROM source_history 
    """
    cur.execute(get_last_key)
    last_key = cur.fetchone()[0]
    nxprint(f"{last_key = }")

    db_connection.commit()
    db_connection.close()

    return None


# ATTEMPT TO IMPORT OTHER MODULES - IN PROGRESS
# FAILS WHEN IT LOOKS FOR 'pywintypes'
# sys.path.insert(0, f"{DATUM_DIR}\\venv\\Lib\\site-packages")
# sys.exec_prefix = f"{DATUM_DIR}\\venv"
# nxprint(f"{sys.exec_prefix = }")
# import xlwings as xw


def get_metadata(nxSession) -> dict:
    """Generate metadata dictionary based on work part user parameters."""
    UNIT_ENUM = {0: "Inches", 1: "Millimeters"}
    metadata = dict()
    workPart = nxSession.Parts.Work
    if "/" in workPart.FullPath:
        metadata["part_name"], metadata["part_rev"] = workPart.FullPath.split("/")
    else:
        metadata["part_name"] = workPart.Name
        metadata["part_rev"] = None
    metadata["part_path"] = workPart.FullPath
    metadata["part_units"] = UNIT_ENUM[int(str(workPart.PartUnits))]
    metadata["retrieval_ts"] = datetime.datetime.today().strftime(DATETIME_FORMAT)
    metadata["user"] = os.getlogin()
    metadata["computer"] = os.environ["COMPUTERNAME"]
    metadata["datum_version"] = datum_version
    metadata["source_type"] = "NX"
    metadata["source_version"] = str(nxSession.ReleaseNumber)

    for key in metadata.keys():
        nxprint(f"{key}: {metadata[key]}")

    return {"METADATA": metadata}


def find_feature_by_name(feature_name):
    theSession = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work

    for feature in workPart.Features:
        if feature.Name == feature_name:
            return feature

    return None


# NOTE: NOT IMPLEMENTED
def check_feature_errors(nxSession=None):  # pragma: no cover
    feature_update_status = nxSession.Parts.Work.FeatureUpdateStatus
    print(feature_update_status)
    print(feature_update_status.Feature.Name)
    print(feature_update_status.Status)


def get_WCS(nxSession):
    """Return the WCS in a dict"""
    wcs = {
        "name": "World Coordinate System",
        "expressions": [
            {
                "name": "WCS",
                "type": "Point",
                "value": {
                    "x": nxSession.Parts.Work.WCS.Origin.X,
                    "y": nxSession.Parts.Work.WCS.Origin.Y,
                    "z": nxSession.Parts.Work.WCS.Origin.Z,
                },
            }
        ],
    }

    return wcs


def export_measurements(json_export_file, nxSession):
    #   Ensure that measruements are updated in the model
    #   Menu: Tools->Update->Interpart Update->Update All
    markId2 = nxSession.SetUndoMark(
        NXOpen.Session.MarkVisibility.Visible, "Update Session"
    )
    # TODO: Error handling of NXOpen.NXException / Update Undo happens
    nxSession.UpdateManager.DoInterpartUpdate(markId2)
    workPart = nxSession.Parts.Work

    # check_feature_errors(nxSession)
    num_measurements_found = 0
    measurement_features = {"measurements": []}
    measurement_features["measurements"].append(get_WCS(nxSession))

    for feature in workPart.Features:
        if feature.Suppressed:
            nxprint(f"Feature {feature.Name} is suppressed.")
            continue
        if "MEASUREMENT" in feature.FeatureType:
            num_measurements_found += 1
            point_count = 0
            current_feature = {"name": feature.Name, "expressions": []}
            for expr in feature.GetExpressions():
                # typical type string: "p7( Face Measure : area )"
                # the regex below extracts "area"
                expr_name = re.search(r"(?<=\d\) )\w+(?=\))", expr.Description)
                if expr_name is None:
                    if expr.Type == "Point":
                        # TODO: If only a single point in expression,
                        # name it "point" instead of "point_1"
                        point_count += 1
                        expr_name = f"point_{point_count}"
                    elif expr.Type == "Number":
                        if expr.Units.Name == "Degrees":
                            expr_name = "angle"
                        else:
                            expr_name = "distance"
                    else:
                        expr_name = "UNKNOWN"
                else:
                    expr_name = expr_name[0]
                # if no expression type, likely a distance measurement.
                # leave this as None / null
                current_expr = {
                    "name": expr_name,
                    "type": expr.Type,
                }

                expr_value = None
                if expr.Type == "Number":
                    expr_value = expr.Value
                    current_expr["units"] = expr.Units.Name
                elif expr.Type == "Point":
                    expr_value = {
                        "x": expr.PointValue.X,
                        "y": expr.PointValue.Y,
                        "z": expr.PointValue.Z,
                    }
                elif expr.Type == "Vector":
                    expr_value = {
                        "x": expr.VectorValue.X,
                        "y": expr.VectorValue.Y,
                        "z": expr.VectorValue.Z,
                    }
                elif expr.Type == "List":
                    expr_value = expr.GetListValue()
                elif expr.Type == "String":
                    expr_value = expr.StringValue
                else:
                    continue

                current_expr["value"] = expr_value

                current_feature["expressions"].append(current_expr)

            measurement_features["measurements"].append(current_feature)

    with open(json_export_file, "w") as json_file:
        measurement_features.update(get_metadata(nxSession))
        json.dump(measurement_features, json_file, indent=4)

    write_metadata_db(get_metadata(nxSession)["METADATA"])
    return num_measurements_found


def get_json_file_path():
    """Opens dialog box for user to choose where to save the measurements.
    Uses default directory in case of failure.

    Requires very hacky work-around of installing
    tk and tcl in NX directories"""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialdir=DATUM_DIR,
            initialfile=JSON_DEFAULT_FILE,
            title="Choose JSON file for measurement export",
        )
        root.destroy()
        if len(file_path) > 0:
            return file_path
        else:
            return None
    except ImportError as error:
        nxprint(error)
        nxprint("Invalid tkinter installation, using default JSON path.")
    except TclError as error:
        nxprint(error)
        nxprint("Invalid TCL installation, using default JSON path.")

    return f"{DATUM_DIR}\\{JSON_DEFAULT_FILE}"


def main():
    nxSession = NXOpen.Session.GetSession()
    nxprint("Measurement Extractor. Using Python Version:")
    nxprint(sys.version)
    json_export_path = get_json_file_path()
    if json_export_path is not None:
        nxprint(f"exporting to {json_export_path}")
        num_measurements = export_measurements(json_export_path, nxSession)
        nxprint(f"found total of {num_measurements} measurement features.")
    else:
        nxprint("no json file specified, exiting...")


if __name__ == "__main__":
    main()
