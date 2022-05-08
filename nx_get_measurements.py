﻿"""
NX Journal for quickly updating assembly level measurements
that are specified in a CSV file.

NX model needs to have named measurements, e.g. "SURFACE_AREA"
that are saved (associative)

TODO: Make this also work for PMI
TODO: Be able to take arguments passed through NX interface
TODO: Implement proper logging
"""
import json
import os
import re
import sys
import datetime
from tkinter import TclError

import NXOpen

from nxmods import nxprint

# TODO: Write a session wrapper that carries around the listing window.
# Change nxprint to be a method of the session so that the LW only
# has to be opened once.

# TODO: Write the following metadata to JSON
# Work part
# Revision
# Date

# user settable defaults for where to save JSON file
JSON_DEFAULT_DIR = f"C:\\Users\\{os.getlogin()}\\Documents\\datum"
JSON_DEFAULT_FILE = "nx_measurements.json"
DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"


def get_metadata(nxSession):
    UNIT_ENUM = {0: "Inches", 1: "Millimeters"}
    metadata = dict()
    workPart = nxSession.Parts.Work
    metadata["part_name"], metadata["part_rev"]\
        = workPart.FullPath.split('/')
    metadata["part_units"] = UNIT_ENUM[int(str(workPart.PartUnits))]
    metadata["retrieval_date"] = datetime.datetime.today().strftime(DATETIME_FORMAT)
    metadata["user"] = os.getlogin()

    for key in metadata.keys():
        nxprint(f"{key}: {metadata[key]}")

    return metadata


def find_feature_by_name(feature_name):
    theSession = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work

    for feature in workPart.Features:
        if feature.Name == feature_name:
            return feature

    return None


def check_feature_errors(nxSession=None):
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
                "description": "WCS",
                "type": "Point",
                "value": {
                    "x": nxSession.Parts.Work.WCS.Origin.X,
                    "y": nxSession.Parts.Work.WCS.Origin.Y,
                    "z": nxSession.Parts.Work.WCS.Origin.Z
                }
            }
        ]
    }

    return wcs

def export_measurements(json_export_file, nxSession=None):
    #   Ensure that measruements are updated in the model
    #   Menu: Tools->Update->Interpart Update->Update All
    if nxSession:
        # TODO: Error handling if no NX session passed to function??
        markId2 = nxSession.SetUndoMark(
            NXOpen.Session.MarkVisibility.Visible, "Update Session"
        )
        nxSession.UpdateManager.DoInterpartUpdate(markId2)
        workPart = nxSession.Parts.Work

    # check_feature_errors(nxSession)
    num_measurements_found = 0
    measurement_features = {"measurements": []}
    measurement_features["measurements"].append(get_WCS(nxSession))

    for feature in workPart.Features:
        if "MEASUREMENT" in feature.FeatureType:
            num_measurements_found += 1
            # nxprint(f'{feature.Name}')
            current_feature = {"name": feature.Name, "expressions": []}
            for expr in feature.GetExpressions():
                # typical type string: "p7( Face Measure : area )"
                # the regex below extracts "area"
                expr_description = re.search(r"(?<=\d\) )\w+(?=\))", expr.Description)
                if expr_description is not None:
                    expr_description = expr_description[0]
                # if no expression type, likely a distance measurement.
                # leave this as None / null
                current_expr = {
                    # "type": expr_type,
                    "description": expr_description,
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
                        "z": expr.PointValue.Z
                    }
                elif expr.Type == "Vector":
                    expr_value = {
                        "x": expr.VectorValue.X,
                        "y": expr.VectorValue.Y,
                        "z": expr.VectorValue.Z
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

    return num_measurements_found


def get_json_file_path():
    """Opens dialog box for user to choose where to save the measurements.
    Uses default directory in case of failure.

    Requires very hacky work-around of installing tk and tcl in NX directories"""
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialdir=JSON_DEFAULT_DIR,
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

    return f"{JSON_DEFAULT_DIR}\\{JSON_DEFAULT_FILE}"


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
