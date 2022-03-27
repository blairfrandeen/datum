"""
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
from tkinter import TclError

import NXOpen

from nxmods import nxprint

# user settable defaults for where to save JSON file
user = os.getlogin()
JSON_DEFAULT_DIR = f"C:\\Users\\{user}\\Documents\\datum\\"
JSON_DEFAULT_FILE = "nx_measurements.json"


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
    for feature in workPart.Features:
        if "MEASUREMENT" in feature.FeatureType:
            num_measurements_found += 1
            # nxprint(f'{feature.Name}')
            current_feature = {"name": feature.Name, "expressions": []}
            for expr in feature.GetExpressions():
                try:
                    # typical type string: "p7( Face Measure : area )"
                    # the regex below extracts "area"
                    expr_type = re.search(r"(?<=\d\) )\w+(?=\))", expr.Description)
                    if expr_type is not None:
                        expr_type = expr_type[0]
                    # if no expression type, likely a distance measurement.
                    # leave this as None / null
                    current_expr = {
                        "type": expr_type,
                        "value": expr.Value,
                        "units": expr.Units.Name,
                    }
                    #  nxprint(
                    #  f'{expr.Description} - {expr.Value} [{expr.Units.Name}]'
                    #  )
                    #  nxprint(
                    #  f'{expr.IsMeasurementExpression}, {expr.Name}, {expr.Tag}, {expr.Type}'
                    #  )
                    current_feature["expressions"].append(current_expr)
                except NXOpen.NXException as nx_except_error:
                    nxprint(nx_except_error)
                    nxprint(
                        f"WARNING: {expr.Description} was not saved in JSON. Cannot handle {expr.Type}."
                    )
                    # JUSTIFICATION:
                    # This exception will be thrown when expr.Value is called on a non-number
                    # expression, such as a point or a vector. Currently this script can
                    # only handle number expressions
                    # TODO: refactor `expr_type` to something that doesn't conflict with expr.Type
                    # TODO: Allow for JSON serialization of List, Vector, and Point types,
                    #   And ensure that these can be accessed.

            measurement_features["measurements"].append(current_feature)

    with open(json_export_file, "w") as json_file:
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
        nxprint(f"Exporting to {json_export_path}")
        num_measurements = export_measurements(json_export_path, nxSession)
        nxprint(f"Found total of {num_measurements} measurement features.")
    else:
        nxprint("No JSON file specified, exiting...")


if __name__ == "__main__":
    main()
