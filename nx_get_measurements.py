"""
NX Journal for quickly updating assembly level measurements that are specified in a CSV file.

NX model needs to have named measurements, e.g. "SURFACE_AREA" that are saved (associative)
CSV file currently has four columns, and should have the first row as a header:
FEATURE_NAME,FEATURE_TYPE,VALUE,UNITS

FEATURE_TYPE is should be left blank for most length measurements, but is required for when
a measurement has more than one value, such as a surface area measurement.

TODO: Set this up to just grab all measurements and dump them in a JSON
TODO: Make this also work for PMI
"""
import NXOpen
import csv
import sys
import json
import re
from nxmods import nxprint


def find_feature_by_name(feature_name):
    theSession  = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work

    for feature in workPart.Features:
        if (feature.Name == feature_name):
            return feature

    return None

def export_measurements(nxSession=None):
    #   Ensure that measruements are updated in the model
    #   Menu: Tools->Update->Interpart Update->Update All
    if nxSession:
        markId2 = nxSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Update Session")
        nxSession.UpdateManager.DoInterpartUpdate(markId2)
        workPart = nxSession.Parts.Work

    num_measurements_found = 0
    measurement_features = []
    for feature in workPart.Features:
        if "MEASUREMENT" in feature.FeatureType:
            num_measurements_found += 1
            # nxprint(f'{feature.Name}')
            current_feature = {'name': feature.Name,\
                'expressions': []}
            for expr in feature.GetExpressions():
                try:
                    expr_type = re.search('(?<=\d\) )\w+(?=\))',expr.Description)
                    if expr_type is not None:
                        expr_type = expr_type[0]
                    current_expr = {\
                        'type': expr_type,
                        'value': expr.Value,
                        'units': expr.Units.Name }
                    # nxprint(f'{expr.Description} - {expr.Value} [{expr.Units.Name}]')
                    # nxprint(f'{expr.IsMeasurementExpression}, {expr.Name}, {expr.Tag}, {expr.Type}')
                    current_feature["expressions"].append(current_expr)
                except NXOpen.NXException:
                    pass
            measurement_features.append(current_feature)

    with open("C:/Users/frandeen/Documents/datum/json_test.json", "w") as json_file:
        json.dump(measurement_features, json_file, indent=4)

    return num_measurements_found

def update_measurements(measurement_database):
    theSession  = NXOpen.Session.GetSession()
    #   Ensure that measruements are updated in the model
    #   Menu: Tools->Update->Interpart Update->Update All
    markId2 = theSession.SetUndoMark(NXOpen.Session.MarkVisibility.Visible, "Update Session")
    theSession.UpdateManager.DoInterpartUpdate(markId2)

    csvfile = open(measurement_database, "r", newline="")
    reader = csv.DictReader(csvfile)
    new_rows_list = []
    for row in reader:
        measurement_feature = find_feature_by_name(row["FEATURE_NAME"])
        if measurement_feature:
            for expr in measurement_feature.GetExpressions():
                if row["FEATURE_TYPE"] in expr.Description:
                    try:
                        row["VALUE"] = expr.Value
                        row["UNITS"] = expr.Units.Abbreviation
                    except:
                        continue
        else:
            nxprint(f'Feature {row["FEATURE_NAME"]} not found!')
        new_rows_list.append(row)
        nxprint(f'{row["FEATURE_NAME"]}\t {row["FEATURE_TYPE"]}\t {row["VALUE"]}\t {row["UNITS"]}')
    csvfile.close()

    csvwrite = open("C:/Users/frandeen/Documents/NX_Journals/test.csv","w", newline="")
    fieldnames = ["FEATURE_NAME", "FEATURE_TYPE", "VALUE", "UNITS"]
    writer = csv.DictWriter(csvwrite, fieldnames)
    writer.writeheader()
    writer.writerows(new_rows_list)

    csvwrite.close()
    # Let's also try this with JSON:
    with open("C:/Users/frandeen/Documents/NX_Journals/json_test.json", "w") as json_file:
        json.dump(new_rows_list, json_file, indent=4)

def main():
    nxSession  = NXOpen.Session.GetSession()
    nxprint("Measurement Extractor. Using Pythong Version:")
    nxprint(sys.version)
    # find_feature_by_name("SURFACE_AREA_PAINTED")
    # update_measurements("C:/Users/frandeen/Documents/NX_Journals/measurements.csv")
    num_measurements = export_measurements(nxSession)
    nxprint(f'Found total of {num_measurements} measurement features.')

if __name__ == '__main__':
    main()
