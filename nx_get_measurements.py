"""
NX Journal for quickly updating assembly level measurements that are specified in a CSV file.

NX model needs to have named measurements, e.g. "SURFACE_AREA" that are saved (associative)
CSV file currently has four columns, and should have the first row as a header:
FEATURE_NAME,FEATURE_TYPE,VALUE,UNITS

FEATURE_TYPE is should be left blank for most length measurements, but is required for when
a measurement has more than one value, such as a surface area measurement.

TODO: Set this up to just grab all measurements and dump them in a JSON
"""
import NXOpen
import NXOpen.Features
import csv
import sys
import json
from nxmods import nxprint


def find_feature_by_name(feature_name):
    theSession  = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work

    for feature in workPart.Features:
        if (feature.Name == feature_name):
            return feature

    return None

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

    theSession  = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work
    nxprint("Measurement Extractor. Using Pythong Version:")
    nxprint(sys.version)
    find_feature_by_name("SURFACE_AREA_PAINTED")
    update_measurements("C:/Users/frandeen/Documents/NX_Journals/measurements.csv")

if __name__ == '__main__':
    main()
