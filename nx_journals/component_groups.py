# NX 1969
# Journal created by frandeen on Wed May  4 14:35:13 2022 Pacific Daylight Time
#
import math

import NXOpen
import NXOpen.UF
import NXOpen.UIStyler

from nxmods import nxdir, nxprint

# Component groups that NX makes by default
DEFAULT_COMPONENT_GROUPS = [
    "AllComponents",
    "LoadedChangedComponents",
    "UnloadedChangedComponents",
    "CurrentComponents"
]

def main(): 
    theSession  = NXOpen.Session.GetSession()
    workPart = theSession.Parts.Work
    displayPart = theSession.Parts.Display
    
    nxprint(workPart)
    # nxdir(workPart)
    # nxdir(workPart.MeasureManager)
    # nxdir(workPart.ComponentGroups)
    nxprint(workPart.ComponentGroups)
    for g in workPart.ComponentGroups:
        if g.Name in DEFAULT_COMPONENT_GROUPS:
            continue
        nxprint(f"COMPONENT GROUP: {g.Name}")
        nxprint(f"{g.Tag = }")
        nxprint(f"{g.OwningComponent = }")
        nxprint(f"{g.OwningPart = }")
        nxprint("-----------------------")
        for c in g.GetComponents():
            nxprint(c.Name)

    nxprint(f"{c.Name = }")
    nxprint(f"{c.Prototype.Name = }")
    nxprint(f"{c.Prototype.Features = }")
    for f in c.Prototype.Features:
        nxprint(f.Name)
    nxprint(f"{c.Prototype.Bodies = }")
    for b in c.Prototype.Bodies:
        nxprint(b.Name)

    
    # nxprint(workPart.ComponentAssembly)
    # nxdir(workPart.ComponentAssembly)
    # nxdir(workPart.Assemblies)

    # bodybuilder = workPart.MeasureManager.CreateMeasureBodyBuilder(c)
    # nxdir(bodybuilder)


if __name__ == '__main__':
    main()
