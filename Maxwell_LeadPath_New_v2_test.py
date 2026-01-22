# -*- coding: utf-8 -*-
"""
Maxwell LeadPath v2 - Minimal Test
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")

print("Hello from LeadPath v2 test!")
print("This is a test message")

oDesktop = oDesktop
print("oDesktop: {}".format(oDesktop))

oProject = oDesktop.GetActiveProject()
print("oProject: {}".format(oProject))

if oProject is None:
    print("No active project! Creating new project...")
    oProject = oDesktop.NewProject()
    print("New project created")

oDesign = oProject.GetActiveDesign()
print("oDesign: {}".format(oDesign))

if oDesign is None:
    print("No active design! Creating new Maxwell 3D design...")
    oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
    oDesign = oProject.GetActiveDesign()
    print("New design created")

oEditor = oDesign.SetActiveEditor("3D Modeler")
print("oEditor: {}".format(oEditor))
print("Maxwell environment initialized successfully!")

# 간단한 Line 하나만 생성
oEditor.CreateLine(
    [
        "NAME:LineParameters",
        "XStart:=", "0mm",
        "YStart:=", "0mm",
        "ZStart:=", "0mm",
        "XEnd:=", "100mm",
        "YEnd:=", "0mm",
        "ZEnd:=", "0mm"
    ],
    [
        "NAME:Attributes",
        "Name:=", "TestLine",
        "Flags:=", "",
        "Color:=", "(143 175 143)",
        "Transparency:=", 0,
        "PartCoordinateSystem:=", "Global",
        "UDMId:=", "",
        "MaterialValue:=", "\"vacuum\"",
        "SurfaceMaterialValue:=", "\"\"",
        "SolveInside:=", True,
        "ShellElement:=", False,
        "ShellElementThickness:=", "0mm",
        "IsMaterialEditable:=", True,
        "UseMaterialAppearance:=", False,
        "IsLightweight:=", False
    ]
)

print("Test line created!")
