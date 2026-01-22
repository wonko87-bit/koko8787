# -*- coding: utf-8 -*-
"""
Maxwell LeadPath v2 - Minimal Test
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")

print("Hello from LeadPath v2 test!")
print("This is a test message")

oDesktop = oDesktop
oProject = oDesktop.GetActiveProject()
oDesign = oProject.GetActiveDesign()
oEditor = oDesign.SetActiveEditor("3D Modeler")

print("Maxwell environment initialized")

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
