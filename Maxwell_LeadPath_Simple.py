# -*- coding: utf-8 -*-
import csv
import os
import ScriptEnv

# Maxwell 환경 초기화
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop = oDesktop
oProject = oDesktop.GetActiveProject()
oDesign = oProject.GetActiveDesign()
oEditor = oDesign.SetActiveEditor("3D Modeler")

print("Maxwell environment initialized!")

# 테스트: Line 하나 생성
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

print("Line created successfully!")
