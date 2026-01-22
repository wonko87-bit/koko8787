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

# 테스트: Polyline으로 Line 생성
point_list = ["NAME:PolylinePoints"]
point_list.append([
    "NAME:PLPoint",
    "X:=", "0mm",
    "Y:=", "0mm",
    "Z:=", "0mm"
])
point_list.append([
    "NAME:PLPoint",
    "X:=", "100mm",
    "Y:=", "0mm",
    "Z:=", "0mm"
])

segment_list = ["NAME:PolylineSegments"]
segment_list.append([
    "NAME:PLSegment",
    "SegmentType:=", "Line",
    "StartIndex:=", 0,
    "NoOfPoints:=", 2
])

oEditor.CreatePolyline(
    [
        "NAME:PolylineParameters",
        "IsPolylineCovered:=", False,
        "IsPolylineClosed:=", False,
        point_list,
        segment_list,
        [
            "NAME:PolylineXSection",
            "XSectionType:=", "None",
            "XSectionOrient:=", "Auto",
            "XSectionWidth:=", "0mm",
            "XSectionTopWidth:=", "0mm",
            "XSectionHeight:=", "0mm",
            "XSectionNumSegments:=", "0",
            "XSectionBendType:=", "Corner"
        ]
    ],
    [
        "NAME:Attributes",
        "Name:=", "TestPolyline",
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

print("Polyline created successfully!")
