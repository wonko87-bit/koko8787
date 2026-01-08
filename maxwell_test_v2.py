# -*- coding: utf-8 -*-
"""
Maxwell V2 스크립트 디버그 테스트
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

print("=" * 60)
print("V2 Script Test Started")
print("=" * 60)

import csv
import os

# 프로젝트 확인
oProject = oDesktop.GetActiveProject()
if oProject is None:
    oProject = oDesktop.NewProject()
    print("새 프로젝트 생성")

oDesign = oProject.GetActiveDesign()
if oDesign is None:
    oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
    oDesign = oProject.GetActiveDesign()
    print("새 디자인 생성")

oEditor = oDesign.SetActiveEditor("3D Modeler")
print("Editor 활성화 완료")

# Rectangle 테스트
print("\nRectangle 생성 테스트...")
oEditor.CreateRectangle(
    [
        "NAME:RectangleParameters",
        "IsCovered:=", True,
        "XStart:=", "0mm",
        "YStart:=", "0mm",
        "ZStart:=", "-100mm",
        "Width:=", "50mm",
        "Height:=", "100mm",
        "WhichAxis:=", "Z"
    ],
    [
        "NAME:Attributes",
        "Name:=", "TestRect",
        "Flags:=", "",
        "Color:=", "(132 132 193)",
        "Transparency:=", 0.4,
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
print("Rectangle 생성 성공!")

# Move 테스트
print("\nMove 테스트...")
oEditor.Move(
    [
        "NAME:Selections",
        "Selections:=", "TestRect",
        "NewPartsModelFlag:=", "Model"
    ],
    [
        "NAME:TranslateParameters",
        "TranslateVectorX:=", "-25mm",
        "TranslateVectorY:=", "-50mm",
        "TranslateVectorZ:=", "0mm"
    ]
)
print("Move 성공!")

# Sweep 테스트
print("\nSweep 테스트...")
oEditor.SweepAlongVector(
    [
        "NAME:Selections",
        "Selections:=", "TestRect",
        "NewPartsModelFlag:=", "Model"
    ],
    [
        "NAME:VectorSweepParameters",
        "DraftAngle:=", "0deg",
        "DraftType:=", "Round",
        "CheckFaceFaceIntersection:=", False,
        "SweepVectorX:=", "0mm",
        "SweepVectorY:=", "0mm",
        "SweepVectorZ:=", "200mm"
    ]
)
print("Sweep 성공!")

oEditor.FitAll()

print("\n" + "=" * 60)
print("테스트 완료! 모든 기능이 정상 작동합니다.")
print("=" * 60)
