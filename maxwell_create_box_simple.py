# -*- coding: utf-8 -*-
"""
Maxwell 3D 박스 생성 - 간단한 버전
===================================

실제 Maxwell 2021 R1에서 녹화된 API를 기반으로 한 간단한 스크립트입니다.
파일 상단의 파라미터만 수정하면 바로 사용할 수 있습니다.

사용 방법:
---------
1. 아래 파라미터 값을 원하는 크기로 수정합니다
2. Maxwell에서 Tools → Run Script로 실행합니다
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

# ============================================================================
# 파라미터 설정 (여기를 수정하세요!)
# ============================================================================
box_width = 20      # X 방향 너비 (mm)
box_depth = 10      # Y 방향 깊이 (mm)
box_height = 5      # Z 방향 높이 (mm)
box_name = "Box1"   # 박스 이름
material = "vacuum" # 재질 (vacuum, aluminum, copper, steel, iron 등)

# ============================================================================
# 박스 생성 실행
# ============================================================================

# 프로젝트와 디자인 설정
oProject = oDesktop.GetActiveProject()
if oProject is None:
    oProject = oDesktop.NewProject()
    AddWarningMessage("새 프로젝트를 생성했습니다.")

oDesign = oProject.GetActiveDesign()
if oDesign is None:
    oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
    oDesign = oProject.SetActiveDesign("Maxwell3DDesign1")
    AddWarningMessage("새 Maxwell 3D 디자인을 생성했습니다.")

# 3D Modeler 에디터 가져오기
oEditor = oDesign.SetActiveEditor("3D Modeler")

# 박스 생성 (Maxwell 2021 R1 녹화 API 기반)
oEditor.CreateBox(
    [
        "NAME:BoxParameters",
        "XPosition:=", "0mm",
        "YPosition:=", "0mm",
        "ZPosition:=", "0mm",
        "XSize:=", str(box_width) + "mm",
        "YSize:=", str(box_depth) + "mm",
        "ZSize:=", str(box_height) + "mm"
    ],
    [
        "NAME:Attributes",
        "Name:=", box_name,
        "Flags:=", "",
        "Color:=", "(143 175 143)",
        "Transparency:=", 0,
        "PartCoordinateSystem:=", "Global",
        "UDMId:=", "",
        "MaterialValue:=", '"{}"'.format(material),
        "SurfaceMaterialValue:=", '""',
        "SolveInside:=", True,
        "ShellElement:=", False,
        "ShellElementThickness:=", "0mm",
        "IsMaterialEditable:=", True,
        "UseMaterialAppearance:=", False,
        "IsLightweight:=", False
    ])

# 뷰 맞추기
oEditor.FitAll()

AddWarningMessage("박스 생성 완료: {} ({}x{}x{} mm, 재질: {})".format(
    box_name, box_width, box_depth, box_height, material))
