# -*- coding: utf-8 -*-
"""
Maxwell 3D - Quarter Arc 테스트
사분원 생성 후 각도만큼 자르기
"""

import csv
import os
import math
import ScriptEnv

# Maxwell 환경 초기화
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop = oDesktop
oProject = oDesktop.GetActiveProject()
oDesign = oProject.GetActiveDesign()
oEditor = oDesign.SetActiveEditor("3D Modeler")


def read_arc_csv(csv_file_path):
    """CSV 파일에서 Arc 데이터를 읽어옵니다."""
    arcs = []

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

    if len(rows) < 2:
        return arcs

    # 데이터 행 처리 (2행부터)
    for i in range(1, len(rows)):
        row = rows[i]
        if not row or all(not cell.strip() for cell in row):
            continue

        try:
            radius = float(row[0].strip())
            angle = float(row[1].strip())

            arcs.append({
                'radius': radius,
                'angle': angle
            })
        except:
            continue

    return arcs


def create_quarter_arc(oEditor, radius, angle, name):
    """
    사분원을 만들고 원하는 각도만큼만 남기기

    1. 원 생성
    2. 사각형 2개로 1사분면만 남기기 (사분원)
    3. 사분원을 (90 - angle)도 회전
    4. 4사분면에 정사각형 생성
    5. Subtract
    """
    center = (0.0, 0.0, 0.0)

    # 1. 원 생성 (XY 평면, 원점 중심)
    circle_name = "{}_Circle".format(name)
    oEditor.CreateCircle(
        [
            "NAME:CircleParameters",
            "IsCovered:=", False,
            "XCenter:=", "0mm",
            "YCenter:=", "0mm",
            "ZCenter:=", "0mm",
            "Radius:=", "{}mm".format(radius),
            "WhichAxis:=", "Z",
            "NumSegments:=", "0"
        ],
        [
            "NAME:Attributes",
            "Name:=", circle_name,
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

    # 2. 1사분면만 남기기 - 사각형 2개로 자르기
    rect_size = radius * 3.0

    # 첫 번째 사각형 - 2,3사분면 덮기 (X < 0)
    rect1_name = "{}_Rect1".format(name)
    oEditor.CreateRectangle(
        [
            "NAME:RectangleParameters",
            "IsCovered:=", True,
            "XStart:=", "{}mm".format(-rect_size),
            "YStart:=", "{}mm".format(-rect_size),
            "ZStart:=", "0mm",
            "Width:=", "{}mm".format(rect_size),
            "Height:=", "{}mm".format(rect_size * 2),
            "WhichAxis:=", "Z"
        ],
        [
            "NAME:Attributes",
            "Name:=", rect1_name,
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

    # 두 번째 사각형 - 3,4사분면 덮기 (Y < 0)
    rect2_name = "{}_Rect2".format(name)
    oEditor.CreateRectangle(
        [
            "NAME:RectangleParameters",
            "IsCovered:=", True,
            "XStart:=", "{}mm".format(-rect_size),
            "YStart:=", "{}mm".format(-rect_size),
            "ZStart:=", "0mm",
            "Width:=", "{}mm".format(rect_size * 2),
            "Height:=", "{}mm".format(rect_size),
            "WhichAxis:=", "Z"
        ],
        [
            "NAME:Attributes",
            "Name:=", rect2_name,
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

    # Subtract로 1사분면만 남기기
    oEditor.Subtract(
        ["NAME:Selections", "Blank Parts:=", circle_name, "Tool Parts:=", "{},{}".format(rect1_name, rect2_name)],
        ["NAME:SubtractParameters", "KeepOriginals:=", False]
    )

    # 3. 사분원을 (90 - angle)도 회전
    rotate_angle = 90.0 - angle
    oEditor.Rotate(
        ["NAME:Selections", "Selections:=", circle_name],
        [
            "NAME:RotateParameters",
            "RotateAxis:=", "Z",
            "RotateAngle:=", "{}deg".format(rotate_angle)
        ]
    )

    # 4. 4사분면에 정사각형 생성
    rect3_name = "{}_Rect3".format(name)
    oEditor.CreateRectangle(
        [
            "NAME:RectangleParameters",
            "IsCovered:=", True,
            "XStart:=", "0mm",
            "YStart:=", "{}mm".format(-rect_size),
            "ZStart:=", "0mm",
            "Width:=", "{}mm".format(rect_size),
            "Height:=", "{}mm".format(rect_size),
            "WhichAxis:=", "Z"
        ],
        [
            "NAME:Attributes",
            "Name:=", rect3_name,
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

    # 5. Subtract로 원하는 각도만 남기기
    oEditor.Subtract(
        ["NAME:Selections", "Blank Parts:=", circle_name, "Tool Parts:=", rect3_name],
        ["NAME:SubtractParameters", "KeepOriginals:=", False]
    )

    # 최종 이름 변경
    try:
        oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Geometry3DAttributeTab",
                    [
                        "NAME:PropServers",
                        circle_name
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Name",
                            "Value:=", name
                        ]
                    ]
                ]
            ]
        )
    except:
        pass


def create_arcs_from_csv(csv_file_path, name_prefix="Arc"):
    """CSV 파일에서 Arc들을 생성"""
    arcs = read_arc_csv(csv_file_path)

    if not arcs:
        return

    for idx, arc_data in enumerate(arcs):
        arc_name = "{}_{}".format(name_prefix, idx + 1)
        create_quarter_arc(oEditor, arc_data['radius'], arc_data['angle'], arc_name)

    # 뷰 맞추기
    try:
        oEditor.FitAll()
    except:
        pass


# 스크립트 실행
try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
except:
    import sys
    if len(sys.argv) > 0:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_dir = os.getcwd()

csv_file = os.path.join(script_dir, "QuarterArcDim.csv")

if os.path.exists(csv_file):
    create_arcs_from_csv(csv_file, "Arc")
