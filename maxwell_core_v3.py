# -*- coding: utf-8 -*-
"""
Maxwell 3D - 변압기 철심 생성 V3 (단순 버전)
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_csv_data(csv_file_path):
    """CSV 파일에서 철심 데이터를 읽어옵니다."""
    data = []
    gap = None  # E1 값

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

        # E1 (row 0, col 4): gap distance
        if len(rows) > 0 and len(rows[0]) > 4:
            gap = float(rows[0][4])

        # 데이터 행 읽기 (A, B, C 열)
        for i, row in enumerate(rows):
            if len(row) >= 3:
                try:
                    a = float(row[0])  # A열: 메인 레그 X
                    b = float(row[1])  # B열: Y (공통)
                    c = float(row[2])  # C열: 사이드 레그 X
                    data.append({'A': a, 'B': b, 'C': c})
                except ValueError:
                    continue

    return data, gap


def create_rectangle(oEditor, x_start, y_start, z_start, width, height, name):
    """XY 평면에 Rectangle 생성"""
    oEditor.CreateRectangle(
        [
            "NAME:RectangleParameters",
            "IsCovered:=", True,
            "XStart:=", "{}mm".format(x_start),
            "YStart:=", "{}mm".format(y_start),
            "ZStart:=", "{}mm".format(z_start),
            "Width:=", "{}mm".format(width),
            "Height:=", "{}mm".format(height),
            "WhichAxis:=", "Z"
        ],
        [
            "NAME:Attributes",
            "Name:=", name,
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
    print("  생성: {}".format(name))


def move_object(oEditor, obj_name, dx, dy, dz):
    """객체를 이동"""
    oEditor.Move(
        [
            "NAME:Selections",
            "Selections:=", obj_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:TranslateParameters",
            "TranslateVectorX:=", "{}mm".format(dx),
            "TranslateVectorY:=", "{}mm".format(dy),
            "TranslateVectorZ:=", "{}mm".format(dz)
        ]
    )
    print("  이동: {} → ({}, {}, {})".format(obj_name, dx, dy, dz))


def create_core_from_csv(csv_file_path, name_prefix="Core"):
    """CSV 파일에서 변압기 철심을 생성"""
    print("=" * 60)
    print("Maxwell 3D - 변압기 철심 생성 V3")
    print("=" * 60)

    # CSV 데이터 읽기
    data, gap = read_csv_data(csv_file_path)

    if not data:
        print("오류: CSV 파일에서 데이터를 읽을 수 없습니다.")
        return

    if gap is None:
        print("오류: E1(gap) 값이 없습니다.")
        return

    print("\nCSV 데이터:")
    print("  Gap (E1): {}mm".format(gap))
    print("  데이터 행 수: {}".format(len(data)))

    # Maxwell 프로젝트 및 디자인 가져오기
    oProject = oDesktop.GetActiveProject()
    if oProject is None:
        oProject = oDesktop.NewProject()
        print("새 프로젝트를 생성했습니다.")

    oDesign = oProject.GetActiveDesign()
    if oDesign is None:
        oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
        oDesign = oProject.GetActiveDesign()
        print("새 Maxwell 3D 디자인을 생성했습니다.")

    oEditor = oDesign.SetActiveEditor("3D Modeler")

    # Z 시작 위치 (일단 0)
    z_start = 0.0

    # 각 레이어마다 3개 직사각형 생성 (모두 원점 중심에 겹쳐짐)
    for i, row_data in enumerate(data):
        a = row_data['A']  # 메인 레그 X
        b = row_data['B']  # Y (공통)
        c = row_data['C']  # 사이드 레그 X

        print("\n레이어 {}: A={}, B={}, C={}".format(i+1, a, b, c))

        # ===== 1. 메인 레그 =====
        main_name = "{}_Layer{}_Main".format(name_prefix, i+1)
        create_rectangle(oEditor, 0, 0, z_start, a, b, main_name)
        # 중심을 원점에
        move_object(oEditor, main_name, -a/2.0, -b/2.0, 0)

        # ===== 2. 사이드 레그 1 (오른쪽) =====
        side1_name = "{}_Layer{}_Side1".format(name_prefix, i+1)
        create_rectangle(oEditor, 0, 0, z_start, c, b, side1_name)
        # 중심을 원점에, gap만큼 오른쪽 이동
        move_object(oEditor, side1_name, -c/2.0 + gap, -b/2.0, 0)

        # ===== 3. 사이드 레그 2 (왼쪽) =====
        side2_name = "{}_Layer{}_Side2".format(name_prefix, i+1)
        create_rectangle(oEditor, 0, 0, z_start, c, b, side2_name)
        # 중심을 원점에, gap만큼 왼쪽 이동 (음수)
        move_object(oEditor, side2_name, -c/2.0 - gap, -b/2.0, 0)

    # 뷰 맞추기
    oEditor.FitAll()

    print("\n" + "=" * 60)
    print("완료! 총 {} 개의 레이어가 생성되었습니다.".format(len(data)))
    print("=" * 60)


# 스크립트 실행
try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
except:
    import sys
    if len(sys.argv) > 0:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_dir = os.getcwd()

csv_file = os.path.join(script_dir, "transformer_core_sample.csv")

print("CSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

# 철심 생성
create_core_from_csv(
    csv_file_path=csv_file,
    name_prefix="Core"
)
