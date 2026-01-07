# -*- coding: utf-8 -*-
"""
Maxwell 3D - CSV 기반 변압기 철심 생성 스크립트 (Yoke 포함)
==========================================================

CSV 파일에서 철심 치수 데이터를 읽어 완전한 1x2 변압기 철심을 생성합니다.

CSV 구조:
- A2:A(x): X1 값 (Main leg의 X 크기)
- B2:B(x): X2 값 (Return leg의 X 크기)
- C2:C(x): Y 값 (Y축 크기)
- E1: Return leg 이격거리 (Main leg 중점 ↔ Return leg 중점)
- E2: 철심 창 높이 (Z축 extrusion 높이)

생성되는 구조:
- Main leg: 1개 (중앙)
- Return legs: 2개 (양쪽)
- Top yoke: 3개 leg를 상단에서 연결
- Bottom yoke: 3개 leg를 하단에서 연결
- 모든 박스는 원점(0,0,0)을 중심으로 배치

Author: Claude
Date: 2026-01-07
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_csv_data(csv_file_path):
    """
    CSV 파일에서 철심 데이터를 읽어옵니다.

    Parameters:
    -----------
    csv_file_path : str
        CSV 파일 경로

    Returns:
    --------
    data : list of dict
        각 행의 데이터 [{X1, X2, Y}, ...]
    gap : float
        Return leg 이격거리 (E1 셀 값)
    window_height : float
        철심 창 높이 (E2 셀 값)
    """

    if not os.path.exists(csv_file_path):
        raise IOError("CSV 파일을 찾을 수 없습니다: {}".format(csv_file_path))

    data = []
    gap = 0
    window_height = 0

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

        # E1, E2 값 읽기 (E열은 인덱스 4)
        if len(rows) > 0 and len(rows[0]) > 4:
            try:
                gap = float(rows[0][4])  # E1
            except:
                pass

        if len(rows) > 1 and len(rows[1]) > 4:
            try:
                window_height = float(rows[1][4])  # E2
            except:
                pass

        # A2:C(x) 데이터 읽기 (헤더 제외하고 2번째 행부터)
        for i in range(1, len(rows)):  # 인덱스 1부터 (A2부터)
            row = rows[i]
            if len(row) >= 3:
                try:
                    x1 = float(row[0])  # A열
                    x2 = float(row[1])  # B열
                    y = float(row[2])   # C열

                    if x1 > 0 and x2 > 0 and y > 0:
                        data.append({'X1': x1, 'X2': x2, 'Y': y})
                except ValueError:
                    # 숫자가 아닌 경우 스킵
                    continue

    return data, gap, window_height


def create_transformer_core_from_csv(csv_file_path, material="steel_1008", name_prefix="Core"):
    """
    CSV 파일로부터 변압기 철심을 생성합니다 (Legs + Yokes).

    Parameters:
    -----------
    csv_file_path : str
        CSV 파일 경로
    material : str
        재질 이름
    name_prefix : str
        박스 이름 접두사

    Returns:
    --------
    created_cores : list
        생성된 철심 이름 목록
    """

    print("=" * 60)
    print("CSV 기반 변압기 철심 생성 (Legs + Yokes)")
    print("=" * 60)

    # CSV 데이터 읽기
    print("\nCSV 파일 읽기 중: {}".format(csv_file_path))
    data, gap, window_height = read_csv_data(csv_file_path)

    if not data:
        raise ValueError("CSV 파일에 유효한 데이터가 없습니다.")

    print("데이터 행 수: {}".format(len(data)))
    print("Return leg 이격거리 (E1): {} mm".format(gap))
    print("철심 창 높이 (E2): {} mm".format(window_height))
    print("-" * 60)

    # 프로젝트 및 디자인 설정
    oProject = oDesktop.GetActiveProject()
    if oProject is None:
        oProject = oDesktop.NewProject()
        print("새 프로젝트를 생성했습니다.")

    oDesign = oProject.GetActiveDesign()
    if oDesign is None:
        oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
        oDesign = oProject.SetActiveDesign("Maxwell3DDesign1")
        print("새 Maxwell 3D 디자인을 생성했습니다.")

    oEditor = oDesign.SetActiveEditor("3D Modeler")

    created_cores = []

    # 각 데이터 행마다 완전한 철심 생성 (3 legs + 2 yokes)
    for i, row_data in enumerate(data):
        x1 = row_data['X1']
        x2 = row_data['X2']
        y = row_data['Y']

        print("\n레이어 {}: X1={}, X2={}, Y={}".format(i+1, x1, x2, y))

        layer_boxes = []

        # ===== 1. Main leg (중앙) =====
        main_name = "{}_Layer{}_Main".format(name_prefix, i+1)
        main_x_pos = -x1 / 2.0
        main_y_pos = -y / 2.0
        main_z_pos = -window_height / 2.0

        oEditor.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:=", "{}mm".format(main_x_pos),
                "YPosition:=", "{}mm".format(main_y_pos),
                "ZPosition:=", "{}mm".format(main_z_pos),
                "XSize:=", "{}mm".format(x1),
                "YSize:=", "{}mm".format(y),
                "ZSize:=", "{}mm".format(window_height)
            ],
            [
                "NAME:Attributes",
                "Name:=", main_name,
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
        layer_boxes.append(main_name)
        print("  Main leg 생성: {}".format(main_name))

        # ===== 2. Left Return leg =====
        left_name = "{}_Layer{}_LeftReturn".format(name_prefix, i+1)
        left_x_pos = -gap - x2 / 2.0
        left_y_pos = -y / 2.0
        left_z_pos = -window_height / 2.0

        oEditor.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:=", "{}mm".format(left_x_pos),
                "YPosition:=", "{}mm".format(left_y_pos),
                "ZPosition:=", "{}mm".format(left_z_pos),
                "XSize:=", "{}mm".format(x2),
                "YSize:=", "{}mm".format(y),
                "ZSize:=", "{}mm".format(window_height)
            ],
            [
                "NAME:Attributes",
                "Name:=", left_name,
                "Flags:=", "",
                "Color:=", "(132 132 193)",
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
        layer_boxes.append(left_name)
        print("  Left Return leg 생성: {}".format(left_name))

        # ===== 3. Right Return leg =====
        right_name = "{}_Layer{}_RightReturn".format(name_prefix, i+1)
        right_x_pos = gap - x2 / 2.0
        right_y_pos = -y / 2.0
        right_z_pos = -window_height / 2.0

        oEditor.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:=", "{}mm".format(right_x_pos),
                "YPosition:=", "{}mm".format(right_y_pos),
                "ZPosition:=", "{}mm".format(right_z_pos),
                "XSize:=", "{}mm".format(x2),
                "YSize:=", "{}mm".format(y),
                "ZSize:=", "{}mm".format(window_height)
            ],
            [
                "NAME:Attributes",
                "Name:=", right_name,
                "Flags:=", "",
                "Color:=", "(132 132 193)",
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
        layer_boxes.append(right_name)
        print("  Right Return leg 생성: {}".format(right_name))

        # ===== 4. Top Yoke =====
        # Yoke 치수 계산
        # Left Return 외곽(-gap - x2/2) ~ Right Return 외곽(gap + x2/2)
        yoke_x_size = 2 * gap + x2  # 정확한 길이 (튀어나오지 않음)
        yoke_y_size = x2  # Leg를 눕힌 것
        yoke_z_size = y   # Leg의 Y 크기

        top_yoke_name = "{}_Layer{}_TopYoke".format(name_prefix, i+1)
        top_yoke_x_pos = -(yoke_x_size) / 2.0
        top_yoke_y_pos = -yoke_y_size / 2.0
        top_yoke_z_pos = window_height / 2.0  # Leg 윗면에서 시작

        oEditor.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:=", "{}mm".format(top_yoke_x_pos),
                "YPosition:=", "{}mm".format(top_yoke_y_pos),
                "ZPosition:=", "{}mm".format(top_yoke_z_pos),
                "XSize:=", "{}mm".format(yoke_x_size),
                "YSize:=", "{}mm".format(yoke_y_size),
                "ZSize:=", "{}mm".format(yoke_z_size)
            ],
            [
                "NAME:Attributes",
                "Name:=", top_yoke_name,
                "Flags:=", "",
                "Color:=", "(193 132 132)",
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
        layer_boxes.append(top_yoke_name)
        print("  Top yoke 생성: {}".format(top_yoke_name))

        # ===== 5. Bottom Yoke =====
        bottom_yoke_name = "{}_Layer{}_BottomYoke".format(name_prefix, i+1)
        bottom_yoke_x_pos = -(yoke_x_size) / 2.0
        bottom_yoke_y_pos = -yoke_y_size / 2.0
        bottom_yoke_z_pos = -window_height / 2.0 - yoke_z_size  # Leg 아래면에서 시작

        oEditor.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:=", "{}mm".format(bottom_yoke_x_pos),
                "YPosition:=", "{}mm".format(bottom_yoke_y_pos),
                "ZPosition:=", "{}mm".format(bottom_yoke_z_pos),
                "XSize:=", "{}mm".format(yoke_x_size),
                "YSize:=", "{}mm".format(yoke_y_size),
                "ZSize:=", "{}mm".format(yoke_z_size)
            ],
            [
                "NAME:Attributes",
                "Name:=", bottom_yoke_name,
                "Flags:=", "",
                "Color:=", "(193 132 132)",
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
        layer_boxes.append(bottom_yoke_name)
        print("  Bottom yoke 생성: {}".format(bottom_yoke_name))

        # ===== 6. Unite all parts into one core =====
        core_name = "{}_Layer{}".format(name_prefix, i+1)
        print("  Unite 연산 중... (5개 박스 -> 1개 철심)")

        # Unite: 첫 번째 박스가 기준, 나머지를 합침
        oEditor.Unite(
            [
                "NAME:Selections",
                "Selections:=", ",".join(layer_boxes)
            ],
            [
                "NAME:UniteParameters",
                "KeepOriginals:=", False
            ])

        # Unite 후 첫 번째 이름으로 남으므로, 원하는 이름으로 변경
        oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Geometry3DAttributeTab",
                    [
                        "NAME:PropServers",
                        layer_boxes[0]  # Unite 후 남은 첫 번째 객체
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Name",
                            "Value:=", core_name
                        ]
                    ]
                ]
            ])

        created_cores.append(core_name)
        print("  완료! 통합된 철심: {}".format(core_name))

    # 뷰 맞추기
    oEditor.FitAll()

    print("\n" + "=" * 60)
    print("완료! 총 {} 개의 철심 레이어가 생성되었습니다.".format(len(created_cores)))
    print("=" * 60)

    return created_cores


# ============================================================================
# 메인 실행 부분
# ============================================================================

if __name__ == "__main__":
    # CSV 파일 경로 설정 (여기를 수정하세요!)
    # 기본값: 스크립트와 같은 폴더의 transformer_core_sample.csv
    import os
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_file = os.path.join(script_dir, "transformer_core_sample.csv")

    # 또는 절대 경로 사용:
    # csv_file = r"C:\your\path\to\transformer_core_data.csv"

    # 재질 설정
    core_material = "steel_1008"  # Maxwell의 재질 라이브러리에 있는 재질 이름

    # 변압기 철심 생성
    try:
        cores = create_transformer_core_from_csv(csv_file, core_material, "Core")
        print("\n생성된 철심 목록:")
        for i, core_name in enumerate(cores, 1):
            print("  {}. {}".format(i, core_name))
    except Exception as e:
        print("\n오류 발생: {}".format(str(e)))
        import traceback
        traceback.print_exc()
