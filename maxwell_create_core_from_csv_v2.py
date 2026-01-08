# -*- coding: utf-8 -*-
"""
Maxwell 3D - CSV 기반 변압기 철심 생성 스크립트 V2 (Rectangle + Sweep 방식)
============================================================================

Rectangle을 먼저 그리고 Move, Rotate 후 Unite하여 Sweep하는 고속 방식

CSV 구조:
- A2:A(x): X1 값 (Main leg의 X 크기)
- B2:B(x): X2 값 (Return leg의 X 크기)
- C2:C(x): Y 값 (Y축 크기)
- E1: Return leg 이격거리 (Main leg 중점 ↔ Return leg 중점)
- E2: 철심 창 높이 (Z축 extrusion 높이)

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
    """
    data = []
    gap = None
    window_height = None

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

        # E1 (row 0, col 4): gap distance
        if len(rows) > 0 and len(rows[0]) > 4:
            gap = float(rows[0][4])

        # E2 (row 1, col 4): window height
        if len(rows) > 1 and len(rows[1]) > 4:
            window_height = float(rows[1][4])

        # 데이터 행 읽기 (A, B, C 열)
        for i, row in enumerate(rows):
            if len(row) >= 3:
                try:
                    x1 = float(row[0])
                    x2 = float(row[1])
                    y = float(row[2])
                    data.append({'X1': x1, 'X2': x2, 'Y': y})
                except ValueError:
                    continue

    return data, gap, window_height


def create_rectangle(oEditor, x_start, y_start, z_start, width, height, name):
    """
    XY 평면에 Rectangle 생성 (Z축 기준)
    """
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
    print("  Created rectangle: {}".format(name))


def move_object(oEditor, obj_name, dx, dy, dz):
    """
    객체를 이동
    """
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
    print("  Moved {} by ({}, {}, {})".format(obj_name, dx, dy, dz))


def rotate_object(oEditor, obj_name, axis, angle_deg):
    """
    객체를 회전 (원점 기준)
    """
    oEditor.Rotate(
        [
            "NAME:Selections",
            "Selections:=", obj_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:RotateParameters",
            "RotateAxis:=", axis,
            "RotateAngle:=", "{}deg".format(angle_deg)
        ]
    )
    print("  Rotated {} by {} degrees around {} axis".format(obj_name, angle_deg, axis))


def sweep_along_z(oEditor, obj_name, sweep_distance):
    """
    Z축 방향으로 Sweep
    """
    oEditor.SweepAlongVector(
        [
            "NAME:Selections",
            "Selections:=", obj_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:VectorSweepParameters",
            "DraftAngle:=", "0deg",
            "DraftType:=", "Round",
            "CheckFaceFaceIntersection:=", False,
            "SweepVectorX:=", "0mm",
            "SweepVectorY:=", "0mm",
            "SweepVectorZ:=", "{}mm".format(sweep_distance)
        ]
    )
    print("  Swept {} along Z axis by {}mm".format(obj_name, sweep_distance))


def create_transformer_core_from_csv(csv_file_path, material="steel_1008", name_prefix="Core"):
    """
    CSV 파일에서 변압기 철심을 생성합니다 (Rectangle + Sweep 방식)

    각 레이어마다:
    1. Main leg (중앙) - Rectangle
    2. Side legs (양쪽 2개) - Rectangle + 90도 회전
    3. Yokes (상단/하단 2개) - Rectangle
    4. 5개 Rectangle을 Unite
    5. Sweep하여 3D화
    """
    print("=" * 60)
    print("Maxwell 3D - 변압기 철심 생성 (Rectangle+Sweep 방식)")
    print("=" * 60)

    # CSV 데이터 읽기
    data, gap, window_height = read_csv_data(csv_file_path)

    if not data:
        print("오류: CSV 파일에서 데이터를 읽을 수 없습니다.")
        return

    if gap is None or window_height is None:
        print("오류: E1(gap) 또는 E2(window_height) 값이 없습니다.")
        return

    print("\nCSV 데이터:")
    print("  Return leg 이격거리 (E1): {}mm".format(gap))
    print("  철심 창 높이 (E2): {}mm".format(window_height))
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

    created_cores = []

    # 화면 업데이트 일시 중지 (성능 향상)
    oEditor.SuspendUpdate()
    print("\n화면 업데이트를 일시 중지했습니다. (성능 향상 모드)")

    # Z 시작 위치 계산 (원점 중심)
    z_start = -window_height / 2.0

    # Y 누적 위치 추적 (적층용)
    cumulative_y = 0.0
    prev_y = 0.0

    # 각 데이터 행마다 완전한 철심 생성
    for i, row_data in enumerate(data):
        x1 = row_data['X1']
        x2 = row_data['X2']
        y = row_data['Y']

        print("\n레이어 {}: X1={}, X2={}, Y={}".format(i+1, x1, x2, y))

        # 적층 Y 좌표 계산 (이전 레이어의 Y/2 + 현재 레이어의 Y/2)
        if i == 0:
            cumulative_y = 0.0  # 첫 번째 레이어는 원점
        else:
            cumulative_y += prev_y / 2.0 + y / 2.0

        print("  적층 Y 위치: {}mm".format(cumulative_y))

        layer_rects = []

        # ===== 1. Main leg (중앙) - XY 평면에 Rectangle =====
        main_name = "{}_Layer{}_Main".format(name_prefix, i+1)
        # 원점을 중심으로 하는 rectangle (0,0)에서 시작하여 x1 x y 크기
        create_rectangle(oEditor, 0, 0, z_start, x1, y, main_name)
        # 중심으로 이동: (-x1/2, cumulative_y - y/2, 0)
        move_object(oEditor, main_name, -x1/2.0, cumulative_y - y/2.0, 0)
        layer_rects.append(main_name)

        # ===== 2. Side legs (양쪽 2개) - X,Y 좌표 교환으로 자동 회전 =====
        # Left side leg
        left_side_name = "{}_Layer{}_LeftSide".format(name_prefix, i+1)
        # Y x X2 rectangle 생성 (X, Y 좌표 바꿈)
        create_rectangle(oEditor, 0, 0, z_start, y, x2, left_side_name)
        # 중심 정렬 후 왼쪽으로 이동: x = -gap, y = cumulative_y - y/2
        move_object(oEditor, left_side_name, -gap - y/2.0, cumulative_y - x2/2.0, 0)
        layer_rects.append(left_side_name)

        # Right side leg
        right_side_name = "{}_Layer{}_RightSide".format(name_prefix, i+1)
        # Y x X2 rectangle 생성 (X, Y 좌표 바꿈)
        create_rectangle(oEditor, 0, 0, z_start, y, x2, right_side_name)
        # 중심 정렬 후 오른쪽으로 이동: x = +gap, y = cumulative_y - y/2
        move_object(oEditor, right_side_name, gap - y/2.0, cumulative_y - x2/2.0, 0)
        layer_rects.append(right_side_name)

        # 다음 레이어를 위해 현재 Y 저장
        prev_y = y

        # ===== 3. Yokes (상단/하단 2개) =====
        yoke_x_size = 2 * gap + x2
        yoke_y_size = x2

        # Top yoke
        top_yoke_name = "{}_Layer{}_TopYoke".format(name_prefix, i+1)
        create_rectangle(oEditor, 0, 0, z_start, yoke_x_size, yoke_y_size, top_yoke_name)
        # 상단으로 이동 (적층 위치 고려)
        move_object(oEditor, top_yoke_name, -yoke_x_size/2.0, cumulative_y + y/2.0, 0)
        layer_rects.append(top_yoke_name)

        # Bottom yoke
        bottom_yoke_name = "{}_Layer{}_BottomYoke".format(name_prefix, i+1)
        create_rectangle(oEditor, 0, 0, z_start, yoke_x_size, yoke_y_size, bottom_yoke_name)
        # 하단으로 이동 (적층 위치 고려)
        move_object(oEditor, bottom_yoke_name, -yoke_x_size/2.0, cumulative_y - y/2.0 - yoke_y_size, 0)
        layer_rects.append(bottom_yoke_name)

        print("  Created 5 rectangles for layer {}".format(i+1))

        # ===== 4. Unite - 5개 Rectangle 병합 =====
        core_name = "{}_Layer{}".format(name_prefix, i+1)

        oEditor.Unite(
            [
                "NAME:Selections",
                "Selections:=", ",".join(layer_rects)
            ],
            [
                "NAME:UniteParameters",
                "KeepOriginals:=", False
            ]
        )
        print("  United 5 rectangles")

        # Unite 후 첫 번째 이름으로 통합됨, core_name으로 변경
        oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Geometry3DAttributeTab",
                    [
                        "NAME:PropServers",
                        layer_rects[0]
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Name",
                            "Value:=", core_name
                        ]
                    ]
                ]
            ]
        )

        # ===== 5. Sweep - Z축으로 확장하여 3D화 =====
        sweep_along_z(oEditor, core_name, window_height)

        # 재질 변경
        oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Geometry3DAttributeTab",
                    [
                        "NAME:PropServers",
                        core_name
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Material",
                            "Value:=", "\"{}\"".format(material)
                        ]
                    ]
                ]
            ]
        )

        created_cores.append(core_name)
        print("  완료! 통합된 철심: {}".format(core_name))

    # 화면 업데이트 재개
    oEditor.ResumeUpdate()
    print("\n화면 업데이트를 재개했습니다.")

    # 뷰 맞추기
    oEditor.FitAll()

    print("\n" + "=" * 60)
    print("완료! 총 {} 개의 철심 레이어가 생성되었습니다.".format(len(created_cores)))
    print("=" * 60)


# 스크립트 실행
if __name__ == "__main__":
    # CSV 파일 경로 (스크립트와 동일한 폴더)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    csv_file = os.path.join(script_dir, "transformer_core_sample.csv")

    # 철심 생성
    create_transformer_core_from_csv(
        csv_file_path=csv_file,
        material="steel_1008",
        name_prefix="Core"
    )
