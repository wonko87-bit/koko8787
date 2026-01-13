# -*- coding: utf-8 -*-
"""
Maxwell 3D - 앵글링 생성 V4 (Boolean 연산 방식)
Circle + Rectangle + Boolean 연산으로 2D Sheet 생성
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_angling_csv(csv_file_path):
    """CSV 파일에서 앵글링 데이터를 읽어옵니다."""
    anglings = []

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

    # 4행부터 시작 (인덱스 3)
    for i in range(3, len(rows)):
        row = rows[i]

        # 빈 행 체크
        if not row or all(not cell.strip() for cell in row):
            break

        try:
            if len(row) < 7:
                print("경고: {}행에 데이터가 부족합니다.".format(i+1))
                continue

            quadrant = int(row[0].strip()) if row[0].strip() else None
            inner_radius = float(row[1].strip()) if row[1].strip() else None
            thickness = float(row[2].strip()) if row[2].strip() else None
            x_bar_length = float(row[3].strip()) if row[3].strip() else None
            z_bar_length = float(row[4].strip()) if row[4].strip() else None
            ref_x = float(row[5].strip()) if row[5].strip() else None
            ref_z = float(row[6].strip()) if row[6].strip() else None

            if None in [quadrant, inner_radius, thickness, x_bar_length, z_bar_length, ref_x, ref_z]:
                print("경고: {}행에 빈 데이터가 있습니다.".format(i+1))
                continue

            if quadrant not in [1, 2, 3, 4]:
                print("경고: {}행의 사분면 값({})이 올바르지 않습니다.".format(i+1, quadrant))
                continue

            anglings.append({
                'quadrant': quadrant,
                'inner_radius': inner_radius,
                'thickness': thickness,
                'x_bar_length': x_bar_length,
                'z_bar_length': z_bar_length,
                'ref_x': ref_x,
                'ref_z': ref_z,
                'row_num': i + 1
            })

        except (ValueError, IndexError) as e:
            print("경고: {}행 데이터 읽기 오류: {}".format(i+1, str(e)))
            continue

    return anglings


def create_circle_xz(oEditor, center_x, center_z, radius, name):
    """XZ 평면에 Circle 생성 (Y축에 수직)"""
    oEditor.CreateCircle(
        [
            "NAME:CircleParameters",
            "IsCovered:=", True,
            "XCenter:=", "{}mm".format(center_x),
            "YCenter:=", "0mm",
            "ZCenter:=", "{}mm".format(center_z),
            "Radius:=", "{}mm".format(radius),
            "WhichAxis:=", "Y",  # XZ 평면
            "NumSegments:=", "0"
        ],
        [
            "NAME:Attributes",
            "Name:=", name,
            "Flags:=", "",
            "Color:=", "(143 175 143)",
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


def create_rectangle_xz(oEditor, x_start, z_start, width, height, name):
    """XZ 평면에 Rectangle 생성"""
    oEditor.CreateRectangle(
        [
            "NAME:RectangleParameters",
            "IsCovered:=", True,
            "XStart:=", "{}mm".format(x_start),
            "YStart:=", "0mm",
            "ZStart:=", "{}mm".format(z_start),
            "Width:=", "{}mm".format(width),
            "Height:=", "{}mm".format(height),
            "WhichAxis:=", "Y"  # XZ 평면
        ],
        [
            "NAME:Attributes",
            "Name:=", name,
            "Flags:=", "",
            "Color:=", "(143 175 143)",
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


def subtract_objects(oEditor, blank_name, tool_name):
    """Boolean Subtract: blank - tool"""
    oEditor.Subtract(
        [
            "NAME:Selections",
            "Blank Parts:=", blank_name,
            "Tool Parts:=", tool_name
        ],
        [
            "NAME:SubtractParameters",
            "KeepOriginals:=", False
        ]
    )


def intersect_objects(oEditor, parts_list, new_name):
    """Boolean Intersect"""
    oEditor.Intersect(
        [
            "NAME:Selections",
            "Selections:=", ",".join(parts_list)
        ],
        [
            "NAME:IntersectParameters",
            "KeepOriginals:=", False
        ]
    )


def unite_objects(oEditor, parts_list):
    """Boolean Unite"""
    oEditor.Unite(
        [
            "NAME:Selections",
            "Selections:=", ",".join(parts_list)
        ],
        [
            "NAME:UniteParameters",
            "KeepOriginals:=", False
        ]
    )


def create_angling_boolean(oEditor, angling_data, name):
    """Boolean 연산으로 앵글링 생성

    방법:
    1. 외측 원 생성
    2. 내측 원 생성
    3. Subtract (외측 - 내측) = 도넛
    4. 사분면 선택을 위한 Rectangle 2개 생성 (큰 크기)
    5. Intersect로 사분면만 남기기
    6. X축, Z축 막대 Rectangle 2개 생성
    7. Unite로 합치기
    """

    quadrant = angling_data['quadrant']
    r_inner = angling_data['inner_radius']
    thickness = angling_data['thickness']
    r_outer = r_inner + thickness
    x_bar_len = angling_data['x_bar_length']
    z_bar_len = angling_data['z_bar_length']
    ref_x = angling_data['ref_x']
    ref_z = angling_data['ref_z']

    print("  사분면: {}, 내측R: {}, 외측R: {}".format(quadrant, r_inner, r_outer))
    print("  X막대: {}, Z막대: {}, 기준점: ({}, 0, {})".format(x_bar_len, z_bar_len, ref_x, ref_z))

    # 사분면별 원 중심 계산
    if quadrant == 1:
        center_x = ref_x - r_outer
        center_z = ref_z - r_outer
    elif quadrant == 2:
        center_x = ref_x - r_outer
        center_z = ref_z + r_outer
    elif quadrant == 3:
        center_x = ref_x + r_outer
        center_z = ref_z + r_outer
    else:  # quadrant == 4
        center_x = ref_x + r_outer
        center_z = ref_z - r_outer

    print("  원 중심: ({}, 0, {})".format(center_x, center_z))

    # 1. 외측 원 생성
    outer_circle_name = name + "_Outer"
    create_circle_xz(oEditor, center_x, center_z, r_outer, outer_circle_name)
    print("  [1] 외측 원 생성: {}".format(outer_circle_name))

    # 2. 내측 원 생성
    inner_circle_name = name + "_Inner"
    create_circle_xz(oEditor, center_x, center_z, r_inner, inner_circle_name)
    print("  [2] 내측 원 생성: {}".format(inner_circle_name))

    # 3. Subtract (도넛 생성)
    subtract_objects(oEditor, outer_circle_name, inner_circle_name)
    donut_name = outer_circle_name  # Subtract 후 blank 이름 유지
    print("  [3] Subtract: 도넛 생성")

    # 4. 사분면 선택용 Rectangle 2개 생성
    # 충분히 큰 크기 (원 반지름의 2배)
    big_size = r_outer * 2

    if quadrant == 1:
        # 1사분면: +X, +Z 방향
        rect1_x = center_x
        rect1_z = center_z
        rect1_w = big_size
        rect1_h = big_size
    elif quadrant == 2:
        # 2사분면: +X, -Z 방향
        rect1_x = center_x
        rect1_z = center_z - big_size
        rect1_w = big_size
        rect1_h = big_size
    elif quadrant == 3:
        # 3사분면: -X, -Z 방향
        rect1_x = center_x - big_size
        rect1_z = center_z - big_size
        rect1_w = big_size
        rect1_h = big_size
    else:  # quadrant == 4
        # 4사분면: -X, +Z 방향
        rect1_x = center_x - big_size
        rect1_z = center_z
        rect1_w = big_size
        rect1_h = big_size

    quarter_rect_name = name + "_Quarter"
    create_rectangle_xz(oEditor, rect1_x, rect1_z, rect1_w, rect1_h, quarter_rect_name)
    print("  [4] 사분면 선택 Rectangle 생성")

    # 5. Intersect (도넛과 사분면 Rectangle)
    intersect_objects(oEditor, [donut_name, quarter_rect_name], name)
    quarter_arc_name = donut_name  # Intersect 후 이름 유지
    print("  [5] Intersect: 사분원 생성")

    # 6. X축, Z축 막대 Rectangle 생성
    # 각 사분면별로 원호 끝에 정확히 연결되도록 계산
    # 막대는 원호의 내측과 외측 사이(두께 구간)에 연결됨

    # X축 막대 (X방향으로 뻗는 수평 막대)
    if quadrant == 1:
        # ┘ 형태: 위쪽(90도)에서 -X 방향으로 연장
        # 원호 끝: center_x 위치, Z는 center_z + r_inner ~ center_z + r_outer
        x_bar_x = center_x - x_bar_len
        x_bar_z = center_z + r_inner
        x_bar_w = thickness  # Z방향 크기
        x_bar_h = x_bar_len  # X방향 크기
    elif quadrant == 2:
        # └ 형태: 아래쪽(270도)에서 -X 방향으로 연장
        # 원호 끝: center_x 위치, Z는 center_z - r_outer ~ center_z - r_inner
        x_bar_x = center_x - x_bar_len
        x_bar_z = center_z - r_outer
        x_bar_w = thickness
        x_bar_h = x_bar_len
    elif quadrant == 3:
        # ┌ 형태: 아래쪽(270도)에서 +X 방향으로 연장
        # 원호 끝: center_x 위치, Z는 center_z - r_outer ~ center_z - r_inner
        x_bar_x = center_x
        x_bar_z = center_z - r_outer
        x_bar_w = thickness
        x_bar_h = x_bar_len
    else:  # quadrant == 4
        # ┐ 형태: 위쪽(90도)에서 +X 방향으로 연장
        # 원호 끝: center_x 위치, Z는 center_z + r_inner ~ center_z + r_outer
        x_bar_x = center_x
        x_bar_z = center_z + r_inner
        x_bar_w = thickness
        x_bar_h = x_bar_len

    x_bar_name = name + "_XBar"
    create_rectangle_xz(oEditor, x_bar_x, x_bar_z, x_bar_w, x_bar_h, x_bar_name)
    print("  [6] X축 막대 생성: XStart={}, ZStart={}, W={}, H={}".format(x_bar_x, x_bar_z, x_bar_w, x_bar_h))

    # Z축 막대 (Z방향으로 뻗는 수직 막대)
    if quadrant == 1:
        # ┘ 형태: 오른쪽(0도)에서 -Z 방향으로 연장
        # 원호 끝: X는 center_x + r_inner ~ center_x + r_outer, center_z 위치
        z_bar_x = center_x + r_inner
        z_bar_z = center_z - z_bar_len
        z_bar_w = z_bar_len  # Z방향 크기
        z_bar_h = thickness  # X방향 크기
    elif quadrant == 2:
        # └ 형태: 오른쪽(0도)에서 +Z 방향으로 연장
        # 원호 끝: X는 center_x + r_inner ~ center_x + r_outer, center_z 위치
        z_bar_x = center_x + r_inner
        z_bar_z = center_z
        z_bar_w = z_bar_len
        z_bar_h = thickness
    elif quadrant == 3:
        # ┌ 형태: 왼쪽(180도)에서 +Z 방향으로 연장
        # 원호 끝: X는 center_x - r_outer ~ center_x - r_inner, center_z 위치
        z_bar_x = center_x - r_outer
        z_bar_z = center_z
        z_bar_w = z_bar_len
        z_bar_h = thickness
    else:  # quadrant == 4
        # ┐ 형태: 왼쪽(180도)에서 -Z 방향으로 연장
        # 원호 끝: X는 center_x - r_outer ~ center_x - r_inner, center_z 위치
        z_bar_x = center_x - r_outer
        z_bar_z = center_z - z_bar_len
        z_bar_w = z_bar_len
        z_bar_h = thickness

    z_bar_name = name + "_ZBar"
    create_rectangle_xz(oEditor, z_bar_x, z_bar_z, z_bar_w, z_bar_h, z_bar_name)
    print("  [7] Z축 막대 생성")

    # 7. Unite (사분원 + X막대 + Z막대)
    unite_objects(oEditor, [quarter_arc_name, x_bar_name, z_bar_name])
    print("  [8] Unite: 최종 완성")

    # 최종 이름 변경
    try:
        oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Geometry3DAttributeTab",
                    [
                        "NAME:PropServers",
                        quarter_arc_name
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
        print("  완성: {}".format(name))
    except:
        print("  완성: {} (이름 변경 실패, 기존 이름 유지)".format(quarter_arc_name))


def create_anglings_from_csv(csv_file_path, name_prefix="Angling"):
    """CSV 파일에서 앵글링들을 생성"""
    print("=" * 60)
    print("Maxwell 3D - 앵글링 생성 V4 (Boolean 연산)")
    print("=" * 60)

    print("\nCSV 파일 읽기 시작...")
    anglings = read_angling_csv(csv_file_path)

    print("읽은 앵글링 수: {}".format(len(anglings)))

    if not anglings:
        print("오류: CSV 파일에서 앵글링 데이터를 읽을 수 없습니다.")
        return

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

    # 각 앵글링 생성
    success_count = 0
    fail_count = 0

    for idx, angling in enumerate(anglings):
        print("\n--- 앵글링 {} (CSV {}행) ---".format(idx + 1, angling['row_num']))

        angling_name = "{}_{}".format(name_prefix, idx + 1)

        try:
            create_angling_boolean(oEditor, angling, angling_name)
            success_count += 1
        except Exception as e:
            print("  [오류] 앵글링 {} 생성 실패: {}".format(idx + 1, str(e)))
            import traceback
            traceback.print_exc()
            print("  앵글링 {}를 건너뛰고 계속 진행합니다.".format(idx + 1))
            fail_count += 1
            continue

    # 뷰 맞추기
    try:
        oEditor.FitAll()
    except:
        pass

    print("\n" + "=" * 60)
    print("완료!")
    print("  성공: {} 개".format(success_count))
    print("  실패: {} 개".format(fail_count))
    print("  전체: {} 개".format(len(anglings)))
    print("=" * 60)


# 스크립트 실행
print("\n" + "=" * 60)
print("스크립트 실행 시작")
print("=" * 60)

try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print("스크립트 디렉토리 (__file__): {}".format(script_dir))
except:
    import sys
    if len(sys.argv) > 0:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        print("스크립트 디렉토리 (sys.argv[0]): {}".format(script_dir))
    else:
        script_dir = os.getcwd()
        print("스크립트 디렉토리 (getcwd): {}".format(script_dir))

csv_file = os.path.join(script_dir, "AnglingDim.csv")

print("\nCSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    print("\n경고: AnglingDim.csv 파일을 찾을 수 없습니다!")
    print("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    print("\n앵글링 생성 함수 호출 중...")
    create_anglings_from_csv(
        csv_file_path=csv_file,
        name_prefix="Angling"
    )
    print("\n스크립트 실행 완료!")
