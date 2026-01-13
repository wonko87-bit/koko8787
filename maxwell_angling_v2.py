# -*- coding: utf-8 -*-
"""
Maxwell 3D - 앵글링 생성 V2
CSV 파일에서 앵글링(사분원 + 막대) 생성
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os
import math


def read_angling_csv(csv_file_path):
    """CSV 파일에서 앵글링 데이터를 읽어옵니다.

    CSV 형식:
    - A4~G4: 첫 번째 앵글링 데이터
    - A5~G5: 두 번째 앵글링 데이터
    - ...
    - 빈 행 또는 파일 끝까지

    변수 순서:
    A: 사분면 (1~4)
    B: 내측 반경
    C: 두께
    D: X축 막대 길이
    E: Z축 막대 길이
    F: 기준점 X좌표
    G: 기준점 Z좌표
    """
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

        # 7개 변수 읽기 (A~G = 인덱스 0~6)
        try:
            if len(row) < 7:
                print("경고: {}행에 데이터가 부족합니다. 건너뜁니다.".format(i+1))
                continue

            quadrant = int(row[0].strip()) if row[0].strip() else None
            inner_radius = float(row[1].strip()) if row[1].strip() else None
            thickness = float(row[2].strip()) if row[2].strip() else None
            x_bar_length = float(row[3].strip()) if row[3].strip() else None
            z_bar_length = float(row[4].strip()) if row[4].strip() else None
            ref_x = float(row[5].strip()) if row[5].strip() else None
            ref_z = float(row[6].strip()) if row[6].strip() else None

            # 데이터 유효성 검사
            if None in [quadrant, inner_radius, thickness, x_bar_length, z_bar_length, ref_x, ref_z]:
                print("경고: {}행에 빈 데이터가 있습니다. 건너뜁니다.".format(i+1))
                continue

            if quadrant not in [1, 2, 3, 4]:
                print("경고: {}행의 사분면 값({})이 올바르지 않습니다. 건너뜁니다.".format(i+1, quadrant))
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


def create_angling_shape(oEditor, angling_data, name):
    """앵글링 형태 생성 (사분원 + 양 끝 막대)

    방법:
    1. XZ 평면에 Polyline으로 외곽선 그리기
    2. Y축 방향으로 두께만큼 Sweep
    """

    quadrant = angling_data['quadrant']
    r_inner = angling_data['inner_radius']
    thickness = angling_data['thickness']
    r_outer = r_inner + thickness
    x_bar_len = angling_data['x_bar_length']
    z_bar_len = angling_data['z_bar_length']
    ref_x = angling_data['ref_x']
    ref_z = angling_data['ref_z']

    print("  사분면: {}, 내측반경: {}, 두께: {}".format(quadrant, r_inner, thickness))
    print("  X막대: {}, Z막대: {}, 기준점: ({}, 0, {})".format(x_bar_len, z_bar_len, ref_x, ref_z))

    # 사분면에 따른 방향 결정
    if quadrant == 1:
        # 1사분면: 기준점에서 -X, -Z 방향
        x_sign = -1
        z_sign = -1
        arc_center_x = ref_x - r_outer
        arc_center_z = ref_z - r_outer
        # 외측 사분원: 중심 기준 0도 → 90도
        arc_start_angle = 0
        arc_end_angle = 90
    elif quadrant == 2:
        # 2사분면: 기준점에서 -X, +Z 방향
        x_sign = -1
        z_sign = 1
        arc_center_x = ref_x - r_outer
        arc_center_z = ref_z + r_outer
        # 외측 사분원: 중심 기준 270도 → 360도
        arc_start_angle = 270
        arc_end_angle = 360
    elif quadrant == 3:
        # 3사분면: 기준점에서 +X, +Z 방향
        x_sign = 1
        z_sign = 1
        arc_center_x = ref_x + r_outer
        arc_center_z = ref_z + r_outer
        # 외측 사분원: 중심 기준 180도 → 270도
        arc_start_angle = 180
        arc_end_angle = 270
    else:  # quadrant == 4
        # 4사분면: 기준점에서 +X, -Z 방향
        x_sign = 1
        z_sign = -1
        arc_center_x = ref_x + r_outer
        arc_center_z = ref_z - r_outer
        # 외측 사분원: 중심 기준 90도 → 180도
        arc_start_angle = 90
        arc_end_angle = 180

    # Polyline으로 외곽선 그리기
    # 순서: 외측 사분원 → 내측 사분원 (역방향) → 막대들로 연결

    # 외측 사분원 점들 생성 (10개 점)
    num_arc_points = 10
    outer_arc_points = []
    for i in range(num_arc_points):
        angle_deg = arc_start_angle + (arc_end_angle - arc_start_angle) * i / (num_arc_points - 1)
        angle_rad = math.radians(angle_deg)
        x = arc_center_x + r_outer * math.cos(angle_rad)
        z = arc_center_z + r_outer * math.sin(angle_rad)
        outer_arc_points.append((x, 0, z))

    # 내측 사분원 점들 생성 (10개 점, 역방향)
    inner_arc_center_x = arc_center_x
    inner_arc_center_z = arc_center_z
    inner_arc_points = []
    for i in range(num_arc_points):
        angle_deg = arc_end_angle - (arc_end_angle - arc_start_angle) * i / (num_arc_points - 1)
        angle_rad = math.radians(angle_deg)
        x = inner_arc_center_x + r_inner * math.cos(angle_rad)
        z = inner_arc_center_z + r_inner * math.sin(angle_rad)
        inner_arc_points.append((x, 0, z))

    # 막대 부분 추가
    # 외측 사분원 끝점 2개 (인덱스 0, -1)
    outer_start = outer_arc_points[0]
    outer_end = outer_arc_points[-1]

    # 내측 사분원 끝점 2개 (인덱스 0, -1)
    inner_start = inner_arc_points[0]
    inner_end = inner_arc_points[-1]

    # Polyline 포인트 리스트 구성
    # 경로: 외측 시작 → 외측 끝 → (막대1) → 내측 끝 → 내측 시작 → (막대2) → 외측 시작
    polyline_points = []

    # 1. 외측 사분원 (시작 → 끝)
    polyline_points.extend(outer_arc_points)

    # 2. 외측 끝 → 막대 연장 → 내측 끝
    # 사분면에 따라 막대 방향 결정
    if quadrant == 1:
        # Z막대: 외측 끝(X방향)에서 -Z 방향
        bar_point = (outer_end[0], 0, outer_end[2] - z_bar_len)
        polyline_points.append(bar_point)
    elif quadrant == 2:
        # Z막대: 외측 끝(X방향)에서 +Z 방향
        bar_point = (outer_end[0], 0, outer_end[2] + z_bar_len)
        polyline_points.append(bar_point)
    elif quadrant == 3:
        # Z막대: 외측 끝(X방향)에서 +Z 방향
        bar_point = (outer_end[0], 0, outer_end[2] + z_bar_len)
        polyline_points.append(bar_point)
    else:  # quadrant == 4
        # Z막대: 외측 끝(X방향)에서 -Z 방향
        bar_point = (outer_end[0], 0, outer_end[2] - z_bar_len)
        polyline_points.append(bar_point)

    # 3. 내측 사분원 (끝 → 시작, 역방향)
    polyline_points.extend(inner_arc_points)

    # 4. 내측 시작 → 막대 연장 → 외측 시작
    if quadrant == 1:
        # X막대: 내측 시작(Z방향)에서 -X 방향
        bar_point = (inner_start[0] - x_bar_len, 0, inner_start[2])
        polyline_points.append(bar_point)
    elif quadrant == 2:
        # X막대: 내측 시작(Z방향)에서 -X 방향
        bar_point = (inner_start[0] - x_bar_len, 0, inner_start[2])
        polyline_points.append(bar_point)
    elif quadrant == 3:
        # X막대: 내측 시작(Z방향)에서 +X 방향
        bar_point = (inner_start[0] + x_bar_len, 0, inner_start[2])
        polyline_points.append(bar_point)
    else:  # quadrant == 4
        # X막대: 내측 시작(Z방향)에서 +X 방향
        bar_point = (inner_start[0] + x_bar_len, 0, inner_start[2])
        polyline_points.append(bar_point)

    # 5. 닫기 (외측 시작점으로 돌아감)
    polyline_points.append(outer_start)

    # Polyline 생성
    polyline_name = name + "_Profile"

    # PolylinePoints 생성
    point_list = []
    for i, (x, y, z) in enumerate(polyline_points):
        point_list.extend([
            "NAME:PLPoint",
            "X:=", "{}mm".format(x),
            "Y:=", "{}mm".format(y),
            "Z:=", "{}mm".format(z)
        ])

    # PolylineSegments 생성 (모두 직선)
    segment_list = []
    for i in range(len(polyline_points) - 1):
        segment_list.extend([
            "NAME:PLSegment",
            "SegmentType:=", "Line",
            "StartIndex:=", i,
            "NoOfPoints:=", 2
        ])

    oEditor.CreatePolyline(
        [
            "NAME:PolylineParameters",
            "IsPolylineCovered:=", True,
            "IsPolylineClosed:=", True,
            [
                "NAME:PolylinePoints"
            ] + point_list,
            [
                "NAME:PolylineSegments"
            ] + segment_list,
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
            "Name:=", polyline_name,
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
    print("  Polyline 생성: {}".format(polyline_name))

    # Y축 방향으로 두께만큼 Sweep
    oEditor.SweepAlongVector(
        [
            "NAME:Selections",
            "Selections:=", polyline_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:VectorSweepParameters",
            "DraftAngle:=", "0deg",
            "DraftType:=", "Round",
            "CheckFaceFaceIntersection:=", False,
            "SweepVectorX:=", "0mm",
            "SweepVectorY:=", "{}mm".format(thickness),
            "SweepVectorZ:=", "0mm"
        ]
    )
    print("  Sweep: Y축 방향 {}mm".format(thickness))

    # 이름 변경
    oEditor.ChangeProperty(
        [
            "NAME:AllTabs",
            [
                "NAME:Geometry3DAttributeTab",
                [
                    "NAME:PropServers",
                    polyline_name
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


def create_anglings_from_csv(csv_file_path, name_prefix="Angling"):
    """CSV 파일에서 앵글링들을 생성"""
    print("=" * 60)
    print("Maxwell 3D - 앵글링 생성 V2")
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
            create_angling_shape(oEditor, angling, angling_name)
            success_count += 1
        except Exception as e:
            print("  [오류] 앵글링 {} 생성 실패: {}".format(idx + 1, str(e)))
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
