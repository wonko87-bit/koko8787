# -*- coding: utf-8 -*-
"""
Maxwell 3D - 앵글링 생성 V3 (2D Sheet with Arc)
CSV 파일에서 앵글링(사분원 Arc + 직사각형 막대) 생성
XZ 평면에 2D Sheet로 생성
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


def create_angling_sheet(oEditor, angling_data, name):
    """앵글링 형태 생성 (2D Sheet: 사분원 Arc + 직사각형 막대)

    방법:
    1. XZ 평면에 Polyline으로 외곽선 그리기 (Arc segment 사용)
    2. IsCovered=True로 Sheet 생성 (Sweep 없음)
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

    # 사분면에 따른 중심점과 각도 계산
    if quadrant == 1:
        # 1사분면: 기준점 우상단, 원 중심은 좌하단
        center_x = ref_x - r_outer
        center_z = ref_z - r_outer
        # 외측 Arc: 0도(+X) → 90도(+Z)
        outer_start_angle = 0
        outer_end_angle = 90
        # X막대 방향: -X, Z막대 방향: -Z
        x_bar_dir = -1
        z_bar_dir = -1
    elif quadrant == 2:
        # 2사분면: 기준점 우하단
        center_x = ref_x - r_outer
        center_z = ref_z + r_outer
        # 외측 Arc: 270도(-Z) → 360도(+X)
        outer_start_angle = 270
        outer_end_angle = 360
        x_bar_dir = -1
        z_bar_dir = 1
    elif quadrant == 3:
        # 3사분면: 기준점 좌하단
        center_x = ref_x + r_outer
        center_z = ref_z + r_outer
        # 외측 Arc: 180도(-X) → 270도(-Z)
        outer_start_angle = 180
        outer_end_angle = 270
        x_bar_dir = 1
        z_bar_dir = 1
    else:  # quadrant == 4
        # 4사분면: 기준점 좌상단
        center_x = ref_x + r_outer
        center_z = ref_z - r_outer
        # 외측 Arc: 90도(+Z) → 180도(-X)
        outer_start_angle = 90
        outer_end_angle = 180
        x_bar_dir = 1
        z_bar_dir = -1

    # 점들 계산
    # 외측 Arc 시작/끝점
    outer_start_x = center_x + r_outer * math.cos(math.radians(outer_start_angle))
    outer_start_z = center_z + r_outer * math.sin(math.radians(outer_start_angle))
    outer_end_x = center_x + r_outer * math.cos(math.radians(outer_end_angle))
    outer_end_z = center_z + r_outer * math.sin(math.radians(outer_end_angle))

    # 내측 Arc 시작/끝점 (역방향)
    inner_end_x = center_x + r_inner * math.cos(math.radians(outer_end_angle))
    inner_end_z = center_z + r_inner * math.sin(math.radians(outer_end_angle))
    inner_start_x = center_x + r_inner * math.cos(math.radians(outer_start_angle))
    inner_start_z = center_z + r_inner * math.sin(math.radians(outer_start_angle))

    # 막대 끝점
    # 외측 끝에서 Z방향 막대
    z_bar_end_x = outer_end_x
    z_bar_end_z = outer_end_z + z_bar_dir * z_bar_len

    # 내측 시작에서 X방향 막대
    x_bar_end_x = inner_start_x + x_bar_dir * x_bar_len
    x_bar_end_z = inner_start_z

    # Polyline 점 리스트 (Y=0 고정)
    # 경로: 외측Arc 시작 → (Arc) → 외측Arc 끝 → Z막대 끝 → 내측Arc 끝 → (Arc 역방향) → 내측Arc 시작 → X막대 끝 → 외측Arc 시작

    points = [
        # P0: 외측 Arc 시작
        (outer_start_x, 0, outer_start_z),
        # P1: 외측 Arc 끝 (Arc segment로 연결)
        (outer_end_x, 0, outer_end_z),
        # P2: Z막대 끝 (Line)
        (z_bar_end_x, 0, z_bar_end_z),
        # P3: 내측 Arc 끝 (Line)
        (inner_end_x, 0, inner_end_z),
        # P4: 내측 Arc 시작 (Arc segment 역방향)
        (inner_start_x, 0, inner_start_z),
        # P5: X막대 끝 (Line)
        (x_bar_end_x, 0, x_bar_end_z),
        # P6: 외측 Arc 시작 (닫기, Line)
        (outer_start_x, 0, outer_start_z)
    ]

    # PolylinePoints 생성
    point_list = []
    for x, y, z in points:
        point_list.append([
            "NAME:PLPoint",
            "X:=", "{}mm".format(x),
            "Y:=", "{}mm".format(y),
            "Z:=", "{}mm".format(z)
        ])

    # PolylineSegments 생성
    segments = []

    # Seg 0: P0 → P1 (외측 Arc)
    segments.append([
        "NAME:PLSegment",
        "SegmentType:=", "Arc",
        "StartIndex:=", 0,
        "NoOfPoints:=", 2,
        "ArcAngle:=", "{}deg".format(outer_end_angle - outer_start_angle),
        "ArcCenterX:=", "{}mm".format(center_x),
        "ArcCenterY:=", "0mm",
        "ArcCenterZ:=", "{}mm".format(center_z),
        "ArcPlane:=", "XZ"
    ])

    # Seg 1: P1 → P2 (Line)
    segments.append([
        "NAME:PLSegment",
        "SegmentType:=", "Line",
        "StartIndex:=", 1,
        "NoOfPoints:=", 2
    ])

    # Seg 2: P2 → P3 (Line)
    segments.append([
        "NAME:PLSegment",
        "SegmentType:=", "Line",
        "StartIndex:=", 2,
        "NoOfPoints:=", 2
    ])

    # Seg 3: P3 → P4 (내측 Arc, 역방향)
    inner_arc_angle = -(outer_end_angle - outer_start_angle)  # 음수로 역방향
    segments.append([
        "NAME:PLSegment",
        "SegmentType:=", "Arc",
        "StartIndex:=", 3,
        "NoOfPoints:=", 2,
        "ArcAngle:=", "{}deg".format(inner_arc_angle),
        "ArcCenterX:=", "{}mm".format(center_x),
        "ArcCenterY:=", "0mm",
        "ArcCenterZ:=", "{}mm".format(center_z),
        "ArcPlane:=", "XZ"
    ])

    # Seg 4: P4 → P5 (Line)
    segments.append([
        "NAME:PLSegment",
        "SegmentType:=", "Line",
        "StartIndex:=", 4,
        "NoOfPoints:=", 2
    ])

    # Seg 5: P5 → P6 (Line, 닫기)
    segments.append([
        "NAME:PLSegment",
        "SegmentType:=", "Line",
        "StartIndex:=", 5,
        "NoOfPoints:=", 2
    ])

    # Polyline 생성
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
            ] + segments,
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
    print("  2D Sheet 생성: {} (Arc + Rectangle)".format(name))


def create_anglings_from_csv(csv_file_path, name_prefix="Angling"):
    """CSV 파일에서 앵글링들을 생성"""
    print("=" * 60)
    print("Maxwell 3D - 앵글링 생성 V3 (2D Sheet with Arc)")
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
            create_angling_sheet(oEditor, angling, angling_name)
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
