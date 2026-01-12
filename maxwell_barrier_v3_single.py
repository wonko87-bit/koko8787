# -*- coding: utf-8 -*-
"""
Maxwell 3D - 배리어 생성 V3 (특정 배리어만 생성)
디버깅용: 13번 배리어만 생성
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_csv_data(csv_file_path):
    """CSV 파일에서 배리어 데이터를 읽어옵니다."""
    groups = []

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

    # 그룹 감지 (빈 행으로 구분)
    i = 0
    while i < len(rows):
        # 빈 행 건너뛰기
        if not rows[i] or all(not cell.strip() for cell in rows[i]):
            i += 1
            continue

        # 그룹 시작 - 4개 행 읽기
        if i + 3 < len(rows):
            row1 = rows[i]      # 내경들
            row2 = rows[i + 1]  # 외경들
            row3 = rows[i + 2]  # z축 이동들
            row4 = rows[i + 3]  # sweep 길이들

            # D열부터 시작 (col_idx=3)
            inner_diameters = []
            outer_diameters = []
            z_offsets = []
            sweep_distances = []

            # 1행에서 원소 개수 파악 (D열=col 3부터)
            for col_idx in range(3, len(row1)):
                try:
                    if row1[col_idx].strip():
                        inner_diameters.append(float(row1[col_idx]))
                    else:
                        break  # 빈 셀 만나면 종료
                except (ValueError, AttributeError):
                    break

            # 원통 개수 = 원소 개수 - 1
            num_cylinders = len(inner_diameters) - 1

            if num_cylinders > 0:
                # 2~4행에서 원통 개수만큼 데이터 읽기
                for j in range(num_cylinders):
                    col_idx = 3 + j  # D열부터 시작

                    # 외경
                    try:
                        if len(row2) > col_idx and row2[col_idx].strip():
                            outer_diameters.append(float(row2[col_idx]))
                        else:
                            outer_diameters.append(None)
                    except (ValueError, AttributeError):
                        outer_diameters.append(None)

                    # z축 이동
                    try:
                        if len(row3) > col_idx and row3[col_idx].strip():
                            z_offsets.append(float(row3[col_idx]))
                        else:
                            z_offsets.append(None)
                    except (ValueError, AttributeError):
                        z_offsets.append(None)

                    # sweep 길이
                    try:
                        if len(row4) > col_idx and row4[col_idx].strip():
                            sweep_distances.append(float(row4[col_idx]))
                        else:
                            sweep_distances.append(None)
                    except (ValueError, AttributeError):
                        sweep_distances.append(None)

                groups.append({
                    'inner_diameters': inner_diameters,
                    'outer_diameters': outer_diameters,
                    'z_offsets': z_offsets,
                    'sweep_distances': sweep_distances,
                    'num_cylinders': num_cylinders
                })

            i += 4  # 다음 그룹으로 (4행 건너뜀)
        else:
            break

    return groups


def create_rectangle_xz(oEditor, inner_radius, outer_radius, height, name):
    """XZ 평면에 직사각형 생성 (도넛 단면) - Polyline 5개 점"""
    # 직사각형 5개 점 (XZ 평면, Y=0, Z=0에서 시작):
    # 점1: (inner_radius, 0, 0)
    # 점2: (inner_radius, 0, height)
    # 점3: (outer_radius, 0, height)
    # 점4: (outer_radius, 0, 0)
    # 점5: (inner_radius, 0, 0) - 닫힌 도형

    points = [
        ["NAME:PLPoint", "X:=", "{}mm".format(inner_radius), "Y:=", "0mm", "Z:=", "0mm"],
        ["NAME:PLPoint", "X:=", "{}mm".format(inner_radius), "Y:=", "0mm", "Z:=", "{}mm".format(height)],
        ["NAME:PLPoint", "X:=", "{}mm".format(outer_radius), "Y:=", "0mm", "Z:=", "{}mm".format(height)],
        ["NAME:PLPoint", "X:=", "{}mm".format(outer_radius), "Y:=", "0mm", "Z:=", "0mm"],
        ["NAME:PLPoint", "X:=", "{}mm".format(inner_radius), "Y:=", "0mm", "Z:=", "0mm"]
    ]

    oEditor.CreatePolyline(
        [
            "NAME:PolylineParameters",
            "IsPolylineCovered:=", True,
            "IsPolylineClosed:=", True,
            ["NAME:PolylinePoints"] + points,
            [
                "NAME:PolylineSegments",
                ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 0, "NoOfPoints:=", 2],
                ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 1, "NoOfPoints:=", 2],
                ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 2, "NoOfPoints:=", 2],
                ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 3, "NoOfPoints:=", 2],
                ["NAME:PLSegment", "SegmentType:=", "Line", "StartIndex:=", 4, "NoOfPoints:=", 2]
            ]
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
    print("  생성: {} (XZ 평면 Polyline 직사각형, 5개 점)".format(name))


def sweep_around_z_axis(oEditor, obj_name, angle_deg=360):
    """Z축 중심으로 회전 (도넛 생성)"""
    oEditor.SweepAroundAxis(
        [
            "NAME:Selections",
            "Selections:=", obj_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:AxisSweepParameters",
            "DraftAngle:=", "0deg",
            "DraftType:=", "Round",
            "CheckFaceFaceIntersection:=", False,
            "SweepAxis:=", "Z",
            "SweepAngle:=", "{}deg".format(angle_deg),
            "NumOfSegments:=", "0"
        ]
    )
    print("  회전: {} → Z축 중심 {}도".format(obj_name, angle_deg))


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
    print("  이동: {} → (dX={}, dY={}, dZ={})".format(obj_name, dx, dy, dz))


def create_single_barrier(csv_file_path, barrier_number=13, name_prefix="Barrier"):
    """특정 배리어 하나만 생성 (디버깅용)"""
    print("=" * 60)
    print("Maxwell 3D - 배리어 생성 V3 (단일 배리어)")
    print("대상 배리어: {}".format(barrier_number))
    print("=" * 60)

    print("\nCSV 파일 읽기 시작...")
    # CSV 데이터 읽기
    groups = read_csv_data(csv_file_path)

    if not groups:
        print("오류: CSV 파일에서 데이터를 읽을 수 없습니다.")
        return

    # 모든 배리어를 리스트로 평탄화
    all_barriers = []
    cylinder_count = 0

    for group_idx, group in enumerate(groups):
        for cyl_idx in range(group['num_cylinders']):
            cylinder_count += 1
            all_barriers.append({
                'number': cylinder_count,
                'group_idx': group_idx,
                'cyl_idx': cyl_idx,
                'inner_dia': group['inner_diameters'][cyl_idx],
                'outer_dia': group['outer_diameters'][cyl_idx],
                'z_offset': group['z_offsets'][cyl_idx],
                'sweep_dist': group['sweep_distances'][cyl_idx]
            })

    print("총 {} 개의 배리어를 찾았습니다.".format(len(all_barriers)))

    # 목표 배리어 찾기
    target_barrier = None
    for b in all_barriers:
        if b['number'] == barrier_number:
            target_barrier = b
            break

    if target_barrier is None:
        print("오류: 배리어 {}를 찾을 수 없습니다.".format(barrier_number))
        return

    print("\n배리어 {} 정보:".format(barrier_number))
    print("  내경: {}mm (반지름: {}mm)".format(target_barrier['inner_dia'], target_barrier['inner_dia']/2.0))
    print("  외경: {}mm (반지름: {}mm)".format(target_barrier['outer_dia'], target_barrier['outer_dia']/2.0))
    print("  높이: {}mm".format(target_barrier['sweep_dist']))
    print("  Z offset: {}mm".format(target_barrier['z_offset']))

    # Maxwell 프로젝트 설정
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

    # 배리어 생성
    barrier_name = "{}_{}".format(name_prefix, barrier_number)
    rect_name = "{}_Rect".format(barrier_name)

    inner_radius = target_barrier['inner_dia'] / 2.0
    outer_radius = target_barrier['outer_dia'] / 2.0
    height = target_barrier['sweep_dist']
    z_offset = target_barrier['z_offset']

    print("\n배리어 생성 시작...")

    try:
        # 1. XZ 평면에 직사각형 생성 (Z=0에서 시작)
        print("  [1] XZ 평면에 직사각형 생성 (Z=0)")
        print("      점1: ({}, 0, 0)".format(inner_radius))
        print("      점2: ({}, 0, {})".format(inner_radius, height))
        print("      점3: ({}, 0, {})".format(outer_radius, height))
        print("      점4: ({}, 0, 0)".format(outer_radius))
        print("      점5: ({}, 0, 0)".format(inner_radius))
        create_rectangle_xz(oEditor, inner_radius, outer_radius, height, rect_name)

        # 2. Z축 중심으로 360도 회전 (도넛 생성)
        print("\n  [2] Z축 중심으로 360도 회전")
        sweep_around_z_axis(oEditor, rect_name, 360)

        # 3. Z축 offset만큼 평행 이동
        print("\n  [3] Z축 offset만큼 평행 이동")
        move_object(oEditor, rect_name, 0, 0, z_offset)

        # 최종 이름 변경
        print("\n  [4] 이름 변경")
        oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Geometry3DAttributeTab",
                    [
                        "NAME:PropServers",
                        rect_name
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Name",
                            "Value:=", barrier_name
                        ]
                    ]
                ]
            ]
        )
        print("  완성: {}".format(barrier_name))

        # 뷰 맞추기
        oEditor.FitAll()

        print("\n" + "=" * 60)
        print("배리어 {} 생성 완료!".format(barrier_number))
        print("=" * 60)

    except Exception as e:
        print("\n오류 발생: {}".format(str(e)))
        print("배리어 {} 생성 실패".format(barrier_number))


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

csv_file = os.path.join(script_dir, "BarrierDim.csv")

print("\nCSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    print("\n오류: BarrierDim.csv 파일을 찾을 수 없습니다!")
    print("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    print("\n배리어 13번 생성 함수 호출 중...")
    # 배리어 13번만 생성
    create_single_barrier(
        csv_file_path=csv_file,
        barrier_number=13,
        name_prefix="Barrier"
    )
    print("\n스크립트 실행 완료!")
