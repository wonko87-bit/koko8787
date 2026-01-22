# -*- coding: utf-8 -*-
"""
Maxwell 3D - Lead Path 생성 v2 (Center Point Arc 방식)
원점(0,0,0)에서 X축 방향으로 시작
Line + Center Point Arc로 경로 생성
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os
import math


def read_leadpath_csv(csv_file_path):
    """
    CSV 파일에서 Lead Path 데이터를 읽어옵니다.

    CSV 구조:
    - Initial_Straight
    - Bend#_Direction, Bend#_Radius, Bend#_Angle, Bend#_Straight (반복)

    Direction: left/right (XY plane), up/down (vertical)

    Returns:
        paths: 경로 정보 리스트
    """
    paths = []

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

    if len(rows) < 2:
        return paths

    # 데이터 행 처리 (2행부터)
    for i in range(1, len(rows)):
        row = rows[i]
        if not row or all(not cell.strip() for cell in row):
            continue

        try:
            # 기본 정보
            initial_straight = float(row[0].strip())

            # Bending 정보 추출 (1열부터 4개씩 묶음)
            bendings = []
            col_idx = 1
            while col_idx + 3 < len(row):
                if not row[col_idx].strip():
                    break

                bend_direction = row[col_idx].strip().lower()
                bend_radius = float(row[col_idx + 1].strip())
                bend_angle = float(row[col_idx + 2].strip())  # 도 단위
                bend_straight = float(row[col_idx + 3].strip())

                bendings.append({
                    'direction': bend_direction,
                    'radius': bend_radius,
                    'angle': bend_angle,
                    'straight': bend_straight
                })

                col_idx += 4

            paths.append({
                'initial_straight': initial_straight,
                'bendings': bendings
            })
        except:
            continue

    return paths


def add_points(p1, p2):
    """두 점 더하기"""
    return (p1[0]+p2[0], p1[1]+p2[1], p1[2]+p2[2])


def scale_vector(v, scale):
    """벡터 스케일링"""
    return (v[0]*scale, v[1]*scale, v[2]*scale)


def normalize_vector(v):
    """벡터 정규화"""
    length = math.sqrt(v[0]**2 + v[1]**2 + v[2]**2)
    if length == 0:
        return (0, 0, 0)
    return (v[0]/length, v[1]/length, v[2]/length)


def cross_product(v1, v2):
    """벡터 외적"""
    return (
        v1[1]*v2[2] - v1[2]*v2[1],
        v1[2]*v2[0] - v1[0]*v2[2],
        v1[0]*v2[1] - v1[1]*v2[0]
    )


def rotate_vector_z(v, angle_deg):
    """Z축 중심으로 벡터 회전"""
    angle_rad = math.radians(angle_deg)
    cos_a = math.cos(angle_rad)
    sin_a = math.sin(angle_rad)

    return (
        v[0]*cos_a - v[1]*sin_a,
        v[0]*sin_a + v[1]*cos_a,
        v[2]
    )


def rotate_vector_axis(v, axis, angle_deg):
    """임의의 축을 중심으로 벡터 회전 (Rodrigues' formula)"""
    angle_rad = math.radians(angle_deg)
    k = normalize_vector(axis)
    cos_a = math.cos(angle_rad)
    sin_a = math.sin(angle_rad)

    dot_kv = k[0]*v[0] + k[1]*v[1] + k[2]*v[2]
    cross_kv = cross_product(k, v)

    v_rot = (
        v[0]*cos_a + cross_kv[0]*sin_a + k[0]*dot_kv*(1-cos_a),
        v[1]*cos_a + cross_kv[1]*sin_a + k[1]*dot_kv*(1-cos_a),
        v[2]*cos_a + cross_kv[2]*sin_a + k[2]*dot_kv*(1-cos_a)
    )

    return v_rot


def create_line(oEditor, pt1, pt2, name):
    """두 점을 연결하는 Line 생성"""
    oEditor.CreateLine(
        [
            "NAME:LineParameters",
            "XStart:=", "{}mm".format(pt1[0]),
            "YStart:=", "{}mm".format(pt1[1]),
            "ZStart:=", "{}mm".format(pt1[2]),
            "XEnd:=", "{}mm".format(pt2[0]),
            "YEnd:=", "{}mm".format(pt2[1]),
            "ZEnd:=", "{}mm".format(pt2[2])
        ],
        [
            "NAME:Attributes",
            "Name:=", name,
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


def create_center_point_arc(oEditor, center, start_pt, angle_deg, name):
    """
    Center Point Arc 생성
    center: 중심점
    start_pt: 시작점
    angle_deg: 회전 각도 (+ 반시계, - 시계)
    """
    oEditor.CreateCircle(
        [
            "NAME:CircleParameters",
            "IsCovered:=", False,
            "XCenter:=", "{}mm".format(center[0]),
            "YCenter:=", "{}mm".format(center[1]),
            "ZCenter:=", "{}mm".format(center[2]),
            "Radius:=", "{}mm".format(math.sqrt(
                (start_pt[0]-center[0])**2 +
                (start_pt[1]-center[1])**2 +
                (start_pt[2]-center[2])**2
            )),
            "WhichAxis:=", "Z",  # 기본값, 상황에 따라 변경 필요
            "NumSegments:=", "0"
        ],
        [
            "NAME:Attributes",
            "Name:=", name,
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


def unite_objects(oEditor, obj_names):
    """Boolean Unite"""
    if len(obj_names) < 2:
        return obj_names[0] if obj_names else None

    oEditor.Unite(
        [
            "NAME:Selections",
            "Selections:=", ",".join(obj_names)
        ],
        [
            "NAME:UniteParameters",
            "KeepOriginals:=", False
        ]
    )
    return obj_names[0]


def create_path_from_data(oEditor, path_data, path_name):
    """경로 데이터로부터 Line과 Arc를 생성"""

    parts = []
    part_counter = 0

    # 시작점: 원점
    current_pos = (0.0, 0.0, 0.0)
    print("Starting position: {}".format(current_pos))

    # 현재 방향: X축
    current_dir = (1.0, 0.0, 0.0)
    print("Starting direction: {}".format(current_dir))

    # 현재 누적 XY 평면 회전 각도 (Z축 기준)
    cumulative_xy_rotation = 0.0

    # 초기 직진
    initial_straight = path_data['initial_straight']
    print("Initial straight distance: {}".format(initial_straight))

    if initial_straight > 0:
        part_counter += 1
        line_name = "{}_Part{}".format(path_name, part_counter)

        next_pos = add_points(current_pos, scale_vector(current_dir, initial_straight))
        print("Creating line '{}' from {} to {}".format(line_name, current_pos, next_pos))
        create_line(oEditor, current_pos, next_pos, line_name)
        parts.append(line_name)
        print("Line created successfully")

        current_pos = next_pos

    # 각 Bending 처리 (임시 비활성화 - 디버깅)
    print("Number of bendings: {} (currently disabled)".format(len(path_data['bendings'])))

    if False:  # 디버깅: bending 비활성화
        for bend_idx, bending in enumerate(path_data['bendings']):
            direction = bending['direction']
            radius = bending['radius']
            angle = bending['angle']
            straight = bending['straight']

            part_counter += 1
            arc_name = "{}_Part{}".format(path_name, part_counter)

            if direction in ['left', 'right']:
                # XY 평면 bending

                # 중심점 결정
                # left: 현재 방향 기준 왼쪽 (Z축 기준 +90도 회전)
                # right: 현재 방향 기준 오른쪽 (Z축 기준 -90도 회전)
                if direction == 'left':
                    perp_dir = rotate_vector_z(current_dir, 90)
                    arc_angle = angle  # + 방향 (반시계)
                else:  # right
                    perp_dir = rotate_vector_z(current_dir, -90)
                    arc_angle = -angle  # - 방향 (시계)

                arc_center = add_points(current_pos, scale_vector(perp_dir, radius))

                # Arc 생성 (일단 전체 원으로, 나중에 부분 호로 수정 필요)
                create_center_point_arc(oEditor, arc_center, current_pos, arc_angle, arc_name)
                parts.append(arc_name)

                # 위치와 방향 업데이트
                current_dir = rotate_vector_z(current_dir, arc_angle)
                cumulative_xy_rotation += arc_angle

                # 새 위치 계산
                radius_vec = (current_pos[0] - arc_center[0],
                             current_pos[1] - arc_center[1],
                             current_pos[2] - arc_center[2])
                rotated_radius = rotate_vector_z(radius_vec, arc_angle)
                current_pos = add_points(arc_center, rotated_radius)

            elif direction in ['up', 'down']:
                # 수직 평면 bending (상대좌표계)

                # 상대좌표계에서:
                # - X'축: 현재 진행 방향
                # - Y'축: 현재 진행 방향을 Z축 기준으로 90도 회전
                # - Z'축: Z축 (변하지 않음)

                # up/down 방향 결정
                if direction == 'up':
                    # 현재 방향과 Z축에 수직인 방향 (위쪽)
                    perp_dir = (0, 0, 1)  # 일단 Z축 방향
                    arc_angle = angle  # + 방향
                else:  # down
                    perp_dir = (0, 0, -1)  # Z축 반대 방향
                    arc_angle = -angle  # - 방향

                # 실제로는 현재 방향에 수직이면서 XY평면에 수직인 방향
                # current_dir과 (0,0,1)의 외적
                lateral_dir = cross_product(current_dir, (0, 0, 1))
                lateral_dir = normalize_vector(lateral_dir)

                # up이면 Z 방향 성분이 +, down이면 -
                if direction == 'up':
                    perp_dir = cross_product(lateral_dir, current_dir)
                else:
                    perp_dir = cross_product(current_dir, lateral_dir)

                perp_dir = normalize_vector(perp_dir)

                arc_center = add_points(current_pos, scale_vector(perp_dir, radius))

                # Arc 생성
                create_center_point_arc(oEditor, arc_center, current_pos, arc_angle, arc_name)
                parts.append(arc_name)

                # 위치와 방향 업데이트 (lateral_dir 축 중심으로 회전)
                current_dir = rotate_vector_axis(current_dir, lateral_dir, arc_angle)

                # 새 위치 계산
                radius_vec = (current_pos[0] - arc_center[0],
                             current_pos[1] - arc_center[1],
                             current_pos[2] - arc_center[2])
                rotated_radius = rotate_vector_axis(radius_vec, lateral_dir, arc_angle)
                current_pos = add_points(arc_center, rotated_radius)

            # Bending 후 직진
            if straight > 0:
                part_counter += 1
                line_name = "{}_Part{}".format(path_name, part_counter)

                next_pos = add_points(current_pos, scale_vector(current_dir, straight))
                create_line(oEditor, current_pos, next_pos, line_name)
                parts.append(line_name)

                current_pos = next_pos

    # 모든 파트 Unite
    if len(parts) > 0:
        final_name = unite_objects(oEditor, parts)

        # 최종 이름 변경
        if final_name:
            try:
                oEditor.ChangeProperty(
                    [
                        "NAME:AllTabs",
                        [
                            "NAME:Geometry3DAttributeTab",
                            [
                                "NAME:PropServers",
                                final_name
                            ],
                            [
                                "NAME:ChangedProps",
                                [
                                    "NAME:Name",
                                    "Value:=", path_name
                                ]
                            ]
                        ]
                    ]
                )
            except:
                pass


def create_leadpaths_from_csv(csv_file_path, name_prefix="LeadPath"):
    """CSV 파일에서 Lead Path들을 생성"""
    print("Reading CSV file: {}".format(csv_file_path))
    paths = read_leadpath_csv(csv_file_path)

    print("Number of paths found: {}".format(len(paths)))
    if not paths:
        print("No paths found in CSV!")
        return

    oProject = oDesktop.GetActiveProject()
    if oProject is None:
        oProject = oDesktop.NewProject()

    oDesign = oProject.GetActiveDesign()
    if oDesign is None:
        oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
        oDesign = oProject.GetActiveDesign()

    oEditor = oDesign.SetActiveEditor("3D Modeler")

    for idx, path_data in enumerate(paths):
        path_name = "{}_{}".format(name_prefix, idx + 1)

        try:
            create_path_from_data(oEditor, path_data, path_name)
        except Exception as e:
            print("Error creating path {}: {}".format(path_name, str(e)))
            continue

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

csv_file = os.path.join(script_dir, "LeadPathDim.csv")

if os.path.exists(csv_file):
    create_leadpaths_from_csv(
        csv_file_path=csv_file,
        name_prefix="LeadPath"
    )
