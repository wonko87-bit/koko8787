# -*- coding: utf-8 -*-
"""
Maxwell 3D - Lead Path 생성 v2 (Center Point Arc 방식)
원점(0,0,0)에서 X축 방향으로 시작
Line + Center Point Arc로 경로 생성
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


def create_line_polyline(oEditor, pt1, pt2, name):
    """두 점을 연결하는 직선 Polyline 생성"""
    point_list = ["NAME:PolylinePoints"]
    point_list.append([
        "NAME:PLPoint",
        "X:=", "{}mm".format(pt1[0]),
        "Y:=", "{}mm".format(pt1[1]),
        "Z:=", "{}mm".format(pt1[2])
    ])
    point_list.append([
        "NAME:PLPoint",
        "X:=", "{}mm".format(pt2[0]),
        "Y:=", "{}mm".format(pt2[1]),
        "Z:=", "{}mm".format(pt2[2])
    ])

    segment_list = ["NAME:PolylineSegments"]
    segment_list.append([
        "NAME:PLSegment",
        "SegmentType:=", "Line",
        "StartIndex:=", 0,
        "NoOfPoints:=", 2
    ])

    oEditor.CreatePolyline(
        [
            "NAME:PolylineParameters",
            "IsPolylineCovered:=", False,
            "IsPolylineClosed:=", False,
            point_list,
            segment_list,
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


def create_arc_polyline(oEditor, points_list, name):
    """여러 점을 연결하는 Polyline 생성 (Arc 근사화용)"""
    point_list = ["NAME:PolylinePoints"]
    for pt in points_list:
        point_list.append([
            "NAME:PLPoint",
            "X:=", "{}mm".format(pt[0]),
            "Y:=", "{}mm".format(pt[1]),
            "Z:=", "{}mm".format(pt[2])
        ])

    segment_list = ["NAME:PolylineSegments"]
    for i in range(len(points_list) - 1):
        segment_list.append([
            "NAME:PLSegment",
            "SegmentType:=", "Line",
            "StartIndex:=", i,
            "NoOfPoints:=", 2
        ])

    oEditor.CreatePolyline(
        [
            "NAME:PolylineParameters",
            "IsPolylineCovered:=", False,
            "IsPolylineClosed:=", False,
            point_list,
            segment_list,
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


def create_arc_xy_plane(oEditor, center, radius, start_angle_deg, arc_angle_deg, name):
    """
    XY 평면에서 Arc 생성 (사분원 방식)
    center: 중심점 (x, y, z)
    radius: 반지름
    start_angle_deg: 시작 각도 (X축 기준, 도 단위)
    arc_angle_deg: 호의 각도 (양수: 반시계, 음수: 시계)
    """
    # 1. 원 생성 (XY 평면, Z축 기준)
    circle_name = "{}_Circle".format(name)
    oEditor.CreateCircle(
        [
            "NAME:CircleParameters",
            "IsCovered:=", False,
            "XCenter:=", "{}mm".format(center[0]),
            "YCenter:=", "{}mm".format(center[1]),
            "ZCenter:=", "{}mm".format(center[2]),
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

    # 2. 사각형 2개 만들어서 사분원으로 만들기
    rect_size = radius * 3.0

    # 첫 번째 사각형 - 원의 왼쪽 절반을 덮음 (X < 0)
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

    # 첫 번째 사각형을 start_angle로 회전 (원점 기준)
    oEditor.Rotate(
        ["NAME:Selections", "Selections:=", rect1_name],
        [
            "NAME:RotateParameters",
            "RotateAxis:=", "Z",
            "RotateAngle:=", "{}deg".format(start_angle_deg)
        ]
    )

    # 첫 번째 사각형을 arc_center로 이동
    oEditor.Move(
        ["NAME:Selections", "Selections:=", rect1_name],
        [
            "NAME:TranslateParameters",
            "TranslateVectorX:=", "{}mm".format(center[0]),
            "TranslateVectorY:=", "{}mm".format(center[1]),
            "TranslateVectorZ:=", "{}mm".format(center[2])
        ]
    )

    # 두 번째 사각형 - 원의 아래쪽 절반을 덮음 (Y < 0)
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

    # 두 번째 사각형을 start_angle + 90도로 회전 (원점 기준)
    oEditor.Rotate(
        ["NAME:Selections", "Selections:=", rect2_name],
        [
            "NAME:RotateParameters",
            "RotateAxis:=", "Z",
            "RotateAngle:=", "{}deg".format(start_angle_deg + 90)
        ]
    )

    # 두 번째 사각형을 arc_center로 이동
    oEditor.Move(
        ["NAME:Selections", "Selections:=", rect2_name],
        [
            "NAME:TranslateParameters",
            "TranslateVectorX:=", "{}mm".format(center[0]),
            "TranslateVectorY:=", "{}mm".format(center[1]),
            "TranslateVectorZ:=", "{}mm".format(center[2])
        ]
    )

    # 3. Subtract로 사분원 만들기
    oEditor.Subtract(
        ["NAME:Selections", "Blank Parts:=", circle_name, "Tool Parts:=", "{},{}".format(rect1_name, rect2_name)],
        ["NAME:SubtractParameters", "KeepOriginals:=", False]
    )

    # 4. 원하는 각도만 남기기
    abs_arc_angle = abs(arc_angle_deg)
    if abs_arc_angle < 90:
        # 회전 각도 = 90 - abs_arc_angle
        rotate_angle = 90 - abs_arc_angle

        # 사분원을 회전 (시계 방향이면 반대로)
        if arc_angle_deg < 0:
            rotate_angle = -rotate_angle

        oEditor.Rotate(
            ["NAME:Selections", "Selections:=", circle_name],
            [
                "NAME:RotateParameters",
                "RotateAxis:=", "Z",
                "RotateAngle:=", "{}deg".format(rotate_angle)
            ]
        )

        # 새 사각형으로 다시 자르기
        rect3_name = "{}_Rect3".format(name)
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

        # 사각형3을 start_angle + 90로 회전 (원점 기준)
        oEditor.Rotate(
            ["NAME:Selections", "Selections:=", rect3_name],
            [
                "NAME:RotateParameters",
                "RotateAxis:=", "Z",
                "RotateAngle:=", "{}deg".format(start_angle_deg + 90)
            ]
        )

        # 사각형3을 arc_center로 이동
        oEditor.Move(
            ["NAME:Selections", "Selections:=", rect3_name],
            [
                "NAME:TranslateParameters",
                "TranslateVectorX:=", "{}mm".format(center[0]),
                "TranslateVectorY:=", "{}mm".format(center[1]),
                "TranslateVectorZ:=", "{}mm".format(center[2])
            ]
        )

        # 다시 Subtract
        oEditor.Subtract(
            ["NAME:Selections", "Blank Parts:=", circle_name, "Tool Parts:=", rect3_name],
            ["NAME:SubtractParameters", "KeepOriginals:=", False]
        )

    # 5. 최종 이름 변경
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
    """경로 데이터로부터 Line과 Arc를 별도 객체로 생성 후 Unite"""

    parts = []
    part_counter = 0

    # 시작점: 원점
    current_pos = (0.0, 0.0, 0.0)

    # 현재 방향: X축
    current_dir = (1.0, 0.0, 0.0)

    # 초기 직진
    initial_straight = path_data['initial_straight']

    if initial_straight > 0:
        part_counter += 1
        line_name = "{}_Part{}".format(path_name, part_counter)

        next_pos = add_points(current_pos, scale_vector(current_dir, initial_straight))
        create_line_polyline(oEditor, current_pos, next_pos, line_name)
        parts.append(line_name)

        current_pos = next_pos

    # 각 Bending 처리
    for bend_idx, bending in enumerate(path_data['bendings']):
        direction = bending['direction']
        radius = bending['radius']
        angle = bending['angle']
        straight = bending['straight']

        # Arc를 작은 선분들로 근사화 (각도에 따라 분할 수 결정)
        num_segments = max(int(abs(angle) / 5), 4)  # 5도마다 하나의 segment, 최소 4개

        if direction in ['left', 'right']:
            # XY 평면 bending

            # 중심점 결정
            if direction == 'left':
                perp_dir = rotate_vector_z(current_dir, 90)
                arc_angle = angle
            else:  # right
                perp_dir = rotate_vector_z(current_dir, -90)
                arc_angle = -angle

            arc_center = add_points(current_pos, scale_vector(perp_dir, radius))

            # 시작 각도 계산 (X축 기준)
            radius_vec = (current_pos[0] - arc_center[0],
                         current_pos[1] - arc_center[1])
            start_angle_deg = math.degrees(math.atan2(radius_vec[1], radius_vec[0]))

            # 실제 Arc 생성
            part_counter += 1
            arc_name = "{}_Part{}".format(path_name, part_counter)
            create_arc_xy_plane(oEditor, arc_center, radius, start_angle_deg, arc_angle, arc_name)
            parts.append(arc_name)

            # 방향과 위치 업데이트
            current_dir = rotate_vector_z(current_dir, arc_angle)

            # 끝 위치 계산
            radius_vec_full = (current_pos[0] - arc_center[0],
                              current_pos[1] - arc_center[1],
                              current_pos[2] - arc_center[2])
            rotated_radius = rotate_vector_z(radius_vec_full, arc_angle)
            current_pos = add_points(arc_center, rotated_radius)

        elif direction in ['up', 'down']:
            # 수직 평면 bending

            # 회전 축: current_dir과 Z축의 외적
            lateral_dir = cross_product(current_dir, (0, 0, 1))
            lateral_dir = normalize_vector(lateral_dir)

            # up/down 방향 결정
            if direction == 'up':
                perp_dir = cross_product(lateral_dir, current_dir)
                arc_angle = angle
            else:  # down
                perp_dir = cross_product(current_dir, lateral_dir)
                arc_angle = -angle

            perp_dir = normalize_vector(perp_dir)
            arc_center = add_points(current_pos, scale_vector(perp_dir, radius))

            # Arc를 작은 선분들로 근사화
            arc_points = [current_pos]
            for i in range(1, num_segments + 1):
                t = float(i) / num_segments
                current_angle = arc_angle * t

                radius_vec = (current_pos[0] - arc_center[0],
                             current_pos[1] - arc_center[1],
                             current_pos[2] - arc_center[2])
                rotated_radius = rotate_vector_axis(radius_vec, lateral_dir, current_angle)
                pt = add_points(arc_center, rotated_radius)
                arc_points.append(pt)

            # Arc Polyline 생성
            part_counter += 1
            arc_name = "{}_Part{}".format(path_name, part_counter)
            create_arc_polyline(oEditor, arc_points, arc_name)
            parts.append(arc_name)

            # 방향과 위치 업데이트
            current_dir = rotate_vector_axis(current_dir, lateral_dir, arc_angle)
            current_pos = arc_points[-1]

        # Bending 후 직진
        if straight > 0:
            part_counter += 1
            line_name = "{}_Part{}".format(path_name, part_counter)

            next_pos = add_points(current_pos, scale_vector(current_dir, straight))
            create_line_polyline(oEditor, current_pos, next_pos, line_name)
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
    paths = read_leadpath_csv(csv_file_path)

    if not paths:
        return

    for idx, path_data in enumerate(paths):
        path_name = "{}_{}".format(name_prefix, idx + 1)
        create_path_from_data(oEditor, path_data, path_name)

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
    create_leadpaths_from_csv(csv_file, "LeadPath")
