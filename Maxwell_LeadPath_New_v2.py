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


def create_polyline_path(oEditor, points, segments, name):
    """점들과 세그먼트 정보로 Polyline 생성"""
    point_list = ["NAME:PolylinePoints"]
    for pt in points:
        point_list.append([
            "NAME:PLPoint",
            "X:=", "{}mm".format(pt[0]),
            "Y:=", "{}mm".format(pt[1]),
            "Z:=", "{}mm".format(pt[2])
        ])

    segment_list = ["NAME:PolylineSegments"]
    for seg in segments:
        segment_list.append(seg)

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


def create_path_from_data(oEditor, path_data, path_name):
    """경로 데이터로부터 하나의 Polyline을 생성 (Line + Arc segments)"""

    points = []    # 모든 점들
    segments = []  # 모든 세그먼트들

    # 시작점: 원점
    current_pos = (0.0, 0.0, 0.0)
    points.append(current_pos)

    # 현재 방향: X축
    current_dir = (1.0, 0.0, 0.0)

    # 초기 직진
    initial_straight = path_data['initial_straight']

    if initial_straight > 0:
        next_pos = add_points(current_pos, scale_vector(current_dir, initial_straight))
        points.append(next_pos)

        # Line segment 추가
        segments.append([
            "NAME:PLSegment",
            "SegmentType:=", "Line",
            "StartIndex:=", 0,
            "NoOfPoints:=", 2
        ])

        current_pos = next_pos

    # 각 Bending 처리
    for bend_idx, bending in enumerate(path_data['bendings']):
        direction = bending['direction']
        radius = bending['radius']
        angle = bending['angle']
        straight = bending['straight']

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

            # Arc의 끝점 계산
            radius_vec = (current_pos[0] - arc_center[0],
                         current_pos[1] - arc_center[1],
                         current_pos[2] - arc_center[2])
            rotated_radius = rotate_vector_z(radius_vec, arc_angle)
            arc_end = add_points(arc_center, rotated_radius)

            # Arc segment 추가
            start_idx = len(points) - 1
            points.append(arc_center)  # 중심점
            points.append(arc_end)     # 끝점

            segments.append([
                "NAME:PLSegment",
                "SegmentType:=", "AngularArc",
                "StartIndex:=", start_idx,
                "NoOfPoints:=", 3,
                "ArcAngle:=", "{}deg".format(arc_angle),
                "ArcCenterX:=", "{}mm".format(arc_center[0]),
                "ArcCenterY:=", "{}mm".format(arc_center[1]),
                "ArcCenterZ:=", "{}mm".format(arc_center[2])
            ])

            # 방향과 위치 업데이트
            current_dir = rotate_vector_z(current_dir, arc_angle)
            current_pos = arc_end

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

            # Arc의 끝점 계산
            radius_vec = (current_pos[0] - arc_center[0],
                         current_pos[1] - arc_center[1],
                         current_pos[2] - arc_center[2])
            rotated_radius = rotate_vector_axis(radius_vec, lateral_dir, arc_angle)
            arc_end = add_points(arc_center, rotated_radius)

            # Arc segment 추가
            start_idx = len(points) - 1
            points.append(arc_center)
            points.append(arc_end)

            segments.append([
                "NAME:PLSegment",
                "SegmentType:=", "AngularArc",
                "StartIndex:=", start_idx,
                "NoOfPoints:=", 3,
                "ArcAngle:=", "{}deg".format(arc_angle),
                "ArcCenterX:=", "{}mm".format(arc_center[0]),
                "ArcCenterY:=", "{}mm".format(arc_center[1]),
                "ArcCenterZ:=", "{}mm".format(arc_center[2])
            ])

            # 방향과 위치 업데이트
            current_dir = rotate_vector_axis(current_dir, lateral_dir, arc_angle)
            current_pos = arc_end

        # Bending 후 직진
        if straight > 0:
            next_pos = add_points(current_pos, scale_vector(current_dir, straight))
            points.append(next_pos)

            start_idx = len(points) - 2
            segments.append([
                "NAME:PLSegment",
                "SegmentType:=", "Line",
                "StartIndex:=", start_idx,
                "NoOfPoints:=", 2
            ])

            current_pos = next_pos

    # 하나의 Polyline 생성
    create_polyline_path(oEditor, points, segments, path_name)


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
