# -*- coding: utf-8 -*-
"""
Maxwell 3D - Lead Path 생성 (원통좌표계 기반)
원통좌표계에서 시작점을 입력받아 연속적인 경로를 생성
초기 직진 + Bending(R-theta 또는 theta-Z) + 직진을 반복
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
    - Start_R, Start_Theta, Start_Z, Initial_Straight
    - Bend#_Type, Bend#_Radius, Bend#_Angle, Bend#_Straight (반복)

    Returns:
        paths: 경로 정보 리스트
    """
    paths = []

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

    if len(rows) < 2:
        return paths

    header = rows[0]

    # 데이터 행 처리 (2행부터)
    for i in range(1, len(rows)):
        row = rows[i]
        if not row or all(not cell.strip() for cell in row):
            continue

        try:
            # 기본 정보
            start_r = float(row[0].strip())
            start_theta = float(row[1].strip())  # 도 단위
            start_z = float(row[2].strip())
            initial_straight = float(row[3].strip())

            # Bending 정보 추출 (4열부터 4개씩 묶음)
            bendings = []
            col_idx = 4
            while col_idx + 3 < len(row):
                if not row[col_idx].strip():
                    break

                bend_type = row[col_idx].strip()
                bend_radius = float(row[col_idx + 1].strip())
                bend_angle = float(row[col_idx + 2].strip())  # 도 단위
                bend_straight = float(row[col_idx + 3].strip())

                bendings.append({
                    'type': bend_type,
                    'radius': bend_radius,
                    'angle': bend_angle,
                    'straight': bend_straight
                })

                col_idx += 4

            paths.append({
                'start_r': start_r,
                'start_theta': start_theta,
                'start_z': start_z,
                'initial_straight': initial_straight,
                'bendings': bendings
            })
        except:
            continue

    return paths


def cylindrical_to_cartesian(r, theta_deg, z):
    """원통좌표 → 직교좌표 변환"""
    theta_rad = math.radians(theta_deg)
    x = r * math.cos(theta_rad)
    y = r * math.sin(theta_rad)
    return (x, y, z)


def get_theta_direction(theta_deg):
    """theta 벡터 방향 (단위벡터) 반환"""
    theta_rad = math.radians(theta_deg)
    dx = -math.sin(theta_rad)
    dy = math.cos(theta_rad)
    dz = 0
    return (dx, dy, dz)


def get_radial_direction(theta_deg):
    """반경 방향 (단위벡터) 반환"""
    theta_rad = math.radians(theta_deg)
    dx = math.cos(theta_rad)
    dy = math.sin(theta_rad)
    dz = 0
    return (dx, dy, dz)


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


def rotate_vector(v, axis, angle_deg):
    """축을 중심으로 벡터 회전 (Rodrigues' rotation formula)"""
    angle_rad = math.radians(angle_deg)
    k = normalize_vector(axis)
    cos_a = math.cos(angle_rad)
    sin_a = math.sin(angle_rad)

    # v_rot = v*cos(a) + (k × v)*sin(a) + k*(k·v)*(1-cos(a))
    dot_kv = k[0]*v[0] + k[1]*v[1] + k[2]*v[2]
    cross_kv = cross_product(k, v)

    v_rot = (
        v[0]*cos_a + cross_kv[0]*sin_a + k[0]*dot_kv*(1-cos_a),
        v[1]*cos_a + cross_kv[1]*sin_a + k[1]*dot_kv*(1-cos_a),
        v[2]*cos_a + cross_kv[2]*sin_a + k[2]*dot_kv*(1-cos_a)
    )

    return v_rot


def add_points(p1, p2):
    """두 점 더하기"""
    return (p1[0]+p2[0], p1[1]+p2[1], p1[2]+p2[2])


def scale_vector(v, scale):
    """벡터 스케일링"""
    return (v[0]*scale, v[1]*scale, v[2]*scale)


def generate_arc_points(center, radius, start_dir, rotation_axis, angle_deg, num_segments=10):
    """
    호(arc) 경로상의 점들을 생성
    center: 호의 중심점
    radius: 곡률 반경
    start_dir: 시작 방향 (정규화된 벡터)
    rotation_axis: 회전축
    angle_deg: 회전 각도
    """
    points = []
    angle_step = angle_deg / num_segments

    for i in range(num_segments + 1):
        current_angle = angle_step * i
        rotated_dir = rotate_vector(start_dir, rotation_axis, current_angle)
        point = add_points(center, scale_vector(rotated_dir, radius))
        points.append(point)

    return points


def generate_path_points(path_data):
    """경로 데이터로부터 3D 경로 점들을 생성"""
    points = []

    # 시작점 (원통좌표 → 직교좌표)
    start_r = path_data['start_r']
    start_theta = path_data['start_theta']
    start_z = path_data['start_z']

    current_pos = cylindrical_to_cartesian(start_r, start_theta, start_z)
    points.append(current_pos)

    # 현재 방향 (초기: theta 방향)
    current_dir = get_theta_direction(start_theta)

    # 초기 직진
    initial_straight = path_data['initial_straight']
    if initial_straight > 0:
        next_pos = add_points(current_pos, scale_vector(current_dir, initial_straight))
        points.append(next_pos)
        current_pos = next_pos

    # 각 Bending 처리
    for bending in path_data['bendings']:
        bend_type = bending['type']
        bend_radius = bending['radius']
        bend_angle = bending['angle']
        bend_straight = bending['straight']

        if bend_type.lower() == 'r-theta':
            # R-theta plane bending (XY 평면, Z축 중심 회전)
            rotation_axis = (0, 0, 1)  # Z축

            # 호의 중심: 현재 위치에서 현재 방향의 수직 방향으로 bend_radius만큼
            # 현재 방향에 수직이면서 Z=0인 방향
            perp_dir = cross_product((0, 0, 1), current_dir)
            perp_dir = normalize_vector(perp_dir)

            arc_center = add_points(current_pos, scale_vector(perp_dir, bend_radius))

            # 호 시작 방향 (중심에서 현재 위치로)
            start_from_center = scale_vector(perp_dir, -1)

            # 호 생성
            arc_points = generate_arc_points(arc_center, bend_radius, start_from_center,
                                            rotation_axis, bend_angle)
            points.extend(arc_points[1:])  # 첫 점은 이미 있으므로 제외
            current_pos = arc_points[-1]

            # 방향 업데이트
            current_dir = rotate_vector(current_dir, rotation_axis, bend_angle)

        elif bend_type.lower() == 'theta-z':
            # theta-Z plane bending (수직 평면)
            # 회전축: 현재 위치의 radial 방향 (현재 theta에서의 접선에 수직)

            # 현재 위치에서 원점까지의 벡터의 XY 성분으로 radial 방향 추정
            # 더 정확하게는 현재 방향에 수직이면서 XY평면에 있는 방향
            radial_dir = cross_product(current_dir, (0, 0, 1))
            radial_dir = normalize_vector(radial_dir)
            rotation_axis = radial_dir

            # 호의 중심: 현재 위치에서 현재 방향의 수직 방향으로 bend_radius만큼
            perp_dir = cross_product(current_dir, rotation_axis)
            perp_dir = normalize_vector(perp_dir)

            arc_center = add_points(current_pos, scale_vector(perp_dir, bend_radius))

            # 호 시작 방향
            start_from_center = scale_vector(perp_dir, -1)

            # 호 생성
            arc_points = generate_arc_points(arc_center, bend_radius, start_from_center,
                                            rotation_axis, bend_angle)
            points.extend(arc_points[1:])
            current_pos = arc_points[-1]

            # 방향 업데이트
            current_dir = rotate_vector(current_dir, rotation_axis, bend_angle)

        # Bending 후 직진
        if bend_straight > 0:
            next_pos = add_points(current_pos, scale_vector(current_dir, bend_straight))
            points.append(next_pos)
            current_pos = next_pos

    return points


def create_polyline(oEditor, points, name):
    """점들을 연결하여 Polyline 생성"""
    if len(points) < 2:
        return

    # PolylinePoints 생성 (각 점을 개별 리스트로)
    point_list = ["NAME:PolylinePoints"]
    for i, pt in enumerate(points):
        point_list.append([
            "NAME:PLPoint",
            "X:=", "{}mm".format(pt[0]),
            "Y:=", "{}mm".format(pt[1]),
            "Z:=", "{}mm".format(pt[2])
        ])

    # PolylineSegments 생성 (각 세그먼트를 개별 리스트로)
    segment_list = ["NAME:PolylineSegments"]
    for i in range(len(points) - 1):
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


def create_leadpaths_from_csv(csv_file_path, name_prefix="LeadPath"):
    """CSV 파일에서 Lead Path들을 생성"""
    paths = read_leadpath_csv(csv_file_path)

    if not paths:
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
            # 경로 점들 생성
            points = generate_path_points(path_data)

            # Polyline 생성
            create_polyline(oEditor, points, path_name)

        except Exception as e:
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
