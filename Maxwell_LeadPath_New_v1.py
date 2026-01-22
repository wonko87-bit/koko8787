# -*- coding: utf-8 -*-
"""
Maxwell 3D - Lead Path 생성 v1 (Line + AngularArc 방식)
원통좌표계에서 시작점을 입력받아 연속적인 경로를 생성
직선은 Line, 곡선은 AngularArc로 생성 후 Unite
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


def dot_product(v1, v2):
    """벡터 내적"""
    return v1[0]*v2[0] + v1[1]*v2[1] + v1[2]*v2[2]


def rotate_vector(v, axis, angle_deg):
    """축을 중심으로 벡터 회전 (Rodrigues' rotation formula)"""
    angle_rad = math.radians(angle_deg)
    k = normalize_vector(axis)
    cos_a = math.cos(angle_rad)
    sin_a = math.sin(angle_rad)

    # v_rot = v*cos(a) + (k × v)*sin(a) + k*(k·v)*(1-cos(a))
    dot_kv = dot_product(k, v)
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


def vector_angle_2d(v):
    """2D 벡터의 각도 (도 단위, XY 평면)"""
    return math.degrees(math.atan2(v[1], v[0]))


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
    """경로 데이터로부터 Line과 AngularArc를 생성하여 Unite"""

    parts = []  # 생성된 파트 이름들
    part_counter = 0

    # 시작점 (원통좌표 → 직교좌표)
    start_r = path_data['start_r']
    start_theta = path_data['start_theta']
    start_z = path_data['start_z']

    current_pos = cylindrical_to_cartesian(start_r, start_theta, start_z)

    # 디버깅: 시작점 출력
    print("Start pos: ({}, {}, {})".format(current_pos[0], current_pos[1], current_pos[2]))

    # 현재 방향 (초기: theta 방향)
    current_dir = get_theta_direction(start_theta)
    print("Start dir: ({}, {}, {})".format(current_dir[0], current_dir[1], current_dir[2]))

    # 초기 직진
    initial_straight = path_data['initial_straight']
    print("Initial straight: {}".format(initial_straight))

    if initial_straight > 0:
        part_counter += 1
        line_name = "{}_Part{}".format(path_name, part_counter)

        next_pos = add_points(current_pos, scale_vector(current_dir, initial_straight))
        print("Creating line from {} to {}".format(current_pos, next_pos))
        create_line(oEditor, current_pos, next_pos, line_name)
        parts.append(line_name)

        current_pos = next_pos

    # 각 Bending 처리 - TODO: Arc 생성 재구현 필요
    print("Number of bendings: {} (currently disabled for debugging)".format(len(path_data['bendings'])))

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
