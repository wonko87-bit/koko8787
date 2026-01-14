# -*- coding: utf-8 -*-
"""
Maxwell 3D - Static Ring 생성 (단순화 버전)
Rectangle + 모든 vertex 한번에 fillet
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_staticring_csv(csv_file_path):
    """CSV 파일에서 Static Ring 데이터를 읽어옵니다.

    변수 순서 (A4~I4):
    1. X 기준 좌표 (큰 직사각형 좌하단 X)
    2. Z 기준 좌표 (큰 직사각형 좌하단 Z)
    3. 두께 (큰 직사각형과 작은 직사각형 사이 거리)
    4. 큰 직사각형 X축 방향 길이
    5. 큰 직사각형 Z축 방향 길이
    6. 1사분면 꼭지점 내부 fillet radius (우상단)
    7. 2사분면 꼭지점 내부 fillet radius (좌상단)
    8. 3사분면 꼭지점 내부 fillet radius (좌하단)
    9. 4사분면 꼭지점 내부 fillet radius (우하단)
    """
    rings = []
    x_offset = 0.0

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

    # B2 셀에서 X축 offset 읽기 (행 인덱스 1, 열 인덱스 1)
    try:
        if len(rows) > 1 and len(rows[1]) > 1 and rows[1][1].strip():
            x_offset = float(rows[1][1].strip())
            print("X축 offset: {}mm".format(x_offset))
    except (ValueError, IndexError):
        print("B2 셀에서 X축 offset을 읽을 수 없습니다. 기본값 0 사용")
        x_offset = 0.0

    # 4행부터 시작 (인덱스 3)
    for i in range(3, len(rows)):
        row = rows[i]

        # 빈 행 체크
        if not row or all(not cell.strip() for cell in row):
            break

        try:
            if len(row) < 9:
                print("경고: {}행에 데이터가 부족합니다. (9개 필요, {}개 있음)".format(i+1, len(row)))
                continue

            ref_x = float(row[0].strip()) if row[0].strip() else None
            ref_z = float(row[1].strip()) if row[1].strip() else None
            thickness = float(row[2].strip()) if row[2].strip() else None
            width = float(row[3].strip()) if row[3].strip() else None
            height = float(row[4].strip()) if row[4].strip() else None
            inner_fillet_q1 = float(row[5].strip()) if row[5].strip() else None
            inner_fillet_q2 = float(row[6].strip()) if row[6].strip() else None
            inner_fillet_q3 = float(row[7].strip()) if row[7].strip() else None
            inner_fillet_q4 = float(row[8].strip()) if row[8].strip() else None

            # 데이터 유효성 검사
            if None in [ref_x, ref_z, thickness, width, height, inner_fillet_q1, inner_fillet_q2, inner_fillet_q3, inner_fillet_q4]:
                print("경고: {}행에 빈 데이터가 있습니다.".format(i+1))
                continue

            rings.append({
                'ref_x': ref_x,
                'ref_z': ref_z,
                'thickness': thickness,
                'width': width,
                'height': height,
                'inner_fillet_q1': inner_fillet_q1,  # 우상단 내부
                'inner_fillet_q2': inner_fillet_q2,  # 좌상단 내부
                'inner_fillet_q3': inner_fillet_q3,  # 좌하단 내부
                'inner_fillet_q4': inner_fillet_q4,  # 우하단 내부
                'row_num': i + 1
            })

        except (ValueError, IndexError) as e:
            print("경고: {}행 데이터 읽기 오류: {}".format(i+1, str(e)))
            continue

    return rings, x_offset


def create_rectangle_xz(oEditor, x_start, z_start, width, height, name):
    """XZ 평면에 Rectangle 생성

    주의: WhichAxis="Y"일 때
    - Width = Z방향 크기
    - Height = X방향 크기
    """
    oEditor.CreateRectangle(
        [
            "NAME:RectangleParameters",
            "IsCovered:=", True,
            "XStart:=", "{}mm".format(x_start),
            "YStart:=", "0mm",
            "ZStart:=", "{}mm".format(z_start),
            "Width:=", "{}mm".format(height),   # Z방향
            "Height:=", "{}mm".format(width),   # X방향
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


def fillet_all_vertices(oEditor, obj_name, radius):
    """객체의 모든 vertex에 동일한 fillet 적용"""
    vertices = oEditor.GetVertexIDsFromObject(obj_name)
    vertex_ids = [int(vid) for vid in vertices]

    oEditor.Fillet(
        [
            "NAME:Selections",
            "Selections:=", obj_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:FilletParameters",
            [
                "NAME:FilletParameters",
                "Edges:=", [],
                "Vertices:=", vertex_ids,
                "Radius:=", "{}mm".format(radius),
                "Setback:=", "0mm"
            ]
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


def create_inner_rectangle(oEditor, ring_data, name, x_offset=0.0):
    """내부 직사각형만 생성 (fillet 미적용)"""
    ref_x = ring_data['ref_x'] + x_offset
    ref_z = ring_data['ref_z']
    thickness = ring_data['thickness']
    width = ring_data['width']
    height = ring_data['height']

    inner_width = width - 2 * thickness
    inner_height = height - 2 * thickness
    inner_x = ref_x + thickness
    inner_z = ref_z + thickness

    if inner_width <= 0 or inner_height <= 0:
        return None

    inner_rect_name = name + "_Inner"
    create_rectangle_xz(oEditor, inner_x, inner_z, inner_width, inner_height, inner_rect_name)
    return inner_rect_name


def create_outer_rectangle(oEditor, ring_data, name, x_offset=0.0):
    """외부 직사각형만 생성 (fillet 미적용)"""
    ref_x = ring_data['ref_x'] + x_offset
    ref_z = ring_data['ref_z']
    width = ring_data['width']
    height = ring_data['height']

    outer_rect_name = name + "_Outer"
    create_rectangle_xz(oEditor, ref_x, ref_z, width, height, outer_rect_name)
    return outer_rect_name


def finalize_static_ring(oEditor, outer_name, inner_name, final_name):
    """Subtract 후 이름 변경"""
    if inner_name is None:
        return

    subtract_objects(oEditor, outer_name, inner_name)

    try:
        oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Geometry3DAttributeTab",
                    [
                        "NAME:PropServers",
                        outer_name
                    ],
                    [
                        "NAME:ChangedProps",
                        [
                            "NAME:Name",
                            "Value:=", final_name
                        ]
                    ]
                ]
            ]
        )
    except:
        pass


def create_staticrings_from_csv(csv_file_path, name_prefix="StaticRing"):
    """CSV 파일에서 Static Ring들을 생성 (fillet 없이 직사각형만)"""
    rings, x_offset = read_staticring_csv(csv_file_path)

    if not rings:
        return

    # Maxwell 프로젝트 및 디자인 가져오기
    oProject = oDesktop.GetActiveProject()
    if oProject is None:
        oProject = oDesktop.NewProject()

    oDesign = oProject.GetActiveDesign()
    if oDesign is None:
        oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
        oDesign = oProject.GetActiveDesign()

    oEditor = oDesign.SetActiveEditor("3D Modeler")

    # 단계 1: 모든 내부 직사각형 생성
    inner_names = []
    for idx, ring in enumerate(rings):
        ring_name = "{}_{}".format(name_prefix, idx + 1)
        try:
            inner_name = create_inner_rectangle(oEditor, ring, ring_name, x_offset)
            inner_names.append(inner_name)
        except:
            inner_names.append(None)

    # 단계 2: 모든 외부 직사각형 생성
    outer_names = []
    for idx, ring in enumerate(rings):
        ring_name = "{}_{}".format(name_prefix, idx + 1)
        try:
            outer_name = create_outer_rectangle(oEditor, ring, ring_name, x_offset)
            outer_names.append(outer_name)
        except:
            outer_names.append(None)

    # 단계 3: 각각 subtract 및 이름 변경
    for idx in range(len(rings)):
        if outer_names[idx] and inner_names[idx]:
            try:
                ring_name = "{}_{}".format(name_prefix, idx + 1)
                finalize_static_ring(oEditor, outer_names[idx], inner_names[idx], ring_name)
            except:
                pass

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

csv_file = os.path.join(script_dir, "StaticRingDim.csv")

if os.path.exists(csv_file):
    create_staticrings_from_csv(
        csv_file_path=csv_file,
        name_prefix="StaticRing"
    )
