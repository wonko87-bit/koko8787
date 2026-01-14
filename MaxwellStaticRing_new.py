# -*- coding: utf-8 -*-
"""
Maxwell 3D - Static Ring 생성 (원 기반 사분원 방식)
각 모서리에 원을 그려 사분원으로 자르고 연결
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_staticring_csv(csv_file_path):
    """CSV 파일에서 Static Ring 데이터를 읽어옵니다."""
    rings = []
    x_offset = 0.0

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

    # B2 셀에서 X축 offset 읽기
    try:
        if len(rows) > 1 and len(rows[1]) > 1 and rows[1][1].strip():
            x_offset = float(rows[1][1].strip())
    except:
        x_offset = 0.0

    # 4행부터 시작
    for i in range(3, len(rows)):
        row = rows[i]
        if not row or all(not cell.strip() for cell in row):
            break

        try:
            if len(row) < 9:
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

            if None in [ref_x, ref_z, thickness, width, height, inner_fillet_q1, inner_fillet_q2, inner_fillet_q3, inner_fillet_q4]:
                continue

            rings.append({
                'ref_x': ref_x,
                'ref_z': ref_z,
                'thickness': thickness,
                'width': width,
                'height': height,
                'r1': inner_fillet_q1,
                'r2': inner_fillet_q2,
                'r3': inner_fillet_q3,
                'r4': inner_fillet_q4
            })
        except:
            continue

    return rings, x_offset


def create_circle_xz(oEditor, center_x, center_z, radius, name):
    """XZ 평면에 Circle 생성"""
    oEditor.CreateCircle(
        [
            "NAME:CircleParameters",
            "IsCovered:=", True,
            "XCenter:=", "{}mm".format(center_x),
            "YCenter:=", "0mm",
            "ZCenter:=", "{}mm".format(center_z),
            "Radius:=", "{}mm".format(radius),
            "WhichAxis:=", "Y",
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


def create_rectangle_xz(oEditor, x_start, z_start, x_size, z_size, name):
    """XZ 평면에 Rectangle 생성 (x_size=X방향, z_size=Z방향)"""
    oEditor.CreateRectangle(
        [
            "NAME:RectangleParameters",
            "IsCovered:=", True,
            "XStart:=", "{}mm".format(x_start),
            "YStart:=", "0mm",
            "ZStart:=", "{}mm".format(z_start),
            "Width:=", "{}mm".format(z_size),   # Z방향
            "Height:=", "{}mm".format(x_size),  # X방향
            "WhichAxis:=", "Y"
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


def intersect_objects(oEditor, blank_name, tool_name):
    """Boolean Intersect"""
    oEditor.Intersect(
        [
            "NAME:Selections",
            "Blank Parts:=", blank_name,
            "Tool Parts:=", tool_name
        ],
        [
            "NAME:IntersectParameters",
            "KeepOriginals:=", False
        ]
    )


def unite_objects(oEditor, obj_names):
    """Boolean Unite"""
    if len(obj_names) < 2:
        return
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


def subtract_objects(oEditor, blank_name, tool_name):
    """Boolean Subtract"""
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


def create_filleted_rectangle(oEditor, ref_x, ref_z, width, height, r1, r2, r3, r4, base_name):
    """
    사분원 기반 filleted rectangle 생성
    width: Z방향, height: X방향
    r1: Q1(우상), r2: Q2(좌상), r3: Q3(좌하), r4: Q4(우하)
    """
    parts = []

    # 각 꼭지점에서 내측으로 반경만큼 이동한 원의 중심
    c_q3 = (ref_x + r3, ref_z + r3)           # 좌하단
    c_q4 = (ref_x + height - r4, ref_z + r4) # 우하단
    c_q2 = (ref_x + r2, ref_z + width - r2)  # 좌상단
    c_q1 = (ref_x + height - r1, ref_z + width - r1)  # 우상단

    # Q3: 1사분면 사분원 (오른쪽 위)
    if r3 > 0:
        circ_name = base_name + "_C_Q3"
        rect_name = base_name + "_R_Q3"
        create_circle_xz(oEditor, c_q3[0], c_q3[1], r3, circ_name)
        create_rectangle_xz(oEditor, c_q3[0], c_q3[1], r3, r3, rect_name)
        intersect_objects(oEditor, circ_name, rect_name)
        parts.append(circ_name)

    # Q4: 2사분면 사분원 (왼쪽 위)
    if r4 > 0:
        circ_name = base_name + "_C_Q4"
        rect_name = base_name + "_R_Q4"
        create_circle_xz(oEditor, c_q4[0], c_q4[1], r4, circ_name)
        create_rectangle_xz(oEditor, c_q4[0] - r4, c_q4[1], r4, r4, rect_name)
        intersect_objects(oEditor, circ_name, rect_name)
        parts.append(circ_name)

    # Q2: 4사분면 사분원 (오른쪽 아래)
    if r2 > 0:
        circ_name = base_name + "_C_Q2"
        rect_name = base_name + "_R_Q2"
        create_circle_xz(oEditor, c_q2[0], c_q2[1], r2, circ_name)
        create_rectangle_xz(oEditor, c_q2[0], c_q2[1] - r2, r2, r2, rect_name)
        intersect_objects(oEditor, circ_name, rect_name)
        parts.append(circ_name)

    # Q1: 3사분면 사분원 (왼쪽 아래)
    if r1 > 0:
        circ_name = base_name + "_C_Q1"
        rect_name = base_name + "_R_Q1"
        create_circle_xz(oEditor, c_q1[0], c_q1[1], r1, circ_name)
        create_rectangle_xz(oEditor, c_q1[0] - r1, c_q1[1] - r1, r1, r1, rect_name)
        intersect_objects(oEditor, circ_name, rect_name)
        parts.append(circ_name)

    # 직선 부분 (4개 사각형)
    # 하단 (Q3와 Q4 사이)
    bottom_x_start = ref_x + r3
    bottom_x_size = height - r3 - r4
    bottom_z_start = ref_z
    bottom_z_size = min(r3, r4)
    if bottom_x_size > 0 and bottom_z_size > 0:
        rect_name = base_name + "_Bottom"
        create_rectangle_xz(oEditor, bottom_x_start, bottom_z_start, bottom_x_size, bottom_z_size, rect_name)
        parts.append(rect_name)

    # 우측 (Q4와 Q1 사이)
    right_x_start = ref_x + height - max(r4, r1)
    right_x_size = max(r4, r1)
    right_z_start = ref_z + r4
    right_z_size = width - r4 - r1
    if right_x_size > 0 and right_z_size > 0:
        rect_name = base_name + "_Right"
        create_rectangle_xz(oEditor, right_x_start, right_z_start, right_x_size, right_z_size, rect_name)
        parts.append(rect_name)

    # 상단 (Q1과 Q2 사이)
    top_x_start = ref_x + r2
    top_x_size = height - r2 - r1
    top_z_start = ref_z + width - min(r1, r2)
    top_z_size = min(r1, r2)
    if top_x_size > 0 and top_z_size > 0:
        rect_name = base_name + "_Top"
        create_rectangle_xz(oEditor, top_x_start, top_z_start, top_x_size, top_z_size, rect_name)
        parts.append(rect_name)

    # 좌측 (Q2와 Q3 사이)
    left_x_start = ref_x
    left_x_size = min(r2, r3)
    left_z_start = ref_z + r3
    left_z_size = width - r3 - r2
    if left_x_size > 0 and left_z_size > 0:
        rect_name = base_name + "_Left"
        create_rectangle_xz(oEditor, left_x_start, left_z_start, left_x_size, left_z_size, rect_name)
        parts.append(rect_name)

    # 중앙 큰 사각형
    center_x_start = ref_x + min(r2, r3)
    center_x_size = height - min(r2, r3) - min(r1, r4)
    center_z_start = ref_z + min(r3, r4)
    center_z_size = width - min(r3, r4) - min(r1, r2)
    if center_x_size > 0 and center_z_size > 0:
        rect_name = base_name + "_Center"
        create_rectangle_xz(oEditor, center_x_start, center_z_start, center_x_size, center_z_size, rect_name)
        parts.append(rect_name)

    # 모든 파트 Unite
    if len(parts) > 0:
        unite_objects(oEditor, parts)
        return parts[0]

    return None


def create_staticrings_from_csv(csv_file_path, name_prefix="StaticRing"):
    """CSV 파일에서 Static Ring들을 생성"""
    rings, x_offset = read_staticring_csv(csv_file_path)

    if not rings:
        return

    oProject = oDesktop.GetActiveProject()
    if oProject is None:
        oProject = oDesktop.NewProject()

    oDesign = oProject.GetActiveDesign()
    if oDesign is None:
        oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
        oDesign = oProject.GetActiveDesign()

    oEditor = oDesign.SetActiveEditor("3D Modeler")

    for idx, ring in enumerate(rings):
        ring_name = "{}_{}".format(name_prefix, idx + 1)

        try:
            ref_x = ring['ref_x'] + x_offset
            ref_z = ring['ref_z']
            width = ring['width']
            height = ring['height']
            thickness = ring['thickness']

            r1 = ring['r1']
            r2 = ring['r2']
            r3 = ring['r3']
            r4 = ring['r4']

            # 외부 filleted rectangle
            outer_name = ring_name + "_Outer"
            outer = create_filleted_rectangle(oEditor, ref_x, ref_z, width, height, r1, r2, r3, r4, outer_name)

            # 내부 filleted rectangle (반경 - 두께)
            inner_r1 = max(0, r1 - thickness)
            inner_r2 = max(0, r2 - thickness)
            inner_r3 = max(0, r3 - thickness)
            inner_r4 = max(0, r4 - thickness)

            inner_x = ref_x + thickness
            inner_z = ref_z + thickness
            inner_width = width - 2 * thickness
            inner_height = height - 2 * thickness

            if inner_width > 0 and inner_height > 0:
                inner_name = ring_name + "_Inner"
                inner = create_filleted_rectangle(oEditor, inner_x, inner_z, inner_width, inner_height,
                                                 inner_r1, inner_r2, inner_r3, inner_r4, inner_name)

                # Subtract
                if outer and inner:
                    subtract_objects(oEditor, outer, inner)

                    # 이름 변경
                    try:
                        oEditor.ChangeProperty(
                            [
                                "NAME:AllTabs",
                                [
                                    "NAME:Geometry3DAttributeTab",
                                    [
                                        "NAME:PropServers",
                                        outer
                                    ],
                                    [
                                        "NAME:ChangedProps",
                                        [
                                            "NAME:Name",
                                            "Value:=", ring_name
                                        ]
                                    ]
                                ]
                            ]
                        )
                    except:
                        pass
        except:
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

csv_file = os.path.join(script_dir, "StaticRingDim.csv")

if os.path.exists(csv_file):
    create_staticrings_from_csv(
        csv_file_path=csv_file,
        name_prefix="StaticRing"
    )
