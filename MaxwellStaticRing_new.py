# -*- coding: utf-8 -*-
"""
Maxwell 3D - Static Ring 생성 (Boolean 연산)
큰 직사각형(fillet) - 작은 직사각형(fillet) = 프레임 형태
내부 fillet 반경 기준으로 입력, 외부 fillet = 내부 fillet + thickness
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os
from collections import Counter


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


def create_rectangle_polyline_xz(oEditor, ref_x, ref_z, width, height, name):
    """XZ 평면에 Polyline으로 Closed Rectangle 생성

    Args:
        ref_x, ref_z: 좌하단 기준점
        width: Z 방향 크기
        height: X 방향 크기

    Points: Q3(좌하) -> Q4(우하) -> Q1(우상) -> Q2(좌상) -> Q3(close)
    """
    # 4개 코너 좌표
    q3 = (ref_x, 0, ref_z)                    # 좌하단
    q4 = (ref_x + height, 0, ref_z)          # 우하단
    q1 = (ref_x + height, 0, ref_z + width)  # 우상단
    q2 = (ref_x, 0, ref_z + width)           # 좌상단

    # PolylinePoints 생성
    point_list = []
    for pt in [q3, q4, q1, q2]:
        point_list.append([
            "NAME:PLPoint",
            "X:=", "{}mm".format(pt[0]),
            "Y:=", "{}mm".format(pt[1]),
            "Z:=", "{}mm".format(pt[2])
        ])

    # PolylineSegments 생성 (모두 Line)
    segment_list = []
    for i in range(4):
        segment_list.append([
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
                "NAME:PolylinePoints",
                point_list
            ],
            [
                "NAME:PolylineSegments",
                segment_list
            ],
            [
                "NAME:PolylineXSection",
                "XSectionType:=", "None"
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


def get_vertex_at_position(oEditor, obj_name, target_x, target_z, tolerance=1.0):
    """특정 위치의 vertex ID 찾기"""
    try:
        vertices = oEditor.GetVertexIDsFromObject(obj_name)
        print("    찾은 vertex 수: {}".format(len(vertices)))

        for vid in vertices:
            pos = oEditor.GetVertexPosition(obj_name, vid)
            x_pos = pos[0]
            y_pos = pos[1]
            z_pos = pos[2]

            print("      Vertex {}: ({:.2f}, {:.2f}, {:.2f})".format(vid, x_pos, y_pos, z_pos))

            # XZ 평면이므로 Y≈0이고, X, Z가 목표 좌표와 일치
            if abs(y_pos) < 0.5 and abs(x_pos - target_x) < tolerance and abs(z_pos - target_z) < tolerance:
                print("    -> 목표 좌표 ({:.2f}, {:.2f})와 일치! Vertex ID={}".format(target_x, target_z, vid))
                return vid

        print("    [경고] 목표 좌표 ({:.2f}, {:.2f})에 해당하는 vertex를 찾지 못했습니다.".format(target_x, target_z))
    except Exception as e:
        print("    [오류] Vertex 찾기 실패: {}".format(str(e)))
        import traceback
        traceback.print_exc()
    return None


def fillet_vertex(oEditor, obj_name, vertex_id, radius):
    """특정 vertex에 fillet 적용"""
    if radius <= 0 or vertex_id is None:
        return False

    try:
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
                    "Vertices:=", [int(vertex_id)],
                    "Radius:=", "{}mm".format(radius),
                    "Setback:=", "0mm"
                ]
            ]
        )
        print("    Fillet 적용: vertex={}, r={}mm".format(vertex_id, radius))
        return True
    except Exception as e:
        print("    [경고] Fillet 적용 실패: vertex={}, r={}, 오류={}".format(vertex_id, radius, str(e)))
        return False


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


def get_common_fillet_value(fillet_q1, fillet_q2, fillet_q3, fillet_q4):
    """4개 fillet 값 중 가장 많이 나타나는 값 반환 (3개 이상 같은 값)"""
    fillets = [fillet_q1, fillet_q2, fillet_q3, fillet_q4]
    counter = Counter(fillets)
    most_common = counter.most_common(1)[0]

    # 가장 많이 나타나는 값이 3개 이상이면 그 값 반환
    if most_common[1] >= 3:
        return most_common[0]
    return None


def create_static_ring(oEditor, ring_data, name, x_offset=0.0):
    """Static Ring 생성

    단계:
    1. 큰 직사각형 생성
    2. 외부 fillet 적용 (내부 fillet + thickness)
    3. 작은 직사각형 생성
    4. 내부 fillet 적용
    5. Subtract (큰 직사각형 - 작은 직사각형)

    주의: 4개 fillet 값 중 3개가 같고 1개가 다른 경우,
          다른 1개는 무시하고 같은 3개만 적용
    """

    ref_x = ring_data['ref_x'] + x_offset  # X축 offset 적용
    ref_z = ring_data['ref_z']
    thickness = ring_data['thickness']
    width = ring_data['width']
    height = ring_data['height']

    # 내부 fillet 반경 (CSV 입력값)
    inner_fillet_q1 = ring_data['inner_fillet_q1']
    inner_fillet_q2 = ring_data['inner_fillet_q2']
    inner_fillet_q3 = ring_data['inner_fillet_q3']
    inner_fillet_q4 = ring_data['inner_fillet_q4']

    # 외부 fillet 반경 = 내부 fillet + thickness
    outer_fillet_q1 = inner_fillet_q1 + thickness
    outer_fillet_q2 = inner_fillet_q2 + thickness
    outer_fillet_q3 = inner_fillet_q3 + thickness
    outer_fillet_q4 = inner_fillet_q4 + thickness

    print("  기준점: ({}, 0, {})".format(ref_x, ref_z))
    print("  큰 직사각형: W={}, H={}".format(width, height))
    print("  두께: {}".format(thickness))
    print("  내부 Fillet: Q1={}, Q2={}, Q3={}, Q4={}".format(
        inner_fillet_q1, inner_fillet_q2, inner_fillet_q3, inner_fillet_q4))
    print("  외부 Fillet: Q1={}, Q2={}, Q3={}, Q4={}".format(
        outer_fillet_q1, outer_fillet_q2, outer_fillet_q3, outer_fillet_q4))

    # 1. 큰 직사각형 생성 (Polyline)
    outer_rect_name = name + "_Outer"
    create_rectangle_polyline_xz(oEditor, ref_x, ref_z, width, height, outer_rect_name)
    print("  [1] 큰 직사각형 생성 (Polyline)")

    # 2. 외부 fillet 적용 (Polyline 순서: Q3->Q4->Q1->Q2)
    # 순차적으로: Q4 -> Q1 -> Q2 -> Q3 (마지막은 vertex 탐색)
    print("  [2] 외부 Fillet 적용 중 (순차적)...")

    # 4개 fillet 값 중 가장 많이 나타나는 값 찾기 (3개 이상)
    outer_common = get_common_fillet_value(outer_fillet_q1, outer_fillet_q2, outer_fillet_q3, outer_fillet_q4)
    if outer_common is not None:
        print("    가장 많이 나타나는 외부 fillet 값: {}mm (나머지는 무시)".format(outer_common))

    # 4개 코너 좌표 계산
    corner_q1 = (ref_x + height, ref_z + width)  # 우상단
    corner_q2 = (ref_x, ref_z + width)           # 좌상단
    corner_q3 = (ref_x, ref_z)                   # 좌하단
    corner_q4 = (ref_x + height, ref_z)          # 우하단

    # 순차 적용: Q4 -> Q1 -> Q2 -> Q3
    # Q4 (우하단) - 첫 번째
    if outer_fillet_q4 > 0 and (outer_common is None or outer_fillet_q4 == outer_common):
        vid = get_vertex_at_position(oEditor, outer_rect_name, corner_q4[0], corner_q4[1])
        if vid:
            fillet_vertex(oEditor, outer_rect_name, vid, outer_fillet_q4)
            print("    Q4 fillet 완료")
        else:
            print("    [경고] Q4 코너 vertex를 찾을 수 없습니다.")
    elif outer_common is not None and outer_fillet_q4 != outer_common:
        print("    Q4 fillet 스킵 (값이 다름: {} != {})".format(outer_fillet_q4, outer_common))

    # Q1 (우상단) - 두 번째
    if outer_fillet_q1 > 0 and (outer_common is None or outer_fillet_q1 == outer_common):
        vid = get_vertex_at_position(oEditor, outer_rect_name, corner_q1[0], corner_q1[1])
        if vid:
            fillet_vertex(oEditor, outer_rect_name, vid, outer_fillet_q1)
            print("    Q1 fillet 완료")
        else:
            print("    [경고] Q1 코너 vertex를 찾을 수 없습니다.")
    elif outer_common is not None and outer_fillet_q1 != outer_common:
        print("    Q1 fillet 스킵 (값이 다름: {} != {})".format(outer_fillet_q1, outer_common))

    # Q2 (좌상단) - 세 번째
    if outer_fillet_q2 > 0 and (outer_common is None or outer_fillet_q2 == outer_common):
        vid = get_vertex_at_position(oEditor, outer_rect_name, corner_q2[0], corner_q2[1])
        if vid:
            fillet_vertex(oEditor, outer_rect_name, vid, outer_fillet_q2)
            print("    Q2 fillet 완료")
        else:
            print("    [경고] Q2 코너 vertex를 찾을 수 없습니다.")
    elif outer_common is not None and outer_fillet_q2 != outer_common:
        print("    Q2 fillet 스킵 (값이 다름: {} != {})".format(outer_fillet_q2, outer_common))

    # Q3 (좌하단) - 마지막, vertex 탐색
    if outer_fillet_q3 > 0 and (outer_common is None or outer_fillet_q3 == outer_common):
        vid = get_vertex_at_position(oEditor, outer_rect_name, corner_q3[0], corner_q3[1])
        if vid:
            fillet_vertex(oEditor, outer_rect_name, vid, outer_fillet_q3)
            print("    Q3 fillet 완료")
        else:
            print("    [경고] Q3 코너 vertex를 찾을 수 없습니다.")
    elif outer_common is not None and outer_fillet_q3 != outer_common:
        print("    Q3 fillet 스킵 (값이 다름: {} != {})".format(outer_fillet_q3, outer_common))

    # 3. 작은 직사각형 생성
    inner_width = width - 2 * thickness
    inner_height = height - 2 * thickness
    inner_x = ref_x + thickness
    inner_z = ref_z + thickness

    if inner_width <= 0 or inner_height <= 0:
        print("  [오류] 두께가 너무 커서 내부 직사각형을 만들 수 없습니다.")
        return

    inner_rect_name = name + "_Inner"
    create_rectangle_polyline_xz(oEditor, inner_x, inner_z, inner_width, inner_height, inner_rect_name)
    print("  [3] 작은 직사각형 생성 (Polyline): W={}, H={}".format(inner_width, inner_height))

    # 4. 내부 fillet 적용 (순차적)
    print("  [4] 내부 Fillet 적용 중 (순차적)...")

    # 4개 fillet 값 중 가장 많이 나타나는 값 찾기 (3개 이상)
    inner_common = get_common_fillet_value(inner_fillet_q1, inner_fillet_q2, inner_fillet_q3, inner_fillet_q4)
    if inner_common is not None:
        print("    가장 많이 나타나는 내부 fillet 값: {}mm (나머지는 무시)".format(inner_common))

    # 내부 직사각형 4개 코너 좌표
    inner_corner_q1 = (inner_x + inner_height, inner_z + inner_width)  # 우상단
    inner_corner_q2 = (inner_x, inner_z + inner_width)                 # 좌상단
    inner_corner_q3 = (inner_x, inner_z)                               # 좌하단
    inner_corner_q4 = (inner_x + inner_height, inner_z)                # 우하단

    # 순차 적용: Q4 -> Q1 -> Q2 -> Q3
    # Q4 (우하단) - 첫 번째
    if inner_fillet_q4 > 0 and (inner_common is None or inner_fillet_q4 == inner_common):
        vid = get_vertex_at_position(oEditor, inner_rect_name, inner_corner_q4[0], inner_corner_q4[1])
        if vid:
            fillet_vertex(oEditor, inner_rect_name, vid, inner_fillet_q4)
            print("    내부 Q4 fillet 완료")
        else:
            print("    [경고] 내부 Q4 코너 vertex를 찾을 수 없습니다.")
    elif inner_common is not None and inner_fillet_q4 != inner_common:
        print("    내부 Q4 fillet 스킵 (값이 다름: {} != {})".format(inner_fillet_q4, inner_common))

    # Q1 (우상단) - 두 번째
    if inner_fillet_q1 > 0 and (inner_common is None or inner_fillet_q1 == inner_common):
        vid = get_vertex_at_position(oEditor, inner_rect_name, inner_corner_q1[0], inner_corner_q1[1])
        if vid:
            fillet_vertex(oEditor, inner_rect_name, vid, inner_fillet_q1)
            print("    내부 Q1 fillet 완료")
        else:
            print("    [경고] 내부 Q1 코너 vertex를 찾을 수 없습니다.")
    elif inner_common is not None and inner_fillet_q1 != inner_common:
        print("    내부 Q1 fillet 스킵 (값이 다름: {} != {})".format(inner_fillet_q1, inner_common))

    # Q2 (좌상단) - 세 번째
    if inner_fillet_q2 > 0 and (inner_common is None or inner_fillet_q2 == inner_common):
        vid = get_vertex_at_position(oEditor, inner_rect_name, inner_corner_q2[0], inner_corner_q2[1])
        if vid:
            fillet_vertex(oEditor, inner_rect_name, vid, inner_fillet_q2)
            print("    내부 Q2 fillet 완료")
        else:
            print("    [경고] 내부 Q2 코너 vertex를 찾을 수 없습니다.")
    elif inner_common is not None and inner_fillet_q2 != inner_common:
        print("    내부 Q2 fillet 스킵 (값이 다름: {} != {})".format(inner_fillet_q2, inner_common))

    # Q3 (좌하단) - 마지막, vertex 탐색
    if inner_fillet_q3 > 0 and (inner_common is None or inner_fillet_q3 == inner_common):
        vid = get_vertex_at_position(oEditor, inner_rect_name, inner_corner_q3[0], inner_corner_q3[1])
        if vid:
            fillet_vertex(oEditor, inner_rect_name, vid, inner_fillet_q3)
            print("    내부 Q3 fillet 완료")
        else:
            print("    [경고] 내부 Q3 코너 vertex를 찾을 수 없습니다.")
    elif inner_common is not None and inner_fillet_q3 != inner_common:
        print("    내부 Q3 fillet 스킵 (값이 다름: {} != {})".format(inner_fillet_q3, inner_common))

    # 5. Subtract
    subtract_objects(oEditor, outer_rect_name, inner_rect_name)
    print("  [5] Subtract: Static Ring 완성")

    # 최종 이름 변경
    try:
        oEditor.ChangeProperty(
            [
                "NAME:AllTabs",
                [
                    "NAME:Geometry3DAttributeTab",
                    [
                        "NAME:PropServers",
                        outer_rect_name
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
    except:
        print("  완성: {} (이름 변경 실패)".format(outer_rect_name))


def create_staticrings_from_csv(csv_file_path, name_prefix="StaticRing"):
    """CSV 파일에서 Static Ring들을 생성"""
    print("=" * 60)
    print("Maxwell 3D - Static Ring 생성")
    print("=" * 60)

    print("\nCSV 파일 읽기 시작...")
    rings, x_offset = read_staticring_csv(csv_file_path)

    print("읽은 Static Ring 수: {}".format(len(rings)))

    if not rings:
        print("오류: CSV 파일에서 Static Ring 데이터를 읽을 수 없습니다.")
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

    # 각 Static Ring 생성
    success_count = 0
    fail_count = 0

    for idx, ring in enumerate(rings):
        print("\n--- Static Ring {} (CSV {}행) ---".format(idx + 1, ring['row_num']))

        ring_name = "{}_{}".format(name_prefix, idx + 1)

        try:
            create_static_ring(oEditor, ring, ring_name, x_offset)
            success_count += 1
        except Exception as e:
            print("  [오류] Static Ring {} 생성 실패: {}".format(idx + 1, str(e)))
            import traceback
            traceback.print_exc()
            print("  Static Ring {}를 건너뛰고 계속 진행합니다.".format(idx + 1))
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
    print("  전체: {} 개".format(len(rings)))
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

csv_file = os.path.join(script_dir, "StaticRingDim.csv")

print("\nCSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    print("\n경고: StaticRingDim.csv 파일을 찾을 수 없습니다!")
    print("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    print("\nStatic Ring 생성 함수 호출 중...")
    create_staticrings_from_csv(
        csv_file_path=csv_file,
        name_prefix="StaticRing"
    )
    print("\n스크립트 실행 완료!")
