# -*- coding: utf-8 -*-
"""
Maxwell 3D - Filleted Rectangle 생성 (Boolean 연산)
큰 직사각형(fillet) - 작은 직사각형 = 프레임 형태
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_frame_csv(csv_file_path):
    """CSV 파일에서 프레임 데이터를 읽어옵니다.

    변수 순서 (A4~I4):
    1. X 기준 좌표 (큰 직사각형 좌하단 X)
    2. Z 기준 좌표 (큰 직사각형 좌하단 Z)
    3. 두께 (큰 직사각형과 작은 직사각형 사이 거리)
    4. 큰 직사각형 X축 방향 길이
    5. 큰 직사각형 Z축 방향 길이
    6. 1사분면 꼭지점 fillet radius (우상단)
    7. 2사분면 꼭지점 fillet radius (좌상단)
    8. 3사분면 꼭지점 fillet radius (좌하단)
    9. 4사분면 꼭지점 fillet radius (우하단)
    """
    frames = []

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

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
            fillet_q1 = float(row[5].strip()) if row[5].strip() else None
            fillet_q2 = float(row[6].strip()) if row[6].strip() else None
            fillet_q3 = float(row[7].strip()) if row[7].strip() else None
            fillet_q4 = float(row[8].strip()) if row[8].strip() else None

            # 데이터 유효성 검사
            if None in [ref_x, ref_z, thickness, width, height, fillet_q1, fillet_q2, fillet_q3, fillet_q4]:
                print("경고: {}행에 빈 데이터가 있습니다.".format(i+1))
                continue

            frames.append({
                'ref_x': ref_x,
                'ref_z': ref_z,
                'thickness': thickness,
                'width': width,
                'height': height,
                'fillet_q1': fillet_q1,  # 우상단
                'fillet_q2': fillet_q2,  # 좌상단
                'fillet_q3': fillet_q3,  # 좌하단
                'fillet_q4': fillet_q4,  # 우하단
                'row_num': i + 1
            })

        except (ValueError, IndexError) as e:
            print("경고: {}행 데이터 읽기 오류: {}".format(i+1, str(e)))
            continue

    return frames


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


def fillet_edges(oEditor, obj_name, edge_ids, radius):
    """엣지에 fillet 적용"""
    if not edge_ids:
        return

    edge_list = ",".join([str(eid) for eid in edge_ids])

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
                "Edges:=", [int(eid) for eid in edge_ids],
                "Vertices:=", [],
                "Radius:=", "{}mm".format(radius),
                "Setback:=", "0mm"
            ]
        ]
    )


def get_corner_edge_id(oEditor, obj_name, corner_x, corner_z, tolerance=0.1):
    """특정 코너(꼭지점)에 인접한 수직 엣지 ID 찾기

    XZ 평면 직사각형의 경우, Y방향(수직) 엣지를 찾아야 함
    """
    # 모든 엣지 가져오기
    edges = oEditor.GetEdgeIDsFromObject(obj_name)

    for edge_id in edges:
        # 엣지의 중점 좌표 가져오기
        edge_center = oEditor.GetEdgePositionAtParameter(obj_name, edge_id, 0.5)

        # XZ 좌표 확인 (Y는 무시)
        x_pos = edge_center[0]
        z_pos = edge_center[2]

        # 코너 근처 엣지인지 확인
        if abs(x_pos - corner_x) < tolerance and abs(z_pos - corner_z) < tolerance:
            return edge_id

    return None


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


def create_filleted_frame(oEditor, frame_data, name):
    """Filleted 프레임 생성

    단계:
    1. 큰 직사각형 생성
    2. 4개 꼭지점에 fillet 적용
    3. 작은 직사각형 생성
    4. Subtract (큰 직사각형 - 작은 직사각형)
    """

    ref_x = frame_data['ref_x']
    ref_z = frame_data['ref_z']
    thickness = frame_data['thickness']
    width = frame_data['width']
    height = frame_data['height']
    fillet_q1 = frame_data['fillet_q1']
    fillet_q2 = frame_data['fillet_q2']
    fillet_q3 = frame_data['fillet_q3']
    fillet_q4 = frame_data['fillet_q4']

    print("  기준점: ({}, 0, {})".format(ref_x, ref_z))
    print("  큰 직사각형: W={}, H={}".format(width, height))
    print("  두께: {}".format(thickness))
    print("  Fillet: Q1={}, Q2={}, Q3={}, Q4={}".format(fillet_q1, fillet_q2, fillet_q3, fillet_q4))

    # 1. 큰 직사각형 생성
    outer_rect_name = name + "_Outer"
    create_rectangle_xz(oEditor, ref_x, ref_z, width, height, outer_rect_name)
    print("  [1] 큰 직사각형 생성")

    # 2. 4개 꼭지점 좌표 계산 (중점 기준 사분면)
    # 좌하단 기준으로 계산
    corner_q1 = (ref_x + width, ref_z + height)   # 우상단 (1사분면)
    corner_q2 = (ref_x, ref_z + height)           # 좌상단 (2사분면)
    corner_q3 = (ref_x, ref_z)                     # 좌하단 (3사분면)
    corner_q4 = (ref_x + width, ref_z)            # 우하단 (4사분면)

    # 3. Fillet 적용 (각 꼭지점에 개별적으로)
    print("  [2] Fillet 적용 중...")

    # Maxwell API에서 Fillet을 적용하려면 EdgeID를 알아야 함
    # 직사각형의 경우 4개 엣지가 있고, 각 코너는 2개 엣지의 만남
    # 여기서는 FaceFillet 또는 전체 선택 방식 사용

    # 대안: Polyline으로 fillet된 직사각형을 직접 그리거나
    # 또는 각 코너별로 순차적으로 fillet 적용

    # 간단한 방법: 한 번에 모든 엣지에 동일한 radius로 fillet
    # 하지만 요구사항은 각 코너마다 다른 radius...

    # Maxwell API 한계로 인해 각 코너별로 다른 fillet을 적용하려면
    # 엣지 선택이 필요합니다. 일단 시도해봅시다.

    try:
        # 모든 엣지 ID 가져오기
        all_edges = oEditor.GetEdgeIDsFromObject(outer_rect_name)
        print("  총 엣지 수: {}".format(len(all_edges)))

        # 각 코너에 fillet 적용 (순차적으로)
        # Q3 (좌하단) - fillet_q3
        if fillet_q3 > 0:
            edge_id = get_corner_edge_id(oEditor, outer_rect_name, corner_q3[0], corner_q3[1])
            if edge_id:
                fillet_edges(oEditor, outer_rect_name, [edge_id], fillet_q3)
                print("  Q3 코너 fillet 적용: r={}".format(fillet_q3))

        # Q4 (우하단) - fillet_q4
        if fillet_q4 > 0:
            edge_id = get_corner_edge_id(oEditor, outer_rect_name, corner_q4[0], corner_q4[1])
            if edge_id:
                fillet_edges(oEditor, outer_rect_name, [edge_id], fillet_q4)
                print("  Q4 코너 fillet 적용: r={}".format(fillet_q4))

        # Q2 (좌상단) - fillet_q2
        if fillet_q2 > 0:
            edge_id = get_corner_edge_id(oEditor, outer_rect_name, corner_q2[0], corner_q2[1])
            if edge_id:
                fillet_edges(oEditor, outer_rect_name, [edge_id], fillet_q2)
                print("  Q2 코너 fillet 적용: r={}".format(fillet_q2))

        # Q1 (우상단) - fillet_q1
        if fillet_q1 > 0:
            edge_id = get_corner_edge_id(oEditor, outer_rect_name, corner_q1[0], corner_q1[1])
            if edge_id:
                fillet_edges(oEditor, outer_rect_name, [edge_id], fillet_q1)
                print("  Q1 코너 fillet 적용: r={}".format(fillet_q1))

    except Exception as e:
        print("  [경고] Fillet 적용 중 오류: {}".format(str(e)))
        print("  Fillet 없이 계속 진행합니다.")

    # 4. 작은 직사각형 생성
    inner_width = width - 2 * thickness
    inner_height = height - 2 * thickness
    inner_x = ref_x + thickness
    inner_z = ref_z + thickness

    if inner_width <= 0 or inner_height <= 0:
        print("  [오류] 두께가 너무 커서 내부 직사각형을 만들 수 없습니다.")
        return

    inner_rect_name = name + "_Inner"
    create_rectangle_xz(oEditor, inner_x, inner_z, inner_width, inner_height, inner_rect_name)
    print("  [3] 작은 직사각형 생성: W={}, H={}".format(inner_width, inner_height))

    # 5. Subtract
    subtract_objects(oEditor, outer_rect_name, inner_rect_name)
    print("  [4] Subtract: 프레임 완성")

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


def create_frames_from_csv(csv_file_path, name_prefix="Frame"):
    """CSV 파일에서 프레임들을 생성"""
    print("=" * 60)
    print("Maxwell 3D - Filleted Frame 생성")
    print("=" * 60)

    print("\nCSV 파일 읽기 시작...")
    frames = read_frame_csv(csv_file_path)

    print("읽은 프레임 수: {}".format(len(frames)))

    if not frames:
        print("오류: CSV 파일에서 프레임 데이터를 읽을 수 없습니다.")
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

    # 각 프레임 생성
    success_count = 0
    fail_count = 0

    for idx, frame in enumerate(frames):
        print("\n--- 프레임 {} (CSV {}행) ---".format(idx + 1, frame['row_num']))

        frame_name = "{}_{}".format(name_prefix, idx + 1)

        try:
            create_filleted_frame(oEditor, frame, frame_name)
            success_count += 1
        except Exception as e:
            print("  [오류] 프레임 {} 생성 실패: {}".format(idx + 1, str(e)))
            import traceback
            traceback.print_exc()
            print("  프레임 {}를 건너뛰고 계속 진행합니다.".format(idx + 1))
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
    print("  전체: {} 개".format(len(frames)))
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

csv_file = os.path.join(script_dir, "FrameDim.csv")

print("\nCSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    print("\n경고: FrameDim.csv 파일을 찾을 수 없습니다!")
    print("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    print("\n프레임 생성 함수 호출 중...")
    create_frames_from_csv(
        csv_file_path=csv_file,
        name_prefix="Frame"
    )
    print("\n스크립트 실행 완료!")
