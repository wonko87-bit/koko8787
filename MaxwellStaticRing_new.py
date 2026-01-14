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


def log(msg):
    """메시지 출력 (Maxwell 메시지 매니저에 표시)"""
    AddMessage(msg)
    print(msg)  # 백업용


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
    try:
        vertices = oEditor.GetVertexIDsFromObject(obj_name)
        log("    모든 vertex에 fillet 적용: {} 개, r={}mm".format(len(vertices), radius))

        # 모든 vertex ID를 정수 리스트로 변환
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
        log("    Fillet 적용 완료")
        return True
    except Exception as e:
        log("    [오류] Fillet 적용 실패: {}".format(str(e)))
        import traceback
        traceback.print_exc()
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


def create_static_ring(oEditor, ring_data, name, x_offset=0.0):
    """Static Ring 생성 - 단순화 버전

    단계:
    1. 큰 직사각형 생성
    2. 모든 vertex에 2mm fillet 한번에 적용
    3. 작은 직사각형 생성
    4. 모든 vertex에 2mm fillet 한번에 적용
    5. Subtract
    """
    ref_x = ring_data['ref_x'] + x_offset
    ref_z = ring_data['ref_z']
    thickness = ring_data['thickness']
    width = ring_data['width']
    height = ring_data['height']

    fillet_radius = 2.0  # 고정값

    log("  기준점: ({}, 0, {})".format(ref_x, ref_z))
    log("  큰 직사각형: W={}, H={}".format(width, height))
    log("  두께: {}".format(thickness))
    log("  모든 모서리 Fillet: {}mm".format(fillet_radius))

    # 1. 큰 직사각형 생성
    outer_rect_name = name + "_Outer"
    create_rectangle_xz(oEditor, ref_x, ref_z, width, height, outer_rect_name)
    log("  [1] 큰 직사각형 생성 완료")

    # 2. 외부 직사각형 모든 vertex에 fillet 적용
    log("  [2] 외부 직사각형 fillet 적용 중...")
    fillet_all_vertices(oEditor, outer_rect_name, fillet_radius)

    # 3. 작은 직사각형 생성
    inner_width = width - 2 * thickness
    inner_height = height - 2 * thickness
    inner_x = ref_x + thickness
    inner_z = ref_z + thickness

    if inner_width <= 0 or inner_height <= 0:
        log("  [오류] 두께가 너무 커서 내부 직사각형을 만들 수 없습니다.")
        return

    inner_rect_name = name + "_Inner"
    create_rectangle_xz(oEditor, inner_x, inner_z, inner_width, inner_height, inner_rect_name)
    log("  [3] 작은 직사각형 생성 완료: W={}, H={}".format(inner_width, inner_height))

    # 4. 내부 직사각형 모든 vertex에 fillet 적용
    log("  [4] 내부 직사각형 fillet 적용 중...")
    fillet_all_vertices(oEditor, inner_rect_name, fillet_radius)

    # 5. Subtract
    log("  [5] Subtract 수행 중...")
    subtract_objects(oEditor, outer_rect_name, inner_rect_name)
    log("  [5] Subtract 완료")

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
        log("  완성: {}".format(name))
    except:
        log("  완성: {} (이름 변경 실패)".format(outer_rect_name))


def create_staticrings_from_csv(csv_file_path, name_prefix="StaticRing"):
    """CSV 파일에서 Static Ring들을 생성"""
    log("=" * 60)
    log("Maxwell 3D - Static Ring 생성")
    log("=" * 60)

    log("\nCSV 파일 읽기 시작...")
    rings, x_offset = read_staticring_csv(csv_file_path)

    log("읽은 Static Ring 수: {}".format(len(rings)))

    if not rings:
        log("오류: CSV 파일에서 Static Ring 데이터를 읽을 수 없습니다.")
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
log("\n" + "=" * 60)
log("스크립트 실행 시작")
log("=" * 60)

try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log("스크립트 디렉토리 (__file__): {}".format(script_dir))
except:
    import sys
    if len(sys.argv) > 0:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        log("스크립트 디렉토리 (sys.argv[0]): {}".format(script_dir))
    else:
        script_dir = os.getcwd()
        log("스크립트 디렉토리 (getcwd): {}".format(script_dir))

csv_file = os.path.join(script_dir, "StaticRingDim.csv")

log("\nCSV 파일 경로: {}".format(csv_file))
log("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    log("\n경고: StaticRingDim.csv 파일을 찾을 수 없습니다!")
    log("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    log("\nStatic Ring 생성 함수 호출 중...")
    create_staticrings_from_csv(
        csv_file_path=csv_file,
        name_prefix="StaticRing"
    )
    log("\n스크립트 실행 완료!")
