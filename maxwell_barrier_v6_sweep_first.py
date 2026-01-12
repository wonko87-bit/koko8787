# -*- coding: utf-8 -*-
"""
Maxwell 3D - 배리어 생성 V6 (Sweep 후 Subtract 방식)
순서 변경: Circle 생성 → Sweep → Subtract
로그 파일 출력 추가
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os
import time

# 로그 파일 설정
log_file = None

def log(msg):
    """로그 메시지 출력 (파일 + 콘솔 + Maxwell)"""
    global log_file

    # 파일에 쓰기
    if log_file:
        log_file.write(msg + "\n")
        log_file.flush()

    # 콘솔 출력
    print(msg)

    # Maxwell Message Manager에 출력
    try:
        oDesktop.AddMessage("", "", 0, msg)
    except:
        pass


def read_csv_data(csv_file_path):
    """CSV 파일에서 배리어 데이터를 읽어옵니다."""
    groups = []

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

    # 그룹 감지 (빈 행으로 구분)
    i = 0
    while i < len(rows):
        # 빈 행 건너뛰기
        if not rows[i] or all(not cell.strip() for cell in rows[i]):
            i += 1
            continue

        # 그룹 시작 - 4개 행 읽기
        if i + 3 < len(rows):
            row1 = rows[i]      # 내경들
            row2 = rows[i + 1]  # 외경들
            row3 = rows[i + 2]  # z축 이동들
            row4 = rows[i + 3]  # sweep 길이들

            # D열부터 시작 (col_idx=3)
            inner_diameters = []
            outer_diameters = []
            z_offsets = []
            sweep_distances = []

            # 1행에서 원소 개수 파악 (D열=col 3부터)
            for col_idx in range(3, len(row1)):
                try:
                    if row1[col_idx].strip():
                        inner_diameters.append(float(row1[col_idx]))
                    else:
                        break  # 빈 셀 만나면 종료
                except (ValueError, AttributeError):
                    break

            # 원통 개수 = 원소 개수 - 1
            num_cylinders = len(inner_diameters) - 1

            if num_cylinders > 0:
                # 2~4행에서 원통 개수만큼 데이터 읽기
                for j in range(num_cylinders):
                    col_idx = 3 + j  # D열부터 시작

                    # 외경
                    try:
                        if len(row2) > col_idx and row2[col_idx].strip():
                            outer_diameters.append(float(row2[col_idx]))
                        else:
                            outer_diameters.append(None)
                    except (ValueError, AttributeError):
                        outer_diameters.append(None)

                    # z축 이동
                    try:
                        if len(row3) > col_idx and row3[col_idx].strip():
                            z_offsets.append(float(row3[col_idx]))
                        else:
                            z_offsets.append(None)
                    except (ValueError, AttributeError):
                        z_offsets.append(None)

                    # sweep 길이
                    try:
                        if len(row4) > col_idx and row4[col_idx].strip():
                            sweep_distances.append(float(row4[col_idx]))
                        else:
                            sweep_distances.append(None)
                    except (ValueError, AttributeError):
                        sweep_distances.append(None)

                groups.append({
                    'inner_diameters': inner_diameters,
                    'outer_diameters': outer_diameters,
                    'z_offsets': z_offsets,
                    'sweep_distances': sweep_distances,
                    'num_cylinders': num_cylinders
                })

            i += 4  # 다음 그룹으로 (4행 건너뜀)
        else:
            break

    return groups


def create_circle(oEditor, x_center, y_center, z_start, radius, name):
    """XY 평면에 Circle 생성"""
    oEditor.CreateCircle(
        [
            "NAME:CircleParameters",
            "IsCovered:=", True,
            "XCenter:=", "{}mm".format(x_center),
            "YCenter:=", "{}mm".format(y_center),
            "ZCenter:=", "{}mm".format(z_start),
            "Radius:=", "{}mm".format(radius),
            "WhichAxis:=", "Z",
            "NumSegments:=", "0"
        ],
        [
            "NAME:Attributes",
            "Name:=", name,
            "Flags:=", "",
            "Color:=", "(132 132 193)",
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
    log("  생성: {}".format(name))


def subtract_objects(oEditor, blank_name, tool_name):
    """Blank에서 Tool을 빼기"""
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
    log("  Subtract: {} - {}".format(blank_name, tool_name))


def move_object(oEditor, obj_name, dx, dy, dz):
    """객체를 이동"""
    oEditor.Move(
        [
            "NAME:Selections",
            "Selections:=", obj_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:TranslateParameters",
            "TranslateVectorX:=", "{}mm".format(dx),
            "TranslateVectorY:=", "{}mm".format(dy),
            "TranslateVectorZ:=", "{}mm".format(dz)
        ]
    )
    log("  이동: {} -> ({}, {}, {})".format(obj_name, dx, dy, dz))


def sweep_along_z(oEditor, obj_name, sweep_distance):
    """Z축 방향으로 Sweep"""
    oEditor.SweepAlongVector(
        [
            "NAME:Selections",
            "Selections:=", obj_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:VectorSweepParameters",
            "DraftAngle:=", "0deg",
            "DraftType:=", "Round",
            "CheckFaceFaceIntersection:=", False,
            "SweepVectorX:=", "0mm",
            "SweepVectorY:=", "0mm",
            "SweepVectorZ:=", "{}mm".format(sweep_distance)
        ]
    )
    log("  Sweep: {} -> Z축 {}mm".format(obj_name, sweep_distance))


def create_barriers_from_csv(csv_file_path, name_prefix="Barrier"):
    """CSV 파일에서 배리어를 생성 (V6 - Sweep 후 Subtract)"""
    log("=" * 60)
    log("Maxwell 3D - 배리어 생성 V6 (Sweep 후 Subtract)")
    log("=" * 60)

    log("\nCSV 파일 읽기 시작...")
    # CSV 데이터 읽기
    groups = read_csv_data(csv_file_path)

    log("읽은 데이터: 그룹 수={}".format(len(groups)))

    if not groups:
        log("오류: CSV 파일에서 데이터를 읽을 수 없습니다.")
        return

    # 각 그룹의 데이터 출력
    for i, group in enumerate(groups):
        log("\n그룹 {}: 원통 개수={}, 내경 개수={}".format(
            i+1,
            group['num_cylinders'],
            len(group['inner_diameters'])
        ))

    # Maxwell 프로젝트 및 디자인 가져오기
    oProject = oDesktop.GetActiveProject()
    if oProject is None:
        oProject = oDesktop.NewProject()
        log("새 프로젝트를 생성했습니다.")

    oDesign = oProject.GetActiveDesign()
    if oDesign is None:
        oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
        oDesign = oProject.GetActiveDesign()
        log("새 Maxwell 3D 디자인을 생성했습니다.")

    oEditor = oDesign.SetActiveEditor("3D Modeler")

    # 각 그룹의 원통 생성
    cylinder_count = 0
    success_count = 0
    fail_count = 0

    for group_idx, group in enumerate(groups):
        log("\n===== 그룹 {} 처리 시작 =====".format(group_idx + 1))

        for cyl_idx in range(group['num_cylinders']):
            cylinder_count += 1

            log("\n--- 배리어 {} (그룹 {}, 원통 {}) ---".format(
                cylinder_count, group_idx + 1, cyl_idx + 1
            ))

            # 데이터 유효성 체크
            inner_dia = group['inner_diameters'][cyl_idx]
            outer_dia = group['outer_diameters'][cyl_idx]
            z_offset = group['z_offsets'][cyl_idx]
            sweep_dist = group['sweep_distances'][cyl_idx]

            if inner_dia is None or outer_dia is None:
                log("  경고: 내경 또는 외경이 없습니다. 건너뜁니다.")
                fail_count += 1
                continue
            if z_offset is None or sweep_dist is None:
                log("  경고: Z 이동 또는 Sweep 거리가 없습니다. 건너뜁니다.")
                fail_count += 1
                continue

            # 내경/외경 크기 체크
            if inner_dia >= outer_dia:
                log("  경고: 내경({})이 외경({})보다 크거나 같습니다. 건너뜁니다.".format(inner_dia, outer_dia))
                fail_count += 1
                continue
            if inner_dia <= 0:
                log("  경고: 내경({})이 0 이하입니다. 건너뜁니다.".format(inner_dia))
                fail_count += 1
                continue

            log("  내경: {}mm".format(inner_dia))
            log("  외경: {}mm".format(outer_dia))
            log("  Z 이동: {}mm".format(z_offset))
            log("  Sweep: {}mm".format(sweep_dist))
            log("  [V6 방식: Sweep 후 Subtract]")

            # 원통 이름
            barrier_name = "{}_{}".format(name_prefix, cylinder_count)
            outer_cylinder_name = "{}_OuterCyl".format(barrier_name)
            inner_cylinder_name = "{}_InnerCyl".format(barrier_name)

            try:
                # === V6 방식: Circle → Move → Sweep → Subtract ===

                # 1. 외경 원 생성
                log("  [1] 외경 원 생성")
                create_circle(oEditor, 0, 0, z_offset, outer_dia/2.0, outer_cylinder_name)

                # 2. 외경 원 Sweep (원통 생성)
                log("  [2] 외경 원 Sweep")
                sweep_along_z(oEditor, outer_cylinder_name, sweep_dist)

                # 3. 내경 원 생성 (같은 Z 위치)
                log("  [3] 내경 원 생성")
                create_circle(oEditor, 0, 0, z_offset, inner_dia/2.0, inner_cylinder_name)

                # 4. 내경 원 Sweep (원통 생성)
                log("  [4] 내경 원 Sweep")
                sweep_along_z(oEditor, inner_cylinder_name, sweep_dist)

                # 5. 외경 원통에서 내경 원통 빼기 (Solid - Solid = Solid)
                log("  [5] Subtract (Solid - Solid)")
                subtract_objects(oEditor, outer_cylinder_name, inner_cylinder_name)

                # 최종 이름 변경
                oEditor.ChangeProperty(
                    [
                        "NAME:AllTabs",
                        [
                            "NAME:Geometry3DAttributeTab",
                            [
                                "NAME:PropServers",
                                outer_cylinder_name
                            ],
                            [
                                "NAME:ChangedProps",
                                [
                                    "NAME:Name",
                                    "Value:=", barrier_name
                                ]
                            ]
                        ]
                    ]
                )
                log("  완성: {}".format(barrier_name))
                success_count += 1

            except Exception as e:
                log("  [오류] 배리어 {} 생성 실패: {}".format(cylinder_count, str(e)))
                log("  배리어 {}를 건너뛰고 계속 진행합니다.".format(cylinder_count))
                fail_count += 1
                continue

    # 뷰 맞추기
    try:
        oEditor.FitAll()
    except:
        pass

    log("\n" + "=" * 60)
    log("완료!")
    log("  성공: {} 개".format(success_count))
    log("  실패: {} 개".format(fail_count))
    log("  전체: {} 개".format(cylinder_count))
    log("=" * 60)


# 스크립트 실행
log_filename = "barrier_v6_log_{}.txt".format(time.strftime("%Y%m%d_%H%M%S"))

try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
except:
    import sys
    if len(sys.argv) > 0:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_dir = os.getcwd()

log_path = os.path.join(script_dir, log_filename)
log_file = open(log_path, 'w')

log("\n" + "=" * 60)
log("스크립트 실행 시작")
log("로그 파일: {}".format(log_path))
log("=" * 60)

csv_file = os.path.join(script_dir, "BarrierDim.csv")

log("\nCSV 파일 경로: {}".format(csv_file))
log("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    log("\n오류: BarrierDim.csv 파일을 찾을 수 없습니다!")
    log("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    log("\n배리어 생성 함수 호출 중...")
    # 배리어 생성
    create_barriers_from_csv(
        csv_file_path=csv_file,
        name_prefix="Barrier"
    )
    log("\n스크립트 실행 완료!")

# 로그 파일 닫기
if log_file:
    log_file.close()
    print("\n로그 파일 저장됨: {}".format(log_path))
