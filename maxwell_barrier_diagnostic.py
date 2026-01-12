# -*- coding: utf-8 -*-
"""
Maxwell 3D - 배리어 생성 진단 버전
Barrier 13 문제 진단을 위한 특수 버전
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


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
    print("  생성: {}".format(name))


def subtract_objects(oEditor, blank_name, tool_name):
    """Blank에서 Tool을 빼기"""
    print("  [진단] Subtract 시작: Blank={}, Tool={}".format(blank_name, tool_name))

    # Subtract 전 객체 존재 확인
    all_objects = oEditor.GetObjectsInGroup("Solids")
    print("  [진단] 현재 Solid 객체 목록: {}".format(all_objects))
    print("  [진단] Blank '{}' 존재: {}".format(blank_name, blank_name in all_objects))
    print("  [진단] Tool '{}' 존재: {}".format(tool_name, tool_name in all_objects))

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
    print("  [진단] Subtract 완료")


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
    print("  이동: {} → ({}, {}, {})".format(obj_name, dx, dy, dz))


def sweep_along_z(oEditor, obj_name, sweep_distance):
    """Z축 방향으로 Sweep (2675mm 제한 회피)"""
    MAX_SWEEP = 2675.0  # Maxwell의 Sweep 제한

    print("  [진단] Sweep 시작: 객체={}, 거리={}mm".format(obj_name, sweep_distance))

    if sweep_distance > MAX_SWEEP:
        # 여러 번 나누어 sweep
        num_sweeps = int(sweep_distance / MAX_SWEEP) + 1
        sweep_per_step = sweep_distance / num_sweeps

        print("  Sweep: {} → Z축 {}mm ({}번 나누어 실행)".format(obj_name, sweep_distance, num_sweeps))

        for i in range(num_sweeps):
            print("  [진단] Sweep 단계 {}/{}: {}mm".format(i+1, num_sweeps, sweep_per_step))
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
                    "SweepVectorZ:=", "{}mm".format(sweep_per_step)
                ]
            )
            print("    단계 {}/{}: {}mm sweep 완료".format(i+1, num_sweeps, sweep_per_step))
    else:
        # 일반 sweep
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
        print("  Sweep: {} → Z축 {}mm 완료".format(obj_name, sweep_distance))


def create_single_barrier(oEditor, barrier_num, inner_dia, outer_dia, z_offset, sweep_dist, name_prefix="Barrier"):
    """단일 배리어 생성 (진단용)"""
    print("\n" + "="*60)
    print("배리어 {} 생성 시작".format(barrier_num))
    print("="*60)
    print("  내경: {}mm (반지름: {}mm)".format(inner_dia, inner_dia/2.0))
    print("  외경: {}mm (반지름: {}mm)".format(outer_dia, outer_dia/2.0))
    print("  Z 이동: {}mm".format(z_offset))
    print("  Sweep: {}mm".format(sweep_dist))
    print("  반지름 차이: {}mm".format((outer_dia - inner_dia)/2.0))

    barrier_name = "{}_{}_TEST".format(name_prefix, barrier_num)
    outer_circle_name = "{}_Outer".format(barrier_name)
    inner_circle_name = "{}_Inner".format(barrier_name)

    # 외경 원 생성
    print("\n[1단계] 외경 원 생성")
    create_circle(oEditor, 0, 0, 0, outer_dia/2.0, outer_circle_name)

    # 내경 원 생성
    print("\n[2단계] 내경 원 생성")
    create_circle(oEditor, 0, 0, 0, inner_dia/2.0, inner_circle_name)

    # Subtract 전 상태 확인
    print("\n[3단계] Subtract 준비")
    print("  외경 원 이름: {}".format(outer_circle_name))
    print("  내경 원 이름: {}".format(inner_circle_name))

    # 외경원에서 내경원 빼기
    subtract_objects(oEditor, outer_circle_name, inner_circle_name)

    # Z축으로 이동
    print("\n[4단계] Z축 이동")
    move_object(oEditor, outer_circle_name, 0, 0, z_offset)

    # Sweep
    print("\n[5단계] Sweep")
    sweep_along_z(oEditor, outer_circle_name, sweep_dist)

    # 최종 이름 변경
    print("\n[6단계] 이름 변경")
    oEditor.ChangeProperty(
        [
            "NAME:AllTabs",
            [
                "NAME:Geometry3DAttributeTab",
                [
                    "NAME:PropServers",
                    outer_circle_name
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
    print("  최종 이름: {}".format(barrier_name))
    print("\n" + "="*60)
    print("배리어 {} 생성 완료!".format(barrier_num))
    print("="*60)


def create_barriers_diagnostic(csv_file_path, target_barrier=13, name_prefix="Barrier"):
    """특정 배리어만 생성하여 진단"""
    print("="*60)
    print("Maxwell 3D - 배리어 진단 모드")
    print("대상 배리어: {}".format(target_barrier))
    print("="*60)

    # CSV 데이터 읽기
    groups = read_csv_data(csv_file_path)

    if not groups:
        print("오류: CSV 파일에서 데이터를 읽을 수 없습니다.")
        return

    # 모든 배리어 데이터를 평탄화
    all_barriers = []
    cylinder_count = 0

    for group_idx, group in enumerate(groups):
        for cyl_idx in range(group['num_cylinders']):
            cylinder_count += 1

            inner_dia = group['inner_diameters'][cyl_idx]
            outer_dia = group['outer_diameters'][cyl_idx]
            z_offset = group['z_offsets'][cyl_idx]
            sweep_dist = group['sweep_distances'][cyl_idx]

            all_barriers.append({
                'num': cylinder_count,
                'inner_dia': inner_dia,
                'outer_dia': outer_dia,
                'z_offset': z_offset,
                'sweep_dist': sweep_dist
            })

    print("\n총 {} 개의 배리어 데이터를 찾았습니다.".format(len(all_barriers)))

    # 목표 배리어 주변 정보 출력
    if target_barrier <= len(all_barriers):
        print("\n배리어 비교:")
        for i in range(max(0, target_barrier-3), min(len(all_barriers), target_barrier+2)):
            b = all_barriers[i]
            marker = " <-- 목표" if b['num'] == target_barrier else ""
            print("  배리어 {}: 내경={}, 외경={}, Z={}, Sweep={}{}".format(
                b['num'], b['inner_dia'], b['outer_dia'], b['z_offset'], b['sweep_dist'], marker
            ))

    # Maxwell 프로젝트 설정
    oProject = oDesktop.GetActiveProject()
    if oProject is None:
        oProject = oDesktop.NewProject()
        print("\n새 프로젝트를 생성했습니다.")

    oDesign = oProject.GetActiveDesign()
    if oDesign is None:
        oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
        oDesign = oProject.GetActiveDesign()
        print("새 Maxwell 3D 디자인을 생성했습니다.")

    oEditor = oDesign.SetActiveEditor("3D Modeler")

    # 목표 배리어만 생성
    if target_barrier <= len(all_barriers):
        b = all_barriers[target_barrier - 1]

        if b['inner_dia'] is None or b['outer_dia'] is None:
            print("\n오류: 배리어 {} 데이터가 불완전합니다.".format(target_barrier))
            return
        if b['z_offset'] is None or b['sweep_dist'] is None:
            print("\n오류: 배리어 {} 데이터가 불완전합니다.".format(target_barrier))
            return
        if b['inner_dia'] >= b['outer_dia']:
            print("\n오류: 배리어 {} 내경이 외경보다 큽니다.".format(target_barrier))
            return

        create_single_barrier(
            oEditor,
            b['num'],
            b['inner_dia'],
            b['outer_dia'],
            b['z_offset'],
            b['sweep_dist'],
            name_prefix
        )
    else:
        print("\n오류: 배리어 {}는 존재하지 않습니다. (최대: {})".format(target_barrier, len(all_barriers)))

    # 뷰 맞추기
    oEditor.FitAll()

    print("\n진단 완료!")


# 스크립트 실행
print("\n" + "="*60)
print("배리어 진단 스크립트 시작")
print("="*60)

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

csv_file = os.path.join(script_dir, "BarrierDim.csv")

print("\nCSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    print("\n오류: BarrierDim.csv 파일을 찾을 수 없습니다!")
    print("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    # 여기서 숫자를 바꿔서 다른 배리어도 테스트 가능
    # 예: create_barriers_diagnostic(csv_file, target_barrier=12)
    create_barriers_diagnostic(csv_file, target_barrier=13, name_prefix="Barrier")
    print("\n스크립트 실행 완료!")
