# -*- coding: utf-8 -*-
"""
Maxwell 3D - 배리어 생성 V5 (Duplicate + Unite 개선)
복사본 이름 자동 감지 및 예외 처리 강화
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
    print("  Subtract: {} - {}".format(blank_name, tool_name))


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
    print("  Sweep: {} → Z축 {}mm".format(obj_name, sweep_distance))


def get_all_objects(oEditor):
    """모든 객체 목록 가져오기"""
    all_objs = []
    try:
        all_objs.extend(oEditor.GetObjectsInGroup("Solids"))
    except:
        pass
    try:
        all_objs.extend(oEditor.GetObjectsInGroup("Sheets"))
    except:
        pass
    try:
        all_objs.extend(oEditor.GetObjectsInGroup("Lines"))
    except:
        pass
    return all_objs


def duplicate_and_unite(oEditor, obj_name, z_offset):
    """객체를 복사하고 이동시킨 뒤 합치기 (개선된 버전)"""

    # Paste 전 객체 목록 저장
    print("  [진단] Paste 전 객체 목록 확인 중...")
    before_objs = get_all_objects(oEditor)
    print("  [진단] Paste 전 객체 수: {}".format(len(before_objs)))

    # 복사
    oEditor.Copy(
        [
            "NAME:Selections",
            "Selections:=", obj_name
        ]
    )
    print("  복사: {}".format(obj_name))

    # 붙여넣기
    oEditor.Paste()
    print("  붙여넣기 완료")

    # Paste 후 객체 목록 확인
    print("  [진단] Paste 후 객체 목록 확인 중...")
    after_objs = get_all_objects(oEditor)
    print("  [진단] Paste 후 객체 수: {}".format(len(after_objs)))

    # 새로 생성된 객체 찾기
    new_objs = [obj for obj in after_objs if obj not in before_objs]

    if len(new_objs) == 0:
        print("  [오류] 붙여넣기 후 새 객체를 찾을 수 없습니다!")
        raise Exception("Paste failed: no new object found")

    copy_name = new_objs[0]
    print("  [진단] 생성된 복사본 이름: {}".format(copy_name))

    # 복사본을 위로 이동
    move_object(oEditor, copy_name, 0, 0, z_offset)

    # 원본과 복사본 합치기
    print("  [진단] Unite 시작: {} + {}".format(obj_name, copy_name))
    oEditor.Unite(
        [
            "NAME:Selections",
            "Selections:=", obj_name + "," + copy_name
        ],
        [
            "NAME:UniteParameters",
            "KeepOriginals:=", False
        ]
    )
    print("  Unite 완료: {} + {} → {}".format(obj_name, copy_name, obj_name))


def create_barriers_from_csv(csv_file_path, name_prefix="Barrier"):
    """CSV 파일에서 배리어를 생성 (V5 - Duplicate + Unite 개선)"""
    print("=" * 60)
    print("Maxwell 3D - 배리어 생성 V5 (Duplicate + Unite 개선)")
    print("=" * 60)

    print("\nCSV 파일 읽기 시작...")
    # CSV 데이터 읽기
    groups = read_csv_data(csv_file_path)

    print("읽은 데이터: 그룹 수={}".format(len(groups)))

    if not groups:
        print("오류: CSV 파일에서 데이터를 읽을 수 없습니다.")
        return

    # 각 그룹의 데이터 출력
    for i, group in enumerate(groups):
        print("\n그룹 {}: 원통 개수={}, 내경 개수={}".format(
            i+1,
            group['num_cylinders'],
            len(group['inner_diameters'])
        ))

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

    # 각 그룹의 원통 생성
    cylinder_count = 0
    success_count = 0
    fail_count = 0

    for group_idx, group in enumerate(groups):
        print("\n===== 그룹 {} 처리 시작 =====".format(group_idx + 1))

        for cyl_idx in range(group['num_cylinders']):
            cylinder_count += 1

            print("\n--- 원통 {} (그룹 {}, 원통 {}) ---".format(
                cylinder_count, group_idx + 1, cyl_idx + 1
            ))

            # 데이터 유효성 체크
            inner_dia = group['inner_diameters'][cyl_idx]
            outer_dia = group['outer_diameters'][cyl_idx]
            z_offset = group['z_offsets'][cyl_idx]
            sweep_dist = group['sweep_distances'][cyl_idx]

            if inner_dia is None or outer_dia is None:
                print("  경고: 내경 또는 외경이 없습니다. 건너뜁니다.")
                fail_count += 1
                continue
            if z_offset is None or sweep_dist is None:
                print("  경고: Z 이동 또는 Sweep 거리가 없습니다. 건너뜁니다.")
                fail_count += 1
                continue

            # 내경/외경 크기 체크
            if inner_dia >= outer_dia:
                print("  경고: 내경({})이 외경({})보다 크거나 같습니다. 건너뜁니다.".format(inner_dia, outer_dia))
                fail_count += 1
                continue
            if inner_dia <= 0:
                print("  경고: 내경({})이 0 이하입니다. 건너뜁니다.".format(inner_dia))
                fail_count += 1
                continue

            print("  내경: {}mm".format(inner_dia))
            print("  외경: {}mm".format(outer_dia))
            print("  Z 이동: {}mm".format(z_offset))
            print("  Sweep: {}mm (V5: 절반씩 2번)".format(sweep_dist))

            # 원통 이름
            barrier_name = "{}_{}".format(name_prefix, cylinder_count)
            outer_circle_name = "{}_Outer".format(barrier_name)
            inner_circle_name = "{}_Inner".format(barrier_name)

            try:
                # 외경 원 생성
                create_circle(oEditor, 0, 0, 0, outer_dia/2.0, outer_circle_name)

                # 내경 원 생성
                create_circle(oEditor, 0, 0, 0, inner_dia/2.0, inner_circle_name)

                # 외경원에서 내경원 빼기 (도넛 형태)
                subtract_objects(oEditor, outer_circle_name, inner_circle_name)

                # Z축으로 이동
                move_object(oEditor, outer_circle_name, 0, 0, z_offset)

                # 절반만큼만 Sweep
                half_sweep = sweep_dist / 2.0
                sweep_along_z(oEditor, outer_circle_name, half_sweep)

                # 복사 + 이동 + 합치기
                duplicate_and_unite(oEditor, outer_circle_name, half_sweep)

                # 최종 이름 변경
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
                print("  완성: {}".format(barrier_name))
                success_count += 1

            except Exception as e:
                print("  [오류] 배리어 {} 생성 실패: {}".format(cylinder_count, str(e)))
                print("  배리어 {}를 건너뛰고 계속 진행합니다.".format(cylinder_count))
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
    print("  전체: {} 개".format(cylinder_count))
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

csv_file = os.path.join(script_dir, "BarrierDim.csv")

print("\nCSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    print("\n오류: BarrierDim.csv 파일을 찾을 수 없습니다!")
    print("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    print("\n배리어 생성 함수 호출 중...")
    # 배리어 생성
    create_barriers_from_csv(
        csv_file_path=csv_file,
        name_prefix="Barrier"
    )
    print("\n스크립트 실행 완료!")
