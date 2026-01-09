# -*- coding: utf-8 -*-
"""
Maxwell 3D - 코일 생성 V1
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_csv_data(csv_file_path):
    """CSV 파일에서 코일 데이터를 읽어옵니다."""
    coil_data = []
    num_coils = None

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

        # C2 셀 (row 1, col 2): 원기둥 수량
        if len(rows) > 1 and len(rows[1]) > 2:
            try:
                if rows[1][2].strip():  # 빈 문자열이 아닌 경우만
                    num_coils = int(float(rows[1][2]))
            except (ValueError, AttributeError):
                print("경고: C2 셀에서 원기둥 수량을 읽을 수 없습니다.")
                return [], None

        if num_coils is None or num_coils <= 0:
            print("경고: 유효한 원기둥 수량이 없습니다.")
            return [], None

        # 각 원기둥 데이터 읽기
        for i in range(num_coils):
            # n번째 원기둥의 열 인덱스 = 2 + (n)*3
            col_idx = 2 + i * 3

            # C9/F9/I9 (row 8, col): Sweep 거리
            sweep_distance = None
            if len(rows) > 8 and len(rows[8]) > col_idx:
                try:
                    if rows[8][col_idx].strip():
                        sweep_distance = float(rows[8][col_idx])
                except (ValueError, AttributeError):
                    pass

            # C10/F10/I10 (row 9, col): Z축 이동 거리
            z_offset = None
            if len(rows) > 9 and len(rows[9]) > col_idx:
                try:
                    if rows[9][col_idx].strip():
                        z_offset = float(rows[9][col_idx])
                except (ValueError, AttributeError):
                    pass

            # C12/F12/I12 (row 11, col): 내경
            inner_diameter = None
            if len(rows) > 11 and len(rows[11]) > col_idx:
                try:
                    if rows[11][col_idx].strip():
                        inner_diameter = float(rows[11][col_idx])
                except (ValueError, AttributeError):
                    pass

            # C13/F13/I13 (row 12, col): 외경
            outer_diameter = None
            if len(rows) > 12 and len(rows[12]) > col_idx:
                try:
                    if rows[12][col_idx].strip():
                        outer_diameter = float(rows[12][col_idx])
                except (ValueError, AttributeError):
                    pass

            coil_data.append({
                'sweep_distance': sweep_distance,
                'z_offset': z_offset,
                'inner_diameter': inner_diameter,
                'outer_diameter': outer_diameter
            })

    return coil_data, num_coils


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


def create_coils_from_csv(csv_file_path, name_prefix="Coil"):
    """CSV 파일에서 코일을 생성 (V1)"""
    print("=" * 60)
    print("Maxwell 3D - 코일 생성 V1")
    print("=" * 60)

    print("\nCSV 파일 읽기 시작...")
    # CSV 데이터 읽기
    coil_data, num_coils = read_csv_data(csv_file_path)

    print("읽은 데이터: num_coils={}, coil_data 길이={}".format(num_coils, len(coil_data) if coil_data else 0))

    if not coil_data or num_coils is None:
        print("오류: CSV 파일에서 데이터를 읽을 수 없습니다.")
        print("  - num_coils: {}".format(num_coils))
        print("  - coil_data: {}".format(coil_data))
        return

    print("\nCSV 데이터:")
    print("  원기둥 수량: {}".format(num_coils))

    # 각 코일의 데이터 출력
    for i, data in enumerate(coil_data):
        print("  원기둥 {}: 내경={}, 외경={}, Z이동={}, Sweep={}".format(
            i+1,
            data.get('inner_diameter'),
            data.get('outer_diameter'),
            data.get('z_offset'),
            data.get('sweep_distance')
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

    # 각 원기둥 생성
    for i, data in enumerate(coil_data):
        print("\n===== 원기둥 {} 생성 =====".format(i+1))

        # 데이터 유효성 체크
        if data['inner_diameter'] is None or data['outer_diameter'] is None:
            print("  경고: 내경 또는 외경이 없습니다. 건너뜁니다.")
            continue
        if data['z_offset'] is None or data['sweep_distance'] is None:
            print("  경고: Z 이동 또는 Sweep 거리가 없습니다. 건너뜁니다.")
            continue

        print("  내경: {}mm".format(data['inner_diameter']))
        print("  외경: {}mm".format(data['outer_diameter']))
        print("  Z 이동: {}mm".format(data['z_offset']))
        print("  Sweep: {}mm".format(data['sweep_distance']))

        # 원기둥 이름
        coil_name = "{}_{}".format(name_prefix, i+1)
        outer_circle_name = "{}_Outer".format(coil_name)
        inner_circle_name = "{}_Inner".format(coil_name)

        print("  외경 원 생성 중...")
        # 외경 원 생성 (중심: 0, 0, 0)
        create_circle(oEditor, 0, 0, 0, data['outer_diameter']/2.0, outer_circle_name)

        print("  내경 원 생성 중...")
        # 내경 원 생성 (중심: 0, 0, 0)
        create_circle(oEditor, 0, 0, 0, data['inner_diameter']/2.0, inner_circle_name)

        print("  Subtract 작업 중...")
        # 외경원에서 내경원 빼기 (도넛 형태)
        subtract_objects(oEditor, outer_circle_name, inner_circle_name)

        print("  Z축 이동 중...")
        # Z축으로 이동
        move_object(oEditor, outer_circle_name, 0, 0, data['z_offset'])

        print("  Sweep 작업 중...")
        # Z축 방향으로 Sweep
        sweep_along_z(oEditor, outer_circle_name, data['sweep_distance'])

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
                            "Value:=", coil_name
                        ]
                    ]
                ]
            ]
        )
        print("  완성: {}".format(coil_name))

    # 뷰 맞추기
    oEditor.FitAll()

    print("\n" + "=" * 60)
    print("완료!")
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

csv_file = os.path.join(script_dir, "CoilDim.csv")

print("\nCSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

if not os.path.exists(csv_file):
    print("\n오류: CoilDim.csv 파일을 찾을 수 없습니다!")
    print("다음 위치에 파일이 있어야 합니다: {}".format(csv_file))
else:
    print("\n코일 생성 함수 호출 중...")
    # 코일 생성
    create_coils_from_csv(
        csv_file_path=csv_file,
        name_prefix="Coil"
    )
    print("\n스크립트 실행 완료!")
