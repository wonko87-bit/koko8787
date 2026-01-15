# -*- coding: utf-8 -*-
"""
Maxwell 3D - 변압기 철심 생성 V4 (Unite + Sweep 방식)
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


def read_csv_data(csv_file_path):
    """CSV 파일에서 철심 데이터를 읽어옵니다."""
    data = []
    gap = None  # E1 값
    window_height = None  # E2 값
    BJR_max = None  # E3 값

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

        # E1 (row 0, col 4): gap distance
        if len(rows) > 0 and len(rows[0]) > 4:
            gap = float(rows[0][4])

        # E2 (row 1, col 4): window height
        if len(rows) > 1 and len(rows[1]) > 4:
            window_height = float(rows[1][4])

        # E3 (row 2, col 4): BJR_max
        if len(rows) > 2 and len(rows[2]) > 4:
            BJR_max = float(rows[2][4])

        # 데이터 행 읽기 (A, B, C 열)
        for i, row in enumerate(rows):
            if len(row) >= 3:
                try:
                    a = float(row[0])  # A열: 메인 레그 X
                    b = float(row[1])  # B열: Y (공통)
                    c = float(row[2])  # C열: 사이드 레그 X
                    data.append({'A': a, 'B': b, 'C': c})
                except ValueError:
                    continue

    return data, gap, window_height, BJR_max


def create_rectangle(oEditor, x_start, y_start, z_start, width, height, name):
    """XY 평면에 Rectangle 생성"""
    oEditor.CreateRectangle(
        [
            "NAME:RectangleParameters",
            "IsCovered:=", True,
            "XStart:=", "{}mm".format(x_start),
            "YStart:=", "{}mm".format(y_start),
            "ZStart:=", "{}mm".format(z_start),
            "Width:=", "{}mm".format(width),
            "Height:=", "{}mm".format(height),
            "WhichAxis:=", "Z"
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


def rotate_object(oEditor, obj_name, axis, angle_deg):
    """객체를 회전 (원점 기준)"""
    oEditor.Rotate(
        [
            "NAME:Selections",
            "Selections:=", obj_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:RotateParameters",
            "RotateAxis:=", axis,
            "RotateAngle:=", "{}deg".format(angle_deg)
        ]
    )
    print("  회전: {} → {} axis, {}deg".format(obj_name, axis, angle_deg))


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


def sweep_along_x(oEditor, obj_name, sweep_distance):
    """X축 방향으로 Sweep"""
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
            "SweepVectorX:=", "{}mm".format(sweep_distance),
            "SweepVectorY:=", "0mm",
            "SweepVectorZ:=", "0mm"
        ]
    )
    print("  Sweep: {} → X축 {}mm".format(obj_name, sweep_distance))


def unite_objects(oEditor, obj_list, result_name):
    """여러 객체를 Unite"""
    if len(obj_list) < 2:
        return

    oEditor.Unite(
        [
            "NAME:Selections",
            "Selections:=", ",".join(obj_list)
        ],
        [
            "NAME:UniteParameters",
            "KeepOriginals:=", False
        ]
    )

    # Unite 후 첫 번째 이름으로 변경
    oEditor.ChangeProperty(
        [
            "NAME:AllTabs",
            [
                "NAME:Geometry3DAttributeTab",
                [
                    "NAME:PropServers",
                    obj_list[0]
                ],
                [
                    "NAME:ChangedProps",
                    [
                        "NAME:Name",
                        "Value:=", result_name
                    ]
                ]
            ]
        ]
    )
    print("  Unite: {} → {}".format(obj_list, result_name))


def duplicate_object(oEditor, obj_name, new_name):
    """객체를 복제"""
    oEditor.Copy(
        [
            "NAME:Selections",
            "Selections:=", obj_name
        ]
    )
    oEditor.Paste()

    # 복사된 객체는 원본 이름 + "1" 형태로 생성됨 (언더스코어 없음)
    copied_name = obj_name + "1"

    # 새 이름으로 변경
    oEditor.ChangeProperty(
        [
            "NAME:AllTabs",
            [
                "NAME:Geometry3DAttributeTab",
                [
                    "NAME:PropServers",
                    copied_name
                ],
                [
                    "NAME:ChangedProps",
                    [
                        "NAME:Name",
                        "Value:=", new_name
                    ]
                ]
            ]
        ]
    )
    print("  복제: {} → {}".format(obj_name, new_name))
    return new_name


def duplicate_and_rotate(oEditor, obj_name, axis, angle_deg, num_clones):
    """객체를 복제하고 회전"""
    oEditor.DuplicateAroundAxis(
        [
            "NAME:Selections",
            "Selections:=", obj_name,
            "NewPartsModelFlag:=", "Model"
        ],
        [
            "NAME:DuplicateAroundAxisParameters",
            "CreateNewObjects:=", True,
            "WhichAxis:=", axis,
            "AngleStr:=", "{}deg".format(angle_deg),
            "NumClones:=", "{}".format(num_clones)
        ],
        [
            "NAME:Options",
            "DuplicateAssignments:=", False
        ]
    )
    print("  복제 및 회전: {} → {} axis, {}deg, {} clones".format(obj_name, axis, angle_deg, num_clones))


def move_faces_along_normal(oEditor, obj_name, face_id, offset_distance):
    """특정 면을 법선 방향으로 이동 (확장)"""
    if face_id is None:
        print("  경고: Face ID가 None이므로 표면 확장을 건너뜁니다.")
        return

    try:
        oEditor.MoveFaces(
            [
                "NAME:Selections",
                "Selections:=", obj_name,
                "NewPartsModelFlag:=", "Model"
            ],
            [
                "NAME:Parameters",
                [
                    "NAME:MoveFacesParameters",
                    "MoveAlongNormalFlag:=", True,
                    "OffsetDistance:=", "{}mm".format(offset_distance),
                    "MoveVectorX:=", "0mm",
                    "MoveVectorY:=", "0mm",
                    "MoveVectorZ:=", "0mm",
                    "FacesToMove:=", [face_id]
                ]
            ]
        )
        print("  표면 확장: {} Face {} → {}mm".format(obj_name, face_id, offset_distance))
    except Exception as e:
        print("  경고: 표면 확장 실패. Object: {}, Face: {}".format(obj_name, face_id))
        print("  에러: {}".format(str(e)))


def get_face_ids(oEditor, obj_name):
    """객체의 모든 면 ID를 가져옵니다"""
    faces = oEditor.GetFaceIDs(obj_name)
    return faces


def get_face_by_position(oEditor, obj_name, x, y, z):
    """특정 위치에 있는 면 ID를 가져옵니다"""
    try:
        face_id = oEditor.GetFaceByPosition(
            [
                "NAME:FaceParameters",
                "BodyName:=", obj_name,
                "XPosition:=", "{}mm".format(x),
                "YPosition:=", "{}mm".format(y),
                "ZPosition:=", "{}mm".format(z)
            ]
        )
        return face_id
    except Exception as e:
        print("  경고: Face를 찾을 수 없습니다. Object: {}, Position: ({}, {}, {})".format(obj_name, x, y, z))
        print("  에러: {}".format(str(e)))
        return None


def create_core_from_csv(csv_file_path, name_prefix="Core"):
    """CSV 파일에서 변압기 철심을 생성 (V5)"""
    print("=" * 60)
    print("Maxwell 3D - 변압기 철심 생성 V5")
    print("=" * 60)

    # CSV 데이터 읽기
    data, gap, window_height, BJR_max = read_csv_data(csv_file_path)

    if not data:
        print("오류: CSV 파일에서 데이터를 읽을 수 없습니다.")
        return

    if gap is None or window_height is None or BJR_max is None:
        print("오류: E1(gap), E2(window_height), E3(BJR_max) 값이 없습니다.")
        return

    print("\nCSV 데이터:")
    print("  Gap (E1): {}mm".format(gap))
    print("  Window Height (E2): {}mm".format(window_height))
    print("  BJR_max (E3): {}mm".format(BJR_max))
    print("  데이터 행 수: {}".format(len(data)))

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

    # Z 시작 위치 (원점)
    z_start = 0.0

    # ===== 1. 각 레이어마다 직사각형 생성 (모두 원점 중심에 겹쳐짐) =====
    main_rects = []
    side_rects = []  # 사이드 레그 (1개만)

    for i, row_data in enumerate(data):
        a = row_data['A']  # 메인 레그 X
        b = row_data['B']  # Y (공통)
        c = row_data['C']  # 사이드 레그 X

        print("\n레이어 {}: A={}, B={}, C={}".format(i+1, a, b, c))

        # 메인 레그
        main_name = "{}_Layer{}_Main".format(name_prefix, i+1)
        create_rectangle(oEditor, 0, 0, z_start, a, b, main_name)
        move_object(oEditor, main_name, -a/2.0, -b/2.0, 0)
        main_rects.append(main_name)

        # 사이드 레그 (원점 중심에 배치)
        side_name = "{}_Layer{}_Side".format(name_prefix, i+1)
        create_rectangle(oEditor, 0, 0, z_start, c, b, side_name)
        move_object(oEditor, side_name, -c/2.0, -b/2.0, 0)
        side_rects.append(side_name)

    # ===== 2. 각 레그별 평면 Unite =====
    print("\n===== Unite 작업 =====")

    # 메인 레그 Unite
    main_united = "{}_MainLeg".format(name_prefix)
    unite_objects(oEditor, main_rects, main_united)

    # 사이드 레그 Unite
    side_united = "{}_SideLeg".format(name_prefix)
    unite_objects(oEditor, side_rects, side_united)

    # ===== 3. Sweep 전에 사이드 레그 평면 복사 (원점 기준 4개 생성) =====
    print("\n===== 사이드 레그 평면 복사 =====")

    # 사이드 레그 2 (반대편)
    side2_united = "{}_Side2Leg".format(name_prefix)
    duplicate_object(oEditor, side_united, side2_united)

    # 상부 요크
    yoke1_united = "{}_Yoke1".format(name_prefix)
    duplicate_object(oEditor, side_united, yoke1_united)

    # 하부 요크
    yoke2_united = "{}_Yoke2".format(name_prefix)
    duplicate_object(oEditor, side_united, yoke2_united)

    # ===== 4. 각 평면을 최종 위치로 이동 및 회전 =====
    print("\n===== 평면 이동 및 회전 =====")

    # 사이드 레그 1 (원본): 오른쪽으로 gap 이동
    move_object(oEditor, side_united, gap, 0, 0)

    # 사이드 레그 2: 왼쪽으로 gap 이동
    move_object(oEditor, side2_united, -gap, 0, 0)

    # 상부 요크: Y축 90도 회전 후 상부로 이동 (Z축)
    rotate_object(oEditor, yoke1_united, "Y", 90)
    yoke1_z = window_height + BJR_max - BJR_max/2
    move_object(oEditor, yoke1_united, 0, 0, yoke1_z)

    # 하부 요크: Y축 -90도 회전 후 하부로 이동 (Z축)
    rotate_object(oEditor, yoke2_united, "Y", -90)
    yoke2_z = - BJR_max/2
    move_object(oEditor, yoke2_united, 0, 0, yoke2_z)

    # ===== 5. Sweep 작업 =====
    print("\n===== Sweep 작업 =====")

    # Main leg와 Side leg: Z축 방향으로 E3*2 + E2 만큼
    leg_sweep_distance = BJR_max * 2.0 + window_height
    sweep_along_z(oEditor, main_united, leg_sweep_distance)
    sweep_along_z(oEditor, side_united, leg_sweep_distance)
    sweep_along_z(oEditor, side2_united, leg_sweep_distance)

    # Yoke: X축 방향으로 2*E1 + E3 만큼
    yoke_sweep_distance = 2.0 * gap + BJR_max
    sweep_along_x(oEditor, yoke1_united, yoke_sweep_distance)
    sweep_along_x(oEditor, yoke2_united, yoke_sweep_distance)

    # ===== 6. Yoke X축 중심 맞추기 =====
    print("\n===== Yoke X축 중심 맞추기 =====")
    yoke_x_offset = -yoke_sweep_distance / 2.0
    move_object(oEditor, yoke1_united, yoke_x_offset, 0, 0)
    move_object(oEditor, yoke2_united, yoke_x_offset, 0, 0)

    # ===== 7. 모든 객체 Z축 중심 맞추기 =====
    #print("\n===== Z축 중심 맞추기 =====")
    leg_z_offset = - BJR_max
    move_object(oEditor, main_united, 0, 0, leg_z_offset)
    move_object(oEditor, side_united, 0, 0, leg_z_offset)
    move_object(oEditor, side2_united, 0, 0, leg_z_offset)
    #move_object(oEditor, yoke1_united, 0, 0, leg_z_offset)
    #move_object(oEditor, yoke2_united, 0, 0, leg_z_offset)


    # 뷰 맞추기
    oEditor.FitAll()

    print("\n" + "=" * 60)
    print("완료!")
    print("=" * 60)


# 스크립트 실행
try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
except:
    import sys
    if len(sys.argv) > 0:
        script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    else:
        script_dir = os.getcwd()

csv_file = os.path.join(script_dir, "CoreDim.csv")

print("CSV 파일 경로: {}".format(csv_file))
print("CSV 파일 존재 여부: {}".format(os.path.exists(csv_file)))

# 철심 생성
create_core_from_csv(
    csv_file_path=csv_file,
    name_prefix="Core"
)
