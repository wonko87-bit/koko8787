# -*- coding: utf-8 -*-
"""
Maxwell 3D Barrier 생성 스크립트
CSV 파일에서 좌표와 크기를 읽어 채워진 직사각형을 XZ 평면에 생성
"""

import csv
import os
import ScriptEnv

# Maxwell 환경 초기화
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop = oDesktop
oProject = oDesktop.GetActiveProject()
oDesign = oProject.GetActiveDesign()
oEditor = oDesign.SetActiveEditor("3D Modeler")


def read_barrier_csv(csv_file_path):
    """
    CSV 파일에서 Barrier 정보를 읽어옴

    CSV 구조:
    - A열: Ref_X
    - B열: Ref_Z
    - C열: X_Length
    - D열: Z_Length
    - B2: X offset (모든 X좌표에 더해짐)
    - 4행부터 데이터 시작 (A4, B4, C4, D4...)

    Returns:
        blocks: 블록 정보 리스트
        x_offset: X offset 값 (B2)
    """
    blocks = []
    x_offset = 0.0

    if not os.path.exists(csv_file_path):
        return blocks, x_offset

    with open(csv_file_path, 'r') as f:
        reader = csv.reader(f)
        rows = list(reader)

        # B2 셀에서 X offset 읽기 (2행, B열=인덱스1)
        if len(rows) > 1 and len(rows[1]) > 1:
            try:
                x_offset = float(rows[1][1])
            except:
                x_offset = 0.0

        # 4행부터 데이터 읽기 (인덱스 3부터)
        for row in rows[3:]:
            if len(row) < 4:
                continue

            try:
                ref_x = float(row[0])
                ref_z = float(row[1])
                x_length = float(row[2])
                z_length = float(row[3])

                blocks.append({
                    'ref_x': ref_x,
                    'ref_z': ref_z,
                    'x_length': x_length,
                    'z_length': z_length
                })
            except:
                continue

    return blocks, x_offset


def create_rectangle_xz(oEditor, x_start, z_start, x_size, z_size, name):
    """XZ 평면에 채워진 Rectangle 생성"""
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


def create_barriers_from_csv(csv_file_path, name_prefix="Barrier"):
    """CSV 파일에서 Barrier들을 생성"""
    blocks, x_offset = read_barrier_csv(csv_file_path)

    if not blocks:
        return

    oProject = oDesktop.GetActiveProject()
    if oProject is None:
        oProject = oDesktop.NewProject()

    oDesign = oProject.GetActiveDesign()
    if oDesign is None:
        oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
        oDesign = oProject.GetActiveDesign()

    oEditor = oDesign.SetActiveEditor("3D Modeler")

    for idx, block in enumerate(blocks):
        block_name = "{}_{}".format(name_prefix, idx + 1)

        try:
            ref_x = block['ref_x'] + x_offset
            ref_z = block['ref_z']
            x_length = block['x_length']
            z_length = block['z_length']

            # 채워진 직사각형 생성
            create_rectangle_xz(oEditor, ref_x, ref_z, x_length, z_length, block_name)

        except Exception as e:
            continue


# 스크립트 실행
if __name__ == "__main__":
    # CSV 파일 경로 (스크립트와 같은 폴더)
    script_dir = os.path.dirname(__file__)
    csv_file = os.path.join(script_dir, "BarrierDim.csv")

    # Barrier 생성
    create_barriers_from_csv(csv_file, "Barrier")
