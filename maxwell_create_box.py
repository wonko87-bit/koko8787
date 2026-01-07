# -*- coding: utf-8 -*-
"""
Ansys Maxwell Electronics Desktop 2021 R1 - 3D Box Modeling Script
===================================================================

This script creates a 3D rectangular box (cuboid) in Maxwell 3D.
The box is created with one corner at the origin (0,0,0) and extends
in the positive direction.

Author: Claude
Date: 2026-01-07

사용 방법:
---------
1. Ansys Maxwell Electronics Desktop을 실행합니다
2. Tools → Record Script to File로 스크립트 녹화를 시작하여 API 확인 가능
3. Tools → Run Script에서 이 파일을 실행합니다

또는 Maxwell 내부 스크립트 콘솔에서 직접 실행 가능합니다.
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()


def create_maxwell_box(width, depth, height, name="Box1", material="vacuum"):
    """
    Maxwell 3D에서 직육면체를 생성합니다.

    Parameters:
    -----------
    width : float
        X 방향 너비 (mm)
    depth : float
        Y 방향 깊이 (mm)
    height : float
        Z 방향 높이 (mm)
    name : str
        박스 이름 (기본값: "Box1")
    material : str
        재질 이름 (기본값: "vacuum")

    Returns:
    --------
    box_name : str
        생성된 박스의 이름
    """

    # 유효성 검사
    if width <= 0 or depth <= 0 or height <= 0:
        raise ValueError("모든 치수는 양수여야 합니다")

    print("=" * 60)
    print("Maxwell 3D - 박스 생성 중...")
    print("=" * 60)
    print("치수:")
    print("  너비 (X):  {} mm".format(width))
    print("  깊이 (Y):  {} mm".format(depth))
    print("  높이 (Z):  {} mm".format(height))
    print("  이름:      {}".format(name))
    print("  재질:      {}".format(material))
    print("-" * 60)

    try:
        # Desktop 객체 가져오기
        oDesktop = oDesktop

        # 활성 프로젝트 가져오기 (없으면 새로 생성)
        oProject = oDesktop.GetActiveProject()
        if oProject is None:
            oProject = oDesktop.NewProject()
            print("새 프로젝트를 생성했습니다.")

        # 활성 디자인 가져오기
        oDesign = oProject.GetActiveDesign()

        # Maxwell 3D 디자인이 없으면 새로 생성
        if oDesign is None:
            oProject.InsertDesign("Maxwell 3D", "Maxwell3DDesign1", "Magnetostatic", "")
            oDesign = oProject.GetActiveDesign()
            print("새 Maxwell 3D 디자인을 생성했습니다.")

        # 3D Modeler 에디터 가져오기
        oEditor = oDesign.SetActiveEditor("3D Modeler")

        # 박스 생성
        # CreateBox 명령 사용
        # 형식: CreateBox [시작점], [크기], [속성]
        oEditor.CreateBox(
            [
                "NAME:BoxParameters",
                "XPosition:=", "0mm",
                "YPosition:=", "0mm",
                "ZPosition:=", "0mm",
                "XSize:=", "{}mm".format(width),
                "YSize:=", "{}mm".format(depth),
                "ZSize:=", "{}mm".format(height)
            ],
            [
                "NAME:Attributes",
                "Name:=", name,
                "Flags:=", "",
                "Color:=", "(143 175 143)",
                "Transparency:=", 0,
                "PartCoordinateSystem:=", "Global",
                "UDMId:=", "",
                "MaterialValue:=", "\"{}\"".format(material),
                "SurfaceMaterialValue:=", "\"\"",
                "SolveInside:=", True,
                "IsMaterialEditable:=", True,
                "UseMaterialAppearance:=", False,
                "IsLightweight:=", False
            ]
        )

        print("\n성공! 박스 '{}' 가 생성되었습니다.".format(name))
        print("=" * 60)

        # 뷰 맞추기
        oEditor.FitAll()

        return name

    except Exception as e:
        print("오류 발생: {}".format(str(e)))
        raise


def create_box_simple():
    """
    기본 설정으로 간단하게 박스를 생성합니다.
    0,0,0 위치에서 20x10x5 크기의 박스를 생성합니다.
    """
    return create_maxwell_box(20, 10, 5, "Box_20x10x5", "vacuum")


def create_multiple_boxes():
    """
    여러 개의 박스를 생성하는 예제입니다.
    """
    boxes = [
        (20, 10, 5, "SmallBox", "vacuum"),
        (50, 50, 50, "Cube", "aluminum"),
        (100, 30, 15, "LargeBox", "copper")
    ]

    created_boxes = []
    for width, depth, height, name, material in boxes:
        box_name = create_maxwell_box(width, depth, height, name, material)
        created_boxes.append(box_name)

    print("\n생성된 박스 목록:")
    for i, box in enumerate(created_boxes, 1):
        print("  {}. {}".format(i, box))

    return created_boxes


# ============================================================================
# 메인 실행 부분
# ============================================================================

if __name__ == "__main__":
    # 기본 예제: 20x10x5 박스 생성
    print("\n" + "=" * 60)
    print("Maxwell 3D 박스 모델링 스크립트")
    print("=" * 60 + "\n")

    # 방법 1: 간단한 박스 생성 (20, 10, 5)
    create_box_simple()

    # 방법 2: 사용자 정의 박스 생성
    # create_maxwell_box(100, 50, 30, "MyCustomBox", "aluminum")

    # 방법 3: 여러 박스 생성
    # create_multiple_boxes()

    print("\n스크립트 실행 완료!")
