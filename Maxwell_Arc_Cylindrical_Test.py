# -*- coding: utf-8 -*-
"""
Maxwell 3D - Cylindrical 좌표로 Arc 생성 테스트
"""

import math
import ScriptEnv

# Maxwell 환경 초기화
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop = oDesktop
oProject = oDesktop.GetActiveProject()
oDesign = oProject.GetActiveDesign()
oEditor = oDesign.SetActiveEditor("3D Modeler")


def create_arc_cylindrical_test1(oEditor, radius, start_angle, end_angle, name):
    """
    방법 1: 좌표를 cylindrical 형식 문자열로 시도
    예: "50mm, 0deg, 0mm" (R, Theta, Z)
    """
    try:
        oEditor.CreatePolyline(
            [
                "NAME:PolylineParameters",
                "IsPolylineCovered:=", False,
                "IsPolylineClosed:=", False,
                [
                    "NAME:PolylinePoints",
                    [
                        "NAME:PLPoint",
                        "X:=", "{}mm, {}deg, 0mm".format(radius, start_angle)  # Cylindrical?
                    ],
                    [
                        "NAME:PLPoint",
                        "X:=", "{}mm, {}deg, 0mm".format(radius, end_angle)
                    ]
                ],
                [
                    "NAME:PolylineSegments",
                    [
                        "NAME:PLSegment",
                        "SegmentType:=", "Arc",
                        "StartIndex:=", 0,
                        "NoOfPoints:=", 2
                    ]
                ],
                [
                    "NAME:PolylineXSection",
                    "XSectionType:=", "None"
                ]
            ],
            [
                "NAME:Attributes",
                "Name:=", name
            ]
        )
        return True
    except Exception as e:
        print("방법 1 실패: {}".format(str(e)))
        return False


def create_arc_cylindrical_test2(oEditor, radius, start_angle, end_angle, name):
    """
    방법 2: R, Theta, Z 파라미터로 시도
    """
    try:
        oEditor.CreatePolyline(
            [
                "NAME:PolylineParameters",
                "IsPolylineCovered:=", False,
                "IsPolylineClosed:=", False,
                [
                    "NAME:PolylinePoints",
                    [
                        "NAME:PLPoint",
                        "R:=", "{}mm".format(radius),
                        "Theta:=", "{}deg".format(start_angle),
                        "Z:=", "0mm"
                    ],
                    [
                        "NAME:PLPoint",
                        "R:=", "{}mm".format(radius),
                        "Theta:=", "{}deg".format(end_angle),
                        "Z:=", "0mm"
                    ]
                ],
                [
                    "NAME:PolylineSegments",
                    [
                        "NAME:PLSegment",
                        "SegmentType:=", "Arc",
                        "StartIndex:=", 0,
                        "NoOfPoints:=", 2
                    ]
                ],
                [
                    "NAME:PolylineXSection",
                    "XSectionType:=", "None"
                ]
            ],
            [
                "NAME:Attributes",
                "Name:=", name
            ]
        )
        return True
    except Exception as e:
        print("방법 2 실패: {}".format(str(e)))
        return False


def create_arc_cylindrical_test3(oEditor, radius, start_angle, end_angle, name):
    """
    방법 3: Cartesian으로 변환 (현재 방식)
    """
    try:
        start_x = radius * math.cos(math.radians(start_angle))
        start_y = radius * math.sin(math.radians(start_angle))

        end_x = radius * math.cos(math.radians(end_angle))
        end_y = radius * math.sin(math.radians(end_angle))

        oEditor.CreatePolyline(
            [
                "NAME:PolylineParameters",
                "IsPolylineCovered:=", False,
                "IsPolylineClosed:=", False,
                [
                    "NAME:PolylinePoints",
                    [
                        "NAME:PLPoint",
                        "X:=", "{}mm".format(start_x),
                        "Y:=", "{}mm".format(start_y),
                        "Z:=", "0mm"
                    ],
                    [
                        "NAME:PLPoint",
                        "X:=", "{}mm".format(end_x),
                        "Y:=", "{}mm".format(end_y),
                        "Z:=", "0mm"
                    ]
                ],
                [
                    "NAME:PolylineSegments",
                    [
                        "NAME:PLSegment",
                        "SegmentType:=", "Arc",
                        "StartIndex:=", 0,
                        "NoOfPoints:=", 2,
                        "ArcAngle:=", "{}deg".format(end_angle - start_angle),
                        "ArcCenterX:=", "0mm",
                        "ArcCenterY:=", "0mm",
                        "ArcCenterZ:=", "0mm"
                    ]
                ],
                [
                    "NAME:PolylineXSection",
                    "XSectionType:=", "None"
                ]
            ],
            [
                "NAME:Attributes",
                "Name:=", name,
                "Color:=", "(143 175 143)",
                "MaterialValue:=", "\"vacuum\"",
                "SolveInside:=", True
            ]
        )
        return True
    except Exception as e:
        print("방법 3 실패: {}".format(str(e)))
        return False


# 테스트 실행
print("=== Cylindrical 좌표 Arc 생성 테스트 ===")
print("")

# 테스트 파라미터
radius = 50
start_angle = 0
end_angle = 45

print("파라미터: R={}, Theta={}~{}".format(radius, start_angle, end_angle))
print("")

# 방법 1 테스트
print("방법 1 테스트: 문자열 형식")
result1 = create_arc_cylindrical_test1(oEditor, radius, start_angle, end_angle, "Arc_Test1")
print("결과: {}".format("성공" if result1 else "실패"))
print("")

# 방법 2 테스트
print("방법 2 테스트: R, Theta, Z 파라미터")
result2 = create_arc_cylindrical_test2(oEditor, radius, start_angle, end_angle + 10, "Arc_Test2")
print("결과: {}".format("성공" if result2 else "실패"))
print("")

# 방법 3 테스트 (기존 방식)
print("방법 3 테스트: Cartesian 변환")
result3 = create_arc_cylindrical_test3(oEditor, radius, start_angle, end_angle + 20, "Arc_Test3")
print("결과: {}".format("성공" if result3 else "실패"))
print("")

print("=== 테스트 완료 ===")
print("성공한 방법:")
if result1:
    print("- 방법 1: 문자열 형식")
if result2:
    print("- 방법 2: R, Theta, Z 파라미터")
if result3:
    print("- 방법 3: Cartesian 변환")

# 뷰 맞추기
try:
    oEditor.FitAll()
except:
    pass
