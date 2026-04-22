# -*- coding: utf-8 -*-
"""
Maxwell 3D - EI 변압기 철심 자동 모델링  (Stage 1)
지원 코어 타입: 1by2 / 2by2 / 3by0 / 3by2

좌표계
  X : 레그 배열 방향 (사이드레그 방향)
  Y : 코어 스택 깊이 방향
  Z : 레그 높이 방향 (수직)

Z 기준
  Z = 0            : 하부 요크 최하단
  Z = Core_Bjr     : 창 하단 (하부 요크 최상단 = Core_HF 기준점)
  Z = Core_Bjr + Core_HF      : 창 상단
  Z = 2*Core_Bjr + Core_HF    : 상부 요크 최상단
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


# ─────────────────────────────────────────────────────────
# CSV 파라미터 읽기
# ─────────────────────────────────────────────────────────

def read_params(csv_path):
    """key-value 형식 CSV 읽기. 주석행(#) 및 빈 행 무시."""
    params = {}
    with open(csv_path, 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) < 3:
                continue
            sec = row[0].strip()
            key = row[1].strip()
            val = row[2].strip()
            if not sec or not key or not val or sec.startswith('#'):
                continue
            params[key] = val
    return params


# ─────────────────────────────────────────────────────────
# Maxwell 오브젝트 헬퍼
# ─────────────────────────────────────────────────────────

def _attr_block(name, color="(128 128 0)", transparency=0.3, material="vacuum"):
    return [
        "NAME:Attributes",
        "Name:=",                    name,
        "Flags:=",                   "",
        "Color:=",                   color,
        "Transparency:=",            transparency,
        "PartCoordinateSystem:=",    "Global",
        "UDMId:=",                   "",
        "MaterialValue:=",           "\"{}\"".format(material),
        "SurfaceMaterialValue:=",    "\"\"",
        "SolveInside:=",             True,
        "ShellElement:=",            False,
        "ShellElementThickness:=",   "0mm",
        "IsMaterialEditable:=",      True,
        "UseMaterialAppearance:=",   False,
        "IsLightweight:=",           False,
    ]


def create_ellipse_xy(oEditor, cx, cy, cz, rx, ry, name, material="vacuum"):
    """
    XY 평면 타원 (WhichAxis=Z).
    rx : X방향 반지름,  ry : Y방향 반지름
    MajRadius = rx (X축),  Ratio = ry/rx
    ※ Maxwell 에서 WhichAxis=Z 일 때 MajRadius 가 X축임을 가정.
      실행 후 단면 방향 확인 필요.
    """
    oEditor.CreateEllipse(
        [
            "NAME:EllipseParameters",
            "IsCovered:=",   True,
            "XCenter:=",     "{}mm".format(cx),
            "YCenter:=",     "{}mm".format(cy),
            "ZCenter:=",     "{}mm".format(cz),
            "MajRadius:=",   "{}mm".format(rx),
            "Ratio:=",       "{}".format(float(ry) / float(rx)),
            "WhichAxis:=",   "Z",
            "NumSegments:=", "0",
        ],
        _attr_block(name, material=material)
    )
    print("  Ellipse(XY): {} cx={} cy={} cz={} rx={} ry={}".format(
        name, cx, cy, cz, rx, ry))


def create_ellipse_yz(oEditor, cx, cy, cz, ry, rz, name, material="vacuum"):
    """
    YZ 평면 타원 (WhichAxis=X).
    ry : Y방향 반지름,  rz : Z방향 반지름
    MajRadius = ry (Y축),  Ratio = rz/ry
    ※ Maxwell 에서 WhichAxis=X 일 때 MajRadius 가 Y축임을 가정.
      실행 후 단면 방향 확인 필요.
    """
    oEditor.CreateEllipse(
        [
            "NAME:EllipseParameters",
            "IsCovered:=",   True,
            "XCenter:=",     "{}mm".format(cx),
            "YCenter:=",     "{}mm".format(cy),
            "ZCenter:=",     "{}mm".format(cz),
            "MajRadius:=",   "{}mm".format(ry),
            "Ratio:=",       "{}".format(float(rz) / float(ry)),
            "WhichAxis:=",   "X",
            "NumSegments:=", "0",
        ],
        _attr_block(name, material=material)
    )
    print("  Ellipse(YZ): {} cx={} cy={} cz={} ry={} rz={}".format(
        name, cx, cy, cz, ry, rz))


def sweep_z(oEditor, name, dist):
    """Z 방향 SweepAlongVector"""
    oEditor.SweepAlongVector(
        [
            "NAME:Selections",
            "Selections:=",        name,
            "NewPartsModelFlag:=", "Model",
        ],
        [
            "NAME:VectorSweepParameters",
            "DraftAngle:=",              "0deg",
            "DraftType:=",               "Round",
            "CheckFaceFaceIntersection:=", False,
            "SweepVectorX:=",            "0mm",
            "SweepVectorY:=",            "0mm",
            "SweepVectorZ:=",            "{}mm".format(dist),
        ]
    )
    print("  SweepZ: {} → {}mm".format(name, dist))


def sweep_x(oEditor, name, dist):
    """X 방향 SweepAlongVector"""
    oEditor.SweepAlongVector(
        [
            "NAME:Selections",
            "Selections:=",        name,
            "NewPartsModelFlag:=", "Model",
        ],
        [
            "NAME:VectorSweepParameters",
            "DraftAngle:=",              "0deg",
            "DraftType:=",               "Round",
            "CheckFaceFaceIntersection:=", False,
            "SweepVectorX:=",            "{}mm".format(dist),
            "SweepVectorY:=",            "0mm",
            "SweepVectorZ:=",            "0mm",
        ]
    )
    print("  SweepX: {} → {}mm".format(name, dist))


def assign_material(oEditor, name, material):
    oEditor.ChangeProperty(
        [
            "NAME:AllTabs",
            [
                "NAME:Geometry3DAttributeTab",
                ["NAME:PropServers", name],
                [
                    "NAME:ChangedProps",
                    ["NAME:Material", "Value:=", "\"{}\"".format(material)],
                ],
            ],
        ]
    )
    print("  Material: {} → {}".format(name, material))


# ─────────────────────────────────────────────────────────
# 레그 배치 계산
# ─────────────────────────────────────────────────────────

def get_leg_layout(core_type, Core_ES, Core_ESR):
    """
    코어 타입별 레그 X 좌표와 종류 반환.
    Returns: list of (x_center, is_main, label)
    Core_ES  : 메인-메인 center to center
    Core_ESR : 메인-사이드 center to center (최외곽 메인 기준)
    """
    if core_type == "1by2":
        return [
            (-Core_ESR,             False, "Side_L"),
            (0.0,                   True,  "Main_1"),
            (+Core_ESR,             False, "Side_R"),
        ]
    elif core_type == "2by2":
        h = Core_ES / 2.0
        return [
            (-(h + Core_ESR),       False, "Side_L"),
            (-h,                    True,  "Main_1"),
            (+h,                    True,  "Main_2"),
            (+(h + Core_ESR),       False, "Side_R"),
        ]
    elif core_type == "3by0":
        return [
            (-Core_ES,              True,  "Main_1"),
            (0.0,                   True,  "Main_2"),
            (+Core_ES,              True,  "Main_3"),
        ]
    elif core_type == "3by2":
        return [
            (-(Core_ES + Core_ESR), False, "Side_L"),
            (-Core_ES,              True,  "Main_1"),
            (0.0,                   True,  "Main_2"),
            (+Core_ES,              True,  "Main_3"),
            (+(Core_ES + Core_ESR), False, "Side_R"),
        ]
    else:
        raise ValueError("지원하지 않는 코어 타입: {}".format(core_type))


def get_yoke_x_range(layout, Core_BS, Core_Bjr):
    """
    요크 X축 시작점과 전체 길이 계산.
    최외곽 레그의 외면 끝까지 커버.
    """
    x_l, is_main_l = layout[0][0],  layout[0][1]
    x_r, is_main_r = layout[-1][0], layout[-1][1]
    r_l = Core_BS / 2.0 if is_main_l else Core_Bjr / 2.0
    r_r = Core_BS / 2.0 if is_main_r else Core_Bjr / 2.0
    x_start = x_l - r_l
    x_end   = x_r + r_r
    return x_start, x_end - x_start   # (시작점, 길이)


# ─────────────────────────────────────────────────────────
# 철심 생성 메인 함수
# ─────────────────────────────────────────────────────────

def create_core(oEditor, core_type,
                Core_DS, Core_BS, Core_SS, Core_Bjr,
                Core_ES, Core_ESR, Core_HF, mat_core):

    print("=" * 60)
    print("EI 철심 생성: {}".format(core_type))
    print("=" * 60)

    # 파생 치수
    h_leg   = Core_HF + 2.0 * Core_Bjr        # 레그 전체 Z 높이
    z_bot_c = Core_Bjr / 2.0                   # 하부 요크 중심 Z
    z_top_c = Core_Bjr + Core_HF + Core_Bjr / 2.0  # 상부 요크 중심 Z

    layout            = get_leg_layout(core_type, Core_ES, Core_ESR)
    yoke_x0, yoke_len = get_yoke_x_range(layout, Core_BS, Core_Bjr)

    print("\n파생 치수:")
    print("  레그 Z높이 = {}mm".format(h_leg))
    print("  요크 Z높이 = {}mm  (= Core_Bjr)".format(Core_Bjr))
    print("  요크 X범위 = {}mm ~ {}mm  (길이 {}mm)".format(
        yoke_x0, yoke_x0 + yoke_len, yoke_len))
    print("\n레그 배치:")
    for x_c, is_main, label in layout:
        print("  {:8s}  X = {:+.1f}mm".format(label, x_c))

    # ── 레그 생성 (XY 타원 → Z sweep) ───────────────────
    print("\n[레그 생성]")
    for x_c, is_main, label in layout:
        rx   = Core_BS / 2.0 if is_main else Core_Bjr / 2.0
        ry   = Core_SS / 2.0
        name = "Core_Leg_{}".format(label)
        create_ellipse_xy(oEditor, x_c, 0.0, 0.0, rx, ry, name)
        sweep_z(oEditor, name, h_leg)
        assign_material(oEditor, name, mat_core)

    # ── 요크 생성 (YZ 타원 → X sweep) ───────────────────
    print("\n[요크 생성]")

    create_ellipse_yz(oEditor,
                      yoke_x0, 0.0, z_bot_c,
                      Core_SS / 2.0, Core_Bjr / 2.0,
                      "Core_Yoke_Bot")
    sweep_x(oEditor, "Core_Yoke_Bot", yoke_len)
    assign_material(oEditor, "Core_Yoke_Bot", mat_core)

    create_ellipse_yz(oEditor,
                      yoke_x0, 0.0, z_top_c,
                      Core_SS / 2.0, Core_Bjr / 2.0,
                      "Core_Yoke_Top")
    sweep_x(oEditor, "Core_Yoke_Top", yoke_len)
    assign_material(oEditor, "Core_Yoke_Top", mat_core)

    oEditor.FitAll()
    print("\n완료!")


# ─────────────────────────────────────────────────────────
# 실행 진입점
# ─────────────────────────────────────────────────────────

try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
except Exception:
    script_dir = os.getcwd()

csv_path = os.path.join(script_dir, "transformer_config.csv")

if os.path.exists(csv_path):
    p         = read_params(csv_path)
    core_type = p.get("core_type", "1by2")
    Core_DS   = float(p.get("Core_DS",   100.0))
    Core_BS   = float(p.get("Core_BS",    80.0))
    Core_SS   = float(p.get("Core_SS",   120.0))
    Core_Bjr  = float(p.get("Core_Bjr",   40.0))
    Core_ES   = float(p.get("Core_ES",   150.0))
    Core_ESR  = float(p.get("Core_ESR",  120.0))
    Core_HF   = float(p.get("Core_HF",   300.0))
    mat_core  = p.get("mat_core", "vacuum")
    print("CSV 로드: {}".format(csv_path))
else:
    print("CSV 없음 → 기본값 사용")
    core_type = "1by2"
    Core_DS   = 100.0
    Core_BS   =  80.0
    Core_SS   = 120.0
    Core_Bjr  =  40.0
    Core_ES   = 150.0
    Core_ESR  = 120.0
    Core_HF   = 300.0
    mat_core  = "vacuum"

oProject = oDesktop.GetActiveProject()
if oProject is None:
    oProject = oDesktop.NewProject()

oDesign = oProject.GetActiveDesign()
if oDesign is None:
    oProject.InsertDesign("Maxwell 3D", "Transformer_Core", "Electrostatic", "")
    oDesign = oProject.GetActiveDesign()

oEditor = oDesign.SetActiveEditor("3D Modeler")

create_core(oEditor, core_type,
            Core_DS, Core_BS, Core_SS, Core_Bjr,
            Core_ES, Core_ESR, Core_HF, mat_core)
