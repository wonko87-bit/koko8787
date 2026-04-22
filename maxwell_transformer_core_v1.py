"""
Maxwell 3D - EI Transformer Core Auto-Modeling  (Stage 1)
Supported core types: 1by2 / 2by2 / 3by0 / 3by2

Coordinate system
  X : leg array direction (toward side legs)
  Y : core stack depth direction
  Z : leg height direction (vertical)

Z reference
  Z = 0                    : bottom face of lower yoke
  Z = Core_Bjr             : window bottom (= top face of lower yoke)
  Z = Core_Bjr + Core_HF   : window top
  Z = 2*Core_Bjr + Core_HF : top face of upper yoke

Cross-section geometry
  All parts start from a circle of diameter Core_DS.
  Flat faces are created by subtracting box-shaped clipping tools.
  Main leg : DS circle clipped to BS(X) x SS(Y)
  Side leg : DS circle clipped to Bjr(X) x SS(Y)
  Yoke     : DS circle (in YZ plane) clipped to SS(Y) x Bjr(Z), swept in X
"""

import ScriptEnv
ScriptEnv.Initialize("Ansoft.ElectronicsDesktop")
oDesktop.RestoreWindow()

import csv
import os


# ---------------------------------------------------------
# CSV parameter reader
# ---------------------------------------------------------

def read_params(csv_path):
    """Read key-value CSV. Skips comment lines (#) and short rows."""
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


# ---------------------------------------------------------
# Temp name generator (IronPython 2.7 compatible - no nonlocal)
# ---------------------------------------------------------

_tmp_idx = [0]

def _tmp():
    _tmp_idx[0] += 1
    return "_clip_{}".format(_tmp_idx[0])


# ---------------------------------------------------------
# Attribute block helper
# ---------------------------------------------------------

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


# ---------------------------------------------------------
# Maxwell primitive helpers
# ---------------------------------------------------------

def create_circle(oEditor, cx, cy, cz, r, which_axis, name, material="vacuum"):
    """
    Create a filled circle (disk).
    which_axis="Z" -> disk in XY plane
    which_axis="X" -> disk in YZ plane
    """
    oEditor.CreateCircle(
        [
            "NAME:CircleParameters",
            "IsCovered:=",   True,
            "XCenter:=",     "{}mm".format(cx),
            "YCenter:=",     "{}mm".format(cy),
            "ZCenter:=",     "{}mm".format(cz),
            "Radius:=",      "{}mm".format(r),
            "WhichAxis:=",   which_axis,
            "NumSegments:=", "0",
        ],
        _attr_block(name, material=material)
    )
    print("  Circle({}): {} c=({},{},{}) r={}".format(which_axis, name, cx, cy, cz, r))


def create_box_tool(oEditor, x, y, z, dx, dy, dz, name):
    """Create a box used as a subtract tool. Material left as default (vacuum)."""
    oEditor.CreateBox(
        [
            "NAME:BoxParameters",
            "XPosition:=", "{}mm".format(x),
            "YPosition:=", "{}mm".format(y),
            "ZPosition:=", "{}mm".format(z),
            "XSize:=",     "{}mm".format(dx),
            "YSize:=",     "{}mm".format(dy),
            "ZSize:=",     "{}mm".format(dz),
        ],
        _attr_block(name)
    )


def subtract(oEditor, blank, tools):
    """Subtract tool solids from blank solid. Tools are consumed."""
    oEditor.Subtract(
        [
            "NAME:Selections",
            "Blank Parts:=", blank,
            "Tool Parts:=",  ",".join(tools),
        ],
        [
            "NAME:SubtractParameters",
            "KeepOriginals:=", False,
        ]
    )
    print("  Subtract: {} - [{}]".format(blank, ", ".join(tools)))


def sweep_z(oEditor, name, dist):
    """Sweep object along +Z by dist [mm]."""
    oEditor.SweepAlongVector(
        [
            "NAME:Selections",
            "Selections:=",        name,
            "NewPartsModelFlag:=", "Model",
        ],
        [
            "NAME:VectorSweepParameters",
            "DraftAngle:=",                "0deg",
            "DraftType:=",                 "Round",
            "CheckFaceFaceIntersection:=", False,
            "SweepVectorX:=",              "0mm",
            "SweepVectorY:=",              "0mm",
            "SweepVectorZ:=",              "{}mm".format(dist),
        ]
    )
    print("  SweepZ: {} {}mm".format(name, dist))


def sweep_x(oEditor, name, dist):
    """Sweep object along +X by dist [mm]."""
    oEditor.SweepAlongVector(
        [
            "NAME:Selections",
            "Selections:=",        name,
            "NewPartsModelFlag:=", "Model",
        ],
        [
            "NAME:VectorSweepParameters",
            "DraftAngle:=",                "0deg",
            "DraftType:=",                 "Round",
            "CheckFaceFaceIntersection:=", False,
            "SweepVectorX:=",              "{}mm".format(dist),
            "SweepVectorY:=",              "0mm",
            "SweepVectorZ:=",              "0mm",
        ]
    )
    print("  SweepX: {} {}mm".format(name, dist))


def rename_object(oEditor, old_name, new_name):
    oEditor.ChangeProperty(
        [
            "NAME:AllTabs",
            [
                "NAME:Geometry3DAttributeTab",
                ["NAME:PropServers", old_name],
                [
                    "NAME:ChangedProps",
                    ["NAME:Name", "Value:=", new_name],
                ],
            ],
        ]
    )
    print("  Rename: {} -> {}".format(old_name, new_name))


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
    print("  Material: {} -> {}".format(name, material))


# ---------------------------------------------------------
# Clipping helpers (subtract flat faces from a swept cylinder)
# ---------------------------------------------------------

def _clip_x(oEditor, solid, x_center, x_size, DS, z_start, z_end):
    """
    Remove material outside +/- x_size/2 in X direction.
    x_center : center X of the solid
    x_size   : target width in X (BS or Bjr)
    DS       : reference circle diameter (original extent in X)
    """
    if x_size >= DS:
        return
    m      = 2.0
    big_y  = DS + 2.0 * m
    big_z  = (z_end - z_start) + 2.0 * m
    excess = DS / 2.0 - x_size / 2.0

    # Left clip
    n = _tmp()
    create_box_tool(oEditor,
                    x_center - DS / 2.0 - m,   # x start (outside left edge)
                    -big_y / 2.0,
                    z_start - m,
                    excess + m,                 # dx: covers excess + margin
                    big_y,
                    big_z,
                    n)
    subtract(oEditor, solid, [n])

    # Right clip
    n = _tmp()
    create_box_tool(oEditor,
                    x_center + x_size / 2.0,   # x start (right flat face)
                    -big_y / 2.0,
                    z_start - m,
                    excess + m,
                    big_y,
                    big_z,
                    n)
    subtract(oEditor, solid, [n])


def _clip_y(oEditor, solid, y_size, DS, x_start, x_end, z_start, z_end):
    """
    Remove material outside +/- y_size/2 in Y direction (centered at Y=0).
    y_size : target depth in Y (SS)
    DS     : reference circle diameter (original extent in Y)
    """
    if y_size >= DS:
        return
    m      = 2.0
    big_x  = (x_end - x_start) + 2.0 * m
    big_z  = (z_end - z_start) + 2.0 * m
    excess = DS / 2.0 - y_size / 2.0

    # Front clip (Y negative side)
    n = _tmp()
    create_box_tool(oEditor,
                    x_start - m,
                    -DS / 2.0 - m,
                    z_start - m,
                    big_x,
                    excess + m,
                    big_z,
                    n)
    subtract(oEditor, solid, [n])

    # Back clip (Y positive side)
    n = _tmp()
    create_box_tool(oEditor,
                    x_start - m,
                    y_size / 2.0,
                    z_start - m,
                    big_x,
                    excess + m,
                    big_z,
                    n)
    subtract(oEditor, solid, [n])


def _clip_z(oEditor, solid, z_center, z_size, DS, x_start, x_end):
    """
    Remove material outside +/- z_size/2 of z_center in Z direction.
    z_center : center Z of the yoke cross-section
    z_size   : target height in Z (Bjr)
    DS       : reference circle diameter (original extent in Z)
    """
    if z_size >= DS:
        return
    m      = 2.0
    big_x  = (x_end - x_start) + 2.0 * m
    big_y  = DS + 2.0 * m
    excess = DS / 2.0 - z_size / 2.0

    # Bottom clip
    n = _tmp()
    create_box_tool(oEditor,
                    x_start - m,
                    -big_y / 2.0,
                    z_center - DS / 2.0 - m,
                    big_x,
                    big_y,
                    excess + m,
                    n)
    subtract(oEditor, solid, [n])

    # Top clip
    n = _tmp()
    create_box_tool(oEditor,
                    x_start - m,
                    -big_y / 2.0,
                    z_center + z_size / 2.0,
                    big_x,
                    big_y,
                    excess + m,
                    n)
    subtract(oEditor, solid, [n])


# ---------------------------------------------------------
# Part creation
# ---------------------------------------------------------

def create_main_leg(oEditor, x_center, DS, BS, SS, h_leg, name, mat_core):
    """
    Main (wound) leg: DS circle clipped to BS(X) x SS(Y), swept h_leg in Z.
    Z base = 0.
    """
    print("  [MainLeg] {}  x={:.1f}  DS={} BS={} SS={}  h={}".format(
        name, x_center, DS, BS, SS, h_leg))
    tmp = name + "_c"
    create_circle(oEditor, x_center, 0.0, 0.0, DS / 2.0, "Z", tmp)
    sweep_z(oEditor, tmp, h_leg)
    _clip_x(oEditor, tmp, x_center, BS, DS, 0.0, h_leg)
    _clip_y(oEditor, tmp, SS, DS,
            x_center - DS / 2.0, x_center + DS / 2.0,
            0.0, h_leg)
    rename_object(oEditor, tmp, name)
    assign_material(oEditor, name, mat_core)


def create_side_leg(oEditor, x_center, DS, Bjr, SS, h_leg, name, mat_core):
    """
    Side (return) leg: ellipse swept h_leg in Z, then clipped to Bjr(X) x SS(Y).
    Ellipse: MajRadius = DS/2 (long axis), minor semi = (Bjr+20)/2 (short axis).
    Clip: long axis direction -> SS, short axis direction -> Bjr.
    DS is used as clip reference extent in both directions (safe upper bound).
    Z base = 0.
    """
    Bjr_ext = Bjr + 20.0  # ellipse short-axis full diameter
    print("  [SideLeg] {}  x={:.1f}  DS={} Bjr_ext={} Bjr={} SS={}  h={}".format(
        name, x_center, DS, Bjr_ext, Bjr, SS, h_leg))
    tmp = name + "_c"
    oEditor.CreateEllipse(
        [
            "NAME:EllipseParameters",
            "IsCovered:=",   True,
            "XCenter:=",     "{}mm".format(x_center),
            "YCenter:=",     "0mm",
            "ZCenter:=",     "0mm",
            "MajRadius:=",   "{}mm".format(DS / 2.0),
            "Ratio:=",       "{}".format(Bjr_ext / DS),
            "WhichAxis:=",   "Z",
            "NumSegments:=", "0",
        ],
        _attr_block(tmp)
    )
    sweep_z(oEditor, tmp, h_leg)
    _clip_y(oEditor, tmp, SS, DS,
            x_center - DS / 2.0, x_center + DS / 2.0,
            0.0, h_leg)
    _clip_x(oEditor, tmp, x_center, Bjr, DS, 0.0, h_leg)
    rename_object(oEditor, tmp, name)
    assign_material(oEditor, name, mat_core)


def create_yoke(oEditor, x_start, yoke_len, DS, SS, Bjr, z_center, name, mat_core):
    """
    Yoke: DS circle in YZ plane clipped to SS(Y) x Bjr(Z), swept yoke_len in X.
    z_center : vertical center of yoke cross-section.
    """
    print("  [Yoke] {}  x_start={:.1f} len={:.1f}  DS={} SS={} Bjr={}  z_center={:.1f}".format(
        name, x_start, yoke_len, DS, SS, Bjr, z_center))
    tmp = name + "_c"
    create_circle(oEditor, x_start, 0.0, z_center, DS / 2.0, "X", tmp)
    sweep_x(oEditor, tmp, yoke_len)
    _clip_y(oEditor, tmp, SS, DS,
            x_start, x_start + yoke_len,
            z_center - DS / 2.0, z_center + DS / 2.0)
    _clip_z(oEditor, tmp, z_center, Bjr, DS, x_start, x_start + yoke_len)
    rename_object(oEditor, tmp, name)
    assign_material(oEditor, name, mat_core)


# ---------------------------------------------------------
# Leg layout calculation
# ---------------------------------------------------------

def get_leg_layout(core_type, Core_ES, Core_ESR):
    """
    Return list of (x_center, is_main, label) for each leg.
    Core_ES  : main-to-main center-to-center spacing
    Core_ESR : outermost-main to side-leg center-to-center spacing
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
        raise ValueError("Unsupported core type: {}".format(core_type))


def get_yoke_x_range(layout, Core_BS, Core_Bjr):
    """Return (x_start, length) of yoke spanning outer edges of all legs."""
    x_l,  is_main_l = layout[0][0],  layout[0][1]
    x_r,  is_main_r = layout[-1][0], layout[-1][1]
    r_l = Core_BS / 2.0 if is_main_l else Core_Bjr / 2.0
    r_r = Core_BS / 2.0 if is_main_r else Core_Bjr / 2.0
    x_start = x_l - r_l
    x_end   = x_r + r_r
    return x_start, x_end - x_start


# ---------------------------------------------------------
# Core assembly
# ---------------------------------------------------------

def create_core(oEditor, core_type,
                Core_DS, Core_BS, Core_SS, Core_Bjr,
                Core_ES, Core_ESR, Core_HF, mat_core):

    print("=" * 60)
    print("EI Core  type={}".format(core_type))
    print("=" * 60)

    h_leg      = Core_HF + 2.0 * Core_Bjr
    z_bot_c    = Core_Bjr / 2.0
    z_top_c    = Core_Bjr + Core_HF + Core_Bjr / 2.0

    layout             = get_leg_layout(core_type, Core_ES, Core_ESR)
    yoke_x0, yoke_len  = get_yoke_x_range(layout, Core_BS, Core_Bjr)

    print("Derived dims:")
    print("  leg  Z-height = {}mm".format(h_leg))
    print("  yoke Z-height = {}mm  (= Core_Bjr)".format(Core_Bjr))
    print("  yoke X-range  = {:.1f}mm ~ {:.1f}mm  (len {:.1f}mm)".format(
        yoke_x0, yoke_x0 + yoke_len, yoke_len))
    print("Leg layout:")
    for x_c, is_main, label in layout:
        tag = "Main" if is_main else "Side"
        print("  {:8s} ({})  X = {:+.1f}mm".format(label, tag, x_c))

    # -- Legs -----------------------------------------------
    print("\n[Legs]")
    for x_c, is_main, label in layout:
        name = "Core_Leg_{}".format(label)
        if is_main:
            create_main_leg(oEditor, x_c, Core_DS, Core_BS, Core_SS,
                            h_leg, name, mat_core)
        else:
            create_side_leg(oEditor, x_c, Core_DS, Core_Bjr, Core_SS,
                            h_leg, name, mat_core)

    # -- Yokes ----------------------------------------------
    print("\n[Yokes]")
    create_yoke(oEditor, yoke_x0, yoke_len,
                Core_DS, Core_SS, Core_Bjr, z_bot_c,
                "Core_Yoke_Bot", mat_core)
    create_yoke(oEditor, yoke_x0, yoke_len,
                Core_DS, Core_SS, Core_Bjr, z_top_c,
                "Core_Yoke_Top", mat_core)

    oEditor.FitAll()
    print("\nDone.")


# ---------------------------------------------------------
# Entry point
# ---------------------------------------------------------

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
    print("CSV loaded: {}".format(csv_path))
else:
    print("CSV not found -> using defaults")
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
