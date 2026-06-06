"""
CSV to AutoCAD Mechanical - Rectangle / Arc / Slot / Sector Drawing Script
===========================================================================

[지원 도형]
  RECTANGLE  직사각형 (회전 각도 지원)
  ARC        원호     (시계/반시계 방향 지원)
  SLOT       장공     (직사각형 양 끝에 반원이 붙은 형태)
  SECTOR     링 섹터  (원호 양 끝에 직사각형 막대가 붙은 형태)

[공통 컬럼]
  shape      : RECTANGLE | ARC | SLOT | SECTOR
  layer      : 레이어 이름 (기본 "0")
  color      : AutoCAD 색상 인덱스 1-256  (0 = BYLAYER)

[RECTANGLE 컬럼]
  x, y       : 기준점 (기본 = 좌하단 CORNER)
  origin     : CORNER(기본) | CENTER
  width      : 가로 길이
  height     : 세로 길이
  angle      : 회전각 (도, 반시계 기준, 기본 0)

[ARC 컬럼]
  x, y        : 원호 중심점
  radius      : 반지름
  start_angle : 시작각 (도, 3시 = 0)
  end_angle   : 끝각   (도)
  direction   : CCW(기본, 반시계) | CW(시계)

[SLOT 컬럼]
  x, y       : 기준점
  origin     : CORNER(기본) | CENTER
  width      : 전체 길이 (긴 방향)
  height     : 전체 폭   (짧은 방향, 반원 지름)
  angle      : 회전각 (도)

[SECTOR 컬럼]
  x, y        : 중심점 (호의 중심)
  r_inner     : 안쪽 반지름 (막대의 안쪽 끝)
  r_outer     : 바깥쪽 반지름 (호의 반지름)
  start_angle : 시작각 (도)  — 이쪽 막대가 해당 방향으로 뻗어나감
  end_angle   : 끝각   (도)  — 이쪽 막대가 해당 방향으로 뻗어나감
  direction   : CCW(기본) | CW
  angle       : 전체 회전 오프셋 (도)

  예) start_angle=0, end_angle=90, direction=CCW 일 때:
                  │  ← end_angle=90쪽 막대 (세로 ㅣ)
                  │
       r_inner    │      r_outer
          ├───────┤──────────────┤
          │  내호  │    외호      │
          ╰───────┴──────────────╯
          ─────────────────────────── ← start_angle=0쪽 막대 (가로 ㅡ)

  실제 닫힌 외곽선 구성:
    ① 외호  (start_angle → end_angle, r_outer)
    ② 막대  (end_angle 방향, r_outer → r_inner)
    ③ 내호  (end_angle → start_angle, r_inner, 역방향)
    ④ 막대  (start_angle 방향, r_inner → r_outer)

[실행]
  pip install pywin32
  python csv_to_autocad.py shapes.csv
"""

import csv
import math
import os
import sys
import traceback

try:
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ---------------------------------------------------------------------------
# 공통 유틸
# ---------------------------------------------------------------------------

def _f(row, key, default=0.0):
    try:
        v = row.get(key, "").strip()
        return float(v) if v else default
    except ValueError:
        return default


def _i(row, key, default=0):
    try:
        v = row.get(key, "").strip()
        return int(v) if v else default
    except ValueError:
        return default


def _pt(x, y, z=0.0):
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z))
    )


def _deg(deg):
    return math.radians(deg)


def _rotate(cx, cy, px, py, rad):
    dx, dy = px - cx, py - cy
    return (
        cx + dx * math.cos(rad) - dy * math.sin(rad),
        cy + dx * math.sin(rad) + dy * math.cos(rad),
    )


def _polar(cx, cy, r, angle_deg):
    """극좌표 → 직교좌표."""
    a = math.radians(angle_deg)
    return cx + r * math.cos(a), cy + r * math.sin(a)


def ensure_layer(doc, name):
    if not name or name == "0":
        return
    try:
        doc.Layers.Item(name)
    except Exception:
        doc.Layers.Add(name)


def apply_props(obj, row):
    layer = row.get("layer", "").strip() or "0"
    obj.Layer = layer
    color = _i(row, "color", 0)
    if color:
        obj.color = color


# ---------------------------------------------------------------------------
# RECTANGLE
# ---------------------------------------------------------------------------

def draw_rectangle(mspace, doc, row):
    x, y   = _f(row, "x"),     _f(row, "y")
    w, h   = _f(row, "width"), _f(row, "height")
    angle  = _f(row, "angle",  0.0)
    origin = row.get("origin", "CORNER").strip().upper()

    if w <= 0 or h <= 0:
        raise ValueError(f"width={w}, height={h} 모두 양수여야 합니다.")

    if origin == "CENTER":
        cx, cy = x, y
        corners = [
            (x - w/2, y - h/2), (x + w/2, y - h/2),
            (x + w/2, y + h/2), (x - w/2, y + h/2),
        ]
    else:  # CORNER
        cx, cy = x + w/2, y + h/2
        corners = [
            (x,     y),     (x + w, y),
            (x + w, y + h), (x,     y + h),
        ]

    rad = _deg(angle)
    rotated = [_rotate(cx, cy, px, py, rad) for px, py in corners]

    flat = []
    for px, py in rotated:
        flat += [px, py, 0.0]
    flat += list(rotated[0]) + [0.0]

    pts = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
    obj = mspace.Add3DPoly(pts)
    apply_props(obj, row)
    return [obj]


# ---------------------------------------------------------------------------
# ARC
# ---------------------------------------------------------------------------

def draw_arc(mspace, doc, row):
    x, y        = _f(row, "x"),          _f(row, "y")
    radius      = _f(row, "radius",      0.0)
    start_angle = _f(row, "start_angle", 0.0)
    end_angle   = _f(row, "end_angle",   90.0)
    direction   = row.get("direction", "CCW").strip().upper()

    if radius <= 0:
        raise ValueError(f"radius={radius} 는 양수여야 합니다.")

    if direction == "CW":
        start_angle, end_angle = -start_angle, -end_angle

    obj = mspace.AddArc(_pt(x, y), radius, _deg(start_angle), _deg(end_angle))
    apply_props(obj, row)
    return [obj]


# ---------------------------------------------------------------------------
# SLOT
# ---------------------------------------------------------------------------

def draw_slot(mspace, doc, row):
    x, y   = _f(row, "x"),     _f(row, "y")
    w, h   = _f(row, "width"), _f(row, "height")
    angle  = _f(row, "angle",  0.0)
    origin = row.get("origin", "CORNER").strip().upper()

    if w <= 0 or h <= 0:
        raise ValueError(f"width={w}, height={h} 모두 양수여야 합니다.")
    if h > w:
        raise ValueError("SLOT: width 가 height 보다 커야 합니다.")

    r = h / 2
    straight = w - h

    if origin == "CENTER":
        cx, cy = x, y
    else:
        cx, cy = x + w/2, y + h/2

    rad = _deg(angle)

    def rot(px, py):
        return _rotate(cx, cy, px, py, rad)

    cr_local = (cx + straight/2, cy)
    cl_local = (cx - straight/2, cy)
    cr = rot(*cr_local)
    cl = rot(*cl_local)

    created = []

    obj1 = mspace.AddArc(_pt(*cr), r, _deg(angle + 270), _deg(angle + 90))
    apply_props(obj1, row); created.append(obj1)

    if straight > 0:
        obj2 = mspace.AddLine(
            _pt(*rot(cr_local[0], cr_local[1] + r)),
            _pt(*rot(cl_local[0], cl_local[1] + r))
        )
        apply_props(obj2, row); created.append(obj2)

    obj3 = mspace.AddArc(_pt(*cl), r, _deg(angle + 90), _deg(angle + 270))
    apply_props(obj3, row); created.append(obj3)

    if straight > 0:
        obj4 = mspace.AddLine(
            _pt(*rot(cl_local[0], cl_local[1] - r)),
            _pt(*rot(cr_local[0], cr_local[1] - r))
        )
        apply_props(obj4, row); created.append(obj4)

    return created


# ---------------------------------------------------------------------------
# SECTOR  (원호 + 양 끝 직사각형 막대)
# ---------------------------------------------------------------------------

def draw_sector(mspace, doc, row):
    """
    링 섹터: 원호(외호)의 양 끝에서 중심 방향으로 직사각형 막대가 뻗는 형태.

    start_angle=0, end_angle=90 예시 (CCW):

           90°
            │  ← 막대(ㅣ)
            │
      ──────┼──────────╮  ← 외호
      막대  │   내호    │
     (ㅡ)  ╰────────── ┤
            0°

    엔티티 4개:
      ① 외호  r_outer, start→end (CCW)
      ② 직선  end_angle 방향, r_outer → r_inner
      ③ 내호  r_inner, end→start (CW, 즉 역방향)
      ④ 직선  start_angle 방향, r_inner → r_outer
    """
    cx, cy      = _f(row, "x"),           _f(row, "y")
    r_inner     = _f(row, "r_inner",      0.0)
    r_outer     = _f(row, "r_outer",      0.0)
    start_angle = _f(row, "start_angle",  0.0)
    end_angle   = _f(row, "end_angle",    90.0)
    direction   = row.get("direction", "CCW").strip().upper()
    offset      = _f(row, "angle",        0.0)   # 전체 회전 오프셋

    if r_inner < 0 or r_outer <= 0:
        raise ValueError(f"r_outer={r_outer} 양수, r_inner={r_inner} 0 이상이어야 합니다.")
    if r_inner >= r_outer:
        raise ValueError(f"r_inner({r_inner}) < r_outer({r_outer}) 이어야 합니다.")

    # 전체 회전 오프셋 적용
    sa = start_angle + offset
    ea = end_angle   + offset

    # CW 지정 시 시작·끝 교환 (내부는 CCW API)
    if direction == "CW":
        sa, ea = ea, sa

    # 4개 꼭짓점 (극좌표 → 직교)
    p_outer_start = _polar(cx, cy, r_outer, sa)
    p_outer_end   = _polar(cx, cy, r_outer, ea)
    p_inner_end   = _polar(cx, cy, r_inner, ea)
    p_inner_start = _polar(cx, cy, r_inner, sa)

    created = []

    # ① 외호 (sa → ea, CCW)
    arc_outer = mspace.AddArc(_pt(cx, cy), r_outer, _deg(sa), _deg(ea))
    apply_props(arc_outer, row)
    created.append(arc_outer)

    # ② end_angle 쪽 막대 (r_outer → r_inner)
    line_end = mspace.AddLine(_pt(*p_outer_end), _pt(*p_inner_end))
    apply_props(line_end, row)
    created.append(line_end)

    # ③ 내호 (ea → sa, CW = AddArc에 순서 뒤집어서)
    if r_inner > 0:
        arc_inner = mspace.AddArc(_pt(cx, cy), r_inner, _deg(ea), _deg(sa))
        apply_props(arc_inner, row)
        created.append(arc_inner)
    else:
        # r_inner=0 이면 중심점으로 수렴 → 직선 2개만 (순수 부채꼴)
        line_inner = mspace.AddLine(_pt(*p_inner_end), _pt(*p_inner_start))
        apply_props(line_inner, row)
        created.append(line_inner)

    # ④ start_angle 쪽 막대 (r_inner → r_outer)
    line_start = mspace.AddLine(_pt(*p_inner_start), _pt(*p_outer_start))
    apply_props(line_start, row)
    created.append(line_start)

    return created


# ---------------------------------------------------------------------------
# CSV 읽기 & 실행
# ---------------------------------------------------------------------------

HANDLERS = {
    "RECTANGLE": draw_rectangle,
    "ARC":       draw_arc,
    "SLOT":      draw_slot,
    "SECTOR":    draw_sector,
}


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        return [{k.strip().lower(): v.strip() for k, v in row.items()}
                for row in reader]


def draw_from_csv(csv_path):
    if not HAS_WIN32:
        print("ERROR: pywin32 없음 →  pip install pywin32")
        sys.exit(1)

    if not os.path.isfile(csv_path):
        print(f"ERROR: 파일 없음 — {csv_path}")
        sys.exit(1)

    rows = read_csv(csv_path)
    print(f"CSV 로드: {len(rows)}행  ({csv_path})")

    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        print("ERROR: AutoCAD Mechanical이 실행 중이지 않습니다.")
        sys.exit(1)

    doc    = acad.ActiveDocument
    mspace = doc.ModelSpace
    print(f"연결됨: {doc.Name}\n")

    ok = err = skip = 0

    for i, row in enumerate(rows, 1):
        shape = row.get("shape", "").upper()
        if not shape:
            print(f"  [{i:03d}] SKIP — shape 컬럼 비어 있음")
            skip += 1
            continue

        handler = HANDLERS.get(shape)
        if handler is None:
            print(f"  [{i:03d}] SKIP — 미지원 도형: '{shape}'")
            skip += 1
            continue

        layer = row.get("layer", "0") or "0"
        ensure_layer(doc, layer)

        try:
            handler(mspace, doc, row)
            print(f"  [{i:03d}] OK    {shape:12s}  layer={layer}")
            ok += 1
        except Exception as e:
            print(f"  [{i:03d}] ERROR {shape:12s}  {e}")
            traceback.print_exc()
            err += 1

    doc.Regen(1)
    print(f"\n완료 — 성공: {ok}  오류: {err}  스킵: {skip}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("사용법:  python csv_to_autocad.py <CSV파일>")
        print("예시:    python csv_to_autocad.py shapes.csv")
        sys.exit(0)
    draw_from_csv(sys.argv[1])
