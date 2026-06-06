"""
CSV to AutoCAD Mechanical - Rectangle / Arc / Slot Drawing Script
==================================================================

[지원 도형]
  RECTANGLE  직사각형 (회전 각도 지원)
  ARC        원호     (시계/반시계 방향 지원)
  SLOT       장공     (직사각형 양 끝에 반원이 붙은 형태)

[공통 컬럼]
  shape      : RECTANGLE | ARC | SLOT
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
  width      : 전체 길이  (긴 방향)
  height     : 전체 폭    (짧은 방향, 반원 지름)
  angle      : 회전각 (도)
  ※ 반원 반지름 = height / 2, 직선부 길이 = width - height

  내부 구조:
    ┌──────── width ────────┐
    │   ╭──────────────╮   │ ← 상단 직선
    │  (  직선부 길이    )  │ ← 양쪽 반원 (radius = height/2)
    │   ╰──────────────╯   │ ← 하단 직선
    height

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
    """점 (px,py)를 중심 (cx,cy) 기준으로 rad 라디안 회전."""
    dx, dy = px - cx, py - cy
    return (
        cx + dx * math.cos(rad) - dy * math.sin(rad),
        cy + dx * math.sin(rad) + dy * math.cos(rad),
    )


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

    # 중심 계산
    if origin == "CENTER":
        cx, cy = x, y
        corners = [
            (x - w / 2, y - h / 2),
            (x + w / 2, y - h / 2),
            (x + w / 2, y + h / 2),
            (x - w / 2, y + h / 2),
        ]
    else:  # CORNER (기본) — x,y = 좌하단
        cx, cy = x + w / 2, y + h / 2
        corners = [
            (x,     y),
            (x + w, y),
            (x + w, y + h),
            (x,     y + h),
        ]

    rad = _deg(angle)
    rotated = [_rotate(cx, cy, px, py, rad) for px, py in corners]

    # 닫힌 3D 폴리라인 (시작점 반복으로 닫기)
    flat = []
    for px, py in rotated:
        flat += [px, py, 0.0]
    flat += list(rotated[0]) + [0.0]   # 마지막 = 첫점

    pts = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
    obj = mspace.Add3DPoly(pts)
    apply_props(obj, row)
    return obj


# ---------------------------------------------------------------------------
# ARC
# ---------------------------------------------------------------------------

def draw_arc(mspace, doc, row):
    x, y        = _f(row, "x"),           _f(row, "y")
    radius      = _f(row, "radius",       0.0)
    start_angle = _f(row, "start_angle",  0.0)
    end_angle   = _f(row, "end_angle",    90.0)
    direction   = row.get("direction", "CCW").strip().upper()

    if radius <= 0:
        raise ValueError(f"radius={radius} 는 양수여야 합니다.")

    # CW 지정 시 각도를 반전하여 AutoCAD CCW API에 맞춤
    if direction == "CW":
        start_angle, end_angle = -start_angle, -end_angle

    obj = mspace.AddArc(_pt(x, y), radius, _deg(start_angle), _deg(end_angle))
    apply_props(obj, row)
    return obj


# ---------------------------------------------------------------------------
# SLOT  (직사각형 + 양 끝 반원)
# ---------------------------------------------------------------------------

def draw_slot(mspace, doc, row):
    """
    장공(슬롯) = 직선부 직사각형 + 양쪽 반원 2개.

    내부 그리기 순서 (angle=0 기준, 좌→우 방향):
      ① 우측 반원 (중심 = Cr, 시작 -90°, 끝 +90°)
      ② 상단 직선 (우상 → 좌상)
      ③ 좌측 반원 (중심 = Cl, 시작 +90°, 끝 -90°  / 즉 90°→270°)
      ④ 하단 직선 (좌하 → 우하)

    회전각(angle)은 각 엔티티를 개별 회전하여 적용.
    """
    x, y   = _f(row, "x"),     _f(row, "y")
    w, h   = _f(row, "width"), _f(row, "height")
    angle  = _f(row, "angle",  0.0)
    origin = row.get("origin", "CORNER").strip().upper()

    if w <= 0 or h <= 0:
        raise ValueError(f"width={w}, height={h} 모두 양수여야 합니다.")
    if h > w:
        raise ValueError(f"SLOT: height({h}) > width({w}) — width 가 긴 방향이어야 합니다.")

    r = h / 2          # 반원 반지름
    straight = w - h   # 직선부 길이

    if straight < 0:
        raise ValueError("SLOT: width 가 height 보다 커야 합니다.")

    # angle=0 기준 좌표계에서 중심 계산
    if origin == "CENTER":
        cx, cy = x, y
    else:  # CORNER — x,y = 바운딩 박스 좌하단
        cx, cy = x + w / 2, y + h / 2

    # 직선부 좌·우 중심
    # (angle=0 기준: 우측 반원 중심 = cx + straight/2, 좌측 = cx - straight/2)
    rad = _deg(angle)

    def rot(px, py):
        return _rotate(cx, cy, px, py, rad)

    cr_local = (cx + straight / 2, cy)   # 우측 반원 중심 (로컬)
    cl_local = (cx - straight / 2, cy)   # 좌측 반원 중심 (로컬)
    cr = rot(*cr_local)
    cl = rot(*cl_local)

    created = []

    # ① 우측 반원 (로컬: 270° → 90°, CCW = 반시계)
    sa_r = _deg(angle + 270)   # 회전 후 각도 = 로컬각 + angle
    ea_r = _deg(angle + 90)
    obj1 = mspace.AddArc(_pt(*cr), r, sa_r, ea_r)
    apply_props(obj1, row)
    created.append(obj1)

    # ② 상단 직선 (우상 → 좌상)
    # 우상 = cr_local + (0, r) 회전
    # 좌상 = cl_local + (0, r) 회전
    top_r = rot(cr_local[0], cr_local[1] + r)
    top_l = rot(cl_local[0], cl_local[1] + r)
    if straight > 0:
        obj2 = mspace.AddLine(_pt(*top_r), _pt(*top_l))
        apply_props(obj2, row)
        created.append(obj2)

    # ③ 좌측 반원 (로컬: 90° → 270°, CCW)
    sa_l = _deg(angle + 90)
    ea_l = _deg(angle + 270)
    obj3 = mspace.AddArc(_pt(*cl), r, sa_l, ea_l)
    apply_props(obj3, row)
    created.append(obj3)

    # ④ 하단 직선 (좌하 → 우하)
    bot_l = rot(cl_local[0], cl_local[1] - r)
    bot_r = rot(cr_local[0], cr_local[1] - r)
    if straight > 0:
        obj4 = mspace.AddLine(_pt(*bot_l), _pt(*bot_r))
        apply_props(obj4, row)
        created.append(obj4)

    return created   # 여러 엔티티 반환


# ---------------------------------------------------------------------------
# CSV 읽기 & 실행
# ---------------------------------------------------------------------------

HANDLERS = {
    "RECTANGLE": draw_rectangle,
    "ARC":       draw_arc,
    "SLOT":      draw_slot,
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
