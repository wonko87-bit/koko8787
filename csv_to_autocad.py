"""
CSV to AutoCAD Mechanical - Rectangle & Arc Drawing Script
===========================================================

AutoCAD Mechanical이 열린 상태에서 실행하면 CSV 데이터를 읽어
모델 공간에 직사각형과 원호를 그립니다.

[CSV 컬럼 설명]
  공통
    shape   : RECTANGLE | ARC
    layer   : 레이어 이름 (없으면 "0")
    color   : AutoCAD 색상 인덱스 1~256 (0 = BYLAYER)

  RECTANGLE 전용
    x, y    : 사각형 기준점 (중심 or 좌하단 — origin 컬럼으로 선택)
    origin  : CENTER(기본) | CORNER  — x,y 기준점 의미
    width   : 가로 길이
    height  : 세로 길이
    angle   : 회전각도 (도, 반시계 방향 기준, 기본 0)

  ARC 전용
    x, y        : 원호 중심점
    radius      : 반지름
    start_angle : 시작각 (도, 3시 방향 = 0, 반시계 방향)
    end_angle   : 끝각   (도)

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
# Utilities
# ---------------------------------------------------------------------------

def _f(row, key, default=0.0):
    val = row.get(key, "").strip()
    try:
        return float(val) if val else default
    except ValueError:
        return default


def _i(row, key, default=0):
    val = row.get(key, "").strip()
    try:
        return int(val) if val else default
    except ValueError:
        return default


def _pt(x, y, z=0.0):
    """VARIANT 3D point (win32com용)."""
    return win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), float(z))
    )


def _deg2rad(deg):
    return math.radians(deg)


def ensure_layer(doc, name):
    if not name or name == "0":
        return
    try:
        doc.Layers.Item(name)
    except Exception:
        doc.Layers.Add(name)


def apply_props(obj, row):
    """레이어·색상 공통 적용."""
    layer = row.get("layer", "").strip() or "0"
    obj.Layer = layer
    color = _i(row, "color", 0)
    if color:
        obj.color = color


# ---------------------------------------------------------------------------
# RECTANGLE
# ---------------------------------------------------------------------------

def _rotate_point(cx, cy, px, py, rad):
    """점 (px,py)을 중심 (cx,cy) 기준으로 rad 라디안 회전."""
    dx, dy = px - cx, py - cy
    rx = cx + dx * math.cos(rad) - dy * math.sin(rad)
    ry = cy + dx * math.sin(rad) + dy * math.cos(rad)
    return rx, ry


def draw_rectangle(mspace, doc, row):
    x, y   = _f(row, "x"),      _f(row, "y")
    w, h   = _f(row, "width"),  _f(row, "height")
    angle  = _f(row, "angle",  0.0)
    origin = row.get("origin", "CENTER").strip().upper()

    if w <= 0 or h <= 0:
        raise ValueError(f"width={w}, height={h} 모두 양수여야 합니다.")

    # 기준점에 따라 좌하단 코너 계산
    if origin == "CORNER":
        cx, cy = x + w / 2, y + h / 2   # 중심
        corners = [
            (x,     y),
            (x + w, y),
            (x + w, y + h),
            (x,     y + h),
        ]
    else:  # CENTER (기본)
        cx, cy = x, y
        corners = [
            (x - w / 2, y - h / 2),
            (x + w / 2, y - h / 2),
            (x + w / 2, y + h / 2),
            (x - w / 2, y + h / 2),
        ]

    # 각도 회전 적용
    rad = _deg2rad(angle)
    rotated = [_rotate_point(cx, cy, px, py, rad) for px, py in corners]

    # 닫힌 폴리라인 (5점, 마지막 = 첫점)
    pts_flat = []
    for px, py in rotated:
        pts_flat += [px, py, 0.0]
    # 닫기: 시작점 반복
    pts_flat += list(rotated[0]) + [0.0]

    pts_var = win32com.client.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8, pts_flat
    )
    obj = mspace.Add3DPoly(pts_var)   # 3D Polyline (각도 회전 후)

    # 닫힌 것처럼 보이게 마지막 선분은 이미 포함됨 (시작점 반복)
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

    if radius <= 0:
        raise ValueError(f"radius={radius} 는 양수여야 합니다.")

    center = _pt(x, y)
    obj = mspace.AddArc(
        center,
        radius,
        _deg2rad(start_angle),
        _deg2rad(end_angle),
    )
    apply_props(obj, row)
    return obj


# ---------------------------------------------------------------------------
# CSV reader & main
# ---------------------------------------------------------------------------

HANDLERS = {
    "RECTANGLE": draw_rectangle,
    "ARC":       draw_arc,
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

    rows = read_csv(csv_path)
    print(f"CSV 로드: {len(rows)}행  ({csv_path})")

    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        print("ERROR: AutoCAD Mechanical이 실행 중이지 않습니다.")
        sys.exit(1)

    doc     = acad.ActiveDocument
    mspace  = doc.ModelSpace
    print(f"연결됨: {doc.Name}\n")

    ok = err = skip = 0

    for i, row in enumerate(rows, 1):
        shape = row.get("shape", "").upper()
        if not shape:
            print(f"  [{i:03d}] SKIP — shape 컬럼 없음")
            skip += 1
            continue

        handler = HANDLERS.get(shape)
        if handler is None:
            print(f"  [{i:03d}] SKIP — 지원하지 않는 도형: '{shape}'")
            skip += 1
            continue

        layer = row.get("layer", "0") or "0"
        ensure_layer(doc, layer)

        try:
            obj = handler(mspace, doc, row)
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
