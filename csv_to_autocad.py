"""
CSV to AutoCAD Mechanical - Shape Drawing Script
=================================================

Reads shape data from a CSV file and draws them in AutoCAD Mechanical
via COM automation (win32com).

CSV Format (header row required):
  shape, x, y, width, height, radius, layer, color

  shape   : LINE | CIRCLE | RECTANGLE | POLYLINE
  x, y    : insertion point (drawing units)
  width   : width  (RECTANGLE / LINE endpoint-X offset)
  height  : height (RECTANGLE / LINE endpoint-Y offset)
  radius  : radius (CIRCLE)
  layer   : target layer name (created if missing)
  color   : AutoCAD color index 0-256 (0 = BYLAYER)

Example CSV rows:
  CIRCLE,100,100,,,50,CENTER,1
  RECTANGLE,200,200,80,40,,WALLS,2
  LINE,0,0,150,150,,,3

Usage (run from Windows with AutoCAD Mechanical open):
  python csv_to_autocad.py shapes.csv

Requirements:
  pip install pywin32
"""

import csv
import sys
import os
import traceback

try:
    import win32com.client
    import pythoncom
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _float(val, default=0.0):
    """Convert string to float safely."""
    try:
        return float(val) if val.strip() != "" else default
    except (ValueError, AttributeError):
        return default


def _int(val, default=0):
    try:
        return int(val) if val.strip() != "" else default
    except (ValueError, AttributeError):
        return default


def point(x, y, z=0.0):
    """Return a VARIANT array for an AutoCAD point."""
    import win32com.client as wcc
    pt = wcc.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))
    return pt


def ensure_layer(doc, layer_name, color_index=7):
    """Create layer if it does not exist, then return it."""
    try:
        return doc.Layers.Item(layer_name)
    except Exception:
        layer = doc.Layers.Add(layer_name)
        layer.color = color_index
        return layer


# ---------------------------------------------------------------------------
# Shape drawing functions
# ---------------------------------------------------------------------------

def draw_circle(mspace, row):
    center = point(_float(row["x"]), _float(row["y"]))
    radius = _float(row.get("radius", 0))
    if radius <= 0:
        print(f"  [SKIP] CIRCLE at ({row['x']},{row['y']}) has radius <= 0")
        return None
    obj = mspace.AddCircle(center, radius)
    return obj


def draw_rectangle(mspace, row):
    x, y = _float(row["x"]), _float(row["y"])
    w, h = _float(row.get("width", 0)), _float(row.get("height", 0))
    if w == 0 or h == 0:
        print(f"  [SKIP] RECTANGLE at ({x},{y}) has zero width/height")
        return None
    import win32com.client as wcc
    pts = wcc.VARIANT(
        pythoncom.VT_ARRAY | pythoncom.VT_R8,
        (x, y, 0,
         x + w, y, 0,
         x + w, y + h, 0,
         x, y + h, 0,
         x, y, 0),
    )
    obj = mspace.AddLightWeightPolyline(pts)
    obj.Closed = True
    return obj


def draw_line(mspace, row):
    x1, y1 = _float(row["x"]), _float(row["y"])
    x2 = x1 + _float(row.get("width", 0))
    y2 = y1 + _float(row.get("height", 0))
    obj = mspace.AddLine(point(x1, y1), point(x2, y2))
    return obj


def draw_polyline(mspace, row):
    """
    POLYLINE shape reads extra columns pts_x and pts_y as
    semicolon-separated coordinate lists.
    e.g.  pts_x = "0;50;100;150"   pts_y = "0;30;0;30"
    """
    raw_x = row.get("pts_x", "")
    raw_y = row.get("pts_y", "")
    xs = [_float(v) for v in raw_x.split(";") if v.strip()]
    ys = [_float(v) for v in raw_y.split(";") if v.strip()]
    if len(xs) < 2 or len(xs) != len(ys):
        print(f"  [SKIP] POLYLINE: need matching pts_x / pts_y with >=2 points")
        return None
    import win32com.client as wcc
    flat = []
    for xi, yi in zip(xs, ys):
        flat += [xi, yi, 0.0]
    pts = wcc.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, flat)
    obj = mspace.AddLightWeightPolyline(pts)
    return obj


SHAPE_HANDLERS = {
    "CIRCLE":    draw_circle,
    "RECTANGLE": draw_rectangle,
    "LINE":      draw_line,
    "POLYLINE":  draw_polyline,
}


# ---------------------------------------------------------------------------
# CSV reader
# ---------------------------------------------------------------------------

def read_csv(filepath):
    """Return list of row dicts with lowercased keys."""
    rows = []
    with open(filepath, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append({k.strip().lower(): v.strip() for k, v in row.items()})
    return rows


# ---------------------------------------------------------------------------
# Main driver
# ---------------------------------------------------------------------------

def draw_from_csv(csv_path):
    if not HAS_WIN32:
        print("ERROR: pywin32 is not installed. Run:  pip install pywin32")
        sys.exit(1)

    if not os.path.isfile(csv_path):
        print(f"ERROR: CSV file not found: {csv_path}")
        sys.exit(1)

    rows = read_csv(csv_path)
    print(f"Loaded {len(rows)} row(s) from '{csv_path}'")

    # Connect to running AutoCAD Mechanical instance
    try:
        acad = win32com.client.GetActiveObject("AutoCAD.Application")
    except Exception:
        print("ERROR: AutoCAD Mechanical is not running or COM registration failed.")
        sys.exit(1)

    doc = acad.ActiveDocument
    mspace = doc.ModelSpace
    print(f"Connected to: {doc.Name}")

    drawn = 0
    skipped = 0

    for i, row in enumerate(rows, start=1):
        shape = row.get("shape", "").upper()
        if not shape:
            print(f"  [SKIP] Row {i}: missing 'shape' column")
            skipped += 1
            continue

        handler = SHAPE_HANDLERS.get(shape)
        if handler is None:
            print(f"  [SKIP] Row {i}: unknown shape '{shape}'")
            skipped += 1
            continue

        try:
            obj = handler(mspace, row)
            if obj is None:
                skipped += 1
                continue

            # Apply layer
            layer_name = row.get("layer", "").strip() or "0"
            ensure_layer(doc, layer_name)
            obj.Layer = layer_name

            # Apply color
            color_idx = _int(row.get("color", "0"))
            if color_idx != 0:
                obj.color = color_idx

            print(f"  [OK]   Row {i}: {shape} on layer '{layer_name}'")
            drawn += 1

        except Exception as exc:
            print(f"  [ERR]  Row {i}: {exc}")
            traceback.print_exc()
            skipped += 1

    doc.Regen(1)  # acRegen = 1
    print(f"\nDone. Drawn: {drawn}  Skipped/Errors: {skipped}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python csv_to_autocad.py <path_to_csv>")
        print("\nExample: python csv_to_autocad.py shapes.csv")
        sys.exit(0)

    draw_from_csv(sys.argv[1])
