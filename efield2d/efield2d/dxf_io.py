"""DXF import -> shapes -> Problem.

Supported entities: LINE, ARC, CIRCLE, LWPOLYLINE, POLYLINE (2D), SPLINE,
ELLIPSE and INSERT (blocks are exploded).  SPLINE / ELLIPSE are approximated
by short straight lines whose deviation from the true curve is at most
`flatten_tol` (default 0.01 mm).  Anything else is skipped and reported in
`DxfModel.warnings`.

How shapes are formed
  * a closed LWPOLYLINE / POLYLINE, a CIRCLE and a closed SPLINE / ELLIPSE are
    one closed shape each (draw every region / electrode as its own closed
    outline - overlapping and touching outlines are fine, see Problem);
  * loose LINE / ARC / open polyline / open spline pieces are joined end to end
    (within `snap`) into chains - closed chains become closed shapes, the rest
    are open lines (usable as Dirichlet lines).  A chain stops at a junction
    where three or more pieces meet.

The model is converted to millimetres using $INSUNITS (override with
`unit_scale`, the factor from DXF units to mm).
"""
from __future__ import annotations

import math
from collections import Counter, defaultdict
from dataclasses import dataclass, field

import numpy as np

from .geometry import Polygon
from .materials import Material
from .problem import Conductor, DirichletLine, Problem, Region

# primitive: (x0, y0, x1, y1, bulge)
_Prim = tuple[float, float, float, float, float]


@dataclass
class Shape:
    id: int
    layer: str
    polygon: Polygon
    closed: bool
    source: str                     # e.g. "LWPOLYLINE", "CIRCLE", "chain"
    area: float = 0.0               # mm^2, closed shapes only

    def bbox(self):
        return self.polygon.bbox()


@dataclass
class DxfModel:
    shapes: list[Shape] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    counts: dict[str, int] = field(default_factory=dict)   # entity type -> count read
    unit_scale: float = 1.0                                 # DXF unit -> mm

    def bounds(self):
        bb = [s.bbox() for s in self.shapes]
        return (min(b[0] for b in bb), min(b[1] for b in bb),
                max(b[2] for b in bb), max(b[3] for b in bb))

    def layers(self) -> list[str]:
        return sorted({s.layer for s in self.shapes})

    def describe(self) -> str:
        lines = [f"entities read: {dict(self.counts)}", f"unit scale to mm: {self.unit_scale}"]
        for s in self.shapes:
            b = s.bbox()
            tag = f"closed, area {s.area:.3f} mm2" if s.closed else "open"
            lines.append(f"  [{s.id}] {s.layer!r:12} {s.source:10} {tag}  "
                         f"bbox=({b[0]:.3f},{b[1]:.3f})-({b[2]:.3f},{b[3]:.3f})  "
                         f"{len(s.polygon.vertices)} vertices")
        lines += [f"WARNING: {w}" for w in self.warnings]
        return "\n".join(lines)


# ---------------------------------------------------------------- conversions
def _poly_from_pts(pts, closed: bool) -> Polygon:
    return Polygon([(float(x), float(y), 0.0) for x, y in pts], closed)


def _arc_bulge(sweep: float) -> float:
    return math.tan(sweep / 4.0)


def _circle_polygon(cx, cy, r) -> Polygon:
    # two CCW semicircles
    return Polygon([(cx + r, cy, 1.0), (cx - r, cy, 1.0)], True)


def _drop_duplicate_end(pts, tol):
    if len(pts) > 2 and math.hypot(pts[0][0] - pts[-1][0], pts[0][1] - pts[-1][1]) <= tol:
        return pts[:-1]
    return pts


_INSUNITS_WARNED = "$INSUNITS is 0 (unitless): assuming millimetres"


def _expand(entities, depth=0):
    """Yield entities, exploding INSERTs recursively."""
    for e in entities:
        if e.dxftype() == "INSERT" and depth < 8:
            try:
                yield from _expand(list(e.virtual_entities()), depth + 1)
            except Exception:  # pragma: no cover - defensive
                yield e
        else:
            yield e


def import_dxf(path: str, *, unit_scale: float | None = None, flatten_tol: float = 0.01,
               snap: float = 1e-3, layers: list[str] | None = None) -> DxfModel:
    """Read `path` and return the shapes (mm). `flatten_tol` and `snap` are in mm."""
    import ezdxf
    from ezdxf import path as ezpath
    from ezdxf import units

    doc = ezdxf.readfile(path)
    model = DxfModel()

    if unit_scale is None:
        ins = int(doc.header.get("$INSUNITS", 0))
        if ins == 0:
            model.warnings.append(_INSUNITS_WARNED)
            unit_scale = 1.0
        else:
            unit_scale = units.conversion_factor(ins, units.MM)
    model.unit_scale = unit_scale
    k = unit_scale
    ftol = flatten_tol / k      # in DXF units, for the flattening step

    closed_shapes: list[tuple[str, str, Polygon]] = []   # (layer, source, polygon)
    prims: list[tuple[str, _Prim]] = []                  # (layer, primitive)
    counts: Counter = Counter()
    skipped: Counter = Counter()

    def add_chain_pts(layer, pts):
        for a, b in zip(pts[:-1], pts[1:]):
            if a != b:
                prims.append((layer, (a[0], a[1], b[0], b[1], 0.0)))

    for e in _expand(list(doc.modelspace())):
        t = e.dxftype()
        layer = e.dxf.layer
        if layers is not None and layer not in layers:
            continue
        counts[t] += 1
        if t == "LINE":
            s, d = e.dxf.start, e.dxf.end
            if (s.x, s.y) != (d.x, d.y):
                prims.append((layer, (s.x, s.y, d.x, d.y, 0.0)))
        elif t == "ARC":
            c, r = e.dxf.center, e.dxf.radius
            a0 = math.radians(e.dxf.start_angle)
            sweep = math.radians((e.dxf.end_angle - e.dxf.start_angle) % 360.0)
            if sweep == 0.0:
                sweep = 2.0 * math.pi
            nseg = 1 if sweep <= math.pi * 1.5 else 2   # keep each arc well below a full turn
            for i in range(nseg):
                s0 = a0 + sweep * i / nseg
                s1 = a0 + sweep * (i + 1) / nseg
                prims.append((layer, (c.x + r * math.cos(s0), c.y + r * math.sin(s0),
                                      c.x + r * math.cos(s1), c.y + r * math.sin(s1),
                                      _arc_bulge(s1 - s0))))
        elif t == "CIRCLE":
            c = e.dxf.center
            closed_shapes.append((layer, t, _circle_polygon(c.x, c.y, e.dxf.radius)))
        elif t == "LWPOLYLINE":
            pts = [(x, y, b) for x, y, b in e.get_points("xyb")]
            if len(pts) < 2:
                continue
            if e.closed:
                closed_shapes.append((layer, t, Polygon(pts, True)))
            else:
                for (x0, y0, b), (x1, y1, _) in zip(pts[:-1], pts[1:]):
                    prims.append((layer, (x0, y0, x1, y1, b)))
        elif t == "POLYLINE":
            if not e.is_2d_polyline:
                skipped["POLYLINE(3D)"] += 1
                continue
            pts = [(v.dxf.location.x, v.dxf.location.y, v.dxf.bulge) for v in e.vertices]
            if len(pts) < 2:
                continue
            if e.is_closed:
                closed_shapes.append((layer, t, Polygon(pts, True)))
            else:
                for (x0, y0, b), (x1, y1, _) in zip(pts[:-1], pts[1:]):
                    prims.append((layer, (x0, y0, x1, y1, b)))
        elif t in ("SPLINE", "ELLIPSE"):
            p = ezpath.make_path(e)
            pts = [(v.x, v.y) for v in p.flattening(ftol, segments=2)]
            if len(pts) < 2:
                continue
            if math.hypot(pts[0][0] - pts[-1][0], pts[0][1] - pts[-1][1]) <= snap / k:
                pts = _drop_duplicate_end(pts, snap / k)
                closed_shapes.append((layer, t, _poly_from_pts(pts, True)))
            else:
                add_chain_pts(layer, pts)
        else:
            skipped[t] += 1

    model.counts = dict(counts)
    for t, n in skipped.items():
        model.warnings.append(f"skipped {n} unsupported {t} entit{'y' if n == 1 else 'ies'}")

    # scale to mm -------------------------------------------------------
    def scale_poly(p: Polygon) -> Polygon:
        return Polygon([(x * k, y * k, b) for x, y, b in p.vertices], p.closed)

    sid = 0
    for layer, src, poly in closed_shapes:
        poly = scale_poly(poly)
        if poly.signed_area() < 0:
            poly = poly.reversed()
        model.shapes.append(Shape(sid, layer, poly, True, src, poly.signed_area()))
        sid += 1

    scaled = [(layer, (a * k, b * k, c * k, d * k, bu)) for layer, (a, b, c, d, bu) in prims]
    for layer, poly, closed in _chains(scaled, snap):
        if closed and poly.signed_area() < 0:
            poly = poly.reversed()
        model.shapes.append(Shape(sid, layer, poly, closed, "chain",
                                  poly.signed_area() if closed else 0.0))
        sid += 1
        if not closed:
            g = math.hypot(poly.vertices[0][0] - poly.vertices[-1][0],
                           poly.vertices[0][1] - poly.vertices[-1][1])
            if g < 1.0 and len(poly.vertices) > 2:
                model.warnings.append(
                    f"open chain [{sid - 1}] on layer {layer!r} has a gap of {g:.4f} mm between its "
                    f"ends - it will not form a closed region (check the drawing or raise `snap`)")
    return model


# ---------------------------------------------------------------- chain joining
def _chains(prims: list[tuple[str, _Prim]], snap: float):
    """Join primitives end to end. Yields (layer, Polygon, closed)."""
    from scipy.spatial import cKDTree

    if not prims:
        return
    ends = np.array([[p[0], p[1]] for _, p in prims] + [[p[2], p[3]] for _, p in prims])
    n = len(prims)
    # cluster end points closer than `snap`
    tree = cKDTree(ends)
    parent = list(range(2 * n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    for i, j in tree.query_pairs(snap):
        parent[find(i)] = find(j)
    node_of = [find(i) for i in range(2 * n)]
    pos = {}
    for i in range(2 * n):
        pos.setdefault(node_of[i], tuple(ends[i]))
    adj = defaultdict(list)             # node -> [(prim index, end 0/1)]
    for i in range(n):
        adj[node_of[i]].append((i, 0))
        adj[node_of[n + i]].append((i, 1))

    used = [False] * n

    def walk(start_node, i0, end0):
        """Follow primitives from start_node via prim i0 (leaving through its
        `end0` side). Returns the list of oriented prims and the final node."""
        chain, node, i, e = [], start_node, i0, end0
        while True:
            used[i] = True
            (x0, y0, x1, y1, b) = prims[i][1]
            if e == 0:
                chain.append((x0, y0, x1, y1, b))
                nxt = node_of[n + i]
            else:
                chain.append((x1, y1, x0, y0, -b))
                nxt = node_of[i]
            node = nxt
            if len(adj[node]) != 2:
                return chain, node
            cand = [(j, ee) for j, ee in adj[node] if not used[j]]
            if not cand:
                return chain, node
            i, e = cand[0]

    def to_poly(chain, closed, layer):
        verts = [(c[0], c[1], c[4]) for c in chain]
        if not closed:
            verts.append((chain[-1][2], chain[-1][3], 0.0))
        return layer, Polygon(verts, closed), closed

    # chains that start at an end or a junction ...
    for node, lst in list(adj.items()):
        if len(lst) == 2:
            continue
        for i, e in lst:
            if used[i]:
                continue
            chain, last = walk(node, i, e)
            yield to_poly(chain, last == node and len(chain) > 1, prims[i][0])
    # ... and what is left are pure loops
    for i in range(n):
        if used[i]:
            continue
        chain, _ = walk(node_of[i], i, 0)
        yield to_poly(chain, True, prims[i][0])


# ---------------------------------------------------------------- assignment
@dataclass
class Assignment:
    """What a shape is: region (with material), conductor (with voltage),
    line (Dirichlet line with voltage) or ignore."""
    kind: str                              # "region" | "conductor" | "line" | "ignore"
    material: Material | None = None
    voltage: complex = 0.0
    name: str = ""


def build_problem(model: DxfModel, assignments: dict[int, Assignment], *,
                  symmetry: str = "planar", **problem_kwargs) -> Problem:
    """Assemble a Problem from DXF shapes and their assignments.

    Regions are ordered by decreasing area so the biggest (e.g. the air box)
    is laid down first and smaller bodies drawn inside it override it.
    Extra keyword arguments go to Problem (mesh_size, edge_size, ...)."""
    by_id = {s.id: s for s in model.shapes}
    regions, conductors, lines = [], [], []
    for sid, a in assignments.items():
        if sid not in by_id:
            raise KeyError(f"no shape with id {sid}")
        s = by_id[sid]
        if a.kind == "ignore":
            continue
        if a.kind in ("region", "conductor") and not s.closed:
            raise ValueError(f"shape [{sid}] is an open line and cannot be a {a.kind}")
        if symmetry == "axisym" and s.bbox()[0] < -1e-6:
            raise ValueError(f"shape [{sid}] extends to x < 0, which is not allowed for an "
                             f"axisymmetric model (x = r)")
        name = a.name or f"{a.kind}{sid}"
        if a.kind == "region":
            if a.material is None:
                raise ValueError(f"region shape [{sid}] needs a material")
            regions.append((s.area, Region(s.polygon, a.material, name)))
        elif a.kind == "conductor":
            conductors.append(Conductor(s.polygon, a.voltage, name))
        elif a.kind == "line":
            lines.append(DirichletLine(s.polygon, a.voltage, name))
        else:
            raise ValueError(f"unknown assignment kind {a.kind!r}")
    regions.sort(key=lambda t: -t[0])
    return Problem(regions=[r for _, r in regions], conductors=conductors, lines=lines,
                   symmetry=symmetry, **problem_kwargs)


if __name__ == "__main__":  # python -m efield2d.dxf_io file.dxf
    import sys
    if len(sys.argv) < 2:
        raise SystemExit("usage: python -m efield2d.dxf_io file.dxf")
    print(import_dxf(sys.argv[1]).describe())
