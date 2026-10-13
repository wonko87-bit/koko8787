"""Mesh generation (gmsh, OpenCASCADE kernel) -> plain numpy arrays."""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np

from .geometry import Polygon
from .problem import Problem


@dataclass
class Mesh:
    nodes: np.ndarray            # (N, 2) in metres
    tris: np.ndarray             # (M, 3) node indices, counter-clockwise
    region: np.ndarray           # (M,) index into Problem.regions
    # Dirichlet nodes: (node index, value) - value taken from the electrode
    # that owns the node (later entity wins at a shared corner).
    bc_nodes: np.ndarray         # (K,) int
    bc_values: np.ndarray        # (K,) complex
    # Per-electrode node sets, for post-processing: name -> node indices
    electrode_nodes: dict[str, np.ndarray] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    @property
    def n_nodes(self) -> int:
        return len(self.nodes)

    @property
    def n_tris(self) -> int:
        return len(self.tris)


def _add_polygon(occ, poly: Polygon, as_surface: bool):
    """Create OCC points / curves for `poly`; return (curve tags, surface tag)."""
    segs = poly.segments()
    pt_cache: dict[tuple[float, float], int] = {}

    def pt(p):
        key = (round(p[0], 9), round(p[1], 9))
        if key not in pt_cache:
            pt_cache[key] = occ.addPoint(p[0], p[1], 0.0)
        return pt_cache[key]

    curves = []
    for s in segs:
        a, b = pt(s.p0), pt(s.p1)
        if s.bulge == 0.0:
            curves.append(occ.addLine(a, b))
        else:
            (cx, cy), _, _, _ = s.arc()
            c = occ.addPoint(cx, cy, 0.0)
            if s.bulge > 0:
                curves.append(occ.addCircleArc(a, c, b))
            else:  # gmsh arcs run counter-clockwise: build b->a and flip
                curves.append(-occ.addCircleArc(b, c, a))
    if not as_surface:
        return curves, None
    loop = occ.addCurveLoop(curves)
    return curves, occ.addPlaneSurface([loop])


def build_mesh(problem: Problem) -> Mesh:
    import gmsh

    if not problem.regions:
        raise ValueError("problem has no regions")

    gmsh.initialize()
    try:
        gmsh.option.setNumber("General.Terminal", 0)
        gmsh.model.add("efield2d")
        occ = gmsh.model.occ

        # ---- inputs, in priority order (later wins) -------------------
        inputs: list[tuple[str, int, int, int]] = []  # (kind, index, dim, tag)
        for i, r in enumerate(problem.regions):
            _, s = _add_polygon(occ, r.polygon, True)
            inputs.append(("region", i, 2, s))
        for i, c in enumerate(problem.conductors):
            _, s = _add_polygon(occ, c.polygon, True)
            inputs.append(("conductor", i, 2, s))
        for i, ln in enumerate(problem.lines):
            curves, _ = _add_polygon(occ, ln.polygon, False)
            for c in curves:
                inputs.append(("line", i, 1, abs(c)))

        _, mapping = occ.fragment([(d, t) for _, _, d, t in inputs], [])

        # precedence of every surviving surface
        owner: dict[int, tuple[str, int]] = {}
        for (kind, idx, d, _), children in zip(inputs, mapping):
            if d != 2:
                continue
            for cd, ct in children:
                if cd == 2:
                    owner[ct] = (kind, idx)  # later input overwrites
        carved = [(2, t) for t, (k, _) in owner.items() if k == "conductor"]
        if carved:
            occ.remove(carved, recursive=False)
        occ.synchronize()

        # ---- size fields ----------------------------------------------
        h = problem.mesh_size
        fields = []
        curves_all = [t for _, t in gmsh.model.getEntities(1)]

        # electrode / Dirichlet outlines (needed before sizing)
        bboxes = [p.polygon.bbox() for p in problem.regions]
        diag = max(np.hypot(b[2] - b[0], b[3] - b[1]) for b in bboxes)
        tol = 1e-6 * diag

        def curve_samples(tag, n=5):
            lo, hi = gmsh.model.getParametrizationBounds(1, tag)
            lo, hi = lo[0], hi[0]
            ts = np.linspace(lo, hi, n)
            return np.array([gmsh.model.getValue(1, tag, [t])[:2] for t in ts])

        curve_owner: dict[int, tuple[str, int]] = {}
        for tag in curves_all:
            pts = curve_samples(tag)
            for kind, items in (("line", problem.lines), ("conductor", problem.conductors)):
                for i, it in enumerate(items):
                    if all(it.polygon.distance(x, y) < tol for x, y in pts):
                        curve_owner[tag] = (kind, i)  # conductors checked last -> win

        if problem.edge_size is not None:
            elec_curves = list(curve_owner.keys())
            if elec_curves:
                f = gmsh.model.mesh.field
                fd = f.add("Distance")
                f.setNumbers(fd, "CurvesList", elec_curves)
                f.setNumber(fd, "Sampling", 100)
                ft = f.add("Threshold")
                f.setNumber(ft, "InField", fd)
                f.setNumber(ft, "SizeMin", problem.edge_size)
                f.setNumber(ft, "SizeMax", h)
                f.setNumber(ft, "DistMin", 0.0)
                f.setNumber(ft, "DistMax", problem.edge_transition or 4.0 * h)
                fields.append(ft)
        for r in problem.refinements:
            f = gmsh.model.mesh.field
            fb = f.add("Ball")
            f.setNumber(fb, "Radius", r.radius)
            f.setNumber(fb, "VIn", r.size)
            f.setNumber(fb, "VOut", h)
            f.setNumber(fb, "Thickness", r.transition if r.transition is not None else 3.0 * h)
            f.setNumber(fb, "XCenter", r.x)
            f.setNumber(fb, "YCenter", r.y)
            f.setNumber(fb, "ZCenter", 0.0)
            fields.append(fb)
        if fields:
            f = gmsh.model.mesh.field
            fm = f.add("Min")
            f.setNumbers(fm, "FieldsList", fields)
            f.setAsBackgroundMesh(fm)
        gmsh.option.setNumber("Mesh.MeshSizeMax", h)
        gmsh.option.setNumber("Mesh.MeshSizeFromPoints", 0)
        gmsh.option.setNumber("Mesh.MeshSizeFromCurvature", 0)
        gmsh.option.setNumber("Mesh.MeshSizeExtendFromBoundary", 0 if fields else 1)

        gmsh.model.mesh.generate(2)

        # ---- extract ---------------------------------------------------
        tags, coords, _ = gmsh.model.mesh.getNodes()
        coords = coords.reshape(-1, 3)[:, :2]
        order = np.argsort(tags)
        tags, coords = tags[order], coords[order]

        tri_nodes, tri_region = [], []
        for dim, stag in gmsh.model.getEntities(2):
            kind, idx = owner.get(stag, (None, None))
            if kind != "region":
                continue
            etypes, _, enodes = gmsh.model.mesh.getElements(dim, stag)
            for et, nn in zip(etypes, enodes):
                if et != 2:  # 3-node triangle
                    continue
                nn = np.asarray(nn).reshape(-1, 3)
                tri_nodes.append(nn)
                tri_region.append(np.full(len(nn), idx, dtype=int))
        if not tri_nodes:
            raise RuntimeError("meshing produced no triangles")
        tri_nodes = np.vstack(tri_nodes)
        tri_region = np.concatenate(tri_region)

        used, inv = np.unique(tri_nodes, return_inverse=True)
        tris = inv.reshape(-1, 3)
        node_xy = coords[np.searchsorted(tags, used)]
        remap = {int(t): i for i, t in enumerate(used)}

        # electrode nodes
        bc: dict[int, complex] = {}
        electrode_nodes: dict[str, set] = {}
        for tag, (kind, i) in sorted(curve_owner.items(), key=lambda kv: (kv[1][0] == "conductor")):
            ent = problem.conductors[i] if kind == "conductor" else problem.lines[i]
            ntags, _, _ = gmsh.model.mesh.getNodes(1, tag, includeBoundary=True)
            idxs = [remap[int(t)] for t in ntags if int(t) in remap]
            name = ent.name or f"{kind}{i}"
            electrode_nodes.setdefault(name, set()).update(idxs)
            for n in idxs:
                bc[n] = complex(ent.voltage)

        # an electrode that ended up with no node in the mesh is not attached to anything
        warns: list[str] = []
        for i, ln in enumerate(problem.lines):
            name = ln.name or f"line{i}"
            if len(electrode_nodes.get(name, ())) == 0:
                raise ValueError(
                    f"Dirichlet line {name!r} does not lie on or inside any region - it would be "
                    f"ignored and the electrode would have no effect. Check that it touches the "
                    f"region outline (tolerance {tol:.1e} mm).")
        for i, c in enumerate(problem.conductors):
            name = c.name or f"conductor{i}"
            if len(electrode_nodes.get(name, ())) == 0:
                warns.append(f"conductor {name!r} does not touch any region and is ignored")

        # counter-clockwise orientation
        p = node_xy[tris]
        area2 = (p[:, 1, 0] - p[:, 0, 0]) * (p[:, 2, 1] - p[:, 0, 1]) \
            - (p[:, 2, 0] - p[:, 0, 0]) * (p[:, 1, 1] - p[:, 0, 1])
        flip = area2 < 0
        tris[flip] = tris[flip][:, [0, 2, 1]]

        bc_nodes = np.array(sorted(bc), dtype=int)
        bc_values = np.array([bc[n] for n in bc_nodes], dtype=complex)
        return Mesh(
            nodes=node_xy * problem.unit,
            tris=tris,
            region=tri_region,
            bc_nodes=bc_nodes,
            bc_values=bc_values,
            electrode_nodes={k: np.array(sorted(v), dtype=int) for k, v in electrode_nodes.items()},
            warnings=warns,
        )
    finally:
        gmsh.finalize()
