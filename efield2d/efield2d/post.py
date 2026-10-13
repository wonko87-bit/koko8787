"""Post-processing: fields, path sampling (creepage), integrals, withstand check.

Conventions
  * lengths of the problem (geometry units, mm) are used for path input,
    everything returned is SI (V/m, F, ...).
  * AC results are peak phasors if the electrode voltages were given as peak
    values; |E| is then the peak field magnitude.
"""
from __future__ import annotations

import numpy as np
from matplotlib.tri import Triangulation

from .fem import Solution, element_geometry, volume_weight
from .geometry import Polygon


# --------------------------------------------------------------------- fields
def element_field(sol: Solution) -> np.ndarray:
    """E = -grad V, constant per element. Shape (M, 2), complex for AC."""
    mesh = sol.mesh
    b, c, area, _ = element_geometry(mesh)
    v = sol.phi[mesh.tris]
    ex = -(b * v).sum(axis=1) / (2.0 * area)
    ey = -(c * v).sum(axis=1) / (2.0 * area)
    return np.stack([ex, ey], axis=1)


def element_magnitude(sol: Solution) -> np.ndarray:
    e = element_field(sol)
    return np.sqrt((np.abs(e) ** 2).sum(axis=1))


def nodal_average(sol: Solution, values: np.ndarray) -> np.ndarray:
    """Area-weighted average of an element quantity onto the nodes (smooths
    across material interfaces - use element values for interface work)."""
    mesh = sol.mesh
    _, _, area, _ = element_geometry(mesh)
    num = np.zeros(mesh.n_nodes, dtype=values.dtype)
    den = np.zeros(mesh.n_nodes)
    for k in range(3):
        np.add.at(num, mesh.tris[:, k], values * area)
        np.add.at(den, mesh.tris[:, k], area)
    return num / den


# ------------------------------------------------------------------ integrals
def energy_integral(sol: Solution) -> complex:
    """Integral of kappa*|E|^2 dV over the model (per metre depth if planar)."""
    e = element_field(sol)
    e2 = (np.abs(e) ** 2).sum(axis=1)
    vol = volume_weight(sol.problem, sol.mesh)
    total = (sol.kappa_region[sol.mesh.region] * e2 * vol).sum()
    return complex(total) if np.iscomplexobj(total) else float(total)


def capacitance(sol: Solution, voltage: float) -> float:
    """Capacitance from stored energy, C = 2W/V^2 (mode 'es')."""
    if sol.mode != "es":
        raise ValueError("capacitance is defined for mode 'es'")
    return energy_integral(sol) / voltage ** 2


def admittance(sol: Solution, voltage: complex) -> complex:
    """Y = (1/V^2) * integral(kappa |E|^2 dV).  Mode 'dc': G; 'ac': G + jwC."""
    if sol.mode == "es":
        raise ValueError("use capacitance() for mode 'es'")
    return energy_integral(sol) / abs(voltage) ** 2


# ------------------------------------------------------------------- sampling
class _Locator:
    def __init__(self, sol: Solution):
        m = sol.mesh
        self.sol = sol
        self.tri = Triangulation(m.nodes[:, 0], m.nodes[:, 1], m.tris)
        self.finder = self.tri.get_trifinder()
        self.cent = m.nodes[m.tris].mean(axis=1)
        self.E = element_field(sol)

    def _bary(self, t, x, y):
        p = self.sol.mesh.nodes[self.sol.mesh.tris[t]]
        T = np.array([[p[0, 0] - p[2, 0], p[1, 0] - p[2, 0]],
                      [p[0, 1] - p[2, 1], p[1, 1] - p[2, 1]]])
        l = np.linalg.solve(T, [x - p[2, 0], y - p[2, 1]])
        return np.array([l[0], l[1], 1.0 - l.sum()])

    def find(self, x, y, region):
        """Triangle at (x, y). With `region`, the triangle of that region
        nearest to the point (so a point on an interface can be evaluated from
        a chosen side)."""
        mesh = self.sol.mesh
        t = int(self.finder(x, y))
        if region is None:
            return t
        if t >= 0 and mesh.region[t] == region:
            return t
        cand = np.flatnonzero(mesh.region == region)
        if len(cand) == 0:
            return -1
        d = np.hypot(self.cent[cand, 0] - x, self.cent[cand, 1] - y)
        near = cand[np.argsort(d)[:12]]
        best, best_score = -1, np.inf
        for k in near:
            lam = self._bary(k, x, y)
            score = max(0.0, -lam.min())
            if score < best_score:
                best, best_score = int(k), score
        return best


def path_points(poly: Polygon, n: int = 200, unit: float = 1.0):
    """Evenly spaced points along a polyline/arc chain (geometry units).
    Returns (xy (n,2), arclength (n,) in metres, tangent (n,2))."""
    pts, ds = [], []
    for s in poly.segments():
        if s.bulge == 0.0:
            L = np.hypot(s.p1[0] - s.p0[0], s.p1[1] - s.p0[1])
            m = max(2, int(np.ceil(n * 1.0)))
            t = np.linspace(0, 1, m)
            seg = np.stack([s.p0[0] + t * (s.p1[0] - s.p0[0]), s.p0[1] + t * (s.p1[1] - s.p0[1])], axis=1)
        else:
            (cx, cy), R, a0, sw = s.arc()
            L = R * sw
            m = max(2, int(np.ceil(n * 1.0)))
            ang = a0 + np.linspace(0, sw, m)
            seg = np.stack([cx + R * np.cos(ang), cy + R * np.sin(ang)], axis=1)
            # arc is stored CCW from its start; reverse for clockwise segments
            if s.bulge < 0:
                seg = seg[::-1]
        pts.append(seg)
        ds.append(L)
    # resample the chain to n points by arclength
    allp = [pts[0]]
    for seg in pts[1:]:
        allp.append(seg[1:])
    P = np.vstack(allp)
    d = np.r_[0.0, np.cumsum(np.hypot(*np.diff(P, axis=0).T))]
    sgrid = np.linspace(0, d[-1], n)
    x = np.interp(sgrid, d, P[:, 0])
    y = np.interp(sgrid, d, P[:, 1])
    xy = np.stack([x, y], axis=1)
    tang = np.gradient(xy, axis=0)
    tang /= np.maximum(np.linalg.norm(tang, axis=1, keepdims=True), 1e-30)
    return xy, sgrid * unit, tang


def sample(sol: Solution, xy: np.ndarray, region: int | None = None):
    """Evaluate phi and E at points `xy` (geometry units). Returns dict with
    'phi' (K,), 'E' (K,2); points outside the mesh give NaN."""
    loc = _Locator(sol)
    u = sol.problem.unit
    mesh = sol.mesh
    phi = np.full(len(xy), np.nan, dtype=sol.phi.dtype)
    E = np.full((len(xy), 2), np.nan, dtype=loc.E.dtype)
    for i, (x, y) in enumerate(xy):
        xm, ym = x * u, y * u
        t = loc.find(xm, ym, region)
        if t < 0:
            continue
        lam = loc._bary(t, xm, ym)
        phi[i] = (lam * sol.phi[mesh.tris[t]]).sum()
        E[i] = loc.E[t]
    return {"phi": phi, "E": E}


def along_path(sol: Solution, poly: Polygon, n: int = 200, region: int | None = None):
    """Potential, tangential and normal field along a path - e.g. a creepage
    path along an insulator surface. Pass `region` (index into
    Problem.regions) to evaluate from that side of an interface; the tangential
    field is continuous across it, the normal field is not."""
    xy, s, tang = path_points(poly, n, unit=sol.problem.unit)
    r = sample(sol, xy, region)
    E = r["E"]
    nrm = np.stack([-tang[:, 1], tang[:, 0]], axis=1)
    return {
        "s": s, "xy": xy, "phi": r["phi"],
        "Et": (E * tang).sum(axis=1), "En": (E * nrm).sum(axis=1),
        "Emag": np.sqrt((np.abs(E) ** 2).sum(axis=1)),
    }


# ------------------------------------------------------------------- withstand
def allowable(curve, x: float = 0.0) -> float:
    """Allowable field (V/m) from a constant or a user curve [(x, E), ...]
    (x = e.g. gap length in mm or stress time; linear interpolation, flat
    outside the range)."""
    if curve is None:
        return float("nan")
    if np.isscalar(curve):
        return float(curve)
    xs, es = zip(*sorted(curve))
    return float(np.interp(x, xs, es))


def utilization(sol: Solution, x: float = 0.0):
    """Per region: peak |E| (V/m), allowable E, utilization = Emax/Eallow.
    For AC the phasors are peak values; applying a Vrms*sqrt(2) voltage is the
    user's responsibility. Edge/corner singularities inflate Emax - look at the
    field away from sharp electrode corners, or round them."""
    em = element_magnitude(sol)
    out = []
    for i, reg in enumerate(sol.problem.regions):
        sel = sol.mesh.region == i
        if not sel.any():
            continue
        emax = float(em[sel].max())
        ea = allowable(reg.material.e_allow, x)
        out.append({"region": i, "name": reg.name or reg.material.name,
                    "Emax": emax, "Eallow": ea,
                    "utilization": emax / ea if ea == ea and ea > 0 else float("nan")})
    return out
