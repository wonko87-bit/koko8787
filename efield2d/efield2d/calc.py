"""Field calculator: evaluate user formulas on a solution (like a FEM tool's field calculator).

Everything is SI (m, V, V/m, ...). Coordinates given in *geometry units* (mm) where noted.

Variables available in a formula
  x, y          position [m]  (aliases r = x, z = y)
  V             potential            (complex for AC; peak phasor)
  Ex, Ey        field E = -grad V    (aliases Er, Ez)
  E             |E| = sqrt(|Ex|^2 + |Ey|^2)
  eps, epsr     absolute / relative permittivity of the region
  sigma         conductivity [S/m]
  kappa         the coefficient used by the analysis (eps | sigma | sigma + j w eps)
  w             angular frequency (2 pi f for AC, else 0)
  Dx, Dy, D     eps * E
  Jcx, Jcy, Jc  sigma * E                     (conduction current density)
  Jx, Jy, J     (sigma + j w eps) * E          (total current density)
  eps0, pi, mm  constants  (mm = 1e-3)
 Along a path (profile / flux) also
  tx, ty        unit tangent;  nx, ny  unit normal on the RIGHT of the travel direction
                (for a counter-clockwise outline this is the outward normal)
  Et, En, Dt, Dn, Jt, Jn, Jct, Jcn   tangential / normal components

Functions: abs re im conj angle sqrt exp log ln log10 sin cos tan atan2 sinh cosh tanh
           mag(a, b, ...) minimum maximum where(cond, a, b) floor ceil sign
Operators: + - * / ** ^ (power), comparisons, unary minus.  Anything else is rejected.
You can add your own named constants ({"a": 5e-3}).

Fields: a P1 element has a constant E.  For points, profiles and fluxes the calculator
uses the recovered (quadratic patch fit of the potential, per region, so material jumps are kept)
and linearly interpolated field - much closer to the true value than the element constant
(`smooth=False` gives the raw element value).  Volume integrals always use the raw element
values, which makes energy / loss identities exact.  V is interpolated at a point.
(A line integral of the recovered E equals the potential drop to ~0.3 %, not exactly; with
`smooth=False` it is the potential drop of the P1 solution.)
"""
from __future__ import annotations

import ast
import math
import operator
from dataclasses import dataclass

import numpy as np

from . import post
from .fem import assemble, element_geometry, volume_weight
from .geometry import Polygon
from .materials import EPS0
from .post import _Locator, path_points


class FormulaError(ValueError):
    """The formula is not valid or uses something that is not available here."""


# ------------------------------------------------------------------ evaluator
_BINOPS = {ast.Add: operator.add, ast.Sub: operator.sub, ast.Mult: operator.mul,
           ast.Div: operator.truediv, ast.Pow: operator.pow, ast.Mod: operator.mod}
_CMPOPS = {ast.Lt: operator.lt, ast.LtE: operator.le, ast.Gt: operator.gt, ast.GtE: operator.ge,
           ast.Eq: operator.eq, ast.NotEq: operator.ne}


def _mag(*a):
    return np.sqrt(sum(np.abs(v) ** 2 for v in a))


FUNCTIONS = {
    "abs": np.abs, "re": np.real, "im": np.imag, "conj": np.conj, "angle": np.angle,
    "sqrt": np.sqrt, "exp": np.exp, "log": np.log, "ln": np.log, "log10": np.log10,
    "sin": np.sin, "cos": np.cos, "tan": np.tan, "atan2": np.arctan2,
    "sinh": np.sinh, "cosh": np.cosh, "tanh": np.tanh, "mag": _mag,
    "minimum": np.minimum, "maximum": np.maximum, "where": np.where,
    "floor": np.floor, "ceil": np.ceil, "sign": np.sign,
}
CONSTANTS = {"pi": math.pi, "eps0": EPS0, "mm": 1e-3, "e": math.e}


def safe_eval(expr: str, names: dict):
    """Evaluate `expr` with only arithmetic, comparisons, whitelisted functions and the
    given names. No attribute access, indexing, lambdas, comprehensions, imports."""
    src = expr.strip().replace("^", "**")
    if not src:
        raise FormulaError("수식이 비어 있습니다")
    try:
        tree = ast.parse(src, mode="eval")
    except SyntaxError as e:
        raise FormulaError(f"수식 문법 오류: {e.msg} (위치 {e.offset})") from None

    def ev(node):
        if isinstance(node, ast.Expression):
            return ev(node.body)
        if isinstance(node, ast.Constant):
            if isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
                return float(node.value)          # never Python big ints (9**9**9 would hang)
            if isinstance(node.value, complex):
                return node.value
            raise FormulaError(f"사용할 수 없는 값: {node.value!r}")
        if isinstance(node, ast.Name):
            if node.id in names:
                return names[node.id]
            if node.id in FUNCTIONS:
                raise FormulaError(f"'{node.id}'는 함수입니다 — {node.id}(...) 형태로 쓰세요")
            raise FormulaError(f"알 수 없는 이름: {node.id!r}")
        if isinstance(node, ast.BinOp) and type(node.op) in _BINOPS:
            return _BINOPS[type(node.op)](ev(node.left), ev(node.right))
        if isinstance(node, ast.UnaryOp):
            if isinstance(node.op, ast.USub):
                return -ev(node.operand)
            if isinstance(node.op, ast.UAdd):
                return ev(node.operand)
        if isinstance(node, ast.Compare) and len(node.ops) == 1 and type(node.ops[0]) in _CMPOPS:
            return _CMPOPS[type(node.ops[0])](ev(node.left), ev(node.comparators[0]))
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and not node.keywords:
            fn = FUNCTIONS.get(node.func.id)
            if fn is None:
                raise FormulaError(f"알 수 없는 함수: {node.func.id!r}")
            return fn(*[ev(a) for a in node.args])
        raise FormulaError(f"지원하지 않는 수식 요소: {type(node).__name__}")

    with np.errstate(all="ignore"):
        try:
            return ev(tree)
        except FormulaError:
            raise
        except (TypeError, ValueError, ZeroDivisionError, OverflowError) as e:
            raise FormulaError(f"계산할 수 없습니다: {e}") from None


def parse_constants(text: str) -> dict:
    """'a=5e-3; b=20*mm' -> {"a": 0.005, "b": 0.02}; later ones may use earlier ones."""
    out: dict = {}
    for part in text.replace("\n", ";").split(";"):
        part = part.strip()
        if not part:
            continue
        if "=" not in part:
            raise FormulaError(f"상수는 이름=값 형태로 쓰세요: {part!r}")
        name, val = (t.strip() for t in part.split("=", 1))
        if not name.isidentifier() or name in FUNCTIONS:
            raise FormulaError(f"상수 이름이 올바르지 않습니다: {name!r}")
        v = safe_eval(val, {**CONSTANTS, **out})
        if np.ndim(v) != 0:
            raise FormulaError(f"상수 {name!r}의 값이 하나의 수가 아닙니다")
        out[name] = v.item() if hasattr(v, "item") else v
    return out


PATH_ONLY = {"tx", "ty", "nx", "ny", "Et", "En", "Dt", "Dn", "Jt", "Jn", "Jct", "Jcn"}


def safe_name(text: str) -> str:
    """A region / electrode name made usable as part of an identifier."""
    out = "".join(ch if (ch.isalnum() or ch == "_") and ord(ch) < 128 else "_" for ch in text.strip())
    return out or "x"


# ----------------------------------------------------------------- calculator
@dataclass
class Result:
    value: object                 # scalar, or array for profiles
    unit: str = ""
    note: str = ""
    n_nan: int = 0


class FieldCalculator:
    def __init__(self, sol, smooth: bool = True):
        self.sol = sol
        self.smooth = smooth
        self._nodal: dict[int, np.ndarray] = {}
        m, prob = sol.mesh, sol.problem
        self.mesh, self.problem = m, prob
        self.unit = prob.unit
        self.axisym = prob.symmetry == "axisym"
        self.w = 2 * math.pi * sol.freq if sol.mode == "ac" else 0.0
        self.E = post.element_field(sol)                                    # (M, 2)
        eps_r = np.array([r.material.eps_r for r in prob.regions])
        sig = np.array([r.material.sigma for r in prob.regions])
        self.eps_el = EPS0 * eps_r[m.region]
        self.sig_el = sig[m.region]
        self.kappa_el = sol.kappa_region[m.region]
        self.vol = volume_weight(prob, m)
        self.cent = m.nodes[m.tris].mean(axis=1)                           # (M, 2) metres
        self.V_cent = sol.phi[m.tris].mean(axis=1)
        self._loc = None

    def auto_constants(self) -> dict:
        """Constants derived from the model, usable in formulas:
        kappa_<region>, eps_<region>, sigma_<region>  (absolute eps, sigma; kappa = analysis coefficient)
        V_<electrode>                                  (applied voltage, peak for AC)
        f                                               (frequency, AC)"""
        out: dict = {}
        for i, r in enumerate(self.problem.regions):
            n = safe_name(r.name or r.material.name)
            out[f"kappa_{n}"] = self.sol.kappa_region[i].item()
            out[f"eps_{n}"] = EPS0 * r.material.eps_r
            out[f"sigma_{n}"] = r.material.sigma
        for e in list(self.problem.conductors) + list(self.problem.lines):
            out[f"V_{safe_name(e.name)}"] = complex(e.voltage) if np.iscomplexobj(self.sol.phi) else float(e.voltage.real)
        if self.sol.mode == "ac":
            out["f"] = self.sol.freq
        return out

    # ---- variable sets -------------------------------------------------
    def _derive(self, base: dict) -> dict:
        v = dict(CONSTANTS)
        v.update(base)
        v["r"], v["z"] = v["x"], v["y"]
        v["Er"], v["Ez"] = v["Ex"], v["Ey"]
        v["E"] = _mag(v["Ex"], v["Ey"])
        v["epsr"] = v["eps"] / EPS0
        v["w"] = self.w
        v["Dx"], v["Dy"] = v["eps"] * v["Ex"], v["eps"] * v["Ey"]
        v["D"] = _mag(v["Dx"], v["Dy"])
        v["Jcx"], v["Jcy"] = v["sigma"] * v["Ex"], v["sigma"] * v["Ey"]
        v["Jc"] = _mag(v["Jcx"], v["Jcy"])
        jt = v["sigma"] + 1j * self.w * v["eps"] if self.w else v["sigma"]
        v["Jx"], v["Jy"] = jt * v["Ex"], jt * v["Ey"]
        v["J"] = _mag(v["Jx"], v["Jy"])
        if "tx" in v:
            t = (v["tx"], v["ty"])
            n = (v["ty"], -v["tx"])
            v["nx"], v["ny"] = n
            for nm, (a, b) in {"E": (v["Ex"], v["Ey"]), "D": (v["Dx"], v["Dy"]),
                               "J": (v["Jx"], v["Jy"]), "Jc": (v["Jcx"], v["Jcy"])}.items():
                v[nm + "t"] = a * t[0] + b * t[1]
                v[nm + "n"] = a * n[0] + b * n[1]
        return v

    def _eval(self, expr: str, base: dict, consts: dict | None, path: bool):
        names = self._derive(base)
        auto = self.auto_constants()
        for k, v in auto.items():
            names.setdefault(k, v)
        if consts:
            clash = set(consts) & (set(names) - set(auto)) | (set(consts) & set(FUNCTIONS))
            if clash:
                raise FormulaError(f"상수 이름이 예약어와 겹칩니다: {', '.join(sorted(clash))}")
            names.update(consts)
        if not path:
            try:
                used = {n.id for n in ast.walk(ast.parse(expr.strip().replace("^", "**"), mode="eval"))
                        if isinstance(n, ast.Name)}
            except SyntaxError:
                used = set()                    # safe_eval below reports the syntax error properly
            bad = sorted(used & PATH_ONLY)
            if bad:
                raise FormulaError(f"{', '.join(bad)}는 경로 계산(프로파일/플럭스)에서만 사용할 수 있습니다")
        return safe_eval(expr, names)

    def _nodal_field(self, region: int) -> np.ndarray:
        """Recovered nodal E = -grad V for one region, (N, 2).

        Polynomial-preserving recovery: around every node, a quadratic is least-squares fitted
        to the *potential* at the nodes of the 2-ring neighbourhood (using only this region's
        elements, so the field jump at a material interface is kept) and its gradient at the
        node is taken. Nodal potentials are far more accurate than element gradients, which
        makes this second order accurate also at boundary / interface nodes and on the aligned
        layered meshes that appear next to electrodes. Nodes without enough neighbours fall
        back to the area-weighted average of the element fields; nodes outside the region
        are NaN."""
        if region in self._nodal:
            return self._nodal[region]
        import scipy.sparse as sp
        m = self.mesh
        n = m.n_nodes
        _, _, area, _ = element_geometry(m)
        sel = np.flatnonzero(m.region == region)
        tri, ar = m.tris[sel], area[sel]
        # average (fallback)
        num = np.zeros((n, 2), dtype=self.E.dtype)
        den = np.zeros(n)
        for k in range(3):
            np.add.at(num, tri[:, k], self.E[sel] * ar[:, None])
            np.add.at(den, tri[:, k], ar)
        out = np.full((n, 2), np.nan, dtype=self.E.dtype)
        has = den > 0
        out[has] = num[has] / den[:, None][has]

        # 2-ring node neighbourhoods within the region
        inc = sp.csr_matrix((np.ones(3 * len(sel)), (np.repeat(np.arange(len(sel)), 3), tri.ravel())),
                            shape=(len(sel), n))
        n1 = (inc.T @ inc).tocsr()
        n1.data[:] = 1.0
        n2 = (n1 @ n1).tocsr()
        n2.data[:] = 1.0
        coo = n2.tocoo()
        i, j = coo.row, coo.col
        scale = float(np.sqrt(np.median(ar))) or 1.0
        d = (m.nodes[j] - m.nodes[i]) / scale
        P = np.stack([np.ones(len(i)), d[:, 0], d[:, 1], d[:, 0] ** 2, d[:, 0] * d[:, 1], d[:, 1] ** 2], axis=1)
        A = np.zeros((n, 6, 6))
        for a_ in range(6):
            for b_ in range(a_, 6):
                A[:, a_, b_] = np.bincount(i, weights=P[:, a_] * P[:, b_], minlength=n)
                A[:, b_, a_] = A[:, a_, b_]
        phi = self.sol.phi
        cplx = np.iscomplexobj(phi)
        B = np.zeros((n, 6), dtype=complex if cplx else float)
        for a_ in range(6):
            B[:, a_] = np.bincount(i, weights=P[:, a_] * phi[j].real, minlength=n)
            if cplx:
                B[:, a_] = B[:, a_] + 1j * np.bincount(i, weights=P[:, a_] * phi[j].imag, minlength=n)
        cnt = np.bincount(i, minlength=n)
        ok = has & (cnt >= 8)                      # 6 unknowns: want some redundancy
        if ok.any():
            coef = np.einsum("nij,nj->ni", np.linalg.pinv(A[ok], rcond=1e-10), B[ok])
            out[ok] = -coef[:, 1:3] / scale
        self._nodal[region] = out
        return out

    def _locator(self):
        if self._loc is None:
            self._loc = _Locator(self.sol)
        return self._loc

    def _sample_base(self, xy_geom: np.ndarray, region: int | None):
        """Per-point base variables at points given in geometry units. Points that do not
        touch the (side) region come back NaN."""
        loc, m, u = self._locator(), self.mesh, self.unit
        K = len(xy_geom)
        cplx = np.iscomplexobj(self.sol.phi)
        dt = complex if cplx else float
        V = np.full(K, np.nan, dtype=dt)
        Ex = np.full(K, np.nan, dtype=dt)
        Ey = np.full(K, np.nan, dtype=dt)
        eps = np.full(K, np.nan)
        sig = np.full(K, np.nan)
        for i, (x, y) in enumerate(xy_geom):
            t = loc.find(x * u, y * u, region)
            if t < 0:
                continue
            lam = np.clip(loc._bary(t, x * u, y * u), 0.0, 1.0)
            lam /= lam.sum()
            V[i] = (lam * self.sol.phi[m.tris[t]]).sum()
            if self.smooth:
                e = (lam[:, None] * self._nodal_field(int(m.region[t]))[m.tris[t]]).sum(axis=0)
            else:
                e = self.E[t]
            Ex[i], Ey[i] = e
            eps[i], sig[i] = self.eps_el[t], self.sig_el[t]
        return {"x": xy_geom[:, 0] * u, "y": xy_geom[:, 1] * u, "V": V, "Ex": Ex, "Ey": Ey,
                "eps": eps, "sigma": sig}

    # ---- operations ----------------------------------------------------
    def at_point(self, expr: str, x_mm: float, y_mm: float, region: int | None = None,
                 consts: dict | None = None) -> Result:
        base = self._sample_base(np.array([[x_mm, y_mm]]), region)
        val = np.asarray(self._eval(expr, base, consts, path=False))
        if np.isnan(base["V"]).any():
            return Result(complex("nan"), n_nan=1,
                          note="이 점은 메시 밖이거나 선택한 영역에 접하지 않습니다")
        return Result(val.reshape(-1)[0] if val.size else val)

    def profile(self, expr: str, poly: Polygon, n: int = 400, region: int | None = None,
                consts: dict | None = None, ref_expr: str | None = None) -> dict:
        """Values along a path: s [m], xy [mm], values, and the line integral (trapezoid)."""
        xy, s, tang = path_points(poly, n, unit=self.unit)
        base = self._sample_base(xy, region)
        base["tx"], base["ty"] = tang[:, 0], tang[:, 1]
        vals = np.broadcast_to(np.asarray(self._eval(expr, base, consts, path=True)), (len(xy),)).copy()
        bad = np.isnan(base["V"])
        out = {"s": s, "xy": xy, "values": vals, "n_nan": int(bad.sum())}
        v0 = np.where(bad, 0.0, vals)
        out["integral"] = np.trapezoid(v0, s)                       # sum of f dl  [unit * m]
        out["mean"] = out["integral"] / (s[-1] - s[0])
        out["cumulative"] = np.concatenate([[0], np.cumsum(0.5 * (v0[1:] + v0[:-1]) * np.diff(s))])
        if ref_expr:
            ref = np.broadcast_to(np.asarray(self._eval(ref_expr, base, consts, path=True)), (len(xy),)).copy()
            out["ref"] = ref
            with np.errstate(all="ignore"):
                ok = ~bad
                den = np.maximum(np.abs(ref[ok]).max(), 1e-300)
                out["max_abs_diff"] = float(np.abs(vals[ok] - ref[ok]).max())
                out["max_rel_diff"] = out["max_abs_diff"] / den     # relative to the largest |ref|
        return out

    def volume_integral(self, expr: str, region: int | None = None, consts: dict | None = None) -> Result:
        """Integral of expr over the model (or one region): sum over elements of f * volume.
        Planar: per metre depth. Axisymmetric: full body (2 pi r)."""
        base = {"x": self.cent[:, 0], "y": self.cent[:, 1], "V": self.V_cent,
                "Ex": self.E[:, 0], "Ey": self.E[:, 1], "eps": self.eps_el, "sigma": self.sig_el}
        val = np.broadcast_to(np.asarray(self._eval(expr, base, consts, path=False)), (self.mesh.n_tris,))
        sel = np.ones(self.mesh.n_tris, bool) if region is None else (self.mesh.region == region)
        total = (val[sel] * self.vol[sel]).sum()
        return Result(total.item() if hasattr(total, "item") else total,
                      note="축대칭 전체 (2πr)" if self.axisym else "깊이 1 m당")

    def flux(self, field: str, poly: Polygon, n: int = 400, region: int | None = None,
             consts: dict | None = None) -> Result:
        """Surface integral of the normal component of D, J, Jc or E through the surface
        swept by the path (axisymmetric: rotated about the axis, weight 2 pi r; planar: 1 m
        deep). The normal is on the right of the travel direction."""
        comp = {"D": "Dn", "J": "Jn", "Jc": "Jcn", "E": "En"}.get(field)
        if comp is None:
            raise FormulaError("field는 D, J, Jc, E 중 하나여야 합니다")
        p = self.profile(comp, poly, n=n, region=region, consts=consts)
        xm = p["xy"][:, 0] * self.unit
        wgt = 2 * math.pi * xm if self.axisym else np.ones_like(xm)
        vals = np.where(np.isnan(p["values"]), 0.0, p["values"])
        total = np.trapezoid(vals * wgt, p["s"])
        unit = {"D": "C", "J": "A", "Jc": "A", "E": "V·m"}[field]
        return Result(total.item() if hasattr(total, "item") else total, unit,
                      note=("축대칭 전체 (2πr)" if self.axisym else "깊이 1 m당") + ", 법선 = 진행 방향의 오른쪽",
                      n_nan=p["n_nan"])

    def electrode_charge(self, name: str) -> Result:
        """Charge [C] (es) or current [A] (dc, ac: complex) flowing out of the electrode into
        the model, from the nodal residual sum(K phi) over its nodes. This is the accurate way
        (it equals 2W/V exactly); a Gauss surface integral of the P1 field is only approximate."""
        if name not in self.mesh.electrode_nodes:
            raise FormulaError(f"전극 {name!r}이 없습니다 (있는 전극: {', '.join(self.mesh.electrode_nodes)})")
        K = assemble(self.problem, self.mesh, self.sol.kappa_region.astype(self.sol.phi.dtype))
        r = K @ self.sol.phi
        val = r[self.mesh.electrode_nodes[name]].sum()
        unit = "C" if self.sol.mode == "es" else "A"
        return Result(val.item() if hasattr(val, "item") else val, unit,
                      note="노드 잔차합 ΣKφ (모델에 포함된 부분만)")

    # ---- review --------------------------------------------------------
    def review(self) -> list[dict]:
        """Self-consistency checks between independent ways of computing the same quantity."""
        sol, m = self.sol, self.mesh
        rows: list[dict] = []

        def add(label, a, b, unit="", note="", scale=None):
            a, b = complex(a), complex(b)
            den = scale if scale else max(abs(a), abs(b), 1e-300)
            rows.append({"label": label, "a": a, "b": b, "unit": unit, "note": note,
                         "rel": abs(a - b) / den})

        names = list(m.electrode_nodes)
        q = {n: self.electrode_charge(n).value for n in names}
        volts = {}
        for n in names:
            volts[n] = complex(np.mean(sol.phi[m.electrode_nodes[n]]))
        qscale = max(max(abs(v) for v in q.values()), 1e-300)
        if sol.mode == "es":
            w_int = 0.5 * self.volume_integral("eps*E**2").value
            # energy = 1/2 sum_i Q_i V_i (all electrodes)
            w_q = 0.5 * sum(q[n] * volts[n] for n in names)
            add("저장 에너지 W: ½∫ε|E|²dV  vs  ½ΣQ·V", w_int, w_q, "J", "전극 전하는 ΣKφ")
            add("전극 전하 합계 ΣQ (0이어야 함)", sum(q.values()), 0.0, "C",
                "모델 경계가 자연경계(Neumann)일 때; 오차는 가장 큰 전극 전하 대비", scale=qscale)
        else:
            p_int = self.volume_integral("sigma*E**2").value
            fac = 0.5 if sol.mode == "ac" else 1.0
            s_q = sum(q[n] * np.conj(volts[n]) for n in names)
            p_q = (0.5 if sol.mode == "ac" else 1.0) * s_q.real
            add("손실 P: " + ("½" if sol.mode == "ac" else "") + "∫σ|E|²dV  vs  "
                + ("½" if sol.mode == "ac" else "") + "Re(ΣI·V*)", fac * p_int, p_q, "W")
            add("전극 전류 합계 ΣI (0이어야 함)", sum(q.values()), 0.0, "A",
                "오차는 가장 큰 전극 전류 대비", scale=qscale)
            if sol.mode == "ac" and self.w > 0:
                we_int = 0.25 * self.volume_integral("eps*E**2").value
                we_q = 0.25 * (sum(q[n] * np.conj(volts[n]) for n in names)).imag / self.w
                add("평균 저장 에너지: ¼∫ε|E|²dV  vs  ¼Im(ΣI·V*)/ω", we_int, abs(we_q), "J",
                    "크기 비교 (부호는 규약에 따름)")
        return rows
