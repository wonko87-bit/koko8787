"""GUI-independent document model: shapes, assignments, materials, settings,
undo/redo, project files and the bridge to the solver."""
from __future__ import annotations

import copy
import json
import math
from dataclasses import dataclass, field, asdict

from ..dxf_io import import_dxf
from ..fem import solve
from ..geometry import Polygon, rect
from ..materials import Material
from ..mesh import build_mesh
from ..problem import Conductor, DirichletLine, Problem, Region

KINDS = ("ignore", "region", "conductor", "line")
KIND_LABEL = {"ignore": "무시", "region": "영역(재료)", "conductor": "도체(전극)", "line": "전위선"}
MODES = ("es", "dc", "ac")
MODE_LABEL = {"es": "정전계 (ε)", "dc": "DC 전도 (σ)", "ac": "AC 복소 (σ+jωε)"}
FILE_VERSION = 1


class DocumentError(Exception):
    """A problem with the model that the user can fix (shown in a message box)."""


@dataclass
class DocShape:
    id: int
    layer: str
    polygon: Polygon
    closed: bool
    source: str = "user"
    kind: str = "ignore"
    material: str = ""
    voltage: float = 0.0       # volts (rms if Settings.rms_input and mode "ac")
    name: str = ""

    @property
    def area(self) -> float:
        return abs(self.polygon.signed_area()) if self.closed else 0.0

    def label(self) -> str:
        return self.name or f"도형{self.id}"


@dataclass
class Settings:
    mode: str = "es"
    freq: float = 50.0
    symmetry: str = "axisym"
    mesh_size: float = 2.0        # mm
    edge_size: float = 0.0        # mm, 0 = off
    edge_transition: float = 0.0  # mm, 0 = 4 x mesh size
    rms_input: bool = True        # AC voltages are rms (solver uses peak = sqrt2 x)
    curve_x: float = 0.0          # x value for allowable-field curves (e.g. gap in mm)


DEFAULT_MATERIALS = {
    # example values only - check against your own data before relying on them
    "air": Material("air", 1.0, 0.0, 3.0e6),
    "oil": Material("oil", 2.2, 1e-12, None),
    "epoxy": Material("epoxy", 4.0, 1e-15, None),
    "pressboard": Material("pressboard", 4.0, 1e-13, None),
}


@dataclass
class Result:
    problem: Problem
    mesh: object
    sol: object
    region_shape_ids: list[int]
    warnings: list[str]
    settings: Settings
    v_scale: float = 1.0
    signature: tuple = ()


def _mat_to_json(m: Material):
    return {"eps_r": m.eps_r, "sigma": m.sigma,
            "e_allow": [list(p) for p in m.e_allow] if isinstance(m.e_allow, (list, tuple)) else m.e_allow}


def _mat_from_json(name, d):
    ea = d.get("e_allow")
    if isinstance(ea, list):
        ea = [tuple(p) for p in ea]
    return Material(name, float(d["eps_r"]), float(d["sigma"]), ea)


class Document:
    def __init__(self):
        self.shapes: list[DocShape] = []
        self.materials: dict[str, Material] = dict(DEFAULT_MATERIALS)
        self.settings = Settings()
        self.path: str | None = None
        self.dirty = False
        self.listeners: list = []
        self._undo: list = []
        self._redo: list = []
        self._next_id = 0

    # ------------------------------------------------------------ plumbing
    def _state(self):
        return copy.deepcopy((self.shapes, self.materials, self.settings, self._next_id))

    def _restore(self, st):
        self.shapes, self.materials, self.settings, self._next_id = copy.deepcopy(st)

    def checkpoint(self):
        self._undo.append(self._state())
        del self._undo[:-100]
        self._redo.clear()

    def changed(self, structural: bool = True):
        self.dirty = True
        for fn in list(self.listeners):
            fn(structural)

    def can_undo(self):
        return bool(self._undo)

    def can_redo(self):
        return bool(self._redo)

    def undo(self):
        if self._undo:
            self._redo.append(self._state())
            self._restore(self._undo.pop())
            self.changed()

    def redo(self):
        if self._redo:
            self._undo.append(self._state())
            self._restore(self._redo.pop())
            self.changed()

    def shape(self, sid: int) -> DocShape:
        for s in self.shapes:
            if s.id == sid:
                return s
        raise KeyError(sid)

    def _new_id(self) -> int:
        self._next_id += 1
        return self._next_id - 1

    def tangent_run(self, sid: int, idx: int, tol_deg: float = 15.0) -> list[int]:
        """Indices of the segments of shape `sid` that continue `idx` smoothly (tangent
        direction changes by less than `tol_deg` at every joint) - like a CAD
        'select tangent chain': a line + fillet arc + line, or a flattened spline."""
        s = self.shape(sid)
        segs = s.polygon.segments()
        n = len(segs)
        cos_tol = math.cos(math.radians(tol_deg))

        def joined(i, j):
            a, b = segs[i].tangents()[1], segs[j].tangents()[0]
            return a[0] * b[0] + a[1] * b[1] >= cos_tol

        run = [idx]
        i = idx
        while True:
            j = (i + 1) % n if s.closed else i + 1
            if j >= n or j in run or not joined(i, j):
                break
            run.append(j)
            i = j
        i = idx
        while True:
            j = (i - 1) % n if s.closed else i - 1
            if j < 0 or j in run or not joined(j, i):
                break
            run.insert(0, j)
            i = j
        return run

    # ------------------------------------------------------------- shapes
    def load_dxf(self, path: str, *, replace: bool = True, **kw):
        model = import_dxf(path, **kw)
        self.checkpoint()
        if replace:
            self.shapes = []
        for sh in model.shapes:
            self.shapes.append(DocShape(self._new_id(), sh.layer, sh.polygon, sh.closed, sh.source))
        self.changed()
        return model

    def add_polygon(self, poly: Polygon, layer: str = "user", source: str = "user") -> int:
        self.checkpoint()
        s = DocShape(self._new_id(), layer, poly, poly.closed, source)
        self.shapes.append(s)
        self.changed()
        return s.id

    def add_rect(self, x0, y0, x1, y1) -> int:
        if x0 == x1 or y0 == y1:
            raise DocumentError("사각형의 폭과 높이는 0이 될 수 없습니다")
        return self.add_polygon(rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)), source="rect")

    def add_circle(self, cx, cy, r) -> int:
        if r <= 0:
            raise DocumentError("반지름은 0보다 커야 합니다")
        return self.add_polygon(Polygon([(cx + r, cy, 1.0), (cx - r, cy, 1.0)], True), source="circle")

    def add_polyline(self, pts, closed: bool) -> int:
        if len(pts) < (3 if closed else 2):
            raise DocumentError("점이 부족합니다 (닫힌 도형 3개, 열린 선 2개 이상)")
        return self.add_polygon(Polygon.from_points(pts, closed), source="polyline")

    def delete(self, ids):
        ids = set(ids)
        if not ids:
            return
        self.checkpoint()
        self.shapes = [s for s in self.shapes if s.id not in ids]
        self.changed()

    def set_vertices(self, sid: int, verts, *, checkpoint: bool = True):
        s = self.shape(sid)
        n_min = 3 if s.closed else 2
        if len(verts) < n_min:
            raise DocumentError(f"꼭짓점이 최소 {n_min}개 필요합니다")
        if checkpoint:
            self.checkpoint()
        s.polygon = Polygon([(float(x), float(y), float(b)) for x, y, b in verts], s.closed)
        self.changed(structural=False)

    def translate(self, ids, dx, dy):
        self.checkpoint()
        for sid in ids:
            s = self.shape(sid)
            s.polygon = Polygon([(x + dx, y + dy, b) for x, y, b in s.polygon.vertices], s.closed)
        self.changed(structural=False)

    def assign(self, ids, kind, *, material="", voltage=0.0, name=""):
        if kind not in KINDS:
            raise DocumentError(f"알 수 없는 종류: {kind}")
        self.checkpoint()
        for sid in ids:
            s = self.shape(sid)
            if kind in ("region", "conductor") and not s.closed:
                raise DocumentError(f"도형 [{sid}]은 열린 선이라 {KIND_LABEL[kind]}로 지정할 수 없습니다 (전위선만 가능)")
            s.kind = kind
            s.material = material if kind == "region" else ""
            s.voltage = float(voltage) if kind in ("conductor", "line") else 0.0
            if name:
                s.name = name if len(ids) == 1 else f"{name}{sid}"
        self.changed(structural=False)

    def set_materials(self, mats: dict[str, Material]):
        self.checkpoint()
        self.materials = dict(mats)
        # a shape that referenced a removed material loses its assignment
        for s in self.shapes:
            if s.kind == "region" and s.material not in self.materials:
                s.kind, s.material = "ignore", ""
        self.changed(structural=False)

    def set_settings(self, **kw):
        self.checkpoint()
        for k, v in kw.items():
            if not hasattr(self.settings, k):
                raise KeyError(k)
            setattr(self.settings, k, v)
        self.changed(structural=False)

    # ----------------------------------------------------------- project IO
    def save(self, path: str):
        data = {
            "version": FILE_VERSION,
            "settings": asdict(self.settings),
            "materials": {k: _mat_to_json(m) for k, m in self.materials.items()},
            "shapes": [{"id": s.id, "layer": s.layer, "closed": s.closed, "source": s.source,
                        "vertices": [list(v) for v in s.polygon.vertices],
                        "kind": s.kind, "material": s.material, "voltage": s.voltage, "name": s.name}
                       for s in self.shapes],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
        self.path, self.dirty = path, False

    @classmethod
    def load(cls, path: str) -> "Document":
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        if data.get("version") != FILE_VERSION:
            raise DocumentError(f"지원하지 않는 프로젝트 파일 버전: {data.get('version')}")
        d = cls()
        d.settings = Settings(**data["settings"])
        d.materials = {k: _mat_from_json(k, v) for k, v in data["materials"].items()}
        for sd in data["shapes"]:
            poly = Polygon([tuple(v) for v in sd["vertices"]], sd["closed"])
            d.shapes.append(DocShape(sd["id"], sd["layer"], poly, sd["closed"], sd["source"],
                                     sd["kind"], sd["material"], sd["voltage"], sd["name"]))
        d._next_id = 1 + max([s.id for s in d.shapes], default=-1)
        d.path, d.dirty = path, False
        return d

    # --------------------------------------------------------------- solver
    def voltage_scale(self) -> float:
        st = self.settings
        return math.sqrt(2.0) if st.mode == "ac" and st.rms_input else 1.0

    def to_problem(self) -> tuple[Problem, list[int]]:
        st = self.settings
        regions, conductors, lines = [], [], []
        k = self.voltage_scale()
        for s in self.shapes:
            if s.kind == "region":
                if s.material not in self.materials:
                    raise DocumentError(f"도형 [{s.id}]의 재료 {s.material!r}가 재료 목록에 없습니다")
                regions.append((s.area, s.id, Region(s.polygon, self.materials[s.material], s.label())))
            elif s.kind == "conductor":
                conductors.append(Conductor(s.polygon, s.voltage * k, s.label()))
            elif s.kind == "line":
                lines.append(DirichletLine(s.polygon, s.voltage * k, s.label()))
        if not regions:
            raise DocumentError("영역(재료)이 지정된 도형이 없습니다")
        if not conductors and not lines:
            raise DocumentError("전극이 없습니다 — 도체 또는 전위선을 하나 이상 지정하세요")
        if st.mode == "ac" and st.freq <= 0:
            raise DocumentError("AC 해석에는 주파수가 필요합니다")
        if st.mode == "dc" and not any(r.material.sigma > 0 for _, _, r in regions):
            raise DocumentError("DC 해석에는 도전율(σ)이 0보다 큰 재료가 필요합니다")
        if st.symmetry == "axisym":
            for s in self.shapes:
                if s.kind != "ignore" and s.polygon.bbox()[0] < -1e-6:
                    raise DocumentError(f"도형 [{s.id}]이 x<0 영역에 있습니다 — 축대칭에서 x는 반지름 r입니다")
        regions.sort(key=lambda t: -t[0])
        prob = Problem(
            regions=[r for _, _, r in regions], conductors=conductors, lines=lines,
            symmetry=st.symmetry, mesh_size=st.mesh_size,
            edge_size=st.edge_size if st.edge_size > 0 else None,
            edge_transition=st.edge_transition if st.edge_transition > 0 else None)
        return prob, [sid for _, sid, _ in regions]

    def signature(self) -> tuple:
        """Everything that changes the solution (not the allowable-field display)."""
        st = self.settings
        shapes = tuple((s.id, s.kind, s.material, s.voltage, tuple(s.polygon.vertices))
                       for s in self.shapes if s.kind != "ignore")
        mats = tuple(sorted((m.name, m.eps_r, m.sigma, repr(m.e_allow)) for m in self.materials.values()))
        return (st.mode, st.freq, st.symmetry, st.mesh_size, st.edge_size, st.edge_transition,
                st.rms_input, shapes, mats)

    def run(self) -> Result:
        prob, ids = self.to_problem()
        try:
            mesh = build_mesh(prob)
        except ValueError as e:
            raise DocumentError(str(e)) from e
        except Exception as e:  # gmsh failures (self-intersection, ...)
            raise DocumentError(f"메시 생성 실패: {e}\n도형이 서로 교차하거나 윤곽이 닫혀 있지 않은지 확인하세요") from e
        sol = solve(prob, mesh, self.settings.mode, self.settings.freq)
        return Result(prob, mesh, sol, ids, list(mesh.warnings), copy.deepcopy(self.settings),
                      self.voltage_scale(), self.signature())
