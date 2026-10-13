"""matplotlib canvas: geometry view with picking / vertex dragging, and the
field result overlay."""
from __future__ import annotations

import math

import numpy as np
import matplotlib
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg
from matplotlib.figure import Figure
from matplotlib.path import Path as MPath
from matplotlib.tri import Triangulation
from PySide6.QtCore import Signal

from .. import post
from ..geometry import outline_xy
from .document import Document, DocShape

# Korean-capable fonts first (Windows, macOS, Linux), then the matplotlib default
matplotlib.rcParams["font.family"] = "sans-serif"
matplotlib.rcParams["font.sans-serif"] = ["Malgun Gothic", "맑은 고딕", "AppleGothic", "NanumGothic",
                                          "Noto Sans CJK KR", "Noto Sans KR", "DejaVu Sans"]
matplotlib.rcParams["axes.unicode_minus"] = False

_FILL = ["#9ecae1", "#fdd0a2", "#c7e9c0", "#dadaeb", "#fcbba1", "#d9d9d9", "#fee391", "#b2df8a"]
CONDUCTOR_FILL = "#e6550d"
LINE_COLOR = "#d6006e"
SELECT_COLOR = "#ffb000"


class Canvas(FigureCanvasQTAgg):
    selectionRequested = Signal(int, bool)   # shape id (-1 = empty space), additive?
    dragFinished = Signal(int)               # shape id whose vertex was dragged
    cursorInfo = Signal(str)

    def __init__(self, doc: Document):
        fig = Figure(figsize=(7, 6), layout="constrained")
        super().__init__(fig)
        self.doc = doc
        gs = fig.add_gridspec(1, 2, width_ratios=[1, 0.035])
        self.ax = fig.add_subplot(gs[0, 0])
        self.cax = fig.add_subplot(gs[0, 1])
        self.cax.set_visible(False)
        self.selected: set[int] = set()
        self.edit_mode = False
        self.snap = 0.0
        self.toolbar = None
        # result state
        self.result = None
        self.quantity = "E"          # "E" | "phi"
        self.vmax = 0.0              # kV/mm or kV, 0 = auto
        self.show_contours = True
        self.show_mesh = False
        self._tri = self._finder = self._elem_E = None
        self._drag = None
        self._have_view = False
        self.mpl_connect("button_press_event", self._on_press)
        self.mpl_connect("motion_notify_event", self._on_motion)
        self.mpl_connect("button_release_event", self._on_release)

    # ---------------------------------------------------------------- state
    def set_selection(self, ids):
        self.selected = set(ids)
        self.redraw()

    def set_result(self, result):
        self.result = result
        self._tri = self._finder = self._elem_E = None
        if result is not None:
            m, u = result.mesh, result.problem.unit
            self._tri = Triangulation(m.nodes[:, 0] / u, m.nodes[:, 1] / u, m.tris)
            self._finder = self._tri.get_trifinder()
            self._elem_E = post.element_magnitude(result.sol) / 1e6      # kV/mm
        self.redraw()

    def fit_view(self):
        shapes = [s for s in self.doc.shapes]
        if not shapes:
            return
        bb = np.array([s.polygon.bbox() for s in shapes])
        x0, y0, x1, y1 = bb[:, 0].min(), bb[:, 1].min(), bb[:, 2].max(), bb[:, 3].max()
        pad = 0.05 * max(x1 - x0, y1 - y0, 1e-6)
        self.ax.set_xlim(x0 - pad, x1 + pad)
        self.ax.set_ylim(y0 - pad, y1 + pad)
        self._have_view = True
        self.draw_idle()

    # --------------------------------------------------------------- drawing
    def _material_color(self, name: str):
        names = sorted(self.doc.materials)
        return _FILL[names.index(name) % len(_FILL)] if name in names else "#cccccc"

    def redraw(self):
        ax, cax = self.ax, self.cax
        lim = (ax.get_xlim(), ax.get_ylim()) if self._have_view else None
        ax.clear()
        cax.clear()
        cax.set_visible(False)
        ax.set_aspect("equal", adjustable="box")
        ax.grid(True, lw=0.3, alpha=0.4)
        ax.set_xlabel("r [mm]" if self.doc.settings.symmetry == "axisym" else "x [mm]")
        ax.set_ylabel("z [mm]" if self.doc.settings.symmetry == "axisym" else "y [mm]")

        res = self.result
        showing = res is not None
        if showing:
            self._draw_result(res)

        order = sorted(self.doc.shapes, key=lambda s: (s.kind != "region", -s.area))
        for s in order:
            self._draw_shape(s, thin=showing)
        if self.doc.settings.symmetry == "axisym" and self.doc.shapes:
            ax.axvline(0, color="gray", lw=0.8, ls="-.")

        if lim is not None:
            ax.set_xlim(*lim[0])
            ax.set_ylim(*lim[1])
        elif self.doc.shapes:
            self.fit_view()
        self.draw_idle()

    def _draw_shape(self, s: DocShape, thin: bool):
        ax = self.ax
        xy = np.array(outline_xy(s.polygon))
        sel = s.id in self.selected
        lw = 1.6 if sel else 0.9
        ec = SELECT_COLOR if sel else "k"
        if s.kind == "ignore":
            ax.plot(xy[:, 0], xy[:, 1], color=SELECT_COLOR if sel else "#888888", lw=lw,
                    ls="--" if not sel else "-", zorder=3)
        elif s.kind == "line":
            ax.plot(xy[:, 0], xy[:, 1], color=SELECT_COLOR if sel else LINE_COLOR,
                    lw=3.0 if sel else 2.2, zorder=5)
        else:
            if not thin:
                fc = self._material_color(s.material) if s.kind == "region" else CONDUCTOR_FILL
                ax.fill(xy[:, 0], xy[:, 1], fc=fc, ec="none", alpha=0.85 if s.kind == "conductor" else 0.7,
                        zorder=1 if s.kind == "region" else 4)
            elif s.kind == "conductor":
                ax.fill(xy[:, 0], xy[:, 1], fc=CONDUCTOR_FILL, ec="none", alpha=0.35, zorder=4)
            ax.plot(xy[:, 0], xy[:, 1], color=ec, lw=lw, zorder=6 if sel else 3)
        if sel:
            v = np.array([(x, y) for x, y, _ in s.polygon.vertices])
            ax.plot(v[:, 0], v[:, 1], "s", ms=5, mfc="white", mec=SELECT_COLOR, zorder=8)
        if s.kind != "ignore":
            c = xy.mean(axis=0)
            txt = s.label() if s.kind == "region" else f"{s.label()}\n{s.voltage / 1e3:g} kV"
            ax.annotate(txt, c, ha="center", va="center", fontsize=7, zorder=9,
                        bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none", alpha=0.6))

    def _draw_result(self, res):
        ax, cax = self.ax, self.cax
        tri = self._tri
        if self.quantity == "E":
            vals = self._elem_E
            vmax = self.vmax if self.vmax > 0 else float(vals.max())
            m = ax.tripcolor(tri, facecolors=vals, cmap="turbo", vmin=0.0, vmax=vmax, zorder=0)
            label = "|E| [kV/mm]" + (" (peak)" if res.settings.mode == "ac" else "")
        else:
            phi = np.abs(res.sol.phi) / 1e3
            vmax = self.vmax if self.vmax > 0 else float(phi.max())
            m = ax.tripcolor(tri, phi, cmap="viridis", vmin=0.0, vmax=vmax, shading="gouraud", zorder=0)
            label = "|V| [kV]" + (" (peak)" if res.settings.mode == "ac" else "")
        if self.show_contours:
            phi = np.abs(res.sol.phi) / 1e3
            if phi.max() > phi.min():
                ax.tricontour(tri, phi, levels=np.linspace(phi.min(), phi.max(), 21)[1:-1],
                              colors="k", linewidths=0.45, zorder=2)
        if self.show_mesh:
            ax.triplot(tri, lw=0.2, color="k", alpha=0.35, zorder=2)
        cax.set_visible(True)
        self.figure.colorbar(m, cax=cax, label=label)

    # ----------------------------------------------------------- hit testing
    def _px_to_data(self, px: float) -> float:
        inv = self.ax.transData.inverted()
        a = inv.transform((0, 0))
        b = inv.transform((px, 0))
        return abs(b[0] - a[0])

    def hit_test(self, x: float, y: float, tol_px: float = 6.0):
        tol = self._px_to_data(tol_px)
        best, best_d = None, tol
        for s in self.doc.shapes:
            d = s.polygon.distance(x, y)
            if d <= best_d:
                best, best_d = s.id, d
        if best is not None:
            return best
        inside = [s for s in self.doc.shapes
                  if s.closed and MPath(np.array(outline_xy(s.polygon))).contains_point((x, y))]
        return min(inside, key=lambda s: s.area).id if inside else None

    def _nearest_vertex(self, sid: int, x: float, y: float, tol_px: float = 9.0):
        s = self.doc.shape(sid)
        tol = self._px_to_data(tol_px)
        best, best_d = None, tol
        for i, (vx, vy, _) in enumerate(s.polygon.vertices):
            d = math.hypot(vx - x, vy - y)
            if d <= best_d:
                best, best_d = i, d
        return best

    def _snapped(self, x, y):
        if self.snap > 0:
            return round(x / self.snap) * self.snap, round(y / self.snap) * self.snap
        return x, y

    # ---------------------------------------------------------------- events
    def _tool_active(self) -> bool:
        return bool(self.toolbar is not None and getattr(self.toolbar, "mode", ""))

    def _on_press(self, ev):
        if ev.inaxes is not self.ax or ev.button != 1 or self._tool_active():
            return
        x, y = ev.xdata, ev.ydata
        if self.edit_mode and len(self.selected) == 1:
            sid = next(iter(self.selected))
            i = self._nearest_vertex(sid, x, y)
            if i is not None:
                self.doc.checkpoint()
                self._drag = (sid, i)
                return
        sid = self.hit_test(x, y)
        additive = ev.key in ("control", "shift")
        self.selectionRequested.emit(-1 if sid is None else sid, additive)

    def _on_motion(self, ev):
        if ev.inaxes is not self.ax or ev.xdata is None:
            return
        x, y = ev.xdata, ev.ydata
        if self._drag is not None:
            sid, i = self._drag
            nx, ny = self._snapped(x, y)
            verts = list(self.doc.shape(sid).polygon.vertices)
            verts[i] = (nx, ny, verts[i][2])
            self.doc.set_vertices(sid, verts, checkpoint=False)
            self.cursorInfo.emit(f"꼭짓점 {i} → ({nx:.3f}, {ny:.3f}) mm")
            return
        self.cursorInfo.emit(self.probe_text(x, y))

    def _on_release(self, ev):
        if self._drag is not None:
            sid, _ = self._drag
            self._drag = None
            self.dragFinished.emit(sid)

    def probe_text(self, x: float, y: float) -> str:
        base = f"({x:.3f}, {y:.3f}) mm"
        res = self.result
        if res is None or self._finder is None:
            return base
        t = int(self._finder(x, y))
        if t < 0:
            return base
        tri = self._tri
        idx = tri.triangles[t]
        p = np.column_stack([tri.x[idx], tri.y[idx]])
        T = np.array([[p[0, 0] - p[2, 0], p[1, 0] - p[2, 0]], [p[0, 1] - p[2, 1], p[1, 1] - p[2, 1]]])
        l = np.linalg.solve(T, [x - p[2, 0], y - p[2, 1]])
        lam = np.array([l[0], l[1], 1 - l.sum()])
        phi = abs((lam * res.sol.phi[idx]).sum())
        reg = res.problem.regions[res.mesh.region[t]]
        return (f"{base}   |V|={phi / 1e3:.3f} kV   |E|={self._elem_E[t]:.4f} kV/mm   "
                f"[{reg.name}]")
