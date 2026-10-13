"""Main window."""
from __future__ import annotations

import csv
import math
import os

import numpy as np
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg, NavigationToolbar2QT
from matplotlib.figure import Figure
from PySide6.QtCore import QSignalBlocker, Qt
from PySide6.QtGui import QAction, QColor, QKeySequence
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QComboBox, QDialog, QDialogButtonBox, QDockWidget,
    QDoubleSpinBox, QFileDialog, QFormLayout, QGroupBox, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QMainWindow, QMessageBox, QPlainTextEdit, QPushButton, QTableWidget,
    QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget)

from .. import post
from ..geometry import Polygon
from ..materials import Material
from .canvas import Canvas
from .document import (KIND_LABEL, KINDS, MODE_LABEL, MODES, Document, DocumentError)


def _spin(lo, hi, val, dec=3, suffix="", step=None):
    s = QDoubleSpinBox()
    s.setRange(lo, hi)
    s.setDecimals(dec)
    s.setValue(val)
    if suffix:
        s.setSuffix(suffix)
    if step:
        s.setSingleStep(step)
    s.setKeyboardTracking(False)
    return s


# ------------------------------------------------------------------ dialogs
class ShapeDialog(QDialog):
    """Numeric entry for a new rectangle / circle / polyline."""

    def __init__(self, kind: str, parent=None):
        super().__init__(parent)
        self.kind = kind
        self.setWindowTitle({"rect": "사각형 추가", "circle": "원 추가", "poly": "폴리라인 추가"}[kind])
        form = QFormLayout(self)
        self.f = {}
        if kind == "rect":
            for k, lab in (("x0", "x0 [mm]"), ("y0", "y0 [mm]"), ("x1", "x1 [mm]"), ("y1", "y1 [mm]")):
                self.f[k] = _spin(-1e6, 1e6, 0.0)
                form.addRow(lab, self.f[k])
        elif kind == "circle":
            for k, lab in (("cx", "중심 x [mm]"), ("cy", "중심 y [mm]"), ("r", "반지름 [mm]")):
                self.f[k] = _spin(-1e6, 1e6, 0.0 if k != "r" else 1.0)
                form.addRow(lab, self.f[k])
        else:
            self.text = QPlainTextEdit("0,0\n10,0\n10,10\n0,10")
            self.closed = QCheckBox("닫힌 도형")
            self.closed.setChecked(True)
            form.addRow(QLabel("한 줄에 한 점 (x,y) [mm]"))
            form.addRow(self.text)
            form.addRow(self.closed)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        form.addRow(bb)

    def apply(self, doc: Document) -> int:
        v = {k: w.value() for k, w in self.f.items()}
        if self.kind == "rect":
            return doc.add_rect(v["x0"], v["y0"], v["x1"], v["y1"])
        if self.kind == "circle":
            return doc.add_circle(v["cx"], v["cy"], v["r"])
        pts = []
        for ln in self.text.toPlainText().splitlines():
            ln = ln.strip().replace(";", ",")
            if not ln:
                continue
            try:
                x, y = (float(t) for t in ln.replace(" ", ",").replace(",,", ",").split(",")[:2])
            except ValueError:
                raise DocumentError(f"점을 읽을 수 없습니다: {ln!r}")
            pts.append((x, y))
        return doc.add_polyline(pts, self.closed.isChecked())


def parse_curve(text: str):
    """'0:20; 10:12' -> [(0, 20e6), (10, 12e6)] (x : kV/mm -> V/m)."""
    pts = []
    for part in text.replace("\n", ";").split(";"):
        part = part.strip()
        if not part:
            continue
        x, e = part.split(":")
        pts.append((float(x), float(e) * 1e6))
    return pts or None


class MaterialDialog(QDialog):
    HEAD = ["이름", "εr", "σ [S/m]", "허용전계 [kV/mm]", "허용전계 곡선  x:kV/mm; x:kV/mm"]

    def __init__(self, materials: dict[str, Material], parent=None):
        super().__init__(parent)
        self.setWindowTitle("재료")
        self.resize(780, 360)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel("기본 재료의 값은 예시입니다. 실제 사용 전 직접 확인하세요.\n"
                             "곡선을 입력하면 상수 허용전계보다 우선하며, x 값은 해석 설정의 '곡선 x'(예: 간격 mm)로 평가됩니다."))
        self.tbl = QTableWidget(0, 5)
        self.tbl.setHorizontalHeaderLabels(self.HEAD)
        self.tbl.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        lay.addWidget(self.tbl)
        for m in materials.values():
            self._add(m)
        row = QHBoxLayout()
        b_add, b_del = QPushButton("추가"), QPushButton("선택 삭제")
        b_add.clicked.connect(lambda: self._add(Material("new", 1.0, 0.0)))
        b_del.clicked.connect(self._delete)
        row.addWidget(b_add)
        row.addWidget(b_del)
        row.addStretch()
        lay.addLayout(row)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)
        self.result_materials: dict[str, Material] = {}

    def _add(self, m: Material):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)
        const, curve = "", ""
        if isinstance(m.e_allow, (list, tuple)):
            curve = "; ".join(f"{x:g}:{e / 1e6:g}" for x, e in m.e_allow)
        elif m.e_allow is not None:
            const = f"{m.e_allow / 1e6:g}"
        for c, t in enumerate([m.name, f"{m.eps_r:g}", f"{m.sigma:g}", const, curve]):
            self.tbl.setItem(r, c, QTableWidgetItem(t))

    def _delete(self):
        for r in sorted({i.row() for i in self.tbl.selectedIndexes()}, reverse=True):
            self.tbl.removeRow(r)

    def _ok(self):
        out = {}
        try:
            for r in range(self.tbl.rowCount()):
                t = [(self.tbl.item(r, c).text() if self.tbl.item(r, c) else "").strip() for c in range(5)]
                name = t[0]
                if not name:
                    raise ValueError("이름이 비어 있는 행이 있습니다")
                if name in out:
                    raise ValueError(f"재료 이름이 중복됩니다: {name}")
                curve = parse_curve(t[4]) if t[4] else None
                ea = curve if curve else (float(t[3]) * 1e6 if t[3] else None)
                eps, sig = float(t[1]), float(t[2] or 0)
                if eps <= 0 or sig < 0:
                    raise ValueError(f"{name}: εr는 0보다 크고 σ는 0 이상이어야 합니다")
                out[name] = Material(name, eps, sig, ea)
        except ValueError as e:
            QMessageBox.warning(self, "재료", f"입력 오류: {e}")
            return
        self.result_materials = out
        self.accept()


# ---------------------------------------------------------------- main window
class MainWindow(QMainWindow):
    def __init__(self, doc: Document | None = None):
        super().__init__()
        self.doc = doc or Document()
        self.result = None
        self.stale = False
        self._syncing = False
        self.setWindowTitle("efield2d")
        self.resize(1500, 900)

        self.canvas = Canvas(self.doc)
        self.nav = NavigationToolbar2QT(self.canvas, self)
        self.canvas.toolbar = self.nav
        self.canvas.selectionRequested.connect(self.on_canvas_select)
        self.canvas.dragFinished.connect(self.on_drag_finished)
        self.canvas.cursorInfo.connect(lambda t: self.statusBar().showMessage(t))

        self.tabs = QTabWidget()
        main = QWidget()
        v = QVBoxLayout(main)
        v.setContentsMargins(0, 0, 0, 0)
        v.addWidget(self.nav)
        v.addWidget(self.canvas, 1)
        self.tabs.addTab(main, "형상 / 결과")
        self._build_path_tab()

        self.setCentralWidget(self.tabs)
        self._build_docks()
        self._build_menu()
        self.doc.listeners.append(self.on_doc_changed)
        self.refresh_all()
        self.statusBar().showMessage("DXF를 열어 시작하세요  (파일 ▸ DXF 열기)")

    # ----------------------------------------------------------------- build
    def _build_docks(self):
        # shapes + assignment ------------------------------------------------
        w = QWidget()
        lay = QVBoxLayout(w)
        self.shape_tbl = QTableWidget(0, 5)
        self.shape_tbl.setHorizontalHeaderLabels(["ID", "레이어", "원본", "크기", "할당"])
        self.shape_tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.shape_tbl.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.shape_tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.shape_tbl.verticalHeader().hide()
        self.shape_tbl.horizontalHeader().setStretchLastSection(True)
        self.shape_tbl.itemSelectionChanged.connect(self.on_table_select)
        lay.addWidget(self.shape_tbl, 1)

        g = QGroupBox("선택한 도형에 할당")
        f = QFormLayout(g)
        self.kind_cb = QComboBox()
        for k in KINDS:
            self.kind_cb.addItem(KIND_LABEL[k], k)
        self.mat_cb = QComboBox()
        self.volt_sp = _spin(-1e5, 1e5, 0.0, 3, " kV")
        self.name_ed = QLineEdit()
        self.apply_btn = QPushButton("적용")
        self.apply_btn.clicked.connect(self.on_apply)
        self.kind_cb.currentIndexChanged.connect(self._update_assign_enable)
        f.addRow("종류", self.kind_cb)
        f.addRow("재료", self.mat_cb)
        self.volt_lbl = QLabel("전압")
        f.addRow(self.volt_lbl, self.volt_sp)
        f.addRow("이름", self.name_ed)
        f.addRow(self.apply_btn)
        lay.addWidget(g)
        d = QDockWidget("도형", self)
        d.setWidget(w)
        self.addDockWidget(Qt.LeftDockWidgetArea, d)

        # shape editing ------------------------------------------------------
        w = QWidget()
        lay = QVBoxLayout(w)
        self.edit_chk = QCheckBox("꼭짓점 드래그 편집 (도형 1개 선택)")
        self.edit_chk.toggled.connect(self.on_edit_toggled)
        self.snap_sp = _spin(0.0, 1000.0, 0.0, 3, " mm", 0.1)
        self.snap_sp.valueChanged.connect(lambda v: setattr(self.canvas, "snap", v))
        row = QHBoxLayout()
        row.addWidget(QLabel("격자 스냅 (0=끔)"))
        row.addWidget(self.snap_sp)
        self.vtx_tbl = QTableWidget(0, 3)
        self.vtx_tbl.setHorizontalHeaderLabels(["x [mm]", "y [mm]", "bulge"])
        self.vtx_tbl.itemChanged.connect(self.on_vertex_edited)
        self.vtx_tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        b_ins, b_del = QPushButton("꼭짓점 삽입"), QPushButton("꼭짓점 삭제")
        b_ins.clicked.connect(self.on_vertex_insert)
        b_del.clicked.connect(self.on_vertex_delete)
        r2 = QHBoxLayout()
        r2.addWidget(b_ins)
        r2.addWidget(b_del)
        self.dx_sp, self.dy_sp = _spin(-1e6, 1e6, 0.0, 3, " mm"), _spin(-1e6, 1e6, 0.0, 3, " mm")
        b_mv = QPushButton("이동")
        b_mv.clicked.connect(self.on_translate)
        r3 = QHBoxLayout()
        for x in (QLabel("Δx"), self.dx_sp, QLabel("Δy"), self.dy_sp, b_mv):
            r3.addWidget(x)
        lay.addWidget(self.edit_chk)
        lay.addLayout(row)
        lay.addWidget(self.vtx_tbl, 1)
        lay.addLayout(r2)
        lay.addLayout(r3)
        d = QDockWidget("형상 편집", self)
        d.setWidget(w)
        self.addDockWidget(Qt.LeftDockWidgetArea, d)

        # analysis -----------------------------------------------------------
        w = QWidget()
        lay = QVBoxLayout(w)
        g = QGroupBox("해석 설정")
        f = QFormLayout(g)
        st = self.doc.settings
        self.mode_cb = QComboBox()
        for m in MODES:
            self.mode_cb.addItem(MODE_LABEL[m], m)
        self.sym_cb = QComboBox()
        self.sym_cb.addItem("축대칭 (x=r, y=z)", "axisym")
        self.sym_cb.addItem("평면 (깊이 1 m)", "planar")
        self.freq_sp = _spin(0.001, 1e9, st.freq, 3, " Hz")
        self.rms_chk = QCheckBox("AC 전압은 rms 값으로 입력 (√2배 하여 peak로 해석)")
        self.mesh_sp = _spin(0.001, 1e4, st.mesh_size, 3, " mm")
        self.edge_sp = _spin(0.0, 1e4, st.edge_size, 3, " mm")
        self.trans_sp = _spin(0.0, 1e4, st.edge_transition, 3, " mm")
        self.curve_sp = _spin(-1e9, 1e9, st.curve_x, 3)
        self.edge_sp.setToolTip("전극/전위선 외곽선 위의 요소 크기. 0이면 사용 안 함")
        self.trans_sp.setToolTip("외곽선에서 전체 요소 크기로 커지는 거리. 0이면 전체 크기의 4배")
        self.curve_sp.setToolTip("허용전계 곡선을 평가할 x 값 (예: 절연 간격 mm)")
        f.addRow("종류", self.mode_cb)
        f.addRow("대칭", self.sym_cb)
        f.addRow("주파수", self.freq_sp)
        f.addRow(self.rms_chk)
        f.addRow("전체 요소 크기", self.mesh_sp)
        f.addRow("전극 가장자리 크기", self.edge_sp)
        f.addRow("가장자리 전이 거리", self.trans_sp)
        f.addRow("허용전계 곡선 x", self.curve_sp)
        for wdg in (self.mode_cb, self.sym_cb):
            wdg.currentIndexChanged.connect(self.on_settings_changed)
        self.rms_chk.toggled.connect(self.on_settings_changed)
        for wdg in (self.freq_sp, self.mesh_sp, self.edge_sp, self.trans_sp, self.curve_sp):
            wdg.editingFinished.connect(self.on_settings_changed)
        lay.addWidget(g)
        self.run_btn = QPushButton("해석 실행 (F5)")
        self.run_btn.clicked.connect(self.on_run)
        lay.addWidget(self.run_btn)

        g = QGroupBox("결과 표시")
        f = QFormLayout(g)
        self.qty_cb = QComboBox()
        self.qty_cb.addItem("전계 |E| [kV/mm]", "E")
        self.qty_cb.addItem("전위 |V| [kV]", "phi")
        self.vmax_sp = _spin(0.0, 1e6, 0.0, 3)
        self.vmax_sp.setToolTip("색 범위 상한 (0 = 자동). 모서리 특이점 때문에 최대값이 지나치게 크면 직접 지정하세요")
        self.cont_chk = QCheckBox("등전위선")
        self.cont_chk.setChecked(True)
        self.mesh_chk = QCheckBox("메시")
        self.hide_btn = QPushButton("결과 숨기기")
        self.hide_btn.clicked.connect(self.on_hide_result)
        f.addRow("표시량", self.qty_cb)
        f.addRow("색 상한", self.vmax_sp)
        f.addRow(self.cont_chk)
        f.addRow(self.mesh_chk)
        f.addRow(self.hide_btn)
        self.qty_cb.currentIndexChanged.connect(self.on_view_changed)
        self.vmax_sp.editingFinished.connect(self.on_view_changed)
        self.cont_chk.toggled.connect(self.on_view_changed)
        self.mesh_chk.toggled.connect(self.on_view_changed)
        lay.addWidget(g)

        self.summary = QLabel("")
        self.summary.setWordWrap(True)
        self.summary.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lay.addWidget(self.summary)
        self.util_tbl = QTableWidget(0, 4)
        self.util_tbl.setHorizontalHeaderLabels(["영역", "Emax [kV/mm]", "허용 [kV/mm]", "이용률"])
        self.util_tbl.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.util_tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        lay.addWidget(self.util_tbl, 1)
        d = QDockWidget("해석 / 결과", self)
        d.setWidget(w)
        self.addDockWidget(Qt.RightDockWidgetArea, d)
        self.resizeDocks([d], [380], Qt.Horizontal)

    def _build_path_tab(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        top = QHBoxLayout()
        self.side_cb = QComboBox()
        b_run = QPushButton("선택한 도형의 윤곽을 따라 해석")
        b_run.clicked.connect(self.on_path)
        b_csv = QPushButton("CSV 저장")
        b_csv.clicked.connect(self.on_path_csv)
        self.path_n_sp = _spin(10, 5000, 300, 0)
        top.addWidget(QLabel("평가 쪽(영역)"))
        top.addWidget(self.side_cb)
        top.addWidget(QLabel("점 수"))
        top.addWidget(self.path_n_sp)
        top.addWidget(b_run)
        top.addWidget(b_csv)
        top.addStretch()
        lay.addLayout(top)
        self.path_fig = Figure(figsize=(7, 5), layout="constrained")
        self.path_canvas = FigureCanvasQTAgg(self.path_fig)
        lay.addWidget(NavigationToolbar2QT(self.path_canvas, w))
        lay.addWidget(self.path_canvas, 1)
        self.path_data = None
        self.tabs.addTab(w, "경로 (연면전계)")

    def _build_menu(self):
        def act(menu, text, fn, shortcut=None):
            a = QAction(text, self)
            if shortcut:
                a.setShortcut(shortcut)
            a.triggered.connect(lambda _=False: fn())
            menu.addAction(a)
            return a
        m = self.menuBar().addMenu("파일")
        act(m, "DXF 열기…", self.on_open_dxf, QKeySequence("Ctrl+O"))
        act(m, "DXF 추가 가져오기…", lambda: self.on_open_dxf(replace=False))
        m.addSeparator()
        act(m, "프로젝트 열기…", self.on_open_project, QKeySequence("Ctrl+Shift+O"))
        act(m, "프로젝트 저장", self.on_save, QKeySequence("Ctrl+S"))
        act(m, "다른 이름으로 저장…", self.on_save_as, QKeySequence("Ctrl+Shift+S"))
        m.addSeparator()
        act(m, "종료", self.close, QKeySequence("Ctrl+Q"))
        m = self.menuBar().addMenu("편집")
        self.undo_act = act(m, "실행 취소", self.doc_undo, QKeySequence("Ctrl+Z"))
        self.redo_act = act(m, "다시 실행", self.doc_redo, QKeySequence("Ctrl+Y"))
        m.addSeparator()
        act(m, "사각형 추가…", lambda: self.on_add("rect"))
        act(m, "원 추가…", lambda: self.on_add("circle"))
        act(m, "폴리라인 추가…", lambda: self.on_add("poly"))
        act(m, "선택 도형 삭제", self.on_delete, QKeySequence("Delete"))
        m.addSeparator()
        act(m, "재료…", self.on_materials)
        m = self.menuBar().addMenu("해석")
        act(m, "실행", self.on_run, QKeySequence("F5"))
        m = self.menuBar().addMenu("보기")
        act(m, "전체 보기", self.canvas.fit_view, QKeySequence("Home"))

    # ------------------------------------------------------------- sync UI
    def selected_ids(self) -> list[int]:
        return sorted(self.canvas.selected)

    def on_doc_changed(self, structural: bool):
        self.refresh_all()
        if self.result is not None:
            self.stale = self.doc.signature() != self.result.signature
            self.fill_results(self.result)      # also re-evaluates allowable-field curves

    def refresh_all(self):
        self._syncing = True
        try:
            st = self.doc.settings
            for cb, val in ((self.mode_cb, st.mode), (self.sym_cb, st.symmetry)):
                with QSignalBlocker(cb):
                    cb.setCurrentIndex(max(0, cb.findData(val)))
            for sp, val in ((self.freq_sp, st.freq), (self.mesh_sp, st.mesh_size), (self.edge_sp, st.edge_size),
                            (self.trans_sp, st.edge_transition), (self.curve_sp, st.curve_x)):
                with QSignalBlocker(sp):
                    sp.setValue(val)
            with QSignalBlocker(self.rms_chk):
                self.rms_chk.setChecked(st.rms_input)
            self.freq_sp.setEnabled(st.mode == "ac")
            self.rms_chk.setEnabled(st.mode == "ac")
            with QSignalBlocker(self.mat_cb):
                cur = self.mat_cb.currentText()
                self.mat_cb.clear()
                self.mat_cb.addItems(list(self.doc.materials))
                if cur in self.doc.materials:
                    self.mat_cb.setCurrentText(cur)
            self._fill_shape_table()
            self.canvas.selected &= {s.id for s in self.doc.shapes}
            self.canvas.redraw()
            self._fill_vertex_table()
            self._update_assign_enable()
            self.undo_act.setEnabled(self.doc.can_undo())
            self.redo_act.setEnabled(self.doc.can_redo())
            self.setWindowTitle(f"efield2d — {os.path.basename(self.doc.path) if self.doc.path else '새 프로젝트'}"
                                f"{' *' if self.doc.dirty else ''}")
        finally:
            self._syncing = False

    def _assign_text(self, s) -> str:
        if s.kind == "region":
            return f"영역: {s.material}"
        if s.kind == "conductor":
            return f"도체: {s.voltage / 1e3:g} kV"
        if s.kind == "line":
            return f"전위선: {s.voltage / 1e3:g} kV"
        return "—"

    def _fill_shape_table(self):
        sel = set(self.canvas.selected)
        with QSignalBlocker(self.shape_tbl):
            self.shape_tbl.setRowCount(len(self.doc.shapes))
            for r, s in enumerate(self.doc.shapes):
                size = f"{s.area:.1f} mm²" if s.closed else "열린 선"
                for c, t in enumerate([str(s.id), s.layer, s.source, size, self._assign_text(s)]):
                    it = QTableWidgetItem(t)
                    if s.kind != "ignore":
                        it.setForeground(QColor("#0b5394"))
                    self.shape_tbl.setItem(r, c, it)
                self.shape_tbl.setRowHidden(r, False)
            self.shape_tbl.clearSelection()
            for r, s in enumerate(self.doc.shapes):
                if s.id in sel:
                    for c in range(5):
                        self.shape_tbl.item(r, c).setSelected(True)
            self.shape_tbl.resizeColumnsToContents()
            self.shape_tbl.horizontalHeader().setStretchLastSection(True)

    def _fill_vertex_table(self):
        ids = self.selected_ids()
        with QSignalBlocker(self.vtx_tbl):
            if len(ids) != 1:
                self.vtx_tbl.setRowCount(0)
                return
            s = self.doc.shape(ids[0])
            self.vtx_tbl.setRowCount(len(s.polygon.vertices))
            for r, (x, y, b) in enumerate(s.polygon.vertices):
                for c, val in enumerate((x, y, b)):
                    self.vtx_tbl.setItem(r, c, QTableWidgetItem(f"{val:.6g}"))

    def _update_assign_enable(self):
        k = self.kind_cb.currentData()
        self.mat_cb.setEnabled(k == "region")
        self.volt_sp.setEnabled(k in ("conductor", "line"))
        self.volt_lbl.setEnabled(k in ("conductor", "line"))

    def _load_assign_panel(self):
        ids = self.selected_ids()
        if len(ids) != 1:
            return
        s = self.doc.shape(ids[0])
        with QSignalBlocker(self.kind_cb):
            self.kind_cb.setCurrentIndex(self.kind_cb.findData(s.kind))
        if s.material:
            self.mat_cb.setCurrentText(s.material)
        self.volt_sp.setValue(s.voltage / 1e3)
        self.name_ed.setText(s.name)
        self._update_assign_enable()

    # --------------------------------------------------------------- events
    def select(self, ids, *, from_table=False):
        ids = set(ids)
        self.canvas.set_selection(ids)
        if not from_table:
            self._syncing = True
            self._fill_shape_table()
            self._syncing = False
        self._fill_vertex_table()
        self._load_assign_panel()
        if len(ids) == 1:
            s = self.doc.shape(next(iter(ids)))
            self.statusBar().showMessage(f"[{s.id}] {s.layer} / {s.source} / {self._assign_text(s)}")

    def on_canvas_select(self, sid: int, additive: bool):
        if sid < 0:
            if not additive:
                self.select([])
            return
        cur = set(self.canvas.selected)
        if additive:
            cur ^= {sid}
        else:
            cur = {sid}
        self.select(cur)

    def on_table_select(self):
        if self._syncing:
            return
        rows = sorted({i.row() for i in self.shape_tbl.selectedIndexes()})
        self.select([self.doc.shapes[r].id for r in rows], from_table=True)

    def on_apply(self):
        ids = self.selected_ids()
        if not ids:
            QMessageBox.information(self, "할당", "먼저 도형을 선택하세요")
            return
        kind = self.kind_cb.currentData()
        if kind == "region" and not self.mat_cb.currentText():
            QMessageBox.warning(self, "할당", "재료를 선택하세요")
            return
        try:
            self.doc.assign(ids, kind, material=self.mat_cb.currentText(),
                            voltage=self.volt_sp.value() * 1e3, name=self.name_ed.text().strip())
        except DocumentError as e:
            QMessageBox.warning(self, "할당", str(e))

    def on_settings_changed(self, *_):
        if self._syncing:
            return
        self.doc.set_settings(
            mode=self.mode_cb.currentData(), symmetry=self.sym_cb.currentData(),
            freq=self.freq_sp.value(), rms_input=self.rms_chk.isChecked(),
            mesh_size=self.mesh_sp.value(), edge_size=self.edge_sp.value(),
            edge_transition=self.trans_sp.value(), curve_x=self.curve_sp.value())

    def on_edit_toggled(self, on: bool):
        self.canvas.edit_mode = on
        self.statusBar().showMessage("꼭짓점을 드래그해서 옮기세요 (사각 마커)" if on else "")

    def on_drag_finished(self, sid: int):
        self._fill_vertex_table()

    def on_vertex_edited(self, item):
        ids = self.selected_ids()
        if self._syncing or len(ids) != 1:
            return
        try:
            verts = [tuple(float(self.vtx_tbl.item(r, c).text()) for c in range(3))
                     for r in range(self.vtx_tbl.rowCount())]
            self.doc.set_vertices(ids[0], verts)
        except (ValueError, DocumentError) as e:
            QMessageBox.warning(self, "꼭짓점", f"입력 오류: {e}")
            self._fill_vertex_table()

    def on_vertex_insert(self):
        ids = self.selected_ids()
        if len(ids) != 1:
            return
        s = self.doc.shape(ids[0])
        r = max(0, self.vtx_tbl.currentRow())
        v = list(s.polygon.vertices)
        if not s.closed and r >= len(v) - 1:
            r = len(v) - 2
        a, b = v[r], v[(r + 1) % len(v)]
        v[r] = (a[0], a[1], 0.0)
        v.insert(r + 1, ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2, 0.0))
        self.doc.set_vertices(ids[0], v)
        self._fill_vertex_table()

    def on_vertex_delete(self):
        ids = self.selected_ids()
        if len(ids) != 1 or self.vtx_tbl.currentRow() < 0:
            return
        v = list(self.doc.shape(ids[0]).polygon.vertices)
        del v[self.vtx_tbl.currentRow()]
        try:
            self.doc.set_vertices(ids[0], v)
        except DocumentError as e:
            QMessageBox.warning(self, "꼭짓점", str(e))
        self._fill_vertex_table()

    def on_translate(self):
        ids = self.selected_ids()
        if ids:
            self.doc.translate(ids, self.dx_sp.value(), self.dy_sp.value())

    def on_add(self, kind):
        dlg = ShapeDialog(kind, self)
        if dlg.exec() == QDialog.Accepted:
            try:
                sid = dlg.apply(self.doc)
            except DocumentError as e:
                QMessageBox.warning(self, "추가", str(e))
                return
            self.canvas.fit_view()
            self.select([sid])

    def on_delete(self):
        self.doc.delete(self.selected_ids())

    def doc_undo(self):
        self.doc.undo()

    def doc_redo(self):
        self.doc.redo()

    def on_materials(self):
        dlg = MaterialDialog(self.doc.materials, self)
        if dlg.exec() == QDialog.Accepted:
            self.doc.set_materials(dlg.result_materials)

    # ---------------------------------------------------------------- files
    def maybe_discard(self) -> bool:
        if not self.doc.dirty or not self.doc.shapes:
            return True
        r = QMessageBox.question(self, "저장하지 않은 변경", "변경 사항을 버리고 계속할까요?")
        return r == QMessageBox.Yes

    def open_dxf(self, path: str, replace: bool = True):
        model = self.doc.load_dxf(path, replace=replace)
        self.canvas._have_view = False
        self.canvas.set_result(None)
        self.result = None
        self.canvas.selected.clear()
        self.refresh_all()
        self.canvas.fit_view()
        msg = f"{os.path.basename(path)}: 도형 {len(model.shapes)}개"
        if model.warnings:
            QMessageBox.information(self, "DXF 가져오기", "\n".join(model.warnings))
            msg += f" (경고 {len(model.warnings)}건)"
        self.statusBar().showMessage(msg)

    def on_open_dxf(self, replace: bool = True):
        if replace and not self.maybe_discard():
            return
        path, _ = QFileDialog.getOpenFileName(self, "DXF 열기", "", "DXF (*.dxf)")
        if path:
            try:
                self.open_dxf(path, replace)
            except Exception as e:
                QMessageBox.critical(self, "DXF 가져오기", f"읽지 못했습니다: {e}")

    def on_open_project(self):
        if not self.maybe_discard():
            return
        path, _ = QFileDialog.getOpenFileName(self, "프로젝트 열기", "", "efield2d 프로젝트 (*.efp);;JSON (*.json)")
        if path:
            self.open_project(path)

    def open_project(self, path: str):
        try:
            new = Document.load(path)
        except Exception as e:
            QMessageBox.critical(self, "프로젝트", f"열 수 없습니다: {e}")
            return
        self.doc.listeners.remove(self.on_doc_changed)
        self.doc = new
        self.doc.listeners.append(self.on_doc_changed)
        self.canvas.doc = new
        self.canvas.selected.clear()
        self.canvas._have_view = False
        self.canvas.set_result(None)
        self.result = None
        self.refresh_all()
        self.canvas.fit_view()

    def on_save(self):
        if not self.doc.path:
            return self.on_save_as()
        self.doc.save(self.doc.path)
        self.refresh_all()

    def on_save_as(self):
        path, _ = QFileDialog.getSaveFileName(self, "프로젝트 저장", "", "efield2d 프로젝트 (*.efp)")
        if path:
            if not path.endswith(".efp"):
                path += ".efp"
            self.doc.save(path)
            self.refresh_all()

    def closeEvent(self, ev):
        if self.maybe_discard():
            ev.accept()
        else:
            ev.ignore()

    # -------------------------------------------------------------- results
    def on_run(self):
        from PySide6.QtWidgets import QApplication
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.statusBar().showMessage("메시 생성 / 해석 중…")
        QApplication.processEvents()
        try:
            res = self.doc.run()
        except DocumentError as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, "해석", str(e))
            self.statusBar().showMessage("해석 실패")
            return
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "해석", f"예기치 않은 오류: {e}")
            return
        QApplication.restoreOverrideCursor()
        self.result = res
        self.stale = False
        self.canvas.quantity = self.qty_cb.currentData()
        self.canvas.vmax = self.vmax_sp.value()
        self.canvas.show_contours = self.cont_chk.isChecked()
        self.canvas.show_mesh = self.mesh_chk.isChecked()
        self.canvas.set_result(res)
        self.fill_results(res)
        self.side_cb.clear()
        self.side_cb.addItem("자동", -1)
        for i, r in enumerate(res.problem.regions):
            self.side_cb.addItem(r.name, i)
        self.statusBar().showMessage(f"완료: 노드 {res.mesh.n_nodes}, 요소 {res.mesh.n_tris}"
                                     + (f"  / {'; '.join(res.warnings)}" if res.warnings else ""))
        self.tabs.setCurrentIndex(0)

    def on_hide_result(self):
        self.result = None
        self.canvas.set_result(None)
        self.summary.setText("")
        self.util_tbl.setRowCount(0)

    def on_view_changed(self, *_):
        self.canvas.quantity = self.qty_cb.currentData()
        self.canvas.vmax = self.vmax_sp.value()
        self.canvas.show_contours = self.cont_chk.isChecked()
        self.canvas.show_mesh = self.mesh_chk.isChecked()
        if self.result is not None:
            self.canvas.redraw()

    def circuit_summary(self, res) -> str:
        """Capacitance / conductance between the two voltage levels, if there are exactly two."""
        volts = sorted({c.voltage for c in list(res.problem.conductors) + list(res.problem.lines)},
                       key=lambda v: (v.real, v.imag))
        if len(volts) != 2:
            return ""
        dv = abs(volts[1] - volts[0])
        if dv == 0:
            return ""
        sol = res.sol
        part = "축대칭 전체" if res.problem.symmetry == "axisym" else "깊이 1 m당"
        try:
            if sol.mode == "es":
                return f"전극 간 정전용량 C = {post.capacitance(sol, dv) * 1e12:.4g} pF  ({part}; 모델에 포함된 부분만)"
            Y = post.admittance(sol, dv)
            if sol.mode == "dc":
                return f"전극 간 컨덕턴스 G = {Y.real:.4g} S" + (f", 저항 R = {1 / Y.real:.4g} Ω" if Y.real > 0 else "") \
                    + f"  ({part})"
            w = 2 * math.pi * sol.freq
            return (f"어드미턴스 Y = {Y.real:.4g} + j{Y.imag:.4g} S  →  C = {Y.imag / w * 1e12:.4g} pF, "
                    f"G = {Y.real:.4g} S  ({part})")
        except Exception as e:  # pragma: no cover
            return f"(요약 계산 실패: {e})"

    def fill_results(self, res):
        sol = res.sol
        em = post.element_magnitude(sol) / 1e6
        lines = []
        if self.stale:
            lines.append("⚠ 해석에 영향을 주는 입력이 바뀌었습니다 — 아래 결과는 이전 입력의 것입니다 (F5로 다시 실행)")
        lines.append(f"최대 |E| = {em.max():.4g} kV/mm" + ("  (peak)" if sol.mode == "ac" else ""))
        s = self.circuit_summary(res)
        if s:
            lines.append(s)
        lines.append("※ 날카로운 전극 모서리의 Emax는 메시에 따라 계속 커집니다 — 곡률을 반영하거나 모서리에서 떨어진 값으로 판정하세요")
        self.summary.setText("\n".join(lines))
        rows = post.utilization(sol, self.doc.settings.curve_x)
        self.util_tbl.setRowCount(len(rows))
        for r, u in enumerate(rows):
            ea = u["Eallow"]
            vals = [u["name"], f"{u['Emax'] / 1e6:.4g}", "—" if ea != ea else f"{ea / 1e6:.4g}",
                    "—" if u["utilization"] != u["utilization"] else f"{u['utilization'] * 100:.1f} %"]
            for c, t in enumerate(vals):
                it = QTableWidgetItem(t)
                if c == 3 and u["utilization"] == u["utilization"]:
                    if u["utilization"] > 1.0:
                        it.setBackground(QColor("#f4a6a6"))
                    elif u["utilization"] > 0.8:
                        it.setBackground(QColor("#ffe0a3"))
                self.util_tbl.setItem(r, c, it)

    # ----------------------------------------------------------------- path
    def on_path(self):
        if self.result is None:
            QMessageBox.information(self, "경로", "먼저 해석을 실행하세요")
            return
        ids = self.selected_ids()
        if len(ids) != 1:
            QMessageBox.information(self, "경로", "경로로 쓸 도형을 하나 선택하세요 (형상 탭에서 클릭)")
            return
        s = self.doc.shape(ids[0])
        region = self.side_cb.currentData()
        region = None if region in (None, -1) else int(region)
        res = self.result
        data = post.along_path(res.sol, s.polygon, n=int(self.path_n_sp.value()), region=region)
        self.path_data = (s, data)
        k = 1e-6  # V/m -> kV/mm
        fig = self.path_fig
        fig.clear()
        a1, a2 = fig.subplots(2, 1, sharex=True)
        sm = data["s"] * 1e3
        a1.plot(sm, np.abs(data["Et"]) * k, label="|E_t| 접선")
        a1.plot(sm, np.abs(data["En"]) * k, label="|E_n| 법선")
        a1.plot(sm, data["Emag"] * k, "k", lw=1.2, label="|E|")
        if region is not None:
            ea = post.allowable(res.problem.regions[region].material.e_allow, res.settings.curve_x)
            if ea == ea:
                a1.axhline(ea * k, color="r", ls="--", label="허용전계")
        a1.set_ylabel("전계 [kV/mm]" + (" (peak)" if res.settings.mode == "ac" else ""))
        a1.legend(loc="best", fontsize=8)
        a1.grid(alpha=0.3)
        a2.plot(sm, np.abs(data["phi"]) / 1e3, color="tab:green")
        a2.set_ylabel("|V| [kV]")
        a2.set_xlabel(f"경로 길이 [mm]  —  [{s.id}] {s.label()}")
        a2.grid(alpha=0.3)
        self.path_canvas.draw_idle()
        bad = int(np.isnan(data["Et"]).sum())
        self.statusBar().showMessage(f"경로 해석: 최대 |E_t| = {np.nanmax(np.abs(data['Et'])) * k:.4g} kV/mm"
                                     + (f" ({bad}개 점이 메시 밖)" if bad else ""))
        self.tabs.setCurrentIndex(1)

    def on_path_csv(self):
        if self.path_data is None:
            QMessageBox.information(self, "CSV", "먼저 경로 해석을 실행하세요")
            return
        path, _ = QFileDialog.getSaveFileName(self, "CSV 저장", "", "CSV (*.csv)")
        if path:
            self.save_path_csv(path)

    def save_path_csv(self, path: str):
        s, d = self.path_data
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["s_mm", "x_mm", "y_mm", "abs_V", "abs_Et_V_per_m", "abs_En_V_per_m", "abs_E_V_per_m"])
            for i in range(len(d["s"])):
                w.writerow([d["s"][i] * 1e3, d["xy"][i, 0], d["xy"][i, 1], abs(d["phi"][i]),
                            abs(d["Et"][i]), abs(d["En"][i]), d["Emag"][i]])
