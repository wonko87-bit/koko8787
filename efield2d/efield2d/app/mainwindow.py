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
from ..calc import FieldCalculator, FormulaError, parse_constants
from ..dxf_io import join_segments
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
        self.canvas.segmentClicked.connect(self.on_segment_clicked)
        self.canvas.cursorInfo.connect(lambda t: self.statusBar().showMessage(t))

        self.tabs = QTabWidget()
        main = QWidget()
        v = QVBoxLayout(main)
        v.setContentsMargins(0, 0, 0, 0)
        v.addWidget(self.nav)
        v.addWidget(self.canvas, 1)
        self.tabs.addTab(main, "형상 / 결과")
        self._build_path_tab()
        self._build_calc_tab()

        self.setCentralWidget(self.tabs)
        self._build_docks()
        self._build_path_dock()
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
        self.dock_results = d

    def _build_path_tab(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        self.path_fig = Figure(figsize=(7, 6), layout="constrained")
        self.path_canvas = FigureCanvasQTAgg(self.path_fig)
        lay.addWidget(NavigationToolbar2QT(self.path_canvas, w))
        lay.addWidget(self.path_canvas, 1)
        self.path_stats = QLabel("")
        self.path_stats.setWordWrap(True)
        self.path_stats.setTextInteractionFlags(Qt.TextSelectableByMouse)
        lay.addWidget(self.path_stats)
        self.path_data = None
        self.path_vcount: dict[int, int] = {}
        self.tabs.addTab(w, "경로 (연면전계)")

    def _build_path_dock(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        g = QGroupBox("경로 지정")
        f = QVBoxLayout(g)
        self.path_mode_chk = QCheckBox("경로 선택 모드 (캔버스에서 선분/곡선 클릭)")
        self.path_mode_chk.setToolTip("클릭: 윤곽의 한 조각(직선 또는 원호) 추가/해제\n"
                                      "Shift+클릭: 접선이 이어지는 구간 전체 (직선+필렛+직선, 스플라인 등)")
        self.path_mode_chk.toggled.connect(self.on_path_mode)
        self.path_info = QLabel("선택한 조각 없음 — 없으면 선택한 도형의 윤곽을 경로로 사용")
        self.path_info.setWordWrap(True)
        b_clear = QPushButton("경로 지우기")
        b_clear.clicked.connect(self.on_path_clear)
        f.addWidget(self.path_mode_chk)
        f.addWidget(self.path_info)
        f.addWidget(b_clear)
        lay.addWidget(g)

        g = QGroupBox("평가")
        f = QFormLayout(g)
        self.side_cb = QComboBox()
        self.path_n_sp = _spin(10, 5000, 300, 0)
        self.flip_chk = QCheckBox("방향 뒤집기")
        self.flip_chk.setToolTip("기본 방향: 전위가 높은 쪽(시작) → 낮은 쪽(끝)")
        self.comp_cb = QComboBox()
        self.comp_cb.addItem("접선 |Et| (연면)", "Et")
        self.comp_cb.addItem("전체 |E|", "E")
        self.thr_sp = _spin(0.0, 1e6, 0.0, 3, " kV/mm")
        self.thr_sp.setToolTip("이 값 이상인 구간의 누적 길이를 셉니다.\n"
                               "0이면 평가 쪽 영역의 허용전계(재료 설정)를 사용하고, 없으면 사용 안 함")
        f.addRow("평가 쪽(영역)", self.side_cb)
        f.addRow("점 수", self.path_n_sp)
        f.addRow(self.flip_chk)
        f.addRow("임계 비교 성분", self.comp_cb)
        f.addRow("임계 전계 (0=허용전계)", self.thr_sp)
        lay.addWidget(g)
        b_run = QPushButton("경로 해석")
        b_run.clicked.connect(self.on_path)
        b_csv = QPushButton("CSV 저장")
        b_csv.clicked.connect(self.on_path_csv)
        lay.addWidget(b_run)
        lay.addWidget(b_csv)
        lay.addStretch()
        d = QDockWidget("연면 경로", self)
        d.setWidget(w)
        self.addDockWidget(Qt.RightDockWidgetArea, d)
        self.tabifyDockWidget(self.dock_results, d)
        self.dock_results.raise_()
        self.dock_path = d

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
        self._prune_path_keys()
        self._update_path_info()
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

    def _reset_path(self):
        self.canvas.path_keys = []
        self.canvas.path_ends = None
        self.path_data = None
        self.path_vcount.clear()
        self.path_stats.setText("")
        self._update_path_info()

    def open_dxf(self, path: str, replace: bool = True):
        model = self.doc.load_dxf(path, replace=replace)
        if replace:
            self._reset_path()
        self.canvas._have_view = False
        self.canvas.set_result(None)
        self.result = None
        self._calc_fill_combos()
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
        self._reset_path()
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
        self._calc_fill_combos()
        self.statusBar().showMessage(f"완료: 노드 {res.mesh.n_nodes}, 요소 {res.mesh.n_tris}"
                                     + (f"  / {'; '.join(res.warnings)}" if res.warnings else ""))
        self.tabs.setCurrentIndex(0)

    def on_hide_result(self):
        self.result = None
        self._calc_fill_combos()
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
    # ---------------------------------------------------------------- calculator
    CALC_OPS = [
        ("point", "한 점의 값"),
        ("profile", "경로 프로파일 + 선적분"),
        ("volume", "영역 체적적분 ∫ f dV"),
        ("flux", "경로 플럭스 ∫ F·n dA (Gauss 등)"),
        ("electrode", "전극 전하 / 전류 (ΣKφ)"),
        ("review", "일관성 검토 요약 (에너지·전하·손실)"),
    ]
    CALC_PRESETS = {
        "point": [("전계 크기 |E|", "E"), ("반경 방향 Er", "Er"), ("축 방향 Ez", "Ez"), ("전위 |V|", "abs(V)"),
                  ("전속밀도 |D|", "D"), ("전류밀도 |J| = |(σ+jωε)E|", "J"), ("에너지 밀도 ½ε|E|²", "0.5*eps*E**2"),
                  ("손실 밀도 σ|E|²", "sigma*E**2")],
        "volume": [("저장 에너지 ½∫ε|E|²dV", "0.5*eps*E**2"), ("손실 ∫σ|E|²dV", "sigma*E**2"),
                   ("AC 평균 저장 에너지 ¼∫ε|E|²dV", "0.25*eps*E**2"), ("AC 평균 손실 ½∫σ|E|²dV", "0.5*sigma*E**2"),
                   ("체적 (축대칭: 전체)", "1")],
    }
    PROFILE_PRESETS = CALC_PRESETS["point"] + [("경로 접선 성분 Et", "Et"), ("경로 법선 성분 En", "En"),
                                              ("법선 전속밀도 Dn", "Dn")]
    REF_PRESETS = [
        ("(없음)", ""),
        ("동축 단층 |E|(r) = V/(r ln(b/a))", "abs(V_inner)/(r*log(b/a))"),
        ("2층 동축 |E|(r) — 모델 A (경계 r=10 mm)",
         "abs(V_inner/(log(2)*(1/kappa_L1+1/kappa_L2)))/(r*abs(where(r < 10*mm, kappa_L1, kappa_L2)))"),
        ("평행판 |E| = V/d", "abs(V_inner)/d"),
    ]

    def _build_calc_tab(self):
        w = QWidget()
        lay = QHBoxLayout(w)
        left = QVBoxLayout()
        g = QGroupBox("수식 계산기")
        f = QFormLayout(g)
        self.calc_op_cb = QComboBox()
        for k, lab in self.CALC_OPS:
            self.calc_op_cb.addItem(lab, k)
        self.calc_preset_cb = QComboBox()
        self.calc_expr = QLineEdit("E")
        self.calc_expr.setPlaceholderText("예: 0.5*eps*E**2   /   abs(V)   /   Er*r")
        self.calc_consts = QLineEdit()
        self.calc_consts.setPlaceholderText("a=5*mm; b=20*mm; V0=1000   (이름=값; …)")
        self.calc_x, self.calc_y = _spin(-1e6, 1e6, 0.0, 4, " mm"), _spin(-1e6, 1e6, 0.0, 4, " mm")
        self.calc_n = _spin(10, 20000, 400, 0)
        self.calc_side_cb = QComboBox()
        self.calc_region_cb = QComboBox()
        self.calc_field_cb = QComboBox()
        for k, lab in (("D", "D 전속밀도 → 전하 [C]"), ("J", "J 전체 전류밀도 → 전류 [A]"),
                       ("Jc", "Jc 전도 전류밀도 → 전류 [A]"), ("E", "E 전계 → [V·m]")):
            self.calc_field_cb.addItem(lab, k)
        self.calc_elec_cb = QComboBox()
        self.calc_ref = QLineEdit()
        self.calc_ref.setPlaceholderText("경로 위에서 계산값과 겹쳐 그릴 해석해 식 (선택)")
        self.calc_ref_cb = QComboBox()
        for lab, ex in self.REF_PRESETS:
            self.calc_ref_cb.addItem(lab, ex)
        self.calc_smooth = QCheckBox("전계 복원(평활) 사용 — 점·경로·플럭스")
        self.calc_smooth.setChecked(True)
        self.calc_smooth.setToolTip("끄면 요소별 상수 전계를 그대로 사용 (계단 모양, 위치 편향 ±수 %).\n"
                                    "체적적분은 항상 요소값을 사용해 에너지/손실 항등식이 정확합니다.")
        self.calc_flip = QCheckBox("경로 방향 뒤집기 (플럭스 법선 반전)")
        self.calc_btn = QPushButton("계산")
        self.calc_btn.clicked.connect(self.on_calc)
        f.addRow("계산 종류", self.calc_op_cb)
        f.addRow("예시 수식", self.calc_preset_cb)
        f.addRow("수식 f", self.calc_expr)
        f.addRow("상수", self.calc_consts)
        self.calc_pt_row = (QLabel("점 (x, y)"), QWidget())
        prow = QHBoxLayout(self.calc_pt_row[1])
        prow.setContentsMargins(0, 0, 0, 0)
        prow.addWidget(self.calc_x)
        prow.addWidget(self.calc_y)
        f.addRow(*self.calc_pt_row)
        f.addRow("경로 점 수", self.calc_n)
        f.addRow("평가 쪽(영역)", self.calc_side_cb)
        f.addRow("영역", self.calc_region_cb)
        f.addRow("플럭스 대상", self.calc_field_cb)
        f.addRow("전극", self.calc_elec_cb)
        f.addRow("해석해 예시", self.calc_ref_cb)
        f.addRow("비교식 (해석해)", self.calc_ref)
        f.addRow(self.calc_smooth)
        f.addRow(self.calc_flip)
        f.addRow(self.calc_btn)
        left.addWidget(g)
        self.calc_hint = QLabel("")
        self.calc_hint.setWordWrap(True)
        self.calc_hint.setTextInteractionFlags(Qt.TextSelectableByMouse)
        left.addWidget(self.calc_hint)
        left.addStretch()
        lw = QWidget()
        lw.setLayout(left)
        lw.setMaximumWidth(470)
        lay.addWidget(lw)

        right = QVBoxLayout()
        self.calc_out = QPlainTextEdit()
        self.calc_out.setReadOnly(True)
        self.calc_out.setMaximumHeight(210)
        right.addWidget(self.calc_out)
        self.calc_fig = Figure(figsize=(6, 5), layout="constrained")
        self.calc_canvas = FigureCanvasQTAgg(self.calc_fig)
        right.addWidget(NavigationToolbar2QT(self.calc_canvas, w))
        right.addWidget(self.calc_canvas, 1)
        lay.addLayout(right, 1)
        self.tabs.addTab(w, "계산기")

        self._calc = None
        self._calc_key = None
        self.calc_op_cb.currentIndexChanged.connect(self._calc_op_changed)
        self.calc_preset_cb.activated.connect(self._calc_preset_picked)
        self.calc_ref_cb.activated.connect(lambda i: self.calc_ref.setText(self.calc_ref_cb.itemData(i) or ""))
        self._calc_op_changed()
        self._calc_fill_combos()

    def _calc_op_changed(self, *_):
        op = self.calc_op_cb.currentData()
        self.calc_preset_cb.clear()
        presets = self.PROFILE_PRESETS if op == "profile" else self.CALC_PRESETS.get(op, [])
        self.calc_preset_cb.addItem("— 선택 —", None)
        for lab, ex in presets:
            self.calc_preset_cb.addItem(lab, ex)
        usesexpr = op in ("point", "profile", "volume")
        for wdg in (self.calc_preset_cb, self.calc_expr, self.calc_consts):
            wdg.setEnabled(usesexpr)
        self.calc_pt_row[0].setEnabled(op == "point")
        self.calc_pt_row[1].setEnabled(op == "point")
        self.calc_n.setEnabled(op in ("profile", "flux"))
        self.calc_side_cb.setEnabled(op in ("point", "profile", "flux"))
        self.calc_region_cb.setEnabled(op == "volume")
        self.calc_field_cb.setEnabled(op == "flux")
        self.calc_elec_cb.setEnabled(op == "electrode")
        self.calc_ref.setEnabled(op == "profile")
        self.calc_ref_cb.setEnabled(op == "profile")
        self.calc_flip.setEnabled(op in ("profile", "flux"))
        hints = {
            "point": "한 점에서 수식 값을 계산합니다. V는 보간, E·D·J는 복원한 전계입니다.",
            "profile": "경로(연면 경로 탭에서 고른 조각 또는 선택한 도형)를 따라 f를 계산하고 선적분 ∫f dl, 평균, 최대를 구합니다. "
                       "비교식을 넣으면 겹쳐 그리고 오차를 표시합니다.",
            "volume": "모델 전체 또는 한 영역에서 ∫ f dV (평면: 깊이 1 m당, 축대칭: 전체 2πr). 요소별 값을 그대로 합산합니다.",
            "flux": "경로를 축 둘레로 회전(축대칭, 2πr)하거나 깊이 1 m로 늘린 면을 통과하는 D, J 등의 법선 성분을 적분합니다. "
                    "법선은 진행 방향의 오른쪽(반시계 윤곽에서는 바깥쪽).",
            "electrode": "전극 노드의 잔차합 ΣKφ: ES는 전하 [C], DC/AC는 전류 [A]. 요소 상수 전계로 구한 Gauss 적분보다 정확합니다.",
            "review": "같은 양을 서로 다른 방법으로 계산해 일치하는지 봅니다 (에너지 ∫ε|E|²dV ↔ ΣQ·V 등).",
        }
        self.calc_hint.setText(hints[op] + "\n\n변수: x y(=r z) V Ex Ey(=Er Ez) E eps epsr sigma kappa w D Dx Dy J Jx Jy Jc … "
                               "경로 계산에서는 Et En Dn Jn 등. 함수: abs re im conj angle sqrt exp log log10 sin cos atan2 "
                               "mag minimum maximum where. 자동 상수: kappa_<영역> eps_<영역> sigma_<영역> V_<전극> f, 그리고 mm, pi, eps0.")

    def _calc_preset_picked(self, i):
        ex = self.calc_preset_cb.itemData(i)
        if ex:
            self.calc_expr.setText(ex)

    def _calc_fill_combos(self):
        res = self.result
        for cb in (self.calc_side_cb, self.calc_region_cb, self.calc_elec_cb):
            cb.clear()
        self.calc_side_cb.addItem("자동", -1)
        self.calc_region_cb.addItem("전체", -1)
        if res is not None:
            for i, r in enumerate(res.problem.regions):
                self.calc_side_cb.addItem(r.name, i)
                self.calc_region_cb.addItem(r.name, i)
            for name in res.mesh.electrode_nodes:
                self.calc_elec_cb.addItem(name, name)

    def get_calc(self) -> FieldCalculator:
        if self.result is None:
            raise DocumentError("먼저 해석을 실행하세요")
        key = (id(self.result), self.calc_smooth.isChecked())
        if self._calc_key != key:
            self._calc = FieldCalculator(self.result.sol, smooth=self.calc_smooth.isChecked())
            self._calc_key = key
        return self._calc

    @staticmethod
    def fmt_val(v, unit: str = "") -> str:
        v = complex(v)
        u = f" {unit}" if unit else ""
        if v != v:
            return "계산할 수 없음 (NaN)"
        if abs(v.imag) <= 1e-12 * max(abs(v.real), 1e-300):
            return f"{v.real:.6g}{u}"
        return (f"{v.real:.6g} {'+' if v.imag >= 0 else '-'} j{abs(v.imag):.6g}{u}   "
                f"(크기 {abs(v):.6g}, 위상 {math.degrees(math.atan2(v.imag, v.real)):.2f}°)")

    def on_calc(self):
        out = []
        self.calc_fig.clear()
        try:
            calc = self.get_calc()
            op = self.calc_op_cb.currentData()
            expr = self.calc_expr.text()
            consts = parse_constants(self.calc_consts.text())
            side = self.calc_side_cb.currentData()
            side = None if side in (None, -1) else int(side)
            if self.stale:
                out.append("⚠ 해석 입력이 바뀌었습니다 — 이전 해석 결과에 대한 계산입니다 (F5로 다시 실행)")
            if op == "point":
                r = calc.at_point(expr, self.calc_x.value(), self.calc_y.value(), side, consts)
                out += [f"f = {expr}", f"점 ({self.calc_x.value():g}, {self.calc_y.value():g}) mm",
                        f"값 = {self.fmt_val(r.value)}"] + ([r.note] if r.note else [])
            elif op == "volume":
                reg = self.calc_region_cb.currentData()
                reg = None if reg in (None, -1) else int(reg)
                r = calc.volume_integral(expr, reg, consts)
                out += [f"∫ ({expr}) dV   영역: {self.calc_region_cb.currentText()}",
                        f"값 = {self.fmt_val(r.value)}", r.note]
            elif op == "electrode":
                name = self.calc_elec_cb.currentData()
                if name is None:
                    raise FormulaError("전극이 없습니다")
                rows = []
                for nm in self.result.mesh.electrode_nodes:
                    r = calc.electrode_charge(nm)
                    rows.append(f"  {nm}: {self.fmt_val(r.value, r.unit)}")
                r = calc.electrode_charge(name)
                out += [f"선택한 전극 {name}: {self.fmt_val(r.value, r.unit)}", r.note, "", "모든 전극 (들어가고 나간 합은 0):"] + rows
            elif op == "review":
                rows = calc.review()
                out.append("서로 다른 방법으로 계산한 값의 일치 여부 (상대 차이는 작을수록 좋음):")
                for r in rows:
                    ok = "✓" if r["rel"] < 1e-6 else ("△" if r["rel"] < 1e-3 else "✗")
                    out.append(f"{ok} {r['label']}\n     {self.fmt_val(r['a'], r['unit'])}   vs   {self.fmt_val(r['b'], r['unit'])}"
                               f"   (차이 {r['rel']:.2e}){'  — ' + r['note'] if r['note'] else ''}")
                cs = self.circuit_summary(self.result)
                if cs:
                    out += ["", "참고 — " + cs]
            else:                                   # profile / flux need a path
                poly, src = self.current_path()
                if self.calc_flip.isChecked():
                    poly = poly.reversed()
                n = int(self.calc_n.value())
                if op == "flux":
                    r = calc.flux(self.calc_field_cb.currentData(), poly, n, side, consts)
                    out += [f"경로: {src}", f"∫ {self.calc_field_cb.currentData()}·n dA = {self.fmt_val(r.value, r.unit)}", r.note]
                    if r.n_nan:
                        out.append(f"⚠ {r.n_nan}개 점은 값을 구하지 못해 0으로 처리됨 (평가 쪽 영역에 접하지 않거나 메시 밖)")
                    el = self.result.mesh.electrode_nodes
                    if el and self.result.sol.mode and (self.calc_field_cb.currentData() in ("D", "J", "Jc")):
                        q = {nm: calc.electrode_charge(nm).value for nm in el}
                        out.append("참고 — 전극 잔차합: " + ", ".join(f"{k}={self.fmt_val(v)}" for k, v in q.items()))
                else:
                    ref = self.calc_ref.text().strip() or None
                    p = calc.profile(expr, poly, n, side, consts, ref)
                    out += self._profile_text(expr, src, p, ref)
                    self._plot_calc_profile(expr, p, ref)
        except (FormulaError, DocumentError) as e:
            out = [f"⚠ {e}"]
        except Exception as e:                       # pragma: no cover - defensive
            out = [f"⚠ 예기치 않은 오류: {e}"]
        self.calc_out.setPlainText("\n".join(x for x in out if x is not None))
        self.calc_canvas.draw_idle()

    def _profile_text(self, expr, src, p, ref):
        v = p["values"]
        a = np.abs(np.nan_to_num(v))
        k = int(np.argmax(a))
        s0 = p["s"][0]
        lines = [f"f = {expr}     경로: {src}  (길이 {(p['s'][-1] - s0) * 1e3:.4g} mm, 점 {len(v)}개)",
                 f"선적분 ∫ f dl = {self.fmt_val(p['integral'])}   [f의 단위 × m]",
                 f"평균 = {self.fmt_val(p['mean'])}",
                 f"|f| 최대 = {a[k]:.6g}  (시작에서 {(p['s'][k] - s0) * 1e3:.4g} mm)"]
        if ref:
            lines.append(f"비교식 대비: 최대 절대 차이 {p['max_abs_diff']:.4g},  최대 상대 차이 {p['max_rel_diff'] * 100:.3g} % (|비교식| 최대값 기준)")
        if p["n_nan"]:
            lines.append(f"⚠ {p['n_nan']}개 점은 값을 구하지 못해 제외/0 처리됨 (평가 쪽 영역에 접하지 않거나 메시 밖)")
        return lines

    def _plot_calc_profile(self, expr, p, ref):
        fig = self.calc_fig
        rows = 2 if ref else 1
        axes = np.atleast_1d(fig.subplots(rows, 1, sharex=True))
        sm = (p["s"] - p["s"][0]) * 1e3
        v = p["values"]
        cplx = np.iscomplexobj(v) and np.abs(np.imag(v)).max() > 1e-12 * max(np.abs(v).max(), 1e-300)
        y = np.abs(v) if cplx else np.real(v)
        axes[0].plot(sm, y, label=f"{'|f|' if cplx else 'f'} = {expr}")
        if ref:
            r = p["ref"]
            axes[0].plot(sm, np.abs(r) if cplx else np.real(r), "--", color="tab:red", label="비교식")
            with np.errstate(all="ignore"):
                err = (np.abs(v) - np.abs(r)) / np.abs(r).max() * 100 if cplx else (np.real(v) - np.real(r)) / np.abs(r).max() * 100
            axes[1].plot(sm, err, color="tab:purple")
            axes[1].axhline(0, color="k", lw=0.5)
            axes[1].set_ylabel("차이 [% of max|비교식|]")
            axes[1].grid(alpha=0.3)
        axes[0].legend(loc="best", fontsize=8)
        axes[0].grid(alpha=0.3)
        axes[-1].set_xlabel("경로 길이 [mm]")

    # ---- path picking
    def on_path_mode(self, on: bool):
        self.canvas.path_mode = on
        self.statusBar().showMessage(
            "경로 선택 모드: 윤곽 조각을 클릭 (Shift+클릭 = 접선이 이어지는 구간 전체, 다시 클릭하면 해제)" if on else "")

    def _prune_path_keys(self):
        """Drop picked segments whose shape was deleted or whose vertex count changed."""
        keep = []
        for sid, idx in self.canvas.path_keys:
            try:
                sh = self.doc.shape(sid)
            except KeyError:
                continue
            if self.path_vcount.get(sid) != len(sh.polygon.vertices):
                continue
            keep.append((sid, idx))
        self.canvas.path_keys = keep

    def on_segment_clicked(self, sid: int, idx: int, whole_run: bool):
        idxs = self.doc.tangent_run(sid, idx) if whole_run else [idx]
        keys = [(sid, i) for i in idxs]
        cur = list(self.canvas.path_keys)
        if (sid, idx) in cur:
            cur = [k for k in cur if k not in keys]
        else:
            cur += [k for k in keys if k not in cur]
        self.path_vcount[sid] = len(self.doc.shape(sid).polygon.vertices)
        self.canvas.path_keys = cur
        self.canvas.path_ends = None
        self._update_path_info()
        self.canvas.redraw()

    def on_path_clear(self):
        self.canvas.path_keys = []
        self.canvas.path_ends = None
        self._update_path_info()
        self.canvas.redraw()

    def _update_path_info(self):
        keys = self.canvas.path_keys
        if not keys:
            self.path_info.setText("선택한 조각 없음 — 없으면 선택한 도형의 윤곽을 경로로 사용")
            return
        total = sum(self.doc.shape(sid).polygon.segments()[i].length() for sid, i in keys)
        self.path_info.setText(f"선택한 조각 {len(keys)}개, 합계 길이 {total:.3f} mm")

    def current_path(self):
        """The path to analyse as one connected Polygon, and a description of its source."""
        segs, src = [], ""
        if self.canvas.path_keys:
            for sid, idx in self.canvas.path_keys:
                segs.append(self.doc.shape(sid).polygon.segments()[idx])
            src = f"선택 조각 {len(segs)}개"
        else:
            ids = self.selected_ids()
            if not ids:
                raise DocumentError("경로가 없습니다 — '경로 선택 모드'에서 윤곽 조각을 클릭하거나, "
                                    "경로로 쓸 도형을 선택하세요")
            for sid in ids:
                segs += self.doc.shape(sid).polygon.segments()
            src = "도형 " + ", ".join(f"[{i}]" for i in ids)
        chains = join_segments(segs, snap=1e-3)
        if len(chains) != 1:
            ends = []
            for c in chains:
                v = c.vertices
                ends.append(f"({v[0][0]:.2f}, {v[0][1]:.2f}) ~ ({v[-1][0]:.2f}, {v[-1][1]:.2f})")
            raise DocumentError(
                f"선택한 조각이 하나의 연속된 경로로 이어지지 않습니다 ({len(chains)}개 덩어리).\n"
                "끊어진 곳이나 갈라지는 곳이 없는지 확인하세요.\n" + "\n".join(f"  · {e}" for e in ends))
        return chains[0], src

    # ---- path analysis
    def on_path(self):
        if self.result is None:
            QMessageBox.information(self, "경로", "먼저 해석을 실행하세요")
            return
        try:
            poly, src = self.current_path()
        except DocumentError as e:
            QMessageBox.warning(self, "경로", str(e))
            return
        res = self.result
        region = self.side_cb.currentData()
        region = None if region in (None, -1) else int(region)
        n = int(self.path_n_sp.value())
        data = post.along_path(res.sol, poly, n=n, region=region)
        # default direction: from the higher-potential end to the lower one
        phi = np.abs(np.nan_to_num(data["phi"]))
        if (phi[0] < phi[-1]) != self.flip_chk.isChecked():
            poly = poly.reversed()
            data = post.along_path(res.sol, poly, n=n, region=region)

        comp = self.comp_cb.currentData()
        thr = self.thr_sp.value() * 1e6
        if thr <= 0 and region is not None:
            ea = post.allowable(res.problem.regions[region].material.e_allow, res.settings.curve_x)
            thr = ea if ea == ea else 0.0
        m = post.path_metrics(data, thr, comp)
        self.path_data = {"label": src, "poly": poly, "data": data, "metrics": m, "threshold": thr,
                          "component": comp}
        self.canvas.path_ends = (tuple(data["xy"][0]), tuple(data["xy"][-1]))
        self.canvas.redraw()
        self._plot_path(res, data, m, thr, comp)
        self.tabs.setCurrentIndex(1)
        self.statusBar().showMessage(f"경로 해석 완료: {src}")

    def _plot_path(self, res, data, m, thr, comp):
        k = 1e-6  # V/m -> kV/mm
        peak = " (peak)" if res.settings.mode == "ac" else ""
        fig = self.path_fig
        fig.clear()
        rows = 3 if thr > 0 else 2
        axes = fig.subplots(rows, 1, sharex=True)
        a1, a2 = axes[0], axes[1]
        sm = (data["s"] - data["s"][0]) * 1e3
        a1.plot(sm, np.abs(data["Et"]) * k, label="|Et| 접선")
        a1.plot(sm, np.abs(data["En"]) * k, label="|En| 법선")
        a1.plot(sm, data["Emag"] * k, "k", lw=1.2, label="|E|")
        if thr > 0:
            a1.axhline(thr * k, color="r", ls="--", label=f"임계 {thr * k:.3g}")
        a1.set_ylabel("전계 [kV/mm]" + peak)
        a1.legend(loc="best", fontsize=8)
        a1.grid(alpha=0.3)
        a2.plot(sm, np.abs(data["phi"]) / 1e3, color="tab:green", label="전위 |V|")
        a2.plot(sm, np.abs(m["cum_V"]) / 1e3, "--", color="tab:purple", label="누적 ∫Et·ds (전압 강하)")
        a2.plot(sm, m["cum_abs"] / 1e3, ":", color="tab:orange", lw=2, label="누적 ∫|Et|·ds")
        a2.set_ylabel("[kV]" + peak)
        a2.legend(loc="best", fontsize=8)
        a2.grid(alpha=0.3)
        if thr > 0:
            a3 = axes[2]
            a3.fill_between(sm, m["cum_len"] * 1e3, step="post", color="tab:red", alpha=0.35)
            a3.plot(sm, m["cum_len"] * 1e3, color="tab:red")
            a3.set_ylabel(f"누적 길이 [mm]\n(|{'Et' if comp == 'Et' else 'E'}| ≥ {thr * k:.3g} kV/mm)")
            a3.grid(alpha=0.3)
        axes[-1].set_xlabel(f"경로 길이 [mm]  —  {self.path_data['label']}  (시작 → 끝)")
        self.path_canvas.draw_idle()

        lines = [f"길이 {m['length'] * 1e3:.3f} mm  /  점 {len(data['s'])}개  /  {self.path_data['label']}"]
        diff = abs(m["dV"] - m["dV_nodal"]) / max(m["dV_nodal"], 1e-30)
        lines.append(f"|ΔV| = |∫Et·ds| = {m['dV'] / 1e3:.4g} kV   (양 끝 전위차 {m['dV_nodal'] / 1e3:.4g} kV, 차이 {diff * 100:.1f} %)"
                     f"   평균 |ΔV|/L = {m['mean_dv'] * k:.4g} kV/mm")
        name = "|Et|" if comp == "Et" else "|E|"
        lines.append(f"평균 |Et| = {m['mean_abs'] * k:.4g} kV/mm   최대 {name} = {m['max'] * k:.4g} kV/mm"
                     f" (시작에서 {m['s_at_max'] * 1e3:.3f} mm)")
        if thr > 0:
            lines.append(f"{name} ≥ {thr * k:.4g} kV/mm 인 구간의 누적 길이 = {m['len_above'] * 1e3:.3f} mm "
                         f"({m['frac_above'] * 100:.1f} %)")
        else:
            lines.append("임계 전계 없음 — 누적 길이 평가를 하려면 임계값을 입력하거나 평가 쪽 영역에 허용전계를 설정하세요")
        if m["n_nan"]:
            lines.append(f"⚠ {m['n_nan']}개 점은 값을 구하지 못해 0으로 처리됨 (메시 밖이거나, 평가 쪽 영역에 접하지 않음) — "
                         f"경로가 모델 안에 있는지, '평가 쪽(영역)'이 경로가 실제로 접한 영역인지 확인하세요 (자동 = 어느 쪽이든)")
        lines.append("※ 날카로운 모서리나 삼중점 근처의 최대값은 메시에 따라 계속 커집니다 — 평균값과 누적량을 함께 보세요")
        self.path_stats.setText("\n".join(lines))

    def on_path_csv(self):
        if self.path_data is None:
            QMessageBox.information(self, "CSV", "먼저 경로 해석을 실행하세요")
            return
        path, _ = QFileDialog.getSaveFileName(self, "CSV 저장", "", "CSV (*.csv)")
        if path:
            self.save_path_csv(path)

    def save_path_csv(self, path: str):
        d, m = self.path_data["data"], self.path_data["metrics"]
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["s_mm", "x_mm", "y_mm", "abs_V", "abs_Et_V_per_m", "abs_En_V_per_m", "abs_E_V_per_m",
                        "cum_abs_int_Et_ds_V", "cum_abs_int_absEt_ds_V", "cum_len_above_threshold_mm"])
            s0 = d["s"][0]
            for i in range(len(d["s"])):
                w.writerow([(d["s"][i] - s0) * 1e3, d["xy"][i, 0], d["xy"][i, 1], abs(d["phi"][i]),
                            abs(d["Et"][i]), abs(d["En"][i]), d["Emag"][i],
                            abs(m["cum_V"][i]), m["cum_abs"][i], m["cum_len"][i] * 1e3])
