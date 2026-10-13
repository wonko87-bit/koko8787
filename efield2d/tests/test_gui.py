"""Headless GUI tests: real mouse events on the canvas, widgets, results."""
import math

import ezdxf
import numpy as np
import pytest

pytest.importorskip("PySide6")
from matplotlib.backend_bases import MouseEvent  # noqa: E402
from PySide6.QtWidgets import QApplication  # noqa: E402

from efield2d import post  # noqa: E402
from efield2d.app.document import Document  # noqa: E402
from efield2d.app.mainwindow import MainWindow, MaterialDialog, parse_curve  # noqa: E402


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


def make_dxf(path):
    f = math.tan(math.pi / 8)
    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 4
    m = doc.modelspace()
    m.add_lwpolyline([(0, 0), (60, 0), (60, 80), (0, 80)], close=True)            # air box
    m.add_lwpolyline([(0, 0), (60, 0), (60, 5), (0, 5)], close=True)              # ground plate
    m.add_lwpolyline([(0, 5), (12, 5), (12, 45), (0, 45)], close=True)            # spacer
    m.add_lwpolyline([(0, 45, 0), (11, 45, f), (14, 48, 0), (14, 54, f), (11, 57, 0), (0, 57, 0)],
                     close=True, format="xyb")                                   # filleted HV electrode
    doc.saveas(path)


@pytest.fixture
def win(app, tmp_path):
    make_dxf(str(tmp_path / "d.dxf"))
    w = MainWindow()
    w.resize(1300, 800)
    w.show()
    w.open_dxf(str(tmp_path / "d.dxf"))
    w.canvas.draw()
    app.processEvents()
    yield w
    w.doc.dirty = False
    w.close()


def event(win, name, x, y, key=None):
    c = win.canvas
    px, py = c.ax.transData.transform((x, y))
    return MouseEvent(name, c, px, py, button=1, key=key)


def click(win, app, x, y, key=None):
    win.canvas._on_press(event(win, "button_press_event", x, y, key))
    app.processEvents()


def assign(win, app, kind, mat="", kv=0.0, name=""):
    win.kind_cb.setCurrentIndex(win.kind_cb.findData(kind))
    if mat:
        win.mat_cb.setCurrentText(mat)
    win.volt_sp.setValue(kv)
    win.name_ed.setText(name)
    win.on_apply()
    app.processEvents()


def setup_model(win, app):
    click(win, app, 30, 60)
    assign(win, app, "region", "air", name="air")
    click(win, app, 6, 25)
    assign(win, app, "region", "epoxy", name="spacer")
    click(win, app, 30, 2.5)
    assign(win, app, "conductor", kv=0.0, name="gnd")
    click(win, app, 7, 51)
    assign(win, app, "conductor", kv=100.0, name="hv")
    win.canvas.selected.clear()


def test_clicking_selects_the_smallest_shape_under_the_cursor(win, app):
    click(win, app, 30, 60)
    assert win.selected_ids() == [0]                 # only the air box is there
    click(win, app, 6, 25)
    assert win.selected_ids() == [2]                 # spacer (inside the air box) wins
    click(win, app, 0.0 + 12, 30)                    # exactly on the spacer outline
    assert win.selected_ids() == [2]
    click(win, app, 30, 60, key="control")           # ctrl adds
    assert win.selected_ids() == [0, 2]
    click(win, app, 30, 60, key="control")           # ctrl toggles off
    assert win.selected_ids() == [2]
    # table selection mirrors the canvas
    rows = sorted({i.row() for i in win.shape_tbl.selectedIndexes()})
    assert [win.doc.shapes[r].id for r in rows] == [2]


def test_table_selection_drives_canvas(win, app):
    win.shape_tbl.selectRow(1)
    app.processEvents()
    assert win.selected_ids() == [win.doc.shapes[1].id]
    assert win.vtx_tbl.rowCount() == 4


def test_assign_via_panel_and_full_run_ac(win, app):
    setup_model(win, app)
    assert [s.kind for s in win.doc.shapes] == ["region", "conductor", "region", "conductor"]
    win.sym_cb.setCurrentIndex(win.sym_cb.findData("axisym"))
    win.mode_cb.setCurrentIndex(win.mode_cb.findData("ac"))
    win.mesh_sp.setValue(3.0)
    win.mesh_sp.editingFinished.emit()
    win.run_btn.click()
    app.processEvents()
    res = win.result
    assert res is not None and res.sol.mode == "ac" and np.iscomplexobj(res.sol.phi)
    assert win.util_tbl.rowCount() == 2
    assert win.util_tbl.item(0, 0).text() == "air"
    assert "C =" in win.summary.text() and "축대칭" in win.summary.text()
    # rms input -> peak: hv electrode is 100 kV rms = 141.4 kV peak
    assert np.abs(res.sol.phi).max() == pytest.approx(100e3 * math.sqrt(2), rel=1e-9)
    assert not win.stale
    # the result view survives a repaint and the probe reads values
    win.canvas.draw()
    txt = win.canvas.probe_text(30, 30)
    assert "kV/mm" in txt and "air" in txt


def test_dc_and_es_modes_run(win, app):
    setup_model(win, app)
    win.doc.set_settings(symmetry="axisym", mesh_size=3.0, mode="es")
    win.on_run()
    assert win.result.sol.mode == "es" and not np.iscomplexobj(win.result.sol.phi)
    # DC needs a conducting material: epoxy sigma > 0 by default
    win.doc.set_settings(mode="dc")
    win.on_run()
    assert win.result.sol.mode == "dc"
    phi = win.result.sol.phi
    assert phi.max() == pytest.approx(100e3) and phi.min() == pytest.approx(0.0, abs=1e-6)


def test_stale_flag_only_for_solver_relevant_changes(win, app):
    setup_model(win, app)
    win.doc.set_settings(symmetry="axisym", mesh_size=3.0)
    win.on_run()
    win.doc.add_polyline([(12, 5), (12, 45)], closed=False)       # scratch path line
    assert not win.stale
    win.doc.set_settings(curve_x=5.0)
    assert not win.stale
    win.doc.translate([3], 0, 1)                                  # move the electrode
    assert win.stale and "이전 입력" in win.summary.text()
    win.on_run()
    assert not win.stale


def test_vertex_drag_edits_geometry_and_undo_restores(win, app):
    click(win, app, 6, 25)                      # spacer
    win.edit_chk.setChecked(True)
    win.snap_sp.setValue(0.5)
    sid = win.selected_ids()[0]
    before = list(win.doc.shape(sid).polygon.vertices)
    c = win.canvas
    c._on_press(event(win, "button_press_event", 12, 45))         # grab the top-right vertex
    assert c._drag == (sid, 2)          # (0,5) (12,5) (12,45) (0,45): index 2
    c._on_motion(event(win, "motion_notify_event", 15.2, 47.9))
    c._on_release(event(win, "button_release_event", 15.2, 47.9))
    after = win.doc.shape(sid).polygon.vertices
    assert after[2][:2] == pytest.approx((15.0, 48.0))             # snapped to 0.5 mm
    assert after[:2] == before[:2] and after[3:] == before[3:]
    assert win.vtx_tbl.item(2, 0).text() == "15"
    win.doc_undo()
    assert win.doc.shape(sid).polygon.vertices == before


def test_vertex_table_edit_insert_delete_translate(win, app):
    click(win, app, 6, 25)
    sid = win.selected_ids()[0]
    win.vtx_tbl.item(1, 0).setText("13")
    assert win.doc.shape(sid).polygon.vertices[1][0] == 13.0
    win.vtx_tbl.setCurrentCell(0, 0)
    win.on_vertex_insert()
    assert len(win.doc.shape(sid).polygon.vertices) == 5
    win.vtx_tbl.setCurrentCell(1, 0)
    win.on_vertex_delete()
    assert len(win.doc.shape(sid).polygon.vertices) == 4
    win.dx_sp.setValue(2.0)
    win.dy_sp.setValue(-1.0)
    x0, y0, _ = win.doc.shape(sid).polygon.vertices[0]
    win.on_translate()
    x1, y1, _ = win.doc.shape(sid).polygon.vertices[0]
    assert (x1 - x0, y1 - y0) == pytest.approx((2.0, -1.0))


def test_add_and_delete_shapes(win, app):
    n = len(win.doc.shapes)
    sid = win.doc.add_circle(30, 30, 3)
    win.select([sid])
    win.on_delete()
    assert len(win.doc.shapes) == n


def test_creepage_path_and_csv(win, app, tmp_path):
    setup_model(win, app)
    win.doc.set_settings(symmetry="axisym", mesh_size=2.5, edge_size=0.5, edge_transition=5.0, mode="es")
    win.on_run()
    pid = win.doc.add_polyline([(12, 5), (12, 44)], closed=False)
    win.select([pid])
    win.side_cb.setCurrentIndex(win.side_cb.findText("air"))
    win.path_n_sp.setValue(120)
    win.on_path()
    s, d = win.path_data
    assert len(d["s"]) == 120 and not np.isnan(d["Et"]).any()
    # along a surface the tangential field dominates and it is continuous across the interface
    side_air = post.along_path(win.result.sol, s.polygon, 120, region=0)
    side_epx = post.along_path(win.result.sol, s.polygon, 120, region=1)
    mid = slice(20, 100)
    assert np.allclose(np.abs(side_air["Et"][mid]), np.abs(side_epx["Et"][mid]), rtol=0.15)
    f = str(tmp_path / "p.csv")
    win.save_path_csv(f)
    rows = open(f, encoding="utf-8-sig").read().splitlines()
    assert rows[0].startswith("s_mm") and len(rows) == 121


def test_open_project_roundtrip_through_window(win, app, tmp_path):
    setup_model(win, app)
    path = str(tmp_path / "m.efp")
    win.doc.save(path)
    win.open_project(path)
    assert [s.kind for s in win.doc.shapes] == ["region", "conductor", "region", "conductor"]
    assert win.shape_tbl.rowCount() == 4 and win.canvas.doc is win.doc
    win.doc.set_settings(symmetry="axisym", mesh_size=3.0)
    win.on_run()
    assert win.result is not None


def test_material_dialog_parses_curves_and_validates(app):
    assert parse_curve("0:20; 10:12") == [(0.0, 20e6), (10.0, 12e6)]
    d = Document()
    dlg = MaterialDialog(d.materials)
    assert dlg.tbl.rowCount() == len(d.materials)
    dlg.tbl.item(0, 4).setText("0:9; 20:3")
    dlg.tbl.item(0, 1).setText("1.0006")
    accepted = []
    dlg.accept = lambda: accepted.append(True)
    dlg._ok()
    assert accepted and dlg.result_materials["air"].e_allow == [(0.0, 9e6), (20.0, 3e6)]
    assert dlg.result_materials["air"].eps_r == pytest.approx(1.0006)
    assert post.allowable(dlg.result_materials["air"].e_allow, 10.0) == pytest.approx(6e6)


def test_run_reports_model_errors_without_crashing(win, app, monkeypatch):
    shown = []
    from efield2d.app import mainwindow
    monkeypatch.setattr(mainwindow.QMessageBox, "warning", lambda *a, **k: shown.append(a[2]))
    win.on_run()                                  # nothing assigned yet
    assert shown and "영역" in shown[0] and win.result is None
