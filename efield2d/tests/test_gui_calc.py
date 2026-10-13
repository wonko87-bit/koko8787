"""GUI calculator driven on benchmark model A."""
import math
import os
import sys

import numpy as np
import pytest

pytest.importorskip("PySide6")
from PySide6.QtWidgets import QApplication  # noqa: E402

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "benchmarks"))
import bench  # noqa: E402

from efield2d import post  # noqa: E402
from efield2d.app.mainwindow import MainWindow  # noqa: E402


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


def make_window(app, mode="es"):
    doc = bench.model_a(mode, 0.5)
    w = MainWindow(doc)
    w.resize(1500, 900)
    w.show()
    w.on_run()
    app.processEvents()
    return w


@pytest.fixture
def win(app):
    w = make_window(app, "es")
    yield w
    w.doc.dirty = False
    w.close()


def set_op(win, key):
    win.calc_op_cb.setCurrentIndex(win.calc_op_cb.findData(key))


def calc(win, op, expr="", **kw):
    set_op(win, op)
    win.calc_expr.setText(expr)
    for k, v in kw.items():
        getattr(win, k).setValue(v) if hasattr(getattr(win, k), "setValue") else getattr(win, k).setText(v)
    win.on_calc()
    return win.calc_out.toPlainText()


def select_path(win, pts):
    sid = win.doc.add_polyline(pts, closed=False)
    win.select([sid])
    return sid


def test_all_operations_are_listed_and_enable_the_right_inputs(win):
    assert [win.calc_op_cb.itemData(i) for i in range(win.calc_op_cb.count())] == \
        ["point", "profile", "volume", "flux", "electrode", "review"]
    set_op(win, "point")
    assert win.calc_x.isEnabled() and not win.calc_region_cb.isEnabled() and not win.calc_ref.isEnabled()
    set_op(win, "profile")
    assert win.calc_ref.isEnabled() and win.calc_n.isEnabled() and not win.calc_x.isEnabled()
    set_op(win, "volume")
    assert win.calc_region_cb.isEnabled() and not win.calc_field_cb.isEnabled()
    set_op(win, "flux")
    assert win.calc_field_cb.isEnabled() and not win.calc_expr.isEnabled()
    set_op(win, "electrode")
    assert win.calc_elec_cb.isEnabled()
    assert [win.calc_region_cb.itemText(i) for i in range(win.calc_region_cb.count())][0] == "전체"
    assert win.calc_elec_cb.count() == 2


def test_point_value_matches_the_closed_form(win):
    _, _, V0 = bench.a_analytic("es")
    out = calc(win, "point", "abs(V)")
    assert "값 =" in out
    win.calc_x.setValue(7.5)
    win.calc_y.setValue(10.0)
    out = calc(win, "point", "abs(V)")
    val = float(out.split("값 =")[1].split()[0])
    assert val == pytest.approx(V0(7.5), rel=1e-3)
    # preset menu fills the formula
    win.calc_op_cb.setCurrentIndex(win.calc_op_cb.findData("point"))
    i = win.calc_preset_cb.findText("에너지 밀도 ½ε|E|²")
    win.calc_preset_cb.activated.emit(i)
    win.calc_preset_cb.setCurrentIndex(i)
    win._calc_preset_picked(i)
    assert win.calc_expr.text() == "0.5*eps*E**2"


def test_energy_and_charge_identities_in_the_gui(win):
    sol = win.result.sol
    C = post.capacitance(sol, 1000.0)
    out = calc(win, "volume", "0.5*eps*E**2")
    assert float(out.split("값 =")[1].split()[0]) == pytest.approx(0.5 * C * 1e6, rel=1e-5)   # printed with 6 digits
    out = calc(win, "electrode")
    assert "inner" in out and "outer" in out
    q = float(out.split("선택한 전극")[1].split(":")[1].split()[0])
    assert abs(q) == pytest.approx(C * 1000.0, rel=1e-5)
    out = calc(win, "review")
    assert "✓" in out and "✗" not in out and "저장 에너지" in out


def test_profile_against_the_analytic_reference_with_plot(win):
    select_path(win, [(5.0, 10.0), (20.0, 10.0)])
    win.calc_ref_cb.setCurrentIndex(2)                    # the 2-layer coax preset
    win.calc_ref.setText(win.calc_ref_cb.itemData(2))
    out = calc(win, "profile", "E")
    assert "비교식 대비" in out
    rel = float(out.split("최대 상대 차이")[1].split("%")[0])
    assert rel < 1.0                                       # percent
    assert len(win.calc_fig.axes) == 2                      # values + difference plot
    # a deliberately wrong reference is flagged
    win.calc_ref.setText("2*" + win.calc_ref_cb.itemData(2))
    out = calc(win, "profile", "E")
    assert float(out.split("최대 상대 차이")[1].split("%")[0]) > 40
    # the line integral of Er is the applied voltage
    win.calc_ref.setText("")
    out = calc(win, "profile", "Er")
    assert float(out.split("∫ f dl =")[1].split()[0]) == pytest.approx(1000.0, rel=5e-3)
    assert len(win.calc_fig.axes) == 1


def test_flux_matches_the_electrode_charge(win):
    q = win._calc_value = None
    select_path(win, [(7.5, 0.0), (7.5, 20.0)])
    win.calc_field_cb.setCurrentIndex(win.calc_field_cb.findData("D"))
    out = calc(win, "flux")
    flux = float(out.split("= ")[1].split()[0])
    sol_q = post.capacitance(win.result.sol, 1000.0) * 1000.0
    assert flux == pytest.approx(sol_q, rel=0.01)
    assert "전극 잔차합" in out
    win.calc_flip.setChecked(True)                         # reversed path: normal flips -> negative
    out = calc(win, "flux")
    assert float(out.split("= ")[1].split()[0]) == pytest.approx(-sol_q, rel=0.01)


def test_user_constants_auto_constants_and_errors_are_reported_inline(win):
    win.calc_x.setValue(7.5)
    win.calc_y.setValue(10.0)
    out = calc(win, "point", "abs(V)/V_inner")             # auto constant for the electrode voltage
    assert 0.5 < float(out.split("값 =")[1].split()[0]) < 0.7
    win.calc_consts.setText("a = 5*mm; b = 20*mm")
    out = calc(win, "point", "x/a")
    assert float(out.split("값 =")[1].split()[0]) == pytest.approx(1.5)
    win.calc_consts.setText("")
    for bad, msg in (("__import__('os')", "⚠"), ("1 +", "문법"), ("nope", "알 수 없는 이름"), ("Et", "경로 계산")):
        out = calc(win, "point", bad)
        assert out.startswith("⚠") and msg in out
    win.calc_consts.setText("V=3")
    assert "예약어" in calc(win, "point", "V")
    win.calc_consts.setText("")
    set_op(win, "profile")
    win.calc_expr.setText("E")
    win.canvas.path_keys = []
    win.select([])
    win.on_calc()
    assert "경로가 없습니다" in win.calc_out.toPlainText()


def test_calculator_without_a_result_and_stale_notice(app):
    w = MainWindow()
    w.show()
    set_op(w, "point")
    w.calc_expr.setText("E")
    w.on_calc()
    assert "먼저 해석을 실행" in w.calc_out.toPlainText()
    w.doc.dirty = False
    w.close()
    w2 = make_window(app, "es")
    w2.doc.translate([w2.doc.shapes[0].id], 0, 0.0)           # same geometry: not stale ...
    assert not w2.stale
    w2.doc.set_settings(mesh_size=0.7)                         # ... a solver setting changes: stale
    set_op(w2, "point")
    w2.calc_expr.setText("E")
    w2.on_calc()
    assert "이전 해석 결과" in w2.calc_out.toPlainText()
    w2.doc.dirty = False
    w2.close()


@pytest.mark.parametrize("mode", ["dc", "ac"])
def test_dc_and_ac_modes_in_the_gui(app, mode):
    w = make_window(app, mode)
    out = calc(w, "review")
    assert "손실" in out and "✗" not in out
    sel = w.doc.add_polyline([(7.5, 0.0), (7.5, 20.0)], closed=False)
    w.select([sel])
    w.calc_field_cb.setCurrentIndex(w.calc_field_cb.findData("J"))
    out = calc(w, "flux")
    assert "A" in out and "전극 잔차합" in out
    y0, _, _ = bench.a_analytic(mode)
    el = w.result.mesh.electrode_nodes
    out = calc(w, "electrode")
    assert "A" in out
    w.doc.dirty = False
    w.close()


def test_profile_csv_export(win, tmp_path):
    select_path(win, [(5.0, 10.0), (20.0, 10.0)])
    win.calc_ref.setText("abs(V_inner/(log(2)*(1/kappa_L1+1/kappa_L2)))/(r*abs(where(r < 10*mm, kappa_L1, kappa_L2)))")
    calc(win, "profile", "E", calc_n=61)
    f = str(tmp_path / "p.csv")
    win.save_calc_csv(f)
    lines = open(f, encoding="utf-8-sig").read().splitlines()
    assert lines[0].startswith("# f = E") and lines[1] == "s_mm,x_mm,y_mm,f,ref"
    assert len(lines) == 2 + 61
    first = [float(x) for x in lines[2].split(",")]
    assert first[1] == pytest.approx(5.0) and first[3] == pytest.approx(first[4], rel=0.01)
