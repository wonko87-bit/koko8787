"""Field calculator, verified on benchmark model A (two-layer coax, closed form)."""
import math
import os
import sys

import numpy as np
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "benchmarks"))
import bench  # noqa: E402

from efield2d import EPS0, Polygon, post  # noqa: E402
from efield2d.calc import FieldCalculator, FormulaError, parse_constants, safe_eval  # noqa: E402

V = 1000.0
RADIAL = Polygon.from_points([(5.0, 10.0), (20.0, 10.0)], closed=False)        # z = 10, r 5 -> 20
UP = lambda r: Polygon.from_points([(r, 0.0), (r, 20.0)], closed=False)        # cylinder at radius r


def idx(c, name):
    """Region index by name (regions are ordered by area, not by drawing order)."""
    return [r.name for r in c.problem.regions].index(name)


@pytest.fixture(scope="module")
def calcs():
    out = {}
    for mode in ("es", "dc", "ac"):
        out[mode] = FieldCalculator(bench.model_a(mode, 0.5).run().sol)
    return out


def consts_for(mode):
    w = 2 * math.pi * 50
    k = {"es": (EPS0 * 2, EPS0 * 4), "dc": (1e-12, 1e-13),
         "ac": (complex(1e-12, w * EPS0 * 2), complex(1e-13, w * EPS0 * 4))}[mode]
    return {"V0": V, "k1": k[0], "k2": k[1]}


# ------------------------------------------------------------- evaluator safety
def test_safe_eval_arithmetic_functions_and_power_caret():
    ns = {"a": np.array([1.0, 4.0]), "x": 3.0}
    assert safe_eval("a**2 + 2*x", ns).tolist() == [7.0, 22.0]
    assert safe_eval("a^2", ns).tolist() == [1.0, 16.0]
    assert safe_eval("sqrt(a) + abs(-x)", ns).tolist() == [4.0, 5.0]
    assert safe_eval("where(a > 2, 1, 0)", ns).tolist() == [0, 1]
    assert safe_eval("1j*x", ns) == 3j
    assert safe_eval("mag(3, 4)", {}) == 5.0


@pytest.mark.parametrize("bad", [
    "__import__('os').system('echo hi')", "x.real", "x[0]", "(lambda: 1)()", "open('f')",
    "[i for i in a]", "a if x else 1", "'text'", "x; y", "abs", "unknown + 1", "foo(1)",
    "sqrt(x=3)", "", "1 +"])
def test_safe_eval_rejects_anything_else(bad):
    with pytest.raises(FormulaError):
        safe_eval(bad, {"a": np.arange(3.0), "x": 3.0})


def test_safe_eval_does_not_hang_on_huge_powers():
    with pytest.raises(FormulaError):                 # reported as an error, instantly
        safe_eval("9**9**9", {})


def test_parse_constants():
    c = parse_constants("a = 5*mm; b=4*a ; V0=1e3")
    assert c == {"a": 0.005, "b": 0.02, "V0": 1000.0}
    for bad in ("a", "1x=3", "sqrt=2", "a=[1]"):
        with pytest.raises(FormulaError):
            parse_constants(bad)


# ------------------------------------------------------- charge / energy identities
@pytest.mark.parametrize("mode", ["es", "dc", "ac"])
def test_electrode_residual_gives_exact_charge_or_current(calcs, mode):
    c = calcs[mode]
    y0, _, _ = bench.a_analytic(mode)
    qi, qo = c.electrode_charge("inner").value, c.electrode_charge("outer").value
    assert qi == pytest.approx(-qo, rel=1e-9)                       # what leaves one electrode enters the other
    assert abs(qi / V / y0 - 1) < 5e-4                              # = Y V (or C V) to the mesh accuracy
    if mode == "es":
        assert qi / V == pytest.approx(post.capacitance(c.sol, V), rel=1e-10)   # 2W/V exactly


def test_volume_integral_energy_equals_half_cv2_and_splits_by_layer(calcs):
    c = calcs["es"]
    w = c.volume_integral("0.5*eps*E**2").value
    assert w == pytest.approx(0.5 * post.capacitance(c.sol, V) * V ** 2, rel=1e-10)
    w1 = c.volume_integral("0.5*eps*E**2", region=idx(c, "L1")).value
    w2 = c.volume_integral("0.5*eps*E**2", region=idx(c, "L2")).value
    assert w1 + w2 == pytest.approx(w, rel=1e-12)
    assert w1 / w == pytest.approx(2.0 / 3.0, abs=1e-3)             # energy ~ 1/eps_r: (1/2) : (1/4)


def test_loss_integral_dc_and_ac(calcs):
    dc = calcs["dc"]
    assert dc.volume_integral("sigma*E**2").value == pytest.approx(
        post.admittance(dc.sol, V).real * V ** 2, rel=1e-10)
    ac = calcs["ac"]
    p = 0.5 * ac.volume_integral("sigma*E**2").value               # time-average, peak phasors
    assert p == pytest.approx(0.5 * post.admittance(ac.sol, V).real * V ** 2, rel=1e-10)


@pytest.mark.parametrize("mode", ["es", "dc", "ac"])
def test_review_checks_agree(calcs, mode):
    rows = calcs[mode].review()
    assert rows and all(r["rel"] < 1e-8 for r in rows), [(r["label"], r["rel"]) for r in rows]


# -------------------------------------------------------------- path operations
@pytest.mark.parametrize("mode", ["es", "dc", "ac"])
def test_radial_line_integral_of_E_is_the_applied_voltage(calcs, mode):
    p = calcs[mode].profile("Er", RADIAL, n=3000)
    assert abs(p["integral"]) == pytest.approx(V, rel=5e-3)         # recovered field: ~0.3 % off
    assert p["n_nan"] == 0
    raw = FieldCalculator(calcs[mode].sol, smooth=False).profile("Er", RADIAL, n=3000)
    assert abs(raw["integral"]) == pytest.approx(V, rel=2e-3)       # P1 gradient: potential drop up to sampling
    assert abs(p["cumulative"][-1]) == pytest.approx(abs(p["integral"]))


@pytest.mark.parametrize("mode", ["es", "dc", "ac"])
def test_profile_against_the_closed_form_reference(calcs, mode):
    ref = "abs(V0/(log(2)*(1/k1+1/k2)))/(r*abs(where(r < 10*mm, k1, k2)))"      # |E|(r) of model A
    p = calcs[mode].profile("E", RADIAL, n=600, consts=consts_for(mode), ref_expr=ref)
    assert p["max_rel_diff"] < 0.01
    # and it really is a comparison: a wrong reference is flagged
    bad = calcs[mode].profile("E", RADIAL, n=600, consts=consts_for(mode), ref_expr="2*(" + ref + ")")
    assert bad["max_rel_diff"] > 0.4


@pytest.mark.parametrize("mode,field", [("es", "D"), ("dc", "Jc"), ("ac", "J")])
def test_gauss_flux_through_cylinders_equals_electrode_charge(calcs, mode, field):
    c = calcs[mode]
    q = c.electrode_charge("inner").value
    for r in (7.5, 15.0):                                           # one in each layer
        f = c.flux(field, UP(r), n=1500)
        assert f.value == pytest.approx(q, rel=0.01)
    # normal D / J is continuous across the interface: both sides agree
    a = c.flux(field, UP(10.0), n=1500, region=idx(c, "L1")).value
    b = c.flux(field, UP(10.0), n=1500, region=idx(c, "L2")).value
    assert a == pytest.approx(b, rel=0.02)
    # the normal is on the right of travel: going down flips the sign
    assert c.flux(field, UP(7.5).reversed(), n=1500).value == pytest.approx(-q, rel=0.02)


def test_normal_and_tangential_components_along_a_path(calcs):
    c = calcs["es"]
    up = c.profile("En", UP(7.5), n=200)                            # normal = radial here
    assert np.allclose(np.abs(up["values"]), np.abs(c.profile("E", UP(7.5), n=200)["values"]), rtol=2e-3)
    et = c.profile("Et", UP(7.5), n=200)["values"]                   # along z: no tangential field
    assert np.abs(et).max() < 0.01 * np.abs(up["values"]).max()


# ----------------------------------------------------------------------- points
def test_values_at_a_point(calcs):
    c = calcs["es"]
    _, E0, V0 = bench.a_analytic("es")
    assert c.at_point("abs(V)", 7.5, 10.0).value == pytest.approx(V0(7.5), rel=1e-3)
    assert c.at_point("epsr", 7.5, 10.0).value == pytest.approx(2.0)
    assert c.at_point("epsr", 15.0, 10.0).value == pytest.approx(4.0)
    assert c.at_point("E", 7.5, 10.0).value == pytest.approx(E0(7.5), rel=0.01)
    assert c.at_point("x/mm + z/mm", 7.5, 10.0).value == pytest.approx(17.5)
    assert math.isnan(c.at_point("abs(V)", 50.0, 10.0).value.real)          # outside the model


def test_region_side_and_error_messages(calcs):
    c = calcs["es"]
    r = c.at_point("abs(V)", 15.0, 10.0, region=idx(c, "L1"))       # L1 does not touch (15, 10)
    assert math.isnan(r.value.real) and "접하지" in r.note
    with pytest.raises(FormulaError, match="경로 계산"):
        c.at_point("Et", 7.5, 10.0)
    with pytest.raises(FormulaError, match="경로 계산"):
        c.volume_integral("En")
    with pytest.raises(FormulaError, match="알 수 없는 이름"):
        c.at_point("Ez2", 7.5, 10.0)
    with pytest.raises(FormulaError, match="예약어"):
        c.at_point("V", 7.5, 10.0, consts={"V": 3.0})
    with pytest.raises(FormulaError, match="전극"):
        c.electrode_charge("nope")
    with pytest.raises(FormulaError):
        c.flux("X", RADIAL)


def test_smoothing_improves_point_and_flux_accuracy_but_not_the_exact_integrals(calcs):
    sol = calcs["es"].sol
    raw, smooth = FieldCalculator(sol, smooth=False), FieldCalculator(sol, smooth=True)
    q = smooth.electrode_charge("inner").value
    err = lambda c: abs(c.flux("D", UP(7.5), n=1500).value / q - 1)
    assert err(smooth) < 0.5 * err(raw) and err(smooth) < 0.01
    _, E0, _ = bench.a_analytic("es")
    perr = lambda c: max(abs(c.at_point("E", r, 10.0).value / E0(r) - 1) for r in (6.0, 7.5, 9.0, 12.0, 15.0, 18.0))
    assert perr(smooth) < perr(raw) and perr(smooth) < 0.01
    # volume integrals do not depend on the smoothing
    assert raw.volume_integral("eps*E**2").value == smooth.volume_integral("eps*E**2").value


def test_syntax_errors_are_formula_errors_in_every_operation(calcs):
    c = calcs["es"]
    for call in (lambda: c.at_point("1 +", 7.5, 10.0), lambda: c.volume_integral("1 +"),
                 lambda: c.profile("1 +", RADIAL)):
        with pytest.raises(FormulaError, match="문법"):
            call()
