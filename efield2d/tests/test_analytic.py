"""Verification against closed-form solutions."""
import numpy as np
import pytest

import efield2d as ef
from efield2d import (Problem, Region, DirichletLine, Polygon, Material, rect,
                      build_mesh, solve, post, EPS0)


def line(p0, p1, v, name=""):
    return DirichletLine(Polygon.from_points([p0, p1], closed=False), v, name)


# ---------------------------------------------------------------- parallel plate
def test_parallel_plate_planar_exact():
    """Linear field is reproduced exactly by linear elements."""
    mat = Material("diel", eps_r=3.0)
    w, d, V = 100.0, 10.0, 1000.0  # mm, mm, volt
    p = Problem(regions=[Region(rect(0, 0, w, d), mat)],
                lines=[line((0, 0), (w, 0), 0.0), line((0, d), (w, d), V)],
                mesh_size=3.0)
    m = build_mesh(p)
    s = solve(p, m, "es")
    E = post.element_field(s)
    assert np.allclose(E[:, 0], 0.0, atol=1e-6)
    assert np.allclose(E[:, 1], -V / (d * 1e-3), rtol=1e-9)  # E points toward the low side
    C_exact = EPS0 * 3.0 * (w * 1e-3) / (d * 1e-3)            # per metre depth
    assert post.capacitance(s, V) == pytest.approx(C_exact, rel=1e-9)


# -------------------------------------------------------------------- coaxial
def test_coax_axisym():
    """Infinite coax modelled as an axisymmetric strip: E = V/(r ln(b/a))."""
    a, b, hgt, V = 5.0, 20.0, 20.0, 1000.0
    p = Problem(regions=[Region(rect(a, 0, b, hgt), Material("diel", 2.0))],
                lines=[line((a, 0), (a, hgt), V), line((b, 0), (b, hgt), 0.0)],
                symmetry="axisym", mesh_size=1.0)
    m = build_mesh(p)
    s = solve(p, m, "es")
    r = m.nodes[:, 0]
    phi_exact = V * np.log(b * 1e-3 / r) / np.log(b / a)
    assert np.max(np.abs(s.phi - phi_exact)) < 0.003 * V
    C_exact = 2 * np.pi * EPS0 * 2.0 * (hgt * 1e-3) / np.log(b / a)
    assert post.capacitance(s, V) == pytest.approx(C_exact, rel=0.003)
    # field at element centroids, away from nothing special (no singularities)
    e = post.element_magnitude(s)
    rc = m.nodes[m.tris].mean(axis=1)[:, 0]
    e_exact = V / (rc * np.log(b / a))
    assert np.max(np.abs(e - e_exact) / e_exact) < 0.05


# ------------------------------------------------------ concentric spheres (arcs)
def _sphere_problem(h, a=10.0, b=30.0, V=1000.0):
    # upper half of the (r, z) quarter plane between the two spheres, with the
    # r = 0 axis and the z = 0 plane as natural (symmetry) boundaries.
    # (a,0) -arc-> (0,a) -line-> (0,b) -arc-> (b,0) -line-> (a,0)
    outline = Polygon([(a, 0, np.tan(np.pi / 8)),    # CCW quarter arc (a,0)->(0,a)
                       (0, a, 0.0),
                       (0, b, -np.tan(np.pi / 8)),   # CW quarter arc (0,b)->(b,0)
                       (b, 0, 0.0)])
    inner = Polygon([(a, 0, np.tan(np.pi / 8)), (0, a, 0.0)], closed=False)
    outer = Polygon([(0, b, -np.tan(np.pi / 8)), (b, 0, 0.0)], closed=False)
    return Problem(regions=[Region(outline, Material("diel", 1.0))],
                   lines=[DirichletLine(inner, V), DirichletLine(outer, 0.0)],
                   symmetry="axisym", mesh_size=h)


def _sphere_C_err(h, a=10.0, b=30.0):
    p = _sphere_problem(h, a, b)
    m = build_mesh(p)
    s = solve(p, m, "es")
    C_half = 2 * np.pi * EPS0 * (a * 1e-3) * (b * 1e-3) / ((b - a) * 1e-3)
    return abs(post.capacitance(s, 1000.0) - C_half) / C_half, s, m


def test_sphere_axisym_arcs():
    err, s, m = _sphere_C_err(2.0)
    assert err < 0.01
    a, b, V = 10.0, 30.0, 1000.0
    R = np.hypot(m.nodes[:, 0], m.nodes[:, 1]) * 1e3
    phi_exact = V * (1 / R - 1 / b) / (1 / a - 1 / b)
    assert np.max(np.abs(s.phi - phi_exact)) < 0.01 * V


def test_sphere_converges():
    errs = [_sphere_C_err(h)[0] for h in (4.0, 2.0, 1.0)]
    assert errs[0] > errs[1] > errs[2]
    assert errs[2] < 0.3 * errs[0]


# ------------------------------------------------------------ two-layer dielectric
def _two_layer(d1=4.0, d2=6.0, w=50.0):
    m1 = Material("A", eps_r=2.0, sigma=1e-9)
    m2 = Material("B", eps_r=4.0, sigma=1e-12)
    p = Problem(regions=[Region(rect(0, 0, w, d1), m1, "A"), Region(rect(0, d1, w, d1 + d2), m2, "B")],
                lines=[line((0, 0), (w, 0), 0.0), line((0, d1 + d2), (w, d1 + d2), 1000.0)],
                mesh_size=2.5)
    return p, m1, m2


def _layer_fields(s):
    E = post.element_field(s)[:, 1]
    reg = s.mesh.region
    return E[reg == 0].mean(), E[reg == 1].mean()


def test_two_layer_es_divides_by_permittivity():
    p, m1, m2 = _two_layer()
    s = solve(p, build_mesh(p), "es")
    e1, e2 = _layer_fields(s)
    assert e1 / e2 == pytest.approx(m2.eps_r / m1.eps_r, rel=1e-9)


def test_two_layer_dc_divides_by_conductivity():
    """DC: the *lower-sigma* layer takes the field - not what eps would say."""
    p, m1, m2 = _two_layer()
    s = solve(p, build_mesh(p), "dc")
    e1, e2 = _layer_fields(s)
    assert e1 / e2 == pytest.approx(m2.sigma / m1.sigma, rel=1e-6)
    assert abs(e2) > 100 * abs(e1)
    # the V drop adds up
    assert abs(e1 * 4e-3 + e2 * 6e-3) == pytest.approx(1000.0, rel=1e-6)


@pytest.mark.parametrize("f", [1e-6, 1e-3, 50.0, 1e6])
def test_two_layer_ac_matches_complex_kappa_ratio(f):
    p, m1, m2 = _two_layer()
    s = solve(p, build_mesh(p), "ac", freq=f)
    e1, e2 = _layer_fields(s)
    k1, k2 = m1.kappa("ac", f), m2.kappa("ac", f)
    assert e1 / e2 == pytest.approx(k2 / k1, rel=1e-6)


def test_ac_limits_reduce_to_es_and_dc():
    p, m1, m2 = _two_layer()
    m = build_mesh(p)
    hi = solve(p, m, "ac", freq=1e9)
    e1, e2 = _layer_fields(hi)
    assert abs(e1 / e2) == pytest.approx(m2.eps_r / m1.eps_r, rel=1e-3)
    lo = solve(p, m, "ac", freq=1e-9)
    e1, e2 = _layer_fields(lo)
    assert abs(e1 / e2) == pytest.approx(m2.sigma / m1.sigma, rel=1e-3)


def test_energy_consistent_with_charge():
    """Admittance of the lossless AC problem gives j*omega*C."""
    p, _, _ = _two_layer()
    m = build_mesh(p)
    # lossless copy of the same model
    es = solve(p, m, "es")
    C = post.capacitance(es, 1000.0)
    p2 = Problem(regions=[Region(r.polygon, Material(r.material.name, r.material.eps_r, 0.0), r.name)
                          for r in p.regions], lines=p.lines, mesh_size=p.mesh_size)
    ac = solve(p2, build_mesh(p2), "ac", freq=50.0)
    Y = post.admittance(ac, 1000.0)
    assert Y.imag / (2 * np.pi * 50.0) == pytest.approx(C, rel=1e-9)
    assert abs(Y.real) < 1e-9 * abs(Y.imag)
