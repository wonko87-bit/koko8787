"""Geometry rules, electrodes, path sampling, refinement, withstand check."""
import numpy as np
import pytest

from efield2d import (Problem, Region, Conductor, DirichletLine, Refinement,
                      Polygon, Material, rect, build_mesh, solve, post)

AIR = Material("air", 1.0, e_allow=3.0e6)
EPOXY = Material("epoxy", 4.0, sigma=1e-14, e_allow=[(0, 20e6), (10, 10e6)])


def test_conductor_is_carved_out_and_held_at_voltage():
    p = Problem(regions=[Region(rect(0, 0, 60, 60), AIR)],
                conductors=[Conductor(rect(10, 10, 20, 20), 1000.0, "hv"),
                            Conductor(rect(40, 40, 50, 50), 0.0, "gnd")],
                mesh_size=4.0)
    m = build_mesh(p)
    c = m.nodes[m.tris].mean(axis=1) * 1e3
    assert not np.any((c[:, 0] > 10.01) & (c[:, 0] < 19.99) & (c[:, 1] > 10.01) & (c[:, 1] < 19.99))
    s = solve(p, m, "es")
    assert np.allclose(s.phi[m.electrode_nodes["hv"]], 1000.0)
    assert np.allclose(s.phi[m.electrode_nodes["gnd"]], 0.0)
    # potential stays between the electrodes (maximum principle)
    assert s.phi.min() >= -1e-9 and s.phi.max() <= 1000.0 + 1e-9


def test_later_region_overrides_earlier():
    """An insulator drawn inside an air box replaces the air there."""
    p = Problem(regions=[Region(rect(0, 0, 60, 60), AIR, "air"),
                         Region(rect(20, 20, 40, 40), EPOXY, "epoxy")],
                conductors=[Conductor(rect(0, 0, 60, 3), 0.0, "gnd"),
                            Conductor(rect(0, 57, 60, 60), 1000.0, "hv")],
                mesh_size=4.0)
    m = build_mesh(p)
    c = m.nodes[m.tris].mean(axis=1) * 1e3
    inside = (c[:, 0] > 20) & (c[:, 0] < 40) & (c[:, 1] > 20) & (c[:, 1] < 40)
    assert np.all(m.region[inside] == 1)
    assert np.all(m.region[~inside] == 0)


def test_creepage_path_on_dielectric_rod_surface():
    """Rod between plates (axisymmetric): E is uniform V/d and purely tangential
    along the rod surface."""
    d, R, V = 20.0, 5.0, 2000.0
    p = Problem(regions=[Region(rect(0, 0, 15, d), AIR, "air"),
                         Region(rect(0, 0, R, d), EPOXY, "rod")],
                lines=[DirichletLine(Polygon.from_points([(0, 0), (15, 0)], closed=False), 0.0),
                       DirichletLine(Polygon.from_points([(0, d), (15, d)], closed=False), V)],
                symmetry="axisym", mesh_size=2.0)
    s = solve(p, build_mesh(p), "es")
    surface = Polygon.from_points([(R, 0.5), (R, d - 0.5)], closed=False)
    for side in (0, 1):  # evaluate from the air side and from the rod side
        r = post.along_path(s, surface, n=40, region=side)
        assert np.allclose(np.abs(r["Et"]), V / (d * 1e-3), rtol=1e-6)
        assert np.max(np.abs(r["En"])) < 1e-6 * V / (d * 1e-3)
    assert r["s"][-1] == pytest.approx((d - 1.0) * 1e-3)


def test_local_refinement_adds_elements_near_point():
    base = Problem(regions=[Region(rect(0, 0, 40, 40), AIR)],
                   lines=[DirichletLine(Polygon.from_points([(0, 0), (40, 0)], closed=False), 0.0),
                          DirichletLine(Polygon.from_points([(0, 40), (40, 40)], closed=False), 1.0)],
                   mesh_size=5.0)
    n0 = build_mesh(base).n_tris
    base.refinements = [Refinement(20, 20, size=0.5, radius=3.0, transition=8.0)]
    m = build_mesh(base)
    assert m.n_tris > 3 * n0
    c = m.nodes[m.tris].mean(axis=1) * 1e3
    near = np.hypot(c[:, 0] - 20, c[:, 1] - 20) < 3.0
    assert near.sum() > 100


def test_edge_refinement_on_electrode():
    mk = lambda **kw: Problem(
        regions=[Region(rect(0, 0, 60, 60), AIR)],
        conductors=[Conductor(rect(20, 20, 40, 40), 1.0), ],
        lines=[DirichletLine(Polygon.from_points([(0, 0), (60, 0)], closed=False), 0.0)],
        mesh_size=6.0, **kw)
    n0 = build_mesh(mk()).n_tris
    n1 = build_mesh(mk(edge_size=1.0, edge_transition=10.0)).n_tris
    assert n1 > 2 * n0


def test_utilization_uses_constant_and_curve_allowables():
    p = Problem(regions=[Region(rect(0, 0, 50, 10), AIR, "air")],
                lines=[DirichletLine(Polygon.from_points([(0, 0), (50, 0)], closed=False), 0.0),
                       DirichletLine(Polygon.from_points([(0, 10), (50, 10)], closed=False), 15e3)],
                mesh_size=3.0)
    s = solve(p, build_mesh(p), "es")
    u = post.utilization(s)[0]
    assert u["Emax"] == pytest.approx(1.5e6, rel=1e-9)         # 15 kV / 10 mm
    assert u["utilization"] == pytest.approx(0.5, rel=1e-9)    # vs 3 MV/m
    assert post.allowable(EPOXY.e_allow, 5.0) == pytest.approx(15e6)
    assert post.allowable(EPOXY.e_allow, 100.0) == pytest.approx(10e6)  # flat outside


def test_dc_needs_conductive_material():
    p = Problem(regions=[Region(rect(0, 0, 10, 10), AIR)],
                lines=[DirichletLine(Polygon.from_points([(0, 0), (10, 0)], closed=False), 0.0)])
    with pytest.raises(ValueError):
        solve(p, build_mesh(p), "dc")


def _line(a, b, v):
    return DirichletLine(Polygon.from_points([a, b], closed=False), v)


def test_open_line_on_part_of_boundary_sets_only_that_part():
    """An electrode drawn as an open line touching just part of the outline."""
    p = Problem(regions=[Region(rect(0, 0, 60, 40), AIR)],
                lines=[_line((0, 0), (60, 0), 0.0), _line((20, 40), (40, 40), 100.0)],
                mesh_size=3.0)
    m = build_mesh(p)
    x, y = (m.nodes * 1e3).T
    on_hv = np.isclose(y, 40.0) & (x >= 20 - 1e-6) & (x <= 40 + 1e-6)
    s = solve(p, m, "es")
    assert on_hv.sum() > 3 and np.allclose(s.phi[on_hv], 100.0)
    free_top = np.isclose(y, 40.0) & ((x < 20 - 1e-6) | (x > 40 + 1e-6))
    assert not np.any(np.isin(m.bc_nodes, np.flatnonzero(free_top)))   # rest of the edge is free
    assert np.all((s.phi[free_top] > 0) & (s.phi[free_top] < 100))


def test_open_line_bending_around_a_corner():
    p = Problem(regions=[Region(rect(0, 0, 60, 40), AIR)],
                lines=[_line((0, 0), (60, 0), 0.0),
                       DirichletLine(Polygon.from_points([(30, 40), (60, 40), (60, 20)], closed=False), 100.0)],
                mesh_size=3.0)
    m = build_mesh(p)
    s = solve(p, m, "es")
    x, y = (m.nodes * 1e3).T
    corner = np.isclose(x, 60.0) & (y >= 20 - 1e-6)
    assert np.allclose(s.phi[corner], 100.0)


def test_thin_electrode_inside_region():
    p = Problem(regions=[Region(rect(0, 0, 60, 40), AIR)],
                lines=[_line((0, 0), (60, 0), 0.0), _line((20, 20), (40, 20), 100.0)],
                mesh_size=3.0)
    m = build_mesh(p)
    s = solve(p, m, "es")
    assert s.phi.max() <= 100.0 + 1e-9
    x, y = (m.nodes * 1e3).T
    assert np.allclose(s.phi[np.isclose(y, 20.0) & (x >= 20 - 1e-6) & (x <= 40 + 1e-6)], 100.0)


def test_line_outside_every_region_is_an_error():
    p = Problem(regions=[Region(rect(0, 0, 60, 40), AIR)],
                lines=[_line((0, 0), (60, 0), 0.0), _line((20, 50), (40, 50), 100.0)],
                mesh_size=3.0)
    with pytest.raises(ValueError, match="does not lie on or inside"):
        build_mesh(p)


def test_conductor_outside_regions_is_reported():
    p = Problem(regions=[Region(rect(0, 0, 60, 40), AIR)],
                lines=[_line((0, 0), (60, 0), 0.0)],
                conductors=[Conductor(rect(100, 100, 110, 110), 5.0, "stray")],
                mesh_size=3.0)
    m = build_mesh(p)
    assert any("stray" in w for w in m.warnings)
