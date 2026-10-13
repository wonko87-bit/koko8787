"""Segment chaining, tangents and cumulative path metrics."""
import math

import numpy as np
import pytest

from efield2d import (Conductor, DirichletLine, Material, Polygon, Problem, Region, build_mesh,
                      post, rect, solve)
from efield2d.dxf_io import join_segments
from efield2d.geometry import Segment


def test_segment_tangents_line_and_arcs():
    (a, b) = Segment((0, 0), (10, 0), 0.0).tangents()
    assert a == b == pytest.approx((1.0, 0.0))
    # CCW quarter circle (10,0) -> (0,10) around the origin: starts going up, ends going left
    t0, t1 = Segment((10, 0), (0, 10), math.tan(math.pi / 8)).tangents()
    assert t0 == pytest.approx((0.0, 1.0), abs=1e-12) and t1 == pytest.approx((-1.0, 0.0), abs=1e-12)
    # the same arc traversed clockwise (0,10) -> (10,0): starts going right, ends going down
    t0, t1 = Segment((0, 10), (10, 0), -math.tan(math.pi / 8)).tangents()
    assert t0 == pytest.approx((1.0, 0.0), abs=1e-12) and t1 == pytest.approx((0.0, -1.0), abs=1e-12)


def test_join_segments_orders_and_orients_pieces():
    segs = [Segment((10, 5), (10, 0), 0.0),            # drawn backwards on purpose
            Segment((0, 0), (10, 0), 0.0),
            Segment((10, 5), (20, 5), 0.0)]
    polys = join_segments(segs)
    assert len(polys) == 1 and not polys[0].closed
    xy = [(x, y) for x, y, _ in polys[0].vertices]
    assert xy in ([(0, 0), (10, 0), (10, 5), (20, 5)], [(20, 5), (10, 5), (10, 0), (0, 0)])
    # a disconnected piece gives a second chain; a branch point splits chains too
    assert len(join_segments(segs + [Segment((50, 50), (60, 50), 0.0)])) == 2
    assert len(join_segments(segs + [Segment((10, 0), (10, -5), 0.0)])) == 3


def _uniform_gap(mesh=2.0):
    """Uniform field 1 MV/m (1 kV / 1 mm) across a 100 x 10 mm slab: 10 kV over 10 mm."""
    p = Problem(regions=[Region(rect(0, 0, 100, 10), Material("m", 2.0))],
                lines=[DirichletLine(Polygon.from_points([(0, 0), (100, 0)], closed=False), 0.0),
                       DirichletLine(Polygon.from_points([(0, 10), (100, 10)], closed=False), 10e3)],
                symmetry="planar", mesh_size=mesh)
    return p, solve(p, build_mesh(p), "es")


def test_cumulative_voltage_equals_potential_difference():
    p, s = _uniform_gap()
    path = Polygon.from_points([(50, 0), (50, 10)], closed=False)           # along the field
    d = post.along_path(s, path, n=101)
    m = post.path_metrics(d)
    assert m["length"] == pytest.approx(0.01)
    assert m["dV"] == pytest.approx(10e3, rel=1e-9) and m["dV_nodal"] == pytest.approx(10e3, rel=1e-9)
    assert m["mean_abs"] == pytest.approx(1e6, rel=1e-9) and m["mean_dv"] == pytest.approx(1e6, rel=1e-9)
    # cumulative curves are monotone and end at the totals; V(0) - V(s) at each point
    assert np.all(np.diff(np.abs(m["cum_V"])) >= -1e-9)
    assert np.allclose(m["cum_V"], d["phi"][0] - d["phi"], rtol=1e-9, atol=1e-6)
    assert m["cum_abs"][-1] == pytest.approx(10e3, rel=1e-9)
    assert m["n_nan"] == 0


def test_cumulative_length_above_threshold():
    p, s = _uniform_gap()
    # path along the electrode direction crossing the field region: field is uniform 1 MV/m
    path = Polygon.from_points([(10, 5), (90, 5)], closed=False)           # perpendicular to the field
    d = post.along_path(s, path, n=81)
    assert np.allclose(np.abs(d["Et"]), 0.0, atol=1e-3)                     # purely normal
    along = post.path_metrics(d, threshold=0.5e6, component="E")           # total field 1 MV/m > 0.5
    assert along["frac_above"] == pytest.approx(1.0)
    assert along["len_above"] == pytest.approx(0.08)
    none = post.path_metrics(d, threshold=2e6, component="E")
    assert none["len_above"] == 0.0 and none["cum_len"][-1] == 0.0
    et = post.path_metrics(d, threshold=0.5e6, component="Et")             # tangential only: 0
    assert et["len_above"] == 0.0
    off = post.path_metrics(d)                                              # no threshold -> nothing flagged
    assert off["len_above"] == 0.0


def test_partial_threshold_exceedance_with_two_materials():
    """High-field half and low-field half on one path: the cumulative length counts only the high half."""
    air, oil = Material("air", 1.0), Material("oil", 4.0)
    p = Problem(regions=[Region(rect(0, 0, 100, 4), oil, "oil"), Region(rect(0, 4, 100, 10), air, "air")],
                lines=[DirichletLine(Polygon.from_points([(0, 0), (100, 0)], closed=False), 0.0),
                       DirichletLine(Polygon.from_points([(0, 10), (100, 10)], closed=False), 10e3)],
                mesh_size=2.0)
    s = solve(p, build_mesh(p), "es")
    path = Polygon.from_points([(50, 0.5), (50, 9.5)], closed=False)
    d = post.along_path(s, path, n=181)
    # E_oil : E_air = 1 : 4 ; V = E_oil*4mm + E_air*6mm -> E_oil = 10k/(4+24)mm, E_air = 4 x that
    e_oil = 10e3 / 28e-3
    thr = 2.0 * e_oil
    m = post.path_metrics(d, threshold=thr, component="Et")
    # path y = 0.5 .. 9.5 mm, interface at y = 4: oil 3.5 mm, air 5.5 mm (the high-field part)
    assert m["len_above"] == pytest.approx(0.0055, abs=1e-4)
    assert m["frac_above"] == pytest.approx(5.5 / 9.0, abs=0.01)
    assert m["max"] == pytest.approx(4 * e_oil, rel=1e-6)
    assert m["s_at_max"] > 0.0035


def test_ac_cumulative_is_complex_and_reversal_is_consistent():
    p = Problem(regions=[Region(rect(0, 0, 100, 10), Material("m", 2.0, 1e-9))],
                lines=[DirichletLine(Polygon.from_points([(0, 0), (100, 0)], closed=False), 0.0),
                       DirichletLine(Polygon.from_points([(0, 10), (100, 10)], closed=False), 1000.0)],
                mesh_size=3.0)
    s = solve(p, build_mesh(p), "ac", freq=50.0)
    up = Polygon.from_points([(50, 0), (50, 10)], closed=False)
    a = post.path_metrics(post.along_path(s, up, 51))
    b = post.path_metrics(post.along_path(s, up.reversed(), 51))
    assert np.iscomplexobj(a["cum_V"]) and a["dV"] == pytest.approx(1000.0, rel=1e-9)
    assert b["dV"] == pytest.approx(a["dV"], rel=1e-9)
    assert b["mean_abs"] == pytest.approx(a["mean_abs"], rel=1e-9)
    # reversing the path flips the sign of the (signed) cumulative voltage drop
    assert b["cum_V"][-1] == pytest.approx(-a["cum_V"][-1], rel=1e-9)


def test_points_on_a_curved_boundary_are_still_found_and_far_ones_are_not():
    """A fillet arc lies a hair outside the straight mesh edges: still inside the mesh.
    A point nowhere near the requested region gets NaN, never a stranger's value."""
    air, oil = Material("air", 1.0), Material("oil", 3.0)
    p = Problem(regions=[Region(rect(0, 0, 40, 40), air, "air"),
                         Region(Polygon([(10, 10, 0.0), (30, 10, 0.0), (30, 30, 0.0), (10, 30, 0.0)]), oil, "oil")],
                conductors=[Conductor(Polygon([(18, 18, math.tan(math.pi / 8)), (22, 22, 0.0), (18, 22, 0.0)]), 1000.0, "pin")],
                lines=[DirichletLine(Polygon.from_points([(0, 0), (40, 0)], closed=False), 0.0)],
                mesh_size=3.0, edge_size=1.0, edge_transition=4.0)
    s = solve(p, build_mesh(p), "es")
    arc = Polygon([(18, 18, math.tan(math.pi / 8)), (22, 22, 0.0)], closed=False)   # the curved side of the pin
    d = post.along_path(s, arc, n=60)                       # side "auto"
    assert not np.isnan(d["Et"]).any()
    assert np.allclose(np.abs(d["phi"]), 1000.0, rtol=2e-3)  # on the electrode: equipotential
    far = Polygon.from_points([(5, 35), (8, 35)], closed=False)                    # in the air, far from the oil
    e = post.along_path(s, far, n=10, region=1)
    assert np.isnan(e["Et"]).all() and np.isnan(e["phi"]).all()
    ok = post.along_path(s, far, n=10, region=0)
    assert not np.isnan(ok["Et"]).any()
