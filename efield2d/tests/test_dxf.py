import math

import ezdxf
import numpy as np
import pytest

from efield2d import (Assignment, Material, build_mesh, build_problem, import_dxf, post,
                      solve, EPS0)


def new_doc(insunits=4):
    doc = ezdxf.new("R2010", setup=True)
    doc.header["$INSUNITS"] = insunits
    return doc, doc.modelspace()


def save(doc, tmp_path, name="t.dxf"):
    p = tmp_path / name
    doc.saveas(p)
    return str(p)


def by_source(model, src):
    return [s for s in model.shapes if s.source == src]


def test_closed_polyline_circle_and_orientation(tmp_path):
    doc, msp = new_doc()
    msp.add_lwpolyline([(0, 0), (0, 10), (20, 10), (20, 0)], close=True)  # clockwise on purpose
    msp.add_circle((50, 5), 4)
    m = import_dxf(save(doc, tmp_path))
    rect_s = by_source(m, "LWPOLYLINE")[0]
    circ = by_source(m, "CIRCLE")[0]
    assert rect_s.closed and rect_s.area == pytest.approx(200.0)       # normalised to CCW
    assert circ.closed and circ.area == pytest.approx(math.pi * 16.0, rel=1e-12)
    assert circ.bbox() == pytest.approx((46, 1, 54, 9))


@pytest.mark.parametrize("ins,scale", [(4, 1.0), (6, 1000.0), (5, 10.0), (1, 25.4)])
def test_units_converted_to_mm(tmp_path, ins, scale):
    doc, msp = new_doc(ins)
    msp.add_lwpolyline([(0, 0), (1, 0), (1, 1), (0, 1)], close=True)
    m = import_dxf(save(doc, tmp_path))
    assert m.shapes[0].area == pytest.approx(scale ** 2)
    assert m.unit_scale == pytest.approx(scale)


def test_unitless_assumes_mm_and_warns_and_override(tmp_path):
    doc, msp = new_doc(0)
    msp.add_lwpolyline([(0, 0), (2, 0), (2, 2), (0, 2)], close=True)
    f = save(doc, tmp_path)
    m = import_dxf(f)
    assert m.shapes[0].area == pytest.approx(4.0) and any("unitless" in w for w in m.warnings)
    assert import_dxf(f, unit_scale=1000.0).shapes[0].area == pytest.approx(4e6)


def test_loose_lines_and_arcs_join_into_closed_loop(tmp_path):
    """D-shape drawn as separate LINE + ARC entities, listed out of order."""
    doc, msp = new_doc()
    msp.add_arc((0, 0), 10, 270, 90)                 # right half circle
    msp.add_line((0, -10), (-15, -10))
    msp.add_line((-15, 10), (0, 10))
    msp.add_line((-15, -10), (-15, 10))
    m = import_dxf(save(doc, tmp_path))
    chains = by_source(m, "chain")
    assert len(chains) == 1 and chains[0].closed
    assert chains[0].area == pytest.approx(15 * 20 + math.pi * 100 / 2, rel=1e-9)


def test_junction_splits_chains_and_gap_is_reported(tmp_path):
    doc, msp = new_doc()
    msp.add_line((0, 0), (10, 0))
    msp.add_line((10, 0), (20, 0))
    msp.add_line((10, 0), (10, 10))        # T junction at (10, 0)
    msp.add_line((30, 0), (30, 10))        # an almost-closed square: gap of 0.2 mm
    msp.add_line((30, 10), (40, 10))
    msp.add_line((40, 10), (40, 0))
    msp.add_line((40, 0), (30.2, 0))
    m = import_dxf(save(doc, tmp_path))
    chains = by_source(m, "chain")
    assert len(chains) == 4 and not any(c.closed for c in chains)
    assert any("gap" in w for w in m.warnings)
    # snapping closes it
    m2 = import_dxf(save(doc, tmp_path), snap=0.5)
    assert any(c.closed for c in by_source(m2, "chain"))


def test_open_spline_is_flattened_within_tolerance(tmp_path):
    doc, msp = new_doc()
    xs = np.linspace(0, 40, 9)
    ys = 0.02 * xs ** 2 / 2                       # points on a parabola
    msp.add_spline([(x, y, 0) for x, y in zip(xs, ys)])
    f = save(doc, tmp_path)
    fine = import_dxf(f, flatten_tol=0.002)
    coarse = import_dxf(f, flatten_tol=0.5)
    sf = fine.shapes[0]
    assert not sf.closed
    assert (len(sf.polygon.vertices) > 3 * len(coarse.shapes[0].polygon.vertices))
    pts = np.array([(x, y) for x, y, _ in sf.polygon.vertices])
    assert pts[0] == pytest.approx([0, 0], abs=1e-9) and pts[-1] == pytest.approx([40, 16], abs=1e-9)
    # reference: the spline evaluated directly, spacing << tolerance
    from scipy.spatial import cKDTree
    bs = msp.query("SPLINE").first.construction_tool()
    ref = np.array([(v.x, v.y) for v in bs.points(np.linspace(0, bs.max_t, 400000))])
    assert np.max(np.hypot(*np.diff(ref, axis=0).T)) < 2e-4
    tree = cKDTree(ref)
    assert tree.query(pts)[0].max() < 1e-4                  # vertices are on the curve
    for res, tol in ((fine, 0.002), (coarse, 0.5)):
        q = np.array([(x, y) for x, y, _ in res.shapes[0].polygon.vertices])
        mid = 0.5 * (q[:-1] + q[1:])
        assert tree.query(mid)[0].max() <= tol              # chord never strays more than tol


def test_ellipse_becomes_closed_shape_with_right_area(tmp_path):
    doc, msp = new_doc()
    msp.add_ellipse((0, 0), major_axis=(20, 0), ratio=0.5)         # a = 20, b = 10
    m = import_dxf(save(doc, tmp_path), flatten_tol=0.002)
    s = m.shapes[0]
    assert s.closed and s.source == "ELLIPSE"
    assert s.area == pytest.approx(math.pi * 20 * 10, rel=1e-3)     # ezdxf approximates ellipses by cubic Beziers


def test_blocks_are_exploded_and_unsupported_reported(tmp_path):
    doc, msp = new_doc()
    blk = doc.blocks.new("PLATE")
    blk.add_lwpolyline([(0, 0), (5, 0), (5, 5), (0, 5)], close=True)
    msp.add_blockref("PLATE", (100, 0))
    msp.add_text("hello")
    m = import_dxf(save(doc, tmp_path))
    assert len(m.shapes) == 1 and m.shapes[0].bbox() == pytest.approx((100, 0, 105, 5))
    assert any("TEXT" in w for w in m.warnings)


def test_layer_filter(tmp_path):
    doc, msp = new_doc()
    doc.layers.add("ELEC")
    doc.layers.add("DIM")
    msp.add_circle((0, 0), 3, dxfattribs={"layer": "ELEC"})
    msp.add_circle((10, 0), 3, dxfattribs={"layer": "DIM"})
    f = save(doc, tmp_path)
    assert len(import_dxf(f).shapes) == 2
    only = import_dxf(f, layers=["ELEC"])
    assert [s.layer for s in only.shapes] == ["ELEC"]


def test_build_problem_orders_regions_and_validates(tmp_path):
    doc, msp = new_doc()
    msp.add_lwpolyline([(0, 0), (6, 0), (6, 6), (0, 6)], close=True)      # small, drawn first
    msp.add_lwpolyline([(-5, -5), (50, -5), (50, 50), (-5, 50)], close=True)
    msp.add_line((0, 0), (6, 0))
    m = import_dxf(save(doc, tmp_path))
    small, big = sorted(m.shapes[:2], key=lambda s: s.area)
    line_shape = m.shapes[-1]
    air, oil = Material("air", 1.0), Material("oil", 2.2)
    p = build_problem(m, {small.id: Assignment("region", oil, name="oil"),
                          big.id: Assignment("region", air, name="air"),
                          line_shape.id: Assignment("line", voltage=5.0)})
    assert [r.name for r in p.regions] == ["air", "oil"]      # big first
    with pytest.raises(ValueError):                           # an open line cannot be a conductor
        build_problem(m, {line_shape.id: Assignment("conductor", voltage=1.0)})
    with pytest.raises(ValueError):                           # axisym needs x >= 0
        build_problem(m, {big.id: Assignment("region", air)}, symmetry="axisym")


def test_dxf_to_solution_coax_from_circles(tmp_path):
    """End to end: planar coax drawn as two circles in a DXF, loose layers,
    solved and compared with E = V / (r ln(b/a))."""
    a, b, V = 5.0, 20.0, 1000.0
    doc, msp = new_doc()
    doc.layers.add("OUTER")
    doc.layers.add("INNER")
    msp.add_circle((0, 0), b, dxfattribs={"layer": "OUTER"})
    msp.add_circle((0, 0), a, dxfattribs={"layer": "INNER"})
    m = import_dxf(save(doc, tmp_path))
    outer = [s for s in m.shapes if s.layer == "OUTER"][0]
    inner = [s for s in m.shapes if s.layer == "INNER"][0]
    eps = Material("oil", 2.2)
    p = build_problem(m, {outer.id: Assignment("region", eps, name="oil"),
                          inner.id: Assignment("conductor", voltage=V, name="core")},
                      mesh_size=1.0, edge_size=0.5, edge_transition=3.0)
    # outer ring at 0 V: a closed Dirichlet line on the outer circle
    from efield2d import DirichletLine
    p.lines.append(DirichletLine(outer.polygon, 0.0, "shield"))
    mesh = build_mesh(p)
    s = solve(p, mesh, "es")
    C = post.capacitance(s, V)
    assert C == pytest.approx(2 * math.pi * EPS0 * 2.2 / math.log(b / a), rel=0.005)
    r = np.hypot(mesh.nodes[:, 0], mesh.nodes[:, 1]) * 1e3
    phi_exact = V * np.log(b / r) / np.log(b / a)
    assert np.max(np.abs(s.phi - phi_exact)) < 0.003 * V
