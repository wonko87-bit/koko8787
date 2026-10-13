import math

import ezdxf
import pytest

from efield2d import EPS0, Material, Polygon, post
from efield2d.app.document import Document, DocumentError


def plate_doc(mode="es"):
    """Parallel plates drawn as bodies: 100 x 10 mm dielectric between two 100 x 3 mm electrodes."""
    d = Document()
    d.set_settings(symmetry="planar", mesh_size=3.0, mode=mode)
    air = d.add_rect(0, 3, 100, 13)
    lo = d.add_rect(0, 0, 100, 3)
    hi = d.add_rect(0, 13, 100, 16)
    d.assign([air], "region", material="oil", name="oil")
    d.assign([lo], "conductor", voltage=0.0, name="gnd")
    d.assign([hi], "conductor", voltage=1000.0, name="hv")
    return d, (air, lo, hi)


def test_run_matches_parallel_plate_formula():
    d, _ = plate_doc()
    res = d.run()
    C = post.capacitance(res.sol, 1000.0)
    assert C == pytest.approx(EPS0 * 2.2 * 0.1 / 0.01, rel=1e-9)


def test_undo_redo_cover_edits_and_assignments():
    d, (air, lo, hi) = plate_doc()
    n = len(d.shapes)
    d.delete([hi])
    assert len(d.shapes) == n - 1
    d.undo()
    assert len(d.shapes) == n and d.shape(hi).kind == "conductor"
    d.undo()                                      # undo the assignment of hi
    assert d.shape(hi).kind == "ignore"
    d.redo()
    assert d.shape(hi).kind == "conductor"
    v = list(d.shape(lo).polygon.vertices)
    d.set_vertices(lo, [(x + 5, y, b) for x, y, b in v])
    assert d.shape(lo).polygon.bbox()[0] == 5
    d.undo()
    assert d.shape(lo).polygon.bbox()[0] == 0


def test_assignment_rules_and_validation_messages():
    d = Document()
    line = d.add_polyline([(0, 0), (10, 0)], closed=False)
    with pytest.raises(DocumentError, match="열린 선"):
        d.assign([line], "region", material="air")
    d.assign([line], "line", voltage=5.0)         # lines are fine
    with pytest.raises(DocumentError, match="영역"):
        d.to_problem()
    r = d.add_rect(0, 0, 10, 10)
    d.assign([r], "region", material="air")
    d.assign([line], "ignore")
    with pytest.raises(DocumentError, match="전극"):
        d.to_problem()
    d.assign([line], "line", voltage=0.0)
    d.set_settings(mode="dc")
    with pytest.raises(DocumentError, match="도전율"):
        d.to_problem()
    d.set_settings(mode="es", symmetry="axisym")
    d.translate([r], -3, 0)
    with pytest.raises(DocumentError, match="x<0"):
        d.to_problem()
    with pytest.raises(DocumentError):
        d.add_rect(0, 0, 0, 5)
    with pytest.raises(DocumentError):
        d.add_circle(0, 0, -1)


def test_ac_rms_input_is_scaled_to_peak():
    d, (_, _, hi) = plate_doc("ac")
    prob, _ = d.to_problem()
    assert max(c.voltage for c in prob.conductors) == pytest.approx(1000.0 * math.sqrt(2))
    d.set_settings(rms_input=False)
    prob, _ = d.to_problem()
    assert max(c.voltage for c in prob.conductors) == pytest.approx(1000.0)
    d.set_settings(mode="es", rms_input=True)     # only AC is affected
    prob, _ = d.to_problem()
    assert max(c.voltage for c in prob.conductors) == pytest.approx(1000.0)


def test_regions_are_laid_down_big_first():
    d = Document()
    small = d.add_rect(2, 2, 4, 4)
    big = d.add_rect(0, 0, 10, 10)
    d.assign([small], "region", material="epoxy", name="small")
    d.assign([big], "region", material="air", name="big")
    d.assign([d.add_rect(0, 0, 10, 1)], "conductor")
    prob, ids = d.to_problem()
    assert [r.name for r in prob.regions] == ["big", "small"] and ids == [big, small]


def test_project_roundtrip_keeps_everything(tmp_path):
    d, ids = plate_doc("ac")
    d.set_materials({**d.materials, "gas": Material("gas", 1.0, 0.0, [(0, 9e6), (10, 3e6)])})
    d.add_circle(50, 50, 4)
    d.set_settings(freq=60.0, curve_x=7.5, edge_size=0.5)
    path = str(tmp_path / "p.efp")
    d.save(path)
    assert not d.dirty
    e = Document.load(path)
    assert e.settings == d.settings
    assert e.materials["gas"].e_allow == [(0, 9e6), (10, 3e6)]
    assert [(s.id, s.kind, s.material, s.voltage, s.name, s.polygon.vertices) for s in e.shapes] == \
           [(s.id, s.kind, s.material, s.voltage, s.name, s.polygon.vertices) for s in d.shapes]
    assert e.signature() == d.signature()                                # same solver input
    assert e.add_rect(0, 0, 1, 1) == max(s.id for s in d.shapes) + 1      # ids keep counting


def test_deleting_a_material_unassigns_regions():
    d, (air, _, _) = plate_doc()
    d.set_materials({k: v for k, v in d.materials.items() if k != "oil"})
    assert d.shape(air).kind == "ignore"


def test_signature_ignores_unassigned_shapes_but_not_real_changes():
    d, (air, lo, hi) = plate_doc()
    sig = d.signature()
    d.add_polyline([(0, 0), (1, 1)], closed=False)      # scratch line, ignored by the solver
    d.set_settings(curve_x=3.0)
    assert d.signature() == sig
    d.translate([hi], 0, 1)
    assert d.signature() != sig


def test_load_dxf_then_assign(tmp_path):
    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 4
    doc.modelspace().add_circle((20, 20), 5)
    f = str(tmp_path / "c.dxf")
    doc.saveas(f)
    d = Document()
    m = d.load_dxf(f)
    assert len(m.shapes) == 1 and d.shapes[0].area == pytest.approx(math.pi * 25)
    d.load_dxf(f, replace=False)
    assert len(d.shapes) == 2 and d.shapes[0].id != d.shapes[1].id
