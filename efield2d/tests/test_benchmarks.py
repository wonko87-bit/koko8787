"""Pin the numbers handed out for cross-checking against Maxwell 2D."""
import os
import sys

import numpy as np
import pytest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "benchmarks"))
import bench  # noqa: E402

from efield2d import Assignment, import_dxf, build_problem, build_mesh, solve, post, Material  # noqa: E402


REF_C, REF_V1040, REF_ET = 1.408, 80.41, 2.298     # values quoted in benchmarks/BENCHMARKS.md


@pytest.mark.parametrize("mode", ["es", "dc", "ac"])
def test_model_a_matches_closed_form(mode):
    y0, E0, V0 = bench.a_analytic(mode)
    n = bench.a_numeric(mode, mesh=0.5)
    assert abs(n["Y"] / y0 - 1) < 5e-4
    for r in (7.5, 15.0):
        assert abs(n[f"V({r:g},10)"] / V0(r) - 1) < 2e-3
        assert abs(n[f"E({r:g},10)"] / E0(r) - 1) < 0.01      # |E| from central differences of V
    assert abs(n["V(10,10)"] / V0(10.0) - 1) < 2e-3


def test_model_a_layer_split_follows_eps_or_sigma():
    """ES / AC: field split by 1/eps_r ; DC: by 1/sigma - the point of having three modes."""
    es, dc = bench.a_numeric("es"), bench.a_numeric("dc")
    # E1/E2 at r = 7.5 vs 15, corrected for the 1/r of the cylinder: (E1*7.5)/(E2*15) = k2/k1
    ratio_es = (es["E(7.5,10)"] * 7.5) / (es["E(15,10)"] * 15)
    ratio_dc = (dc["E(7.5,10)"] * 7.5) / (dc["E(15,10)"] * 15)
    assert ratio_es == pytest.approx(4.0 / 2.0, rel=0.03)
    assert ratio_dc == pytest.approx(1e-13 / 1e-12, rel=0.03)


def test_model_b_is_mesh_converged_where_it_should_be():
    a = bench.b_numeric(2.0, 0.5, 5.0)
    b = bench.b_numeric(1.0, 0.25, 5.0)
    assert abs(a["C"] / b["C"] - 1) < 0.005
    for key in ("V(6,25)", "V(14,25)", "V(20,50)", "V(30,30)", "V(10,20)", "V(10,40)"):
        assert abs(a[key] / b[key] - 1) < 0.005
    assert abs(a["Et_mean(z5-40)"] / b["Et_mean(z5-40)"] - 1) < 0.01
    for key in ("E(6,25)", "E(14,25)", "E(20,50)", "E(30,30)", "E(40,70)"):
        assert abs(a[key] / b[key] - 1) < 0.02
    # reference values quoted in BENCHMARKS.md
    assert b["C"] * 1e12 == pytest.approx(REF_C, abs=0.005)
    assert b["V(10,40)"] / 1e3 == pytest.approx(REF_V1040, abs=0.2)
    assert b["Et_mean(z5-40)"] / 1e6 == pytest.approx(REF_ET, abs=0.01)


def test_generated_dxf_gives_the_same_answer_as_the_gui_model(tmp_path):
    bench.write_models(str(tmp_path))
    model = import_dxf(str(tmp_path / "modelA_coax2layer.dxf"))
    by_layer = {s.layer: s for s in model.shapes}
    assert set(by_layer) == {"LAYER1", "LAYER2", "INNER", "OUTER"}
    p = build_problem(model, {
        by_layer["LAYER1"].id: Assignment("region", Material("L1", 2.0, 1e-12), name="L1"),
        by_layer["LAYER2"].id: Assignment("region", Material("L2", 4.0, 1e-13), name="L2"),
        by_layer["INNER"].id: Assignment("line", voltage=1000.0, name="inner"),
        by_layer["OUTER"].id: Assignment("line", voltage=0.0, name="outer")},
        symmetry="axisym", mesh_size=0.5)
    s = solve(p, build_mesh(p), "es")
    assert post.capacitance(s, 1000.0) == pytest.approx(bench.a_analytic("es")[0], rel=5e-4)
    # model B dxf: four closed shapes, the electrode has its two fillet arcs
    mb = import_dxf(str(tmp_path / "modelB_spacer.dxf"))
    hv = [s for s in mb.shapes if s.layer == "HV"][0]
    assert sum(1 for v in hv.polygon.vertices if v[2] != 0.0) == 2 and len(mb.shapes) == 4
    # the .efp projects load and run
    from efield2d.app.document import Document
    d = Document.load(str(tmp_path / "modelB_es.efp"))
    assert d.run().sol.mode == "es"
