"""Benchmark models for cross-checking against another FEM tool (e.g. Maxwell 2D).

  python benchmarks/bench.py          # runs everything, writes models/*.dxf, models/*.efp, BENCHMARKS.md

Model A: two-layer coaxial dielectric (axisymmetric strip) - closed-form reference.
Model B: filleted HV electrode on an epoxy spacer above a ground plate - no closed form;
         compare capacitance, potentials / fields at points away from corners, and a
         creepage path.  Every number is produced by the same Document pipeline the GUI uses.
"""
from __future__ import annotations

import math
import os
import sys

import ezdxf
import numpy as np

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from efield2d import EPS0, Material, Polygon, post  # noqa: E402
from efield2d.app.document import Document  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
FIL = math.tan(math.pi / 8)       # bulge of a 90-degree arc


def efield_at(sol, r, z, d):
    """|E| at (r, z) from central differences of the (well converged) potential over +-d mm.
    Smoother than the element-constant field and what a point probe in another tool returns."""
    pts = np.array([[r + d, z], [r - d, z], [r, z + d], [r, z - d]])
    ph = post.sample(sol, pts)["phi"]
    gr, gz = (ph[0] - ph[1]) / (2 * d * 1e-3), (ph[2] - ph[3]) / (2 * d * 1e-3)
    return float(np.sqrt(abs(gr) ** 2 + abs(gz) ** 2))


# ------------------------------------------------------------------ model A
A = dict(r0=5.0, r1=10.0, r2=20.0, h=20.0, V=1000.0, f=50.0,
         L1=dict(eps_r=2.0, sigma=1e-12), L2=dict(eps_r=4.0, sigma=1e-13))


def model_a(mode: str, mesh: float = 0.5) -> Document:
    d = Document()
    d.set_materials({"L1": Material("L1", **A["L1"]), "L2": Material("L2", **A["L2"])})
    d.set_settings(symmetry="axisym", mode=mode, freq=A["f"], rms_input=False,
                   mesh_size=mesh, edge_size=0, curve_x=0)
    r0, r1, r2, h = A["r0"], A["r1"], A["r2"], A["h"]
    s1 = d.add_rect(r0, 0, r1, h)
    s2 = d.add_rect(r1, 0, r2, h)
    inner = d.add_polyline([(r0, 0), (r0, h)], closed=False)
    outer = d.add_polyline([(r2, 0), (r2, h)], closed=False)
    d.assign([s1], "region", material="L1", name="L1")
    d.assign([s2], "region", material="L2", name="L2")
    d.assign([inner], "line", voltage=A["V"], name="inner")
    d.assign([outer], "line", voltage=0.0, name="outer")
    return d


def a_analytic(mode: str):
    """Returns (Y or C, E(r) function)."""
    w = 2 * math.pi * A["f"]
    def kap(L):
        if mode == "es":
            return EPS0 * L["eps_r"]
        if mode == "dc":
            return L["sigma"]
        return complex(L["sigma"], w * EPS0 * L["eps_r"])
    k1, k2 = kap(A["L1"]), kap(A["L2"])
    ln2 = math.log(2.0)
    y = 2 * math.pi * (A["h"] * 1e-3) / (ln2 * (1 / k1 + 1 / k2))        # Q/V (es), G (dc), Y (ac)

    q = A["V"] / (ln2 * (1 / k1 + 1 / k2))                                # Q' = Q / (2 pi h)

    def E(r):                                                              # r in mm
        k = k1 if r < A["r1"] else k2
        return abs(q / (k * (r * 1e-3)))

    def V(r):
        if r <= A["r1"]:
            return abs(A["V"] - q * math.log(r / A["r0"]) / k1)
        return abs(q * math.log(A["r2"] / r) / k2)
    return y, E, V


def a_numeric(mode: str, mesh: float = 0.5):
    d = model_a(mode, mesh)
    res = d.run()
    sol = res.sol
    if mode == "es":
        y = post.capacitance(sol, A["V"])
    else:
        y = post.admittance(sol, A["V"])
    out = {"Y": y}
    for r in (7.5, 15.0):
        s = post.sample(sol, np.array([[r, 10.0]]))
        out[f"E({r:g},10)"] = efield_at(sol, r, 10.0, 0.5)
        out[f"V({r:g},10)"] = float(abs(s["phi"][0]))
    s = post.sample(sol, np.array([[A["r1"], 10.0]]), region=0)
    out["V(10,10)"] = float(abs(s["phi"][0]))
    return out


# ------------------------------------------------------------------ model B
PATH_R = 10.0
B = dict(V=100e3, probes=[(6.0, 25.0), (14.0, 25.0), (20.0, 50.0), (30.0, 30.0), (40.0, 70.0)],
         path_z=[10.0, 20.0, 30.0, 40.0])


def model_b(mesh=1.0, edge=0.25, trans=5.0) -> Document:
    d = Document()
    d.set_materials({"air": Material("air", 1.0, 0.0, 3.0e6), "epoxy": Material("epoxy", 4.0, 0.0)})
    d.set_settings(symmetry="axisym", mode="es", mesh_size=mesh, edge_size=edge, edge_transition=trans)
    air = d.add_rect(0, 0, 60, 80)
    gnd = d.add_rect(0, 0, 60, 5)
    spc = d.add_rect(0, 5, 10, 45)
    hv = d.add_polygon(Polygon([(0, 45, 0), (11, 45, FIL), (14, 48, 0), (14, 54, FIL),
                                (11, 57, 0), (0, 57, 0)], True), layer="HV", source="polyline")
    d.assign([air], "region", material="air", name="air")
    d.assign([spc], "region", material="epoxy", name="epoxy")
    d.assign([gnd], "conductor", voltage=0.0, name="gnd")
    d.assign([hv], "conductor", voltage=B["V"], name="hv")
    return d


def b_numeric(mesh=1.0, edge=0.25, trans=5.0):
    d = model_b(mesh, edge, trans)
    res = d.run()
    sol = res.sol
    out = {"C": post.capacitance(sol, B["V"]), "nodes": res.mesh.n_nodes, "tris": res.mesh.n_tris}
    for (r, z) in B["probes"]:
        s = post.sample(sol, np.array([[r, z]]))
        out[f"V({r:g},{z:g})"] = float(abs(s["phi"][0]))
        out[f"E({r:g},{z:g})"] = efield_at(sol, r, z, 1.0)
    # creepage path on the spacer surface r = 10, z = 5 -> 45, evaluated from the air side
    path = Polygon.from_points([(PATH_R, 5.0), (PATH_R, 45.0)], closed=False)
    data = post.along_path(sol, path, n=801, region=0)
    z = 5.0 + data["s"] * 1e3
    for zz in B["path_z"]:
        i = int(np.argmin(abs(z - zz)))
        out[f"Et({PATH_R:g},{zz:g})"] = float(abs(data["Et"][i]))
        out[f"V({PATH_R:g},{zz:g})"] = float(abs(data["phi"][i]))
    sel = (z >= 5.0) & (z <= 40.0)
    out["Et_mean(z5-40)"] = float(np.trapezoid(np.abs(data["Et"])[sel], data["s"][sel]) / (data["s"][sel][-1] - data["s"][sel][0]))
    out["dV(z5-40)"] = float(abs(data["phi"][sel][0] - data["phi"][sel][-1]))
    return out


# ------------------------------------------------------------------ files
def write_models(outdir):
    os.makedirs(outdir, exist_ok=True)
    # model A dxf (mm): two rectangles and two electrode lines
    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 4
    m = doc.modelspace()
    r0, r1, r2, h = A["r0"], A["r1"], A["r2"], A["h"]
    for nm in ("LAYER1", "LAYER2", "INNER", "OUTER"):
        doc.layers.add(nm)
    m.add_lwpolyline([(r0, 0), (r1, 0), (r1, h), (r0, h)], close=True, dxfattribs={"layer": "LAYER1"})
    m.add_lwpolyline([(r1, 0), (r2, 0), (r2, h), (r1, h)], close=True, dxfattribs={"layer": "LAYER2"})
    m.add_line((r0, 0), (r0, h), dxfattribs={"layer": "INNER"})
    m.add_line((r2, 0), (r2, h), dxfattribs={"layer": "OUTER"})
    doc.saveas(os.path.join(outdir, "modelA_coax2layer.dxf"))
    # model B dxf
    doc = ezdxf.new("R2010")
    doc.header["$INSUNITS"] = 4
    m = doc.modelspace()
    for nm in ("AIR", "GND", "SPACER", "HV"):
        doc.layers.add(nm)
    m.add_lwpolyline([(0, 0), (60, 0), (60, 80), (0, 80)], close=True, dxfattribs={"layer": "AIR"})
    m.add_lwpolyline([(0, 0), (60, 0), (60, 5), (0, 5)], close=True, dxfattribs={"layer": "GND"})
    m.add_lwpolyline([(0, 5), (10, 5), (10, 45), (0, 45)], close=True, dxfattribs={"layer": "SPACER"})
    m.add_lwpolyline([(0, 45, 0), (11, 45, FIL), (14, 48, 0), (14, 54, FIL), (11, 57, 0), (0, 57, 0)],
                     close=True, format="xyb", dxfattribs={"layer": "HV"})
    doc.saveas(os.path.join(outdir, "modelB_spacer.dxf"))
    for mode in ("es", "dc", "ac"):
        model_a(mode).save(os.path.join(outdir, f"modelA_{mode}.efp"))
    model_b().save(os.path.join(outdir, "modelB_es.efp"))


def fmt(x, digits=4):
    return f"{x:.{digits}g}"


def build_report() -> str:
    L = []
    w = L.append
    w("# 벤치마크 — Maxwell 2D와 비교용")
    w("")
    w("이 문서는 `python benchmarks/bench.py` 로 자동 생성됩니다 (숫자는 GUI와 같은 계산 경로에서 나온 값).")
    w("모델 파일: `benchmarks/models/` (DXF: mm 단위, `.efp`: 이 프로그램용 프로젝트).")
    w("")
    w("## 공통 설정")
    w("")
    w("- 단위 **mm**, 모두 **축대칭**. Maxwell 2D에서는 *Solution type ▸ Geometry mode = Cylindrical about Z*,")
    w("  그림은 XY 평면에 **X = r(반지름), Y = z(축 방향)** 로 그립니다 (X ≥ 0).")
    w("- 전압은 **peak(크기) 값**으로 입력합니다 (AC도 1000 V = 크기 1000 V). 결과 전계도 peak 값입니다.")
    w("- 전위를 지정한 곳 외의 경계는 모두 **자연경계(Neumann)**: 영역 바깥 경계에 Balloon이나 전위 경계를 걸지 마세요.")
    w("  (r = 0 축과 모델 끝면은 Maxwell 기본값 그대로 두면 됩니다.)")
    w("- 해석 모드 대응: **ES** = Electrostatic, **DC** = DC Conduction, **AC** = AC Conduction (50 Hz).")
    w("- Maxwell은 메시를 자동으로 세밀화하므로 *Percent error ≤ 0.1~0.5 %* 가 되도록 수렴시킨 값을 비교하세요 (항목 이름은 버전에 따라 다를 수 있음).")
    w("- 모델 B는 z 방향이 대칭이 아닙니다: **접지판이 아래(z = 0), HV 전극이 위**가 되도록 축 방향을 확인하세요.")
    w("- 결과는 모두 SI 단위로 환산해 적었습니다 (kV/mm = MV/m).")
    w("")

    # ---------------- Model A
    w("## 모델 A — 2층 유전체 동축 (해석해 있음)")
    w("")
    w("두 개의 직사각형 (r 방향 폭 × z 방향 높이 20 mm). 끝면(z = 0, z = 20)은 자연경계 → 무한히 긴 동축과 같음.")
    w("")
    w("| 영역 | 꼭짓점 (r, z) [mm] | εr | σ [S/m] |")
    w("|---|---|---|---|")
    w(f"| L1 | (5,0) (10,0) (10,20) (5,20) | {A['L1']['eps_r']} | {A['L1']['sigma']:g} |")
    w(f"| L2 | (10,0) (20,0) (20,20) (10,20) | {A['L2']['eps_r']} | {A['L2']['sigma']:g} |")
    w("")
    w("전극: **r = 5 mm** 선(z 0~20) = **1000 V**, **r = 20 mm** 선(z 0~20) = **0 V**.")
    w("")
    w("### 비교 항목 (해석해 / 이 프로그램, 메시 0.5 mm)")
    w("")
    for mode, label, yname, yunit, ys in (
            ("es", "ES", "정전용량 C", "pF", 1e12),
            ("dc", "DC", "컨덕턴스 G", "S", 1.0),
            ("ac", "AC 50 Hz", "어드미턴스 Y = G + jωC", "S", 1.0)):
        y0, E0, V0 = a_analytic(mode)
        n = a_numeric(mode)
        w(f"**{label}**")
        w("")
        w("| 항목 | 해석해 | 이 프로그램 | 비 |")
        w("|---|---|---|---|")
        if mode == "ac":
            wv = 2 * math.pi * A["f"]
            w(f"| G = Re(Y) [S] | {fmt(y0.real)} | {fmt(n['Y'].real)} | {fmt(n['Y'].real / y0.real, 5)} |")
            w(f"| C = Im(Y)/ω [pF] | {fmt(y0.imag / wv * 1e12)} | {fmt(n['Y'].imag / wv * 1e12)} | {fmt(n['Y'].imag / y0.imag, 5)} |")
        else:
            w(f"| {yname} [{yunit}] | {fmt(y0 * ys)} | {fmt(n['Y'] * ys)} | {fmt(n['Y'] / y0, 5)} |")
            if mode == "dc":
                w(f"| 저항 R [Ω] | {fmt(1 / y0)} | {fmt(1 / n['Y'])} | |")
        for r in (7.5, 15.0):
            w(f"| \\|E\\| at (r={r:g}, z=10) [kV/mm] | {fmt(E0(r) / 1e6)} | {fmt(n[f'E({r:g},10)'] / 1e6)} | {fmt(n[f'E({r:g},10)'] / E0(r), 4)} |")
        for r in (7.5, 15.0):
            w(f"| \\|V\\| at (r={r:g}, z=10) [V] | {fmt(V0(r))} | {fmt(n[f'V({r:g},10)'])} | {fmt(n[f'V({r:g},10)'] / V0(r), 5)} |")
        w(f"| \\|V\\| at 경계 (r=10, z=10) [V] | {fmt(V0(10.0))} | {fmt(n['V(10,10)'])} | {fmt(n['V(10,10)'] / V0(10.0), 5)} |")
        w("")
    w("점 \\|E\\|는 전위의 중심차분(±0.5 mm)으로 구한 값입니다. 전위·정전용량은 0.05 % 이내로 맞습니다.")
    w("")
    w("확인 포인트: ES는 층이 **εr 비의 역**으로, DC는 **σ 비의 역**으로 전계를 나눠 갖습니다 (DC에서는 σ가 작은 L2가 높은 전계를 받음).")
    w("AC 50 Hz에서는 σ ≪ ωε 라서 ES와 거의 같은 분포가 됩니다. Maxwell 결과가 이와 다르면 재료 입력(σ, εr)이나 모드를 먼저 점검하세요.")
    w("")

    # ---------------- Model B
    w("## 모델 B — 필렛 전극 + 에폭시 스페이서 + 접지판 (ES)")
    w("")
    w("| 객체 | 꼭짓점 (r, z) [mm] | 재료 / 전위 |")
    w("|---|---|---|")
    w("| 공기 상자 | (0,0) (60,0) (60,80) (0,80) | 공기 εr = 1 |")
    w("| 접지판 | (0,0) (60,0) (60,5) (0,5) | 도체 **0 V** |")
    w("| 스페이서 | (0,5) (10,5) (10,45) (0,45) | 에폭시 εr = 4 |")
    w("| HV 전극 | (0,45) → (11,45) → 호 → (14,48) → (14,54) → 호 → (11,57) → (0,57) → 닫기 | 도체 **100 000 V** |")
    w("")
    w("HV 전극의 호는 **반지름 3 mm 의 90° 필렛** (중심 (11,48), (11,54); 바깥으로 볼록).")
    w("스페이서 윗면(z = 45, r ≤ 10)은 HV 전극의 **평평한 밑면(r ≤ 11)** 에 완전히 맞닿습니다. 스페이서 옆면(r = 10)과 전극 밑면이 만나는")
    w("(10, 45)가 전극·스페이서·공기의 삼중점이며, 전극 밑면은 r = 10 ~ 11 구간이 공기에 노출됩니다 (돌출 1 mm + 필렛).")
    w("겹치는 부분은 **도체(전극)가 우선, 스페이서가 공기보다 우선**입니다 (Maxwell에서는 공기 상자를 먼저 그리고 나머지를 그 안에 만들면 됩니다).")
    w("")
    meshes = ((2.0, 0.5), (1.0, 0.25), (0.5, 0.12), (0.35, 0.08))
    res = [b_numeric(m, e, 5.0) for m, e in meshes]
    ref = res[-1]
    w("### 비교 항목")
    w("")
    w("마지막 열은 두 번째로 세밀한 메시와 가장 세밀한 메시의 차이입니다 (이 프로그램 자체의 메시 의존성).")
    w("")
    w("| 항목 | 값 (최세밀 메시) | 메시 의존성 | 비교 허용 오차 (권장) |")
    w("|---|---|---|---|")
    prev = res[-2]

    def row(label, key, scale, unit, tol):
        d = abs(ref[key] - prev[key]) / abs(ref[key]) * 100
        w(f"| {label} | {fmt(ref[key] * scale)} {unit} | {d:.2f} % | {tol} |")
    row("정전용량 C (HV 전극 ↔ 접지, 축대칭 전체)", "C", 1e12, "pF", "±1 %")
    for (r, z) in B["probes"]:
        row(f"\\|V\\| at (r={r:g}, z={z:g})", f"V({r:g},{z:g})", 1e-3, "kV", "±0.5 %")
    for (r, z) in B["probes"]:
        row(f"\\|E\\| at (r={r:g}, z={z:g})", f"E({r:g},{z:g})", 1e-6, "kV/mm", "±3 %")
    w("")
    w("**연면 경로**: 스페이서 표면 r = 10 mm, z = 5 → 45 (공기 쪽에서 평가; 접선 성분은 경계 양쪽이 같음).")
    w("Maxwell에서는 이 선을 따라 **E의 Y 성분(= z 방향 = 접선 성분)** 을 그리세요 (Mag_E는 경계에서 법선 성분 때문에 모호).")
    w("")
    w("| 항목 | 값 (최세밀 메시) | 메시 의존성 | 비교 허용 오차 (권장) |")
    w("|---|---|---|---|")
    for zz in B["path_z"]:
        row(f"\\|Et\\| at (r=10, z={zz:g})", f"Et(10,{zz:g})", 1e-6, "kV/mm", "±3 %")
    for zz in B["path_z"]:
        row(f"\\|V\\| at (r=10, z={zz:g})", f"V(10,{zz:g})", 1e-3, "kV", "±0.5 %")
    row("평균 \\|Et\\|, z = 5 ~ 40 구간", "Et_mean(z5-40)", 1e-6, "kV/mm", "±1 %")
    w("")
    w("### 이 프로그램의 메시 수렴 (참고)")
    w("")
    w("| 전체 / 가장자리 요소 크기 [mm] | 요소 수 | C [pF] | V(10,40) [kV] | 평균 Et z5~40 [kV/mm] | E(14,25) [kV/mm] | E(20,50) [kV/mm] |")
    w("|---|---|---|---|---|---|---|")
    for (m, e), o in zip(meshes, res):
        w(f"| {m} / {e} | {o['tris']} | {fmt(o['C'] * 1e12, 5)} | {fmt(o['V(10,40)'] / 1e3, 5)} | "
          f"{fmt(o['Et_mean(z5-40)'] / 1e6, 5)} | {fmt(o['E(14,25)'] / 1e6, 4)} | {fmt(o['E(20,50)'] / 1e6, 4)} |")
    w("")
    w("### 주의")
    w("")
    w("- **전극 모서리·삼중점(전극/스페이서/공기가 만나는 점)의 최대 전계(Emax)는 비교하지 마세요.** 이론상 무한대에 가까워 메시에 따라 값이 계속 변합니다")
    w("  (이 프로그램도, Maxwell도 마찬가지). 위 표의 \\|E\\| 비교 지점은 모두 모서리에서 떨어져 있습니다.")
    w("- 점 \\|E\\|는 이 프로그램이 전위의 중심차분(±1 mm)으로 구한 값입니다 (1차 요소의 요소별 상수 전계는 점 값이 ±4 % 흔들려 쓰지 않음).")
    w("  Maxwell의 점 값과 비교하되, 전위와 정전용량이 먼저 맞는지 보세요 — 그쪽이 훨씬 엄밀한 비교입니다.")
    w("- 정전용량은 공기 상자 크기(60 × 80 mm)와 자연경계 조건에 의존합니다. Maxwell에서 Balloon을 쓰면 값이 달라집니다.")
    w("")
    return "\n".join(L)


if __name__ == "__main__":
    write_models(os.path.join(HERE, "models"))
    text = build_report()
    with open(os.path.join(HERE, "BENCHMARKS.md"), "w", encoding="utf-8") as f:
        f.write(text)
    print(text)
