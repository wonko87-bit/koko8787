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


def a_expected():
    """Closed-form reference numbers of model A (what a converged FEM tool should give)."""
    V = A["V"]
    w = 2 * math.pi * A["f"]
    out = {}
    y, E, Vr = a_analytic("es")
    out["es"] = dict(C=y, Q=y * V, W=0.5 * y * V * V, W1=(2 / 3) * 0.5 * y * V * V, W2=(1 / 3) * 0.5 * y * V * V,
                     E75=E(7.5), E15=E(15.0), V75=Vr(7.5), V10=Vr(10.0), V15=Vr(15.0))
    y, E, Vr = a_analytic("dc")
    out["dc"] = dict(G=y, R=1 / y, I=y * V, P=y * V * V, E75=E(7.5), E15=E(15.0), V75=Vr(7.5), V10=Vr(10.0), V15=Vr(15.0))
    y, E, Vr = a_analytic("ac")
    out["ac"] = dict(G=y.real, C=y.imag / w, Y=abs(y), I=abs(y) * V, P=0.5 * y.real * V * V,
                     We=0.25 * (y.imag / w) * V * V, E75=E(7.5), E15=E(15.0), V75=Vr(7.5), V10=Vr(10.0), V15=Vr(15.0))
    return out


def write_profiles(outdir):
    """E(r) and V(r) along z = 10 mm (r = 5 .. 20 mm, step 0.25) - analytic and this program."""
    import csv
    from efield2d.calc import FieldCalculator
    rad = Polygon.from_points([(A["r0"], 10.0), (A["r2"], 10.0)], closed=False)
    for mode in ("es", "dc", "ac"):
        c = FieldCalculator(model_a(mode, 0.25).run().sol)
        n = int(round((A["r2"] - A["r0"]) / 0.25)) + 1
        pe = c.profile("E", rad, n=n)
        pv = c.profile("abs(V)", rad, n=n)
        _, E0, V0 = a_analytic(mode)
        with open(os.path.join(outdir, f"modelA_profile_{mode}.csv"), "w", newline="", encoding="utf-8-sig") as f:
            wr = csv.writer(f)
            wr.writerow(["r_mm", "E_analytic_kV_per_mm", "E_this_program_kV_per_mm", "V_analytic_V", "V_this_program_V"])
            for i in range(n):
                r = pe["xy"][i, 0]
                wr.writerow([f"{r:.4f}", f"{E0(r) / 1e6:.6g}", f"{abs(pe['values'][i]) / 1e6:.6g}",
                             f"{V0(r):.6g}", f"{pv['values'][i].real:.6g}"])


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
    write_profiles(outdir)


def build_maxwell_guide() -> str:
    e = a_expected()
    es, dc, ac = e["es"], e["dc"], e["ac"]
    g = lambda x, d=4: f"{x:.{d}g}"
    L = []
    w = L.append
    w("# Maxwell 2D 에서 모델 A 결과 확인하기")
    w("")
    w("이 문서는 `python benchmarks/bench.py` 로 자동 생성됩니다 (기대값은 **해석해**).")
    w("")
    w("> **주의** — 메뉴 이름·버튼 위치는 Maxwell 2D / Electronics Desktop 버전에 따라 조금 다릅니다. 아래는 일반적인 흐름과 의도입니다.")
    w("> 화면과 다르면 **기대값 표의 숫자가 맞는지**로 판정하세요 (절차보다 숫자가 기준). 막히는 단계는 화면을 보내 주시면 맞춰 드립니다.")
    w("")
    w("## 0. 기대값 (해석해) — Maxwell 결과가 이 값에 수렴해야 합니다")
    w("")
    w("모델: 축대칭, r 5~10 mm 층 L1 (εr 2, σ 1e-12), r 10~20 mm 층 L2 (εr 4, σ 1e-13), 높이 20 mm, 전극 r=5 → 1000 V, r=20 → 0 V (자세한 형상은 `BENCHMARKS.md`).")
    w("")
    w("| 확인 항목 | ES (정전계) | DC 전도 | AC 전도 (50 Hz) |")
    w("|---|---|---|---|")
    w(f"| 매트릭스 (Parameters ▸ Matrix) | C = **{g(es['C'] * 1e12, 5)} pF** | G = **{g(dc['G'])} S** (R = {g(dc['R'])} Ω) | G = **{g(ac['G'])} S**, C = **{g(ac['C'] * 1e12, 5)} pF** (\\|Y\\| = {g(ac['Y'])} S) |")
    w(f"| 전극 전하 / 전류 | Q = **{g(es['Q'] * 1e9, 5)} nC** | I = **{g(dc['I'])} A** | \\|I\\| = **{g(ac['I'])} A** |")
    w(f"| 에너지 / 손실 | 저장 에너지 W = **{g(es['W'] * 1e6, 5)} µJ** (L1 {g(es['W1'] * 1e6, 5)}, L2 {g(es['W2'] * 1e6, 5)}) | 손실 P = **{g(dc['P'])} W** | 평균 손실 ½GV² = **{g(ac['P'])} W**, 평균 저장 에너지 ¼CV² = **{g(ac['We'] * 1e6, 5)} µJ** |")
    w(f"| \\|E\\| at r=7.5 mm, z=10 | **{g(es['E75'] / 1e6)} kV/mm** | **{g(dc['E75'] / 1e6)} kV/mm** | **{g(ac['E75'] / 1e6)} kV/mm** (peak) |")
    w(f"| \\|E\\| at r=15 mm, z=10 | **{g(es['E15'] / 1e6)} kV/mm** | **{g(dc['E15'] / 1e6)} kV/mm** | **{g(ac['E15'] / 1e6)} kV/mm** (peak) |")
    w(f"| \\|V\\| at r=7.5 / 10 / 15 mm | {g(es['V75'])} / {g(es['V10'])} / {g(es['V15'])} V | {g(dc['V75'])} / {g(dc['V10'])} / {g(dc['V15'])} V | {g(ac['V75'])} / {g(ac['V10'])} / {g(ac['V15'])} V |")
    w("| ∫E·dl (r 5→20 mm) | 1000 V | 1000 V | 1000 V |")
    w("")
    w("층 사이 전계 분배: ES·AC는 **L1(εr 작음)** 이 크고, DC는 **L2(σ 작음)** 가 큽니다 — 모드를 바꿨는데 분포가 안 바뀌면 재료 입력을 먼저 의심하세요.")
    w("")
    w("메시 수렴된 이 프로그램 값, 그리고 r = 5 ~ 20 mm 를 0.25 mm 간격으로 적은 E(r), V(r) 표는 `models/modelA_profile_{es,dc,ac}.csv` 에 있습니다 (엑셀에서 Maxwell 내보내기 데이터와 겹쳐 비교).")
    w("")
    w("## 1. 모델 만들기 (ES / DC / AC 공통)")
    w("")
    w("1. Maxwell 2D 디자인 추가. **Maxwell 2D ▸ Solution Type** — *Electric* 아래에서 `Electrostatic`, `DC Conduction`, `AC Conduction` 중 하나, **Geometry mode = Cylindrical about Z** (축대칭).")
    w("2. **Modeler ▸ Units = mm**. 그림은 XY 평면에 그립니다 — 이 모드에서 **X = r(반지름), Y = z(축 방향)**.")
    w("3. **Draw ▸ Rectangle** 두 개: L1 = (X 5, Y 0) 크기 (5, 20), L2 = (X 10, Y 0) 크기 (10, 20). 이름은 L1, L2.")
    w("4. 재료: **Assign Material ▸ Add Material** 로 새 재료를 만들고 *Relative Permittivity*, *Bulk Conductivity* 를 표의 값으로 입력.")
    w("   L1: εr 2 / σ 1e-12 S/m,  L2: εr 4 / σ 1e-13 S/m. (ES는 σ 무시, DC는 εr 무시, AC는 둘 다 사용.) 영역 바깥의 배경(Region)은 만들지 마세요 — 모델은 L1, L2 두 개가 전부입니다.")
    w("5. 전극: **Edit ▸ Select ▸ Edges** 로 L1 의 안쪽 변(r = 5)을 고르고 **Excitations ▸ Assign ▸ Voltage** = **1000 V**, 바깥 변(L2의 r = 20)에 **0 V**.")
    w("   AC Conduction 은 크기 1000 V(**peak**), 위상 0° 로 넣고 Setup 에서 주파수 **50 Hz**.")
    w("6. 다른 경계(위·아래 끝면, r 축)는 **아무것도 지정하지 않습니다** (자연경계 Neumann). Balloon 금지.")
    w("7. **Analysis ▸ Add Solution Setup** (*Percent Error* 0.1 % 정도, 패스 충분히) ▸ **Validate** ▸ **Analyze**.")
    w("")
    w("## 2. 계산기 없이 바로 확인 (먼저 이것부터)")
    w("")
    w("- **정전용량 / 컨덕턴스**: **Maxwell 2D ▸ Parameters ▸ Assign ▸ Matrix** 에서 두 Voltage 선택 후 Analyze, 이어서 결과 창(*Solution Data*)의 **Matrix** 탭. "
      "ES 는 Capacitance, DC 는 Conductance, AC 는 Conductance 와 Capacitance 가 나옵니다. → 위 표의 C, G 와 비교.")
    w("- **에너지**: **Results ▸ Solution Data ▸ Convergence** 탭에 *Energy* (정전계: 저장 에너지) 가 나옵니다. → W = ½CV² 와 비교. "
      "이 값이 위 표와 같으면 축대칭 전체(360°) 모델로 계산되고 있는 것입니다 (절반이면 대칭 승수/모드 확인).")
    w("- **전계·전위 분포**: **Fields ▸ Plot Fields ▸ Mag_E / Voltage** 로 영역을 골라 그려 보고, 층 경계에서 전계가 점프하는지 확인.")
    w("")
    w("## 3. Field Calculator 로 확인")
    w("")
    w("**Maxwell 2D ▸ Fields ▸ Calculator** 를 엽니다. 계산기는 **스택(RPN) 방식**입니다 — 값을 쌓고(`Input` 에서 Quantity, Geometry, Number…) 연산 버튼을 누르면 위의 값들을 소비합니다. "
      "마지막에 **Eval** 로 숫자를 보고, 필요하면 **Add** 로 이름을 붙여 보관합니다.")
    w("")
    w("### 3-A. 반경 방향 선 위의 E(r), V(r) 그래프 → 해석해와 비교")
    w("1. 확인용 선을 만듭니다: **Draw ▸ Line** 으로 **(5, 10) → (20, 10)** (z = 10 mm, r 5→20) 폴리라인. 속성에서 **Model 체크 해제** 권장 (형상 해석에 포함되지 않게).")
    w("2. **Results ▸ Create Fields Report ▸ Rectangular Plot**. *Geometry* = 그 Polyline, *Category* = Calculator Expressions(또는 Fields), *Quantity* = **Mag_E**, "
      "다시 한 번 *Quantity* = **Voltage** 로 추가. X축은 선을 따른 거리 (Distance).")
    w("3. 그래프 위에서 우클릭 ▸ **Export Data** 로 CSV 저장 → 엑셀에서 `models/modelA_profile_*.csv` 의 `E_analytic` 열과 겹쳐 그립니다. "
      "거리 d 와 r 의 관계는 r = 5 + d (mm).")
    w("4. 합격 기준: 전계 **±1 %**, 전위 **±0.2 %** (경계 r = 10 에서 전계가 점프하는 것도 확인).")
    w("")
    w("### 3-B. ∫E·dl = 인가 전압 (전위차) 확인")
    w("1. 스택에 **Quantity ▸ E**, **Geometry ▸ Line ▸ (그 Polyline)** 을 올립니다.")
    w("2. 선의 접선 단위벡터를 올리고 (**Unit Vec** 메뉴의 *Tangent* — 버전에 따라 위치가 다름) **Dot** (E·t), 이어서 **∫** (적분) 후 **Eval**.")
    w("3. 기대값 **1000 V** (ES, DC, AC 모두 크기 기준). 부호는 선의 방향에 따라 ±.")
    w("   *Tangent 를 못 찾으면*: E 의 X 성분만 쓰면 됩니다 (이 선은 r 방향) — **Quantity ▸ E ▸ Scal? ▸ ScalarX**, **Geometry ▸ Line**, **∫**, **Eval**.")
    w("")
    w("### 3-C. 층별 에너지 ½∫D·E dV (ES)")
    w("1. **Quantity ▸ D**, **Quantity ▸ E**, **Dot**, **Number ▸ Scalar 0.5**, **×** (곱).")
    w("2. **Geometry ▸ Volume ▸ L1**, **∫**, **Eval** → 기대 L1 = {0} µJ.".format(g(es['W1'] * 1e6, 5)))
    w("3. 같은 방법으로 L2 → 기대 {0} µJ. 합 = {1} µJ = ½CV² (Convergence 탭의 Energy 와도 같아야 함).".format(g(es['W2'] * 1e6, 5), g(es['W'] * 1e6, 5)))
    w("   *값이 이 값의 2π 분의 1* 이면 Volume 적분이 회전 방향을 포함하지 않은 것 (모드/대칭 승수 확인), *½ 배* 이면 에너지 정의(½ 누락)부터 확인.")
    w("")
    w("### 3-D. 손실 ∫J·E dV (DC / AC)")
    w("1. **Quantity ▸ J**, **Quantity ▸ E**, **Dot**, (AC 는 복소 → **Real** 또는 크기 연산 주의) **Geometry ▸ Volume ▸ L1**, **∫**, **Eval**; L2 도 같은 방법으로 두 값을 더합니다.")
    w("2. 기대: DC **{0} W** (= G·V²),  AC 평균 손실 **{1} W** (= ½GV², peak 기준 — Maxwell 이 peak/RMS 중 무엇으로 보여 주는지 확인해서 √2², 즉 2배 차이를 감안).".format(g(dc['P']), g(ac['P'])))
    w("   (σ 가 작아 값이 매우 작습니다 — 표시 자릿수를 늘리세요.)")
    w("")
    w("### 3-E. (선택) Gauss 법칙으로 전하 — 축대칭에서는 까다로움")
    w("전극 전하는 3-절의 **Matrix 의 C·V** 가 가장 쉽고 정확합니다 (기대 {0} nC). 계산기로 하려면 반지름 r 의 원통면을 지나는 D 의 법선 성분을 적분해야 하며, ".format(g(es['Q'] * 1e9, 5)) +
      "2D 축대칭에서는 선 적분에 **2π·r** 가중을 직접 넣어야 합니다: **D·Normal·2π·X** 를 선 (r = 7.5 mm, z 0→20) 위에서 적분. "
      "되면 기대값은 위 Q 와 같고 (두 층 어디서 잡아도 같음 — D 의 법선 성분은 연속), 이 방식이 막히면 건너뛰어도 됩니다.")
    w("")
    w("## 4. 비교가 안 맞을 때 점검 순서")
    w("")
    w("1. **C 가 크게 다르다** → 지오메트리 모드(축대칭이 아니라 Cartesian 이면 '깊이 1 m당' 값으로 다름), 단위(mm), 재료 εr 입력, Balloon 사용 여부.")
    w("2. **C 는 맞는데 전계 분포가 다르다** → 선/점의 위치 (r = 5 + d), 경계가 아니라 층 내부에서 읽었는지, 수렴 에러(Percent Error)가 큰지.")
    w("3. **DC 만 다르다** → 재료 σ 입력 (Bulk Conductivity), 해석 타입이 DC Conduction 인지 (ES 로 풀면 εr 분배가 나옴).")
    w("4. **AC 크기가 √2 배 / ½ 배** → peak 와 RMS 해석 차이: 우리는 **peak** (1000 V = 크기), Maxwell 도 입력한 값을 크기로 쓰는지, 결과 표시가 RMS 인지 확인.")
    w("5. 에너지·손실 적분이 2π 또는 ½ 어긋나면 위 3-C 의 주석.")
    w("6. **위 순서로도 안 풀리면**: 확인 항목별 Maxwell 값, 해석 설정(Solution type, Percent Error, 패스 수)을 알려 주세요. 우리 쪽 코드는 해석해와 일치함이 테스트로 고정되어 있어, 보통 설정 차이입니다.")
    w("")
    w("## 5. 알려 주실 때")
    w("")
    w("```")
    w("모드: ES / DC / AC")
    w("Matrix:    C = ...  또는 G = ...")
    w("에너지(Convergence 탭):  ... J        (ES)")
    w("E at r=7.5 / 15 (z=10):  ... / ... kV/mm")
    w("∫E·dl (r 5→20):  ... V")
    w("층별 에너지 L1 / L2:  ... / ... µJ")
    w("Percent Error, 패스 수, 요소 수:  ...")
    w("```")
    w("")
    return "\n".join(L)


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
    with open(os.path.join(HERE, "MAXWELL_GUIDE.md"), "w", encoding="utf-8") as f:
        f.write(build_maxwell_guide())
    print(text)
