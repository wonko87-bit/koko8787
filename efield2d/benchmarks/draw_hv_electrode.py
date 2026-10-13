"""Schematic of benchmark model B (axisymmetric): whole model + zoom on the HV electrode.
   python benchmarks/draw_hv_electrode.py  ->  benchmarks/modelB_hv_electrode.png"""
import math
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon as MplPoly, Rectangle

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from efield2d import Polygon  # noqa: E402
from efield2d.geometry import outline_xy  # noqa: E402

matplotlib.rcParams["font.sans-serif"] = ["Malgun Gothic", "AppleGothic", "NanumGothic", "Noto Sans CJK KR", "DejaVu Sans"]
matplotlib.rcParams["axes.unicode_minus"] = False
F = math.tan(math.pi / 8)
HV = Polygon([(0, 45, 0), (11, 45, F), (14, 48, 0), (14, 54, F), (11, 57, 0), (0, 57, 0)], True)
RED, BLUE, GRAY, GREEN = "#d6452a", "#2b6cb0", "#8a8f98", "#2f855a"

fig = plt.figure(figsize=(13.5, 7.6), layout="constrained")
gs = fig.add_gridspec(1, 2, width_ratios=[0.8, 1.5])

# ------------------------------------------------------------------ (a) whole model
a = fig.add_subplot(gs[0, 0])
a.add_patch(Rectangle((0, 0), 60, 80, fc="#f4f6f8", ec=GRAY, lw=1))
a.add_patch(Rectangle((0, 0), 60, 5, fc=GRAY, ec="k", lw=1))
a.add_patch(Rectangle((0, 5), 10, 40, fc="#cfe3f5", ec="k", lw=1))
a.add_patch(MplPoly(outline_xy(HV), fc=RED, ec="k", lw=1))
a.axvline(0, color="k", ls="-.", lw=0.8)
a.text(30, 2.5, "접지판 0 V  (r 0~60, z 0~5)", ha="center", va="center", color="w", fontsize=9)
a.text(5, 25, "스페이서\n에폭시\nεr 4\nr 0~10\nz 5~45", ha="center", va="center", fontsize=9)
a.text(40, 22, "공기 상자\nεr 1\nr 0~60, z 0~80\n(바깥 경계는 자연경계)", ha="center", va="center", fontsize=9)
a.annotate("HV 전극\n100 kV", (7, 51), (30, 62), ha="center", fontsize=10, color=RED, fontweight="bold",
           arrowprops=dict(arrowstyle="->", color=RED))
a.text(-1.2, 40, "회전축 r = 0", rotation=90, ha="right", va="center", fontsize=8)
a.set_xlim(-6, 62), a.set_ylim(-2, 82), a.set_aspect("equal")
a.set_xlabel("r [mm]"), a.set_ylabel("z [mm]")
a.set_title("(a) 모델 B 전체 단면 (축대칭, r–z 평면)", fontsize=11)
a.add_patch(Rectangle((-1, 36), 24, 24, fill=False, ec=BLUE, lw=1.4, ls="--"))
a.text(11, 61, "(b) 확대 범위", color=BLUE, fontsize=9, ha="center")

# ------------------------------------------------------------------ (b) zoom
b = fig.add_subplot(gs[0, 1])
b.add_patch(Rectangle((0, 36), 10, 9, fc="#cfe3f5", ec="k", lw=1.2))
b.text(5, 40.5, "스페이서\n(에폭시)", ha="center", va="center", fontsize=10)
xy = outline_xy(HV)
b.add_patch(MplPoly(xy, fc=RED, ec="k", lw=1.5, alpha=0.9))
b.add_patch(MplPoly([(-x, y) for x, y in xy], fc=RED, ec="k", lw=0.8, alpha=0.18, ls="--"))   # revolved half
b.axvline(0, color="k", ls="-.", lw=1)
b.text(-0.3, 61.6, "회전축 (r = 0)", ha="right", fontsize=8)
b.text(-7.5, 51, "← 축 반대쪽 (회전해서 생기는 면,\n  같은 모양)", fontsize=8, color=GRAY, ha="left", va="center")
b.text(7.5, 51, "HV 전극\n(회전체: 두꺼운 원판)\n100 000 V", ha="center", va="center", fontsize=11, color="w", fontweight="bold")

pts = {"A": (0, 45), "B": (11, 45), "C": (14, 48), "D": (14, 54), "E": (11, 57), "F": (0, 57)}
for k, (x, y) in pts.items():
    b.plot([x], [y], "o", ms=6, mfc="w", mec="k", zorder=6)
off = {"A": (-0.4, -1.0, "right"), "B": (-0.5, 1.1, "right"), "C": (0.5, -0.8, "left"), "D": (0.5, 0.4, "left"),
       "E": (0.0, 0.7, "center"), "F": (-0.4, 0.5, "right")}
for k, (x, y) in pts.items():
    dx, dy, ha = off[k]
    b.text(x + dx, y + dy, f"{k} ({x}, {y})", ha=ha, fontsize=9, fontweight="bold",
           bbox=dict(fc="w", ec="none", alpha=0.75, pad=1) if k == "B" else None)

# fillet centres and radii
for (cx, cy, a0, a1) in ((11, 48, -90, 0), (11, 54, 0, 90)):
    b.plot([cx], [cy], "+", color="k", ms=9, zorder=6)
    for ang in (a0, a1):
        b.plot([cx, cx + 3 * math.cos(math.radians(ang))], [cy, cy + 3 * math.sin(math.radians(ang))],
               color="k", lw=0.7, ls=":")
b.annotate("필렛 중심 (11, 48)\nR = 3 mm", (11, 48), (16.2, 45.2), fontsize=9, arrowprops=dict(arrowstyle="->"))
b.annotate("필렛 중심 (11, 54)\nR = 3 mm", (11, 54), (16.2, 57.4), fontsize=9, arrowprops=dict(arrowstyle="->"))

# dimensions
def dim(p, q, text, side=(0, 0), ha="center", va="center"):
    b.annotate("", p, q, arrowprops=dict(arrowstyle="<->", color=BLUE, lw=1.1))
    b.text((p[0] + q[0]) / 2 + side[0], (p[1] + q[1]) / 2 + side[1], text, color=BLUE, fontsize=9, ha=ha, va=va,
           bbox=dict(fc="w", ec="none", alpha=0.85, pad=1))
dim((0, 60.2), (14, 60.2), "외경 r = 14", side=(0, 0.1))
dim((0, 34.2), (11, 34.2), "전극 밑면은 평평 (r 0 → 11)", side=(0, 0.9))
dim((19.5, 45), (19.5, 57), "두께 12\n(z 45 → 57)", side=(1.8, 0))
dim((10, 43.4), (14, 43.4), "돌출 4", side=(1.6, -1.0))
b.plot([10, 10], [36, 45], color=GREEN, lw=3, solid_capstyle="butt", zorder=5)
b.annotate("연면 경로 r = 10\n(z 5 → 45, 공기 쪽 평가)", (10, 39), (12.2, 36.9), color=GREEN, fontsize=9,
           arrowprops=dict(arrowstyle="->", color=GREEN))
b.plot([10], [45], "s", ms=9, mfc="#ffd24d", mec="k", zorder=7)
b.annotate("삼중점 (10, 45)\n전극·스페이서·공기 — 전계 집중\n(Emax 비교 금지)", (10, 45), (-8.5, 41.2), fontsize=9,
           arrowprops=dict(arrowstyle="->"))
b.annotate("r 10 ~ 11 구간의 평평한 밑면은\n공기에 노출 (1 mm)", (10.6, 45.0), (15.0, 40.6), fontsize=9,
           arrowprops=dict(arrowstyle="->"))

b.set_xlim(-9, 23), b.set_ylim(33, 62.5), b.set_aspect("equal")
b.set_xlabel("r [mm]"), b.set_ylabel("z [mm]")
b.set_title("(b) HV 전극 확대 — A→B→호→C→D→호→E→F→A  (닫힌 윤곽 1개)", fontsize=11)
b.grid(alpha=0.25)

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "modelB_hv_electrode.png")
fig.savefig(out, dpi=130)
print(out)
