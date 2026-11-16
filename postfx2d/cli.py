"""명령줄 실행: python cli.py 입력.png [출력.png] [--groups N]"""
import argparse
import os

from postfx2d import core

ap = argparse.ArgumentParser(description="최소안전율부위 후처리")
ap.add_argument("input")
ap.add_argument("output", nargs="?")
ap.add_argument("--groups", type=int, default=0, help="그룹 수 (0=자동)")
a = ap.parse_args()
out = a.output or os.path.splitext(a.input)[0] + "_minSF.png"
res = core.process(core.imread(a.input), core.Settings(group_count=a.groups))
core.imwrite(out, res.image)
core.save_csv(res.rows, os.path.splitext(out)[0] + "_summary.csv")
for r in res.rows:
    print(r)
for w in res.warnings:
    print("경고:", w)
print("저장:", out)
