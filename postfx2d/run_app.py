"""PostFX 2D 실행 진입점 (PyInstaller 빌드 대상).

  PostFX2D.exe                     GUI 실행
  PostFX2D.exe --selftest IN OUT   GUI 없이 처리 (빌드 검증용, 실패 시 종료코드 1)
"""
import sys


def selftest(inp, out):
    from postfx2d import core
    res = core.process(core.imread(inp))
    core.imwrite(out, res.image)
    core.save_csv(res.rows, out.rsplit(".", 1)[0] + "_summary.csv")
    glob = sum(r["하이라이트"] == "글로벌" for r in res.rows)
    vals = sum(p.value is not None for p in res.points)
    idx = sum(p.index is not None for p in res.points)
    ok = len(res.groups) == 6 and glob == 1 and vals == idx == len(res.points)
    report = (f"groups={len(res.groups)} global={glob} read={vals}/{len(res.points)} index={idx} "
              f"warnings={res.warnings}\n" + "\n".join(map(str, res.rows)))
    if sys.stdout:                     # 창 모드 exe 에서는 stdout 이 없음
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            print(report)
        except Exception:
            pass
    with open(out.rsplit(".", 1)[0] + "_report.txt", "w", encoding="utf-8") as f:
        f.write(report)
    return 0 if (ok and field_selftest(inp) and paths_selftest(inp)) else 1


def field_selftest(inp):
    """모드 2: 합성 rainbow 맵을 알려진 정답과 비교 (평균 오차 < 1% of range)."""
    import os
    from postfx2d import core, fieldmap
    path = os.path.join(os.path.dirname(os.path.abspath(inp)), "field_sample.png")
    if not os.path.exists(path):
        return True
    img = core.imread(path)
    fm = fieldmap.FieldMap(fieldmap.builtin_lut("Rainbow"), 0, 5000)
    # (x, y, 정답) - tools/make_field_sample.py 기준
    cases = [(273, 240, 4344.7), (537, 350, 3046.3), (405, 262, 1183.3)]
    err = max(abs(fm.read(img, x, y).value - t) for x, y, t in cases)
    return err < 50


def paths_selftest(inp):
    """모드 3: 경로 3장 취합 → 각 경로 검출 + 배경에 빨간 선 잔상 없음."""
    import os
    from postfx2d import core, paths
    d = os.path.join(os.path.dirname(os.path.abspath(inp)), "path_samples")
    if not os.path.isdir(d):
        return True
    ps = [paths.make_path(f"p{i}", core.imread(os.path.join(d, f"path{i}.png"))) for i in (1, 2, 3)]
    bg = paths.restore_background(ps)
    out = paths.merge(ps)
    ok = all(p.pixels > 500 for p in ps) and paths.red_mask(bg).sum() == 0 and out.shape == bg.shape
    eqp = os.path.join(d, "equipotential.png")
    if ok and os.path.exists(eqp):            # 등전위선 겹치기: 선이 검출되고 결과가 달라져야 함
        eq = core.imread(eqp)
        ok = (paths.equipotential_alpha(eq, bg) > 0.5).sum() > 5000 and \
            not (paths.merge(ps, eq_img=eq) == out).all()
    return bool(ok)


if __name__ == "__main__":
    if len(sys.argv) >= 4 and sys.argv[1] == "--selftest":
        sys.exit(selftest(sys.argv[2], sys.argv[3]))
    from postfx2d.gui import main
    main()
