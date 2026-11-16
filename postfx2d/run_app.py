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
    return 0 if ok else 1


if __name__ == "__main__":
    if len(sys.argv) >= 4 and sys.argv[1] == "--selftest":
        sys.exit(selftest(sys.argv[2], sys.argv[3]))
    from postfx2d.gui import main
    main()
