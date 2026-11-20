import base64, pathlib, re
src = pathlib.Path(__file__).resolve().parent.parent / "src"
css = (src/"style.css").read_text(encoding="utf-8")
font = base64.b64encode((src/"fonts/PretendardVariable.woff2").read_bytes()).decode()
css = css.replace("url('fonts/PretendardVariable.woff2')", f"url(data:font/woff2;base64,{font})")
html = (src/"index.html").read_text(encoding="utf-8")
seed = r"""
window.PLANNER_PREVIEW = true;
window.PLANNER_SEED = function (D) {
  const t = new Date(); const base = new Date(t.getFullYear(), t.getMonth(), t.getDate());
  const k = (n) => { const d = new Date(base.getFullYear(), base.getMonth(), base.getDate() + n); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
  D.days[k(-2)] = [{ s: 'done', t: '09:00', x: '주간 업무회의', c: '업무' }, { s: 'done', t: '14:00', x: '설계 검토 회의', c: '업무' }, { s: '', x: '출장 정산', c: '업무' }];
  D.days[k(-1)] = [{ s: 'done', t: '10:00', x: '해석 결과 정리', c: '업무' }, { s: 'carry', x: '보고서 초안 작성', c: '업무', link: { id: 'd:' + k(0), i: 1, x: '보고서 초안 작성' } }, { s: '', t: '19:30', x: '아이 학원 픽업', c: '가족' }];
  D.days[k(0)] = [{ s: 'done', t: '08:30', x: '메일 확인', c: '업무' }, { s: '', x: '보고서 초안 작성', c: '업무', from: '어제' }, { s: '', t: '13:30', x: '팀 커피 주문', c: '업무' }, { s: 'cancel', t: '18:00', x: '헬스장', c: '개인' }, { s: '', t: '19:00', x: '저녁 약속', c: '약속' }];
  D.days[k(1)] = [{ s: '', t: '10:00', x: '과제 진도 점검', c: '업무' }];
  D.days[k(2)] = [{ s: '', t: '11:00', x: '가족 나들이', c: '가족' }];
};
"""
html = html.replace('<link rel="stylesheet" href="style.css" />', f"<style>\n{css}\n</style>")
html = html.replace('<script src="holidays.js"></script>', "<script>\n" + (src/"holidays.js").read_text(encoding="utf-8") + "\n" + seed + "\n</script>")
html = html.replace('<script src="sync.js"></script>', "<script>\n" + (src/"sync.js").read_text(encoding="utf-8") + "\n</script>")
html = html.replace('<script src="app.js"></script>', "<script>\n" + (src/"app.js").read_text(encoding="utf-8") + "\n</script>")
html = html.replace("<title>Planner</title>", "<title>Planner 미리보기</title>")
out = pathlib.Path(__file__).resolve().parent.parent / "preview" / "planner-preview.html"
out.parent.mkdir(exist_ok=True)
out.write_text(html, encoding="utf-8")
print(out, out.stat().st_size)
