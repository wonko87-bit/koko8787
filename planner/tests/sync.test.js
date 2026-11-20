'use strict';
// 실행: node --test planner/tests
const test = require('node:test');
const assert = require('node:assert/strict');
const S = require('../src/sync.js');

const clone = (x) => JSON.parse(JSON.stringify(x));
const empty = () => ({ version: 1, settings: { weekStart: 0, theme: 'light', customHolidays: {} }, days: {}, weeks: {}, months: {} });

/** 한 기기 흉내: 데이터 + 마지막 도장 시점 */
function device(id, data) {
  const d = data || empty();
  d.settings.deviceId = id;
  S.migrate(d);
  return { id, data: d, prev: S.flatten(d) };
}
function edit(dev, now, fn) {
  fn(dev.data);
  const r = S.stamp(dev.data, dev.prev, dev.id, now);
  dev.prev = r.flat;
  return r.changed;
}
/** 앱의 syncNow와 같은 순서: 받기 → 합치기 → 반영 → 올리기 */
function sync(dev, drive, now) {
  const m = S.mergeCells(S.toCells(dev.data), drive.cells || {}, now);
  if (m.localChanged) {
    dev.data = S.fromCells(m.cells, dev.data);
    dev.prev = S.flatten(dev.data);
  }
  if (m.remoteChanged || !drive.cells) drive.cells = clone(m.cells);
  return m;
}
const view = (dev) => S.flatten(dev.data);

test('평탄화 ↔ 복원이 내용을 보존한다', () => {
  const d = empty();
  d.days['2026-11-19'] = [{ s: 'done', t: '09:00', x: '회의', c: '업무' }, null, { x: '보고서', t: '', s: '', c: '' }];
  d.weeks['2026-11-19'] = { goals: ['목표1', '', '목표3'], memo: '메모', next: [{ x: '넘길 일' }] };
  d.months['2026-11'] = { goals: ['월 목표'], todos: [null, { x: '할일', d: '2026-11-30' }] };
  d.settings.customHolidays = { '2027-03-15': '창립기념일' };
  d.settings.weekStart = 1;
  const f = S.flatten(d);
  assert.equal(f['d/2026-11-19/2'].x, '보고서');
  assert.equal(f['d/2026-11-19/1'], undefined);
  assert.deepEqual(S.flatten(S.unflatten(f, d)), f);
  assert.equal(S.unflatten(f, d).settings.theme, 'light', '기기별 설정은 유지');
});

test('옛 데이터는 모든 칸에 오래된 도장이 찍힌다', () => {
  const d = empty();
  d.days['2026-11-19'] = [{ x: 'a' }];
  assert.equal(S.migrate(d), true);
  assert.deepEqual(d.meta['d/2026-11-19/0'], [1, 'migrated']);
  assert.equal(S.migrate(d), false, '두 번 변환하지 않음');
});

test('서로 다른 날을 고치면 둘 다 살아남는다', () => {
  const drive = {};
  const A = device('A'), B = device('B');
  edit(A, 1000, (d) => { d.days['2026-11-17'] = [{ x: '화요일 일' }]; });
  edit(B, 1100, (d) => { d.days['2026-11-19'] = [{ x: '목요일 일' }]; });
  sync(A, drive, 2000); sync(B, drive, 2100); sync(A, drive, 2200);
  for (const dev of [A, B]) {
    assert.equal(view(dev)['d/2026-11-17/0'].x, '화요일 일');
    assert.equal(view(dev)['d/2026-11-19/0'].x, '목요일 일');
  }
  assert.deepEqual(view(A), view(B));
});

test('같은 칸을 양쪽에서 고치면 나중 수정이 이긴다 (순서와 무관)', () => {
  const A = device('A'), B = device('B');
  edit(A, 1000, (d) => { d.days['2026-11-19'] = [{ x: 'A 버전' }]; });
  edit(B, 1500, (d) => { d.days['2026-11-19'] = [{ x: 'B 버전' }]; });
  const ab = S.mergeCells(S.toCells(A.data), S.toCells(B.data), 2000).cells;
  const ba = S.mergeCells(S.toCells(B.data), S.toCells(A.data), 2000).cells;
  assert.deepEqual(ab, ba, '병합은 교환법칙을 만족');
  assert.equal(ab['d/2026-11-19/0'][2].x, 'B 버전');
});

test('지운 칸은 다른 기기의 옛 값으로 되살아나지 않는다', () => {
  const drive = {};
  const A = device('A'), B = device('B');
  edit(A, 1000, (d) => { d.days['2026-11-19'] = [{ x: '지울 일' }]; });
  sync(A, drive, 1100); sync(B, drive, 1200);
  assert.equal(view(B)['d/2026-11-19/0'].x, '지울 일');
  edit(A, 2000, (d) => { d.days['2026-11-19'][0] = null; });
  sync(A, drive, 2100); sync(B, drive, 2200); sync(A, drive, 2300);
  assert.equal(view(A)['d/2026-11-19/0'], undefined);
  assert.equal(view(B)['d/2026-11-19/0'], undefined);
});

test('지운 뒤 더 나중에 다시 쓰면 다시 쓴 값이 이긴다', () => {
  const drive = {};
  const A = device('A'), B = device('B');
  edit(A, 1000, (d) => { d.days['2026-11-19'] = [{ x: '처음' }]; });
  sync(A, drive, 1100); sync(B, drive, 1200);
  edit(A, 2000, (d) => { d.days['2026-11-19'][0] = null; });
  edit(B, 3000, (d) => { d.days['2026-11-19'][0] = { x: '다시 씀' }; });
  sync(A, drive, 3100); sync(B, drive, 3200); sync(A, drive, 3300);
  assert.equal(view(A)['d/2026-11-19/0'].x, '다시 씀');
  assert.deepEqual(view(A), view(B));
});

test('상태·구분 변경도 칸 단위로 합쳐진다 (이월 링크 포함)', () => {
  const drive = {};
  const A = device('A'), B = device('B');
  edit(A, 1000, (d) => { d.days['2026-11-19'] = [{ x: '보고서', s: '' }]; });
  sync(A, drive, 1100); sync(B, drive, 1200);
  edit(B, 2000, (d) => {
    d.days['2026-11-19'][0] = { x: '보고서', s: 'carry', link: { id: 'd:2026-11-20', i: 0, x: '보고서' } };
    d.days['2026-11-20'] = [{ x: '보고서', s: '', from: '11/19' }];
  });
  sync(B, drive, 2100); sync(A, drive, 2200);
  assert.equal(view(A)['d/2026-11-19/0'].s, 'carry');
  assert.equal(view(A)['d/2026-11-20/0'].from, '11/19');
});

test('여러 번 주고받아도 같은 상태로 수렴하고, 더 바뀔 것이 없으면 업로드하지 않는다', () => {
  const drive = {};
  const A = device('A'), B = device('B');
  edit(A, 1000, (d) => { d.weeks['2026-11-19'] = { goals: ['A 목표'], memo: 'A 메모' }; d.months['2026-11'] = { todos: [{ x: '월간 A' }] }; });
  edit(B, 1200, (d) => { d.weeks['2026-11-19'] = { goals: ['', 'B 목표'] }; d.settings.customHolidays = { '2027-03-15': '창립기념일' }; });
  sync(A, drive, 2000); sync(B, drive, 2100); sync(A, drive, 2200); sync(B, drive, 2300);
  assert.deepEqual(view(A), view(B));
  const w = A.data.weeks['2026-11-19'];
  assert.deepEqual(w.goals, ['A 목표', 'B 목표']);
  assert.equal(w.memo, 'A 메모');
  assert.equal(A.data.settings.customHolidays['2027-03-15'], '창립기념일');
  const again = sync(A, drive, 2400);
  assert.equal(again.localChanged, false);
  assert.equal(again.remoteChanged, false);
});

test('한 주 시작 요일은 동기화되고 테마는 기기별로 남는다', () => {
  const drive = {};
  const A = device('A'), B = device('B');
  B.data.settings.theme = 'dark';
  edit(A, 1000, (d) => { d.settings.weekStart = 1; d.settings.theme = 'light'; });
  sync(A, drive, 1100); sync(B, drive, 1200);
  assert.equal(B.data.settings.weekStart, 1);
  assert.equal(B.data.settings.theme, 'dark');
  assert.equal(B.data.settings.deviceId, 'B');
});

test('시계가 뒤로 가도 수정 도장은 항상 앞으로 간다', () => {
  const A = device('A');
  edit(A, 5000, (d) => { d.days['2026-11-19'] = [{ x: '1' }]; });
  edit(A, 3000, (d) => { d.days['2026-11-19'][0].x = '2'; });
  assert.equal(A.data.meta['d/2026-11-19/0'][0], 5001);
});

test('90일 지난 묘비는 정리된다', () => {
  const drive = {};
  const A = device('A');
  edit(A, 1000, (d) => { d.days['2026-11-19'] = [{ x: 'x' }]; });
  edit(A, 2000, (d) => { d.days['2026-11-19'][0] = null; });
  sync(A, drive, 3000);
  assert.ok(drive.cells['d/2026-11-19/0']);
  sync(A, drive, 2000 + S.TOMBSTONE_TTL + 1);
  assert.equal(drive.cells['d/2026-11-19/0'], undefined);
  assert.equal(A.data.meta['d/2026-11-19/0'], undefined);
});

test('옛 데이터끼리 처음 합칠 때도 양쪽 내용을 모두 가져온다', () => {
  const drive = {};
  const a = empty(); a.days['2026-11-18'] = [{ x: '회사 PC에서 적은 일' }];
  const b = empty(); b.days['2026-11-20'] = [{ x: '노트북에서 적은 일' }];
  const A = device('A', a), B = device('B', b);
  sync(A, drive, 1000); sync(B, drive, 1100); sync(A, drive, 1200);
  assert.equal(view(A)['d/2026-11-20/0'].x, '노트북에서 적은 일');
  assert.equal(view(B)['d/2026-11-18/0'].x, '회사 PC에서 적은 일');
});
