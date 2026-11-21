'use strict';
/* Planner 동기화 엔진 — 칸(셀) 단위 Last-Writer-Wins 병합
 *
 * 플래너 데이터를 잘게 나눈 "칸"마다 [수정시각(ms), 기기ID] 도장을 붙여 두고,
 * 두 기기의 데이터를 합칠 때 칸마다 더 최근 도장을 고른다.
 * 지운 칸은 값 null(묘비)로 남겨 다른 기기의 옛 값이 되살아나지 않게 한다.
 *
 * 칸 경로
 *   d/<날짜>/<줄>           하루 할일 한 줄
 *   w/<목요일>/g/<n>        주 목표
 *   w/<목요일>/memo         주간 메모
 *   w/<목요일>/n/<줄>       다음 주로 넘길 일
 *   m/<YYYY-MM>/g/<n>       월 목표
 *   m/<YYYY-MM>/t/<줄>      월간 할일
 *   s/weekStart             한 주 시작 요일
 *   s/h/<날짜>              사용자 휴일
 * 테마·화면 배율·기기ID는 기기별 설정이라 동기화하지 않는다.
 */
(function (root) {
  const LOCAL_SETTINGS = ['theme', 'scale', 'deviceId'];
  const ITEM_KEYS = ['s', 't', 'x', 'c', 'd', 'from', 'link'];
  const TOMBSTONE_TTL = 90 * 24 * 3600 * 1000;
  const MIGRATED = [1, 'migrated'];

  const isBlank = (v) => v == null || String(v).trim() === '';

  /** 할일 한 줄을 비교 가능한 값으로 (빈 줄은 null, 키 순서 고정) */
  function itemVal(it) {
    if (!it || (isBlank(it.x) && isBlank(it.t))) return null;
    const o = {};
    for (const k of ITEM_KEYS) if (it[k] != null && it[k] !== '') o[k] = it[k];
    return o;
  }

  /** 플래너 데이터 → { 경로: 값 } (빈 칸은 빠짐) */
  function flatten(data) {
    const f = {};
    const put = (p, v) => { if (v != null && v !== '') f[p] = v; };
    for (const [k, arr] of Object.entries(data.days || {})) (arr || []).forEach((it, i) => put(`d/${k}/${i}`, itemVal(it)));
    for (const [k, w] of Object.entries(data.weeks || {})) {
      if (!w) continue;
      (w.goals || []).forEach((g, i) => put(`w/${k}/g/${i}`, isBlank(g) ? null : g));
      put(`w/${k}/memo`, isBlank(w.memo) ? null : w.memo);
      (w.next || []).forEach((it, i) => put(`w/${k}/n/${i}`, itemVal(it)));
    }
    for (const [k, m] of Object.entries(data.months || {})) {
      if (!m) continue;
      (m.goals || []).forEach((g, i) => put(`m/${k}/g/${i}`, isBlank(g) ? null : g));
      (m.todos || []).forEach((it, i) => put(`m/${k}/t/${i}`, itemVal(it)));
    }
    const s = data.settings || {};
    put('s/weekStart', s.weekStart == null ? 0 : s.weekStart);
    for (const [k, v] of Object.entries(s.customHolidays || {})) put(`s/h/${k}`, isBlank(v) ? null : v);
    return f;
  }

  function setIdx(arr, i, v, fill) {
    while (arr.length < i) arr.push(fill);
    arr[i] = v;
  }

  /** { 경로: 값 } → 플래너 데이터. 기기별 설정은 base에서 가져온다 */
  function unflatten(flat, base) {
    const f = JSON.parse(JSON.stringify(flat));
    const bs = (base && base.settings) || {};
    const d = { version: 2, settings: { weekStart: 0, customHolidays: {} }, days: {}, weeks: {}, months: {} };
    for (const k of LOCAL_SETTINGS) if (bs[k] !== undefined) d.settings[k] = bs[k];
    for (const [p, v] of Object.entries(f)) {
      if (v == null) continue;
      const a = p.split('/');
      if (a[0] === 'd') setIdx(d.days[a[1]] || (d.days[a[1]] = []), +a[2], v, null);
      else if (a[0] === 'w') {
        const w = d.weeks[a[1]] || (d.weeks[a[1]] = {});
        if (a[2] === 'g') setIdx(w.goals || (w.goals = []), +a[3], v, '');
        else if (a[2] === 'memo') w.memo = v;
        else if (a[2] === 'n') setIdx(w.next || (w.next = []), +a[3], v, null);
      } else if (a[0] === 'm') {
        const m = d.months[a[1]] || (d.months[a[1]] = {});
        if (a[2] === 'g') setIdx(m.goals || (m.goals = []), +a[3], v, '');
        else if (a[2] === 't') setIdx(m.todos || (m.todos = []), +a[3], v, null);
      } else if (a[0] === 's') {
        if (a[1] === 'weekStart') d.settings.weekStart = v;
        else if (a[1] === 'h') d.settings.customHolidays[a[2]] = v;
      }
    }
    return d;
  }

  const same = (a, b) => JSON.stringify(a === undefined ? null : a) === JSON.stringify(b === undefined ? null : b);

  /** 도장이 없는 옛 데이터: 모든 칸에 "아주 오래된" 도장을 찍어 둔다 (어떤 실제 수정도 이것보다 최신) */
  function migrate(data) {
    if (data.meta) return false;
    data.meta = {};
    for (const p of Object.keys(flatten(data))) data.meta[p] = MIGRATED.slice();
    data.version = 2;
    return true;
  }

  /**
   * 지난번 도장 이후 바뀐 칸에 [now, deviceId] 도장을 찍는다.
   * prevFlat: 지난번 flatten 결과. 반환: { flat, changed: [경로] }
   */
  function stamp(data, prevFlat, deviceId, now) {
    const meta = data.meta || (data.meta = {});
    const cur = flatten(data);
    const changed = [];
    for (const p of new Set([...Object.keys(cur), ...Object.keys(prevFlat || {})])) {
      if (same(cur[p], prevFlat[p])) continue;
      const prev = meta[p];
      meta[p] = [Math.max(now, prev ? prev[0] + 1 : 0), deviceId];
      changed.push(p);
    }
    return { flat: cur, changed };
  }

  /** 로컬 데이터 → { 경로: [시각, 기기, 값|null] } (묘비 포함) */
  function toCells(data) {
    const f = flatten(data);
    const m = data.meta || {};
    const c = {};
    for (const p of new Set([...Object.keys(m), ...Object.keys(f)])) {
      const st = m[p] || MIGRATED;
      c[p] = [st[0], st[1], f[p] === undefined ? null : f[p]];
    }
    return c;
  }

  function newer(a, b) {
    if (a[0] !== b[0]) return a[0] > b[0];
    if (a[1] !== b[1]) return String(a[1]) > String(b[1]);
    return JSON.stringify(a[2]) > JSON.stringify(b[2]); // 같은 도장인데 값이 다르면 결정적으로 하나 고름
  }
  const sameCell = (a, b) => !!a && !!b && a[0] === b[0] && a[1] === b[1] && same(a[2], b[2]);

  /**
   * 칸 단위 병합.
   * 반환: { cells, localChanged, remoteChanged }
   *  - localChanged: 로컬에 반영할 것이 있음
   *  - remoteChanged: 드라이브에 올릴 것이 있음
   */
  function mergeCells(local, remote, now) {
    const cells = {};
    let localChanged = false;
    let remoteChanged = false;
    for (const p of new Set([...Object.keys(local || {}), ...Object.keys(remote || {})])) {
      const l = local[p];
      const r = remote[p];
      const pick = !r ? l : !l ? r : newer(r, l) ? r : l;
      // 오래된 묘비 정리
      if (pick[2] === null && now - pick[0] > TOMBSTONE_TTL) {
        if (l) localChanged = true;
        if (r) remoteChanged = true;
        continue;
      }
      cells[p] = pick;
      if (!sameCell(pick, l)) localChanged = true;
      if (!sameCell(pick, r)) remoteChanged = true;
    }
    return { cells, localChanged, remoteChanged };
  }

  /** 병합 결과 칸들 → 새 플래너 데이터 (meta 포함) */
  function fromCells(cells, base) {
    const flat = {};
    const meta = {};
    for (const [p, c] of Object.entries(cells)) {
      meta[p] = [c[0], c[1]];
      if (c[2] != null) flat[p] = c[2];
    }
    const d = unflatten(flat, base);
    d.meta = meta;
    return d;
  }

  function newDeviceId() {
    const rnd = Math.random().toString(36).slice(2, 8);
    return `dev-${Date.now().toString(36)}-${rnd}`;
  }

  const api = { flatten, unflatten, migrate, stamp, toCells, mergeCells, fromCells, newDeviceId, itemVal, TOMBSTONE_TTL };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.PlannerSync = api;
})(typeof window !== 'undefined' ? window : globalThis);
