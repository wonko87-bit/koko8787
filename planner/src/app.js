'use strict';
/* Planner — 주간·월간·연간 할일/일정 플래너
 * 데이터는 하나의 JSON 객체(state.data)이며
 *  - Tauri 앱: 문서\Planner\planner-data.json 에 저장 (Rust 쪽 save_data)
 *  - 브라우저 미리보기: localStorage 에 저장
 */

const TAURI = window.__TAURI__ || null;
const invoke = TAURI ? TAURI.core.invoke : null;

const DOW = ['일', '월', '화', '수', '목', '금', '토'];
const CATS = ['업무', '개인', '가족', '약속'];
const STATUS_CYCLE = ['', 'done', 'carry', 'cancel'];
const STATUS_LABEL = { '': '', done: '✔ 완료', carry: '→ 이월', cancel: '✕ 취소' };
const LEN = { d: 11, n: 11, m: 60 };
const WEEK_GOALS = 4;
const MONTH_GOALS = 6;
const TODO_ROW_H = 30;
const LS_KEY = 'planner-data-v1';

const ICONS = {
  done: '<svg class="i-done" viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>',
  carry: '<svg class="i-carry" viewBox="0 0 24 24"><path d="M5 12h13M13 6l6 6-6 6"/></svg>',
  cancel: '<svg class="i-cancel" viewBox="0 0 24 24"><path d="M7 7l10 10M17 7L7 17"/></svg>',
  plus: '<svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>',
  close: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6L6 18"/></svg>',
  left: '<svg viewBox="0 0 24 24"><path d="M15 6l-6 6 6 6"/></svg>',
  right: '<svg viewBox="0 0 24 24"><path d="M9 6l6 6-6 6"/></svg>',
};

/* ───────────────────────── utilities ───────────────────────── */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
function h(tag, attrs, ...kids) {
  const el = document.createElement(tag);
  if (attrs) {
    for (const [k, v] of Object.entries(attrs)) {
      if (v == null || v === false) continue;
      if (k === 'class') el.className = v;
      else if (k === 'html') el.innerHTML = v;
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else if (k === 'dataset') Object.assign(el.dataset, v);
      else el.setAttribute(k, v === true ? '' : v);
    }
  }
  for (const k of kids.flat()) {
    if (k == null || k === false) continue;
    el.append(k.nodeType ? k : document.createTextNode(String(k)));
  }
  return el;
}
const pad = (n) => String(n).padStart(2, '0');
const dkey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const mkey = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
const parseKey = (k) => { const [y, m, d] = k.split('-').map(Number); return new Date(y, m - 1, d || 1); };
const addDays = (d, n) => new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
const addMonths = (d, n) => new Date(d.getFullYear(), d.getMonth() + n, 1);
const today = () => { const t = new Date(); return new Date(t.getFullYear(), t.getMonth(), t.getDate()); };
const sameDay = (a, b) => dkey(a) === dkey(b);

function weekStartOf(d) {
  const ws = state.data.settings.weekStart;
  return addDays(d, -((d.getDay() - ws + 7) % 7));
}
function thursdayOf(start) { return addDays(start, (4 - start.getDay() + 7) % 7); }
function weekInfo(d) {
  const start = weekStartOf(d);
  const th = thursdayOf(start);
  return { start, end: addDays(start, 6), th, y: th.getFullYear(), m: th.getMonth() + 1, k: Math.floor((th.getDate() - 1) / 7) + 1, key: dkey(th) };
}
function holidayOf(k) { return (state.data.settings.customHolidays || {})[k] || (window.KR_HOLIDAYS || {})[k] || ''; }
function dayTone(d) {
  if (holidayOf(dkey(d)) || d.getDay() === 0) return 'sun';
  if (d.getDay() === 6) return 'sat';
  return '';
}
const isEmpty = (it) => !it || (!(it.x || '').trim() && !(it.t || '').trim());

/* ───────────────────────── state & storage ───────────────────────── */
const state = {
  data: null,
  view: 'week',
  cursor: today(),
  drawerKey: null,
  history: [],
  future: [],
  saveTimer: null,
};

function defaultData() {
  return { version: 1, settings: { weekStart: 0, theme: 'system', scale: 'auto', customHolidays: {} }, days: {}, weeks: {}, months: {} };
}
function normalize(d) {
  const base = defaultData();
  d = d && typeof d === 'object' ? d : base;
  d.settings = Object.assign(base.settings, d.settings || {});
  for (const k of ['days', 'weeks', 'months']) if (!d[k] || typeof d[k] !== 'object') d[k] = {};
  return d;
}
async function loadData() {
  try {
    if (invoke) {
      const s = await invoke('load_data');
      return normalize(s ? JSON.parse(s) : null);
    }
    return normalize(JSON.parse(localStorage.getItem(LS_KEY) || 'null'));
  } catch (e) {
    console.error(e);
    toast('데이터를 읽지 못해 빈 플래너로 시작합니다');
    return defaultData();
  }
}
function compact(d) {
  // 빈 칸 정리 후 저장
  const out = JSON.parse(JSON.stringify(d));
  for (const [k, arr] of Object.entries(out.days)) {
    const a = arr.map((x) => (isEmpty(x) ? null : x));
    if (a.every((x) => !x)) delete out.days[k]; else out.days[k] = a;
  }
  for (const [k, w] of Object.entries(out.weeks)) {
    if (w.next) { w.next = w.next.map((x) => (isEmpty(x) ? null : x)); if (w.next.every((x) => !x)) delete w.next; }
    if (w.goals && w.goals.every((x) => !x)) delete w.goals;
    if (!w.memo) delete w.memo;
    if (!Object.keys(w).length) delete out.weeks[k];
  }
  for (const [k, m] of Object.entries(out.months)) {
    if (m.todos) { m.todos = m.todos.map((x) => (isEmpty(x) ? null : x)); if (m.todos.every((x) => !x)) delete m.todos; }
    if (m.goals && m.goals.every((x) => !x)) delete m.goals;
    if (!Object.keys(m).length) delete out.months[k];
  }
  return out;
}
async function persistNow() {
  clearTimeout(state.saveTimer);
  state.saveTimer = null;
  const json = JSON.stringify(compact(state.data));
  try {
    if (invoke) await invoke('save_data', { json });
    else localStorage.setItem(LS_KEY, json);
  } catch (e) {
    console.error(e);
    toast('저장 실패: ' + e);
  }
}
function markDirty() {
  clearTimeout(state.saveTimer);
  state.saveTimer = setTimeout(persistNow, 450);
}
function snapshot() {
  state.history.push(JSON.stringify(state.data));
  if (state.history.length > 80) state.history.shift();
  state.future = [];
}
function undo() {
  if (!state.history.length) return toast('되돌릴 작업이 없어요');
  state.future.push(JSON.stringify(state.data));
  state.data = JSON.parse(state.history.pop());
  markDirty(); render(); toast('되돌렸어요');
}
function redo() {
  if (!state.future.length) return;
  state.history.push(JSON.stringify(state.data));
  state.data = JSON.parse(state.future.pop());
  markDirty(); render(); toast('다시 실행했어요');
}

/* stores: d:YYYY-MM-DD (하루 10줄) / n:<목요일키> (다음 주로 넘길 일) / m:YYYY-MM (월간 할일) */
function getStore(id, create) {
  const [t, k] = [id[0], id.slice(2)];
  let arr;
  if (t === 'd') {
    arr = state.data.days[k];
    if (!arr && create) arr = state.data.days[k] = [];
  } else if (t === 'n') {
    const w = state.data.weeks[k] || (create ? (state.data.weeks[k] = {}) : null);
    arr = w && (w.next || (create ? (w.next = []) : null));
  } else {
    const m = state.data.months[k] || (create ? (state.data.months[k] = {}) : null);
    arr = m && (m.todos || (create ? (m.todos = []) : null));
  }
  return arr || [];
}
function storeLabel(id) {
  const k = id.slice(2);
  if (id[0] === 'd') { const d = parseKey(k); return `${d.getMonth() + 1}/${d.getDate()}`; }
  if (id[0] === 'n') { const w = weekInfo(parseKey(k)); return `${w.m}월 ${w.k}주`; }
  return `${Number(k.slice(5))}월`;
}
function nextStoreId(id) {
  const k = id.slice(2);
  if (id[0] === 'd') return 'd:' + dkey(addDays(parseKey(k), 1));
  if (id[0] === 'n') return 'd:' + dkey(addDays(weekInfo(parseKey(k)).start, 7));
  return 'm:' + mkey(addMonths(parseKey(k + '-01'), 1));
}
function setItem(id, i, patch) {
  const arr = getStore(id, true);
  arr[i] = Object.assign(arr[i] || {}, patch);
  return arr[i];
}

function applyStatus(id, i, s) {
  const arr = getStore(id, true);
  const it = arr[i] || (arr[i] = {});
  const prev = it.s || '';
  if (prev === s) return;
  if (prev === 'carry') undoCarry(it);
  it.s = s;
  if (s === 'carry' && !isEmpty(it)) doCarry(id, it);
}
function doCarry(id, it) {
  const tid = nextStoreId(id);
  const tarr = getStore(tid, true);
  const max = LEN[tid[0]];
  let j = -1;
  for (let x = 0; x < max; x++) if (isEmpty(tarr[x])) { j = x; break; }
  if (j < 0) { toast(`${storeLabel(tid)} 칸이 가득 차서 이월하지 못했어요`); return; }
  tarr[j] = { x: it.x || '', t: it.t || '', c: it.c || '', d: it.d || '', s: '', from: storeLabel(id) };
  it.link = { id: tid, i: j, x: it.x || '' };
  toast(`${storeLabel(tid)}(으)로 이월했어요`);
}
function undoCarry(it) {
  const L = it.link;
  if (!L) return;
  const arr = getStore(L.id, false);
  const t = arr[L.i];
  if (t && (t.x || '') === L.x && !t.s) arr[L.i] = null;
  delete it.link;
}

/* ───────────────────────── stats ───────────────────────── */
function statsOf(arrs) {
  let total = 0, done = 0;
  for (const arr of arrs) for (const it of arr || []) {
    if (isEmpty(it) || it.s === 'cancel') continue;
    total++;
    if (it.s === 'done') done++;
  }
  return { total, done };
}
function dayStats(k) { return statsOf([state.data.days[k]]); }

/* ───────────────────────── row component ───────────────────────── */
function normTime(v) {
  v = (v || '').trim();
  if (!v) return '';
  let m = v.match(/^(\d{1,2})[:시.\s]?(\d{0,2})분?$/);
  if (!m) return v;
  let [hh, mm] = [m[1], m[2]];
  if (/^\d{3,4}$/.test(v)) { hh = v.slice(0, v.length - 2); mm = v.slice(-2); }
  const H = Number(hh), M = Number(mm || 0);
  if (H > 23 || M > 59) return v;
  return `${pad(H)}:${pad(M)}`;
}
function parseDue(v, ctxMonthKey) {
  v = (v || '').trim();
  if (!v) return '';
  let m = v.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})$/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;
  m = v.match(/^(\d{1,2})[-./\s월]+(\d{1,2})일?$/);
  if (!m) return null;
  const base = parseKey(ctxMonthKey + '-01');
  let y = base.getFullYear();
  const mo = Number(m[1]);
  if (mo < base.getMonth() + 1 - 6) y++;
  return `${y}-${pad(mo)}-${pad(m[2])}`;
}
const fmtDue = (k) => { if (!k) return ''; const d = parseKey(k); return `${d.getMonth() + 1}/${d.getDate()}`; };

/**
 * kind: 'day' (상태·시간·할일·구분) | 'simple' (상태·내용) | 'todo' (상태·할일·마감·구분)
 */
function itemRow(id, i, kind, onChange) {
  const it = (getStore(id, false)[i]) || {};
  const row = h('div', { class: `row ${kind === 'day' ? '' : kind}` });
  const setCls = () => {
    const cur = getStore(id, false)[i] || {};
    row.classList.remove('s-done', 's-carry', 's-cancel');
    if (cur.s) row.classList.add('s-' + cur.s);
    const fromEl = $('.from', row);
    if (fromEl) fromEl.remove();
    if (cur.from && !isEmpty(cur)) row.append(h('span', { class: 'from', title: `${cur.from}에서 이월됨` }, '↪ ' + cur.from));
  };

  const st = h('button', { class: 'st', title: '클릭: 완료 → 이월 → 취소 → 비움\n우클릭: 비움', html: ICONS.done + ICONS.carry + ICONS.cancel, tabindex: '-1' });
  const cycle = (dir = 1) => {
    const cur = getStore(id, false)[i] || {};
    if (isEmpty(cur) && !cur.s) { toast('내용을 먼저 적어 주세요'); return; }
    snapshot();
    const idx = STATUS_CYCLE.indexOf(cur.s || '');
    applyStatus(id, i, STATUS_CYCLE[(idx + dir + STATUS_CYCLE.length) % STATUS_CYCLE.length]);
    setCls();
    st.classList.remove('pop'); void st.offsetWidth; st.classList.add('pop');
    markDirty(); onChange && onChange();
  };
  st.addEventListener('click', (e) => cycle(e.shiftKey ? -1 : 1));
  st.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    const cur = getStore(id, false)[i] || {};
    if (!cur.s) return;
    snapshot(); applyStatus(id, i, ''); setCls(); markDirty(); onChange && onChange();
  });
  row.append(st);

  let typed = false;
  const textInput = (cls, field, ph, transform) => {
    const inp = h('input', { class: `plain ${cls}`, value: it[field] ? (field === 'd' ? fmtDue(it[field]) : it[field]) : '', placeholder: ph || '', spellcheck: 'false' });
    inp.addEventListener('focus', () => { typed = false; });
    inp.addEventListener('input', () => {
      if (!typed) { snapshot(); typed = true; }
      if (field === 'd') return; // 마감일은 blur 때 확정
      setItem(id, i, { [field]: inp.value });
      if (field === 'x') inp.title = inp.value;
      markDirty(); onChange && onChange();
    });
    inp.addEventListener('blur', () => {
      if (transform) {
        const v = transform(inp.value);
        if (v === null) { toast('마감일은 12/5 같은 형식으로 적어 주세요'); inp.value = fmtDue((getStore(id, false)[i] || {}).d); return; }
        if (field === 'd') { setItem(id, i, { d: v }); inp.value = fmtDue(v); updateDue(); }
        else { setItem(id, i, { [field]: v }); inp.value = v; }
        markDirty(); onChange && onChange();
      }
      setCls();
    });
    inp.addEventListener('keydown', (e) => rowKeys(e, row, cycle));
    return inp;
  };

  let dueInput;
  const updateDue = () => {
    if (!dueInput) return;
    const cur = getStore(id, false)[i] || {};
    const over = cur.d && cur.d < dkey(today()) && cur.s !== 'done' && cur.s !== 'cancel';
    dueInput.classList.toggle('overdue', !!over);
  };

  if (kind === 'day') row.append(textInput('tm', 't', '', normTime));
  const tx = textInput('tx', 'x', '');
  tx.title = it.x || '';
  row.append(tx);
  if (kind === 'todo') {
    dueInput = textInput('due', 'd', '', (v) => parseDue(v, id.slice(2)));
    row.append(dueInput);
    updateDue();
  }
  if (kind !== 'simple') {
    const ct = h('button', { class: 'ct', tabindex: '-1', title: '구분 선택 (우클릭: 지우기)' });
    const paintCat = () => {
      const c = (getStore(id, false)[i] || {}).c;
      if (c) { ct.dataset.cat = c; ct.textContent = c; } else { delete ct.dataset.cat; ct.textContent = ''; }
    };
    paintCat();
    ct.addEventListener('click', (e) => {
      e.stopPropagation();
      openCatPopover(ct, (c) => { snapshot(); setItem(id, i, { c }); paintCat(); markDirty(); onChange && onChange(); });
    });
    ct.addEventListener('contextmenu', (e) => { e.preventDefault(); snapshot(); setItem(id, i, { c: '' }); paintCat(); markDirty(); onChange && onChange(); });
    row.append(ct);
  }
  setCls();
  return row;
}

function rowKeys(e, row, cycle) {
  const inp = e.target;
  const cls = inp.classList.contains('tm') ? '.tm' : inp.classList.contains('due') ? '.due' : '.tx';
  const move = (dir) => {
    let r = dir > 0 ? row.nextElementSibling : row.previousElementSibling;
    if (r && r.classList.contains('row')) { const t = $(cls, r) || $('.tx', r); t && t.focus(); return true; }
    return false;
  };
  if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); cycle(1); }
  else if (e.key === 'Enter' || e.key === 'ArrowDown') { if (move(1)) e.preventDefault(); else if (e.key === 'Enter') inp.blur(); }
  else if (e.key === 'ArrowUp') { if (move(-1)) e.preventDefault(); }
  else if (e.key === 'Escape') inp.blur();
}

/* ───────────────────────── popover / toast / modal ───────────────────────── */
function openCatPopover(anchor, onPick) {
  const pop = $('#popover');
  pop.innerHTML = '';
  for (const c of CATS) pop.append(h('button', { class: 'chip', dataset: { cat: c }, onclick: () => { onPick(c); closePopover(); } }, c));
  pop.append(h('button', { class: 'chip none', onclick: () => { onPick(''); closePopover(); } }, '없음'));
  pop.hidden = false;
  // 화면 배율(zoom) 보정: 화면 좌표 → CSS 좌표
  const z = Number(document.documentElement.style.zoom || 1);
  const r = anchor.getBoundingClientRect();
  const pw = pop.offsetWidth * z, ph = pop.offsetHeight * z;
  const left = Math.max(8, Math.min(window.innerWidth - pw - 8, r.right - pw));
  const top = r.bottom + ph + 12 > window.innerHeight ? r.top - ph - 6 : r.bottom + 6;
  pop.style.left = left / z + 'px';
  pop.style.top = top / z + 'px';
}
function closePopover() { $('#popover').hidden = true; }
document.addEventListener('mousedown', (e) => { if (!e.target.closest('#popover')) closePopover(); });

function toast(msg) {
  const t = h('div', { class: 'toast' }, msg);
  $('#toasts').append(t);
  setTimeout(() => { t.classList.add('out'); setTimeout(() => t.remove(), 320); }, 2200);
}

function openModal(title, body, footer) {
  const m = $('#modal');
  m.innerHTML = '';
  const box = h('div', { class: 'modal-box' },
    h('div', { class: 'modal-h' }, title, h('button', { class: 'icon-btn', html: ICONS.close, onclick: closeModal })),
    h('div', { class: 'modal-b' }, body),
    footer ? h('div', { class: 'modal-f' }, footer) : null);
  m.append(box);
  m.hidden = false;
  m.onclick = (e) => { if (e.target === m) closeModal(); };
}
function closeModal() { $('#modal').hidden = true; }

/* ───────────────────────── header ───────────────────────── */
function updateSeg() {
  const seg = $('#viewSeg');
  $$('button', seg).forEach((b) => b.classList.toggle('on', b.dataset.view === state.view));
  const on = $('button.on', seg);
  const ink = $('.seg-ink', seg);
  ink.style.left = on.offsetLeft + 'px';
  ink.style.width = on.offsetWidth + 'px';
}
function setRate(st, label) {
  const pct = st.total ? Math.round((st.done / st.total) * 100) : 0;
  $('#rateNum').textContent = st.total ? pct + '%' : '–';
  $('#rateSub').textContent = st.total ? `${label} ${st.done}/${st.total}` : `${label} 할일 없음`;
  $('#ringFg').style.strokeDashoffset = String(94.25 * (1 - pct / 100));
}

/* ───────────────────────── views ───────────────────────── */
function render(anim = true) {
  applyScale();
  updateSeg();
  applyTheme();
  const v = $('#view');
  v.innerHTML = '';
  const node = state.view === 'week' ? renderWeek() : state.view === 'month' ? renderMonth() : renderYear();
  if (anim) node.classList.add('view-enter');
  v.append(node);
  if (node._mount) node._mount();
}

function weekStats(w) {
  const arrs = [];
  for (let i = 0; i < 7; i++) arrs.push(state.data.days[dkey(addDays(w.start, i))]);
  return statsOf(arrs);
}

function renderWeek() {
  const w = weekInfo(state.cursor);
  $('#periodTitle').innerHTML = '';
  $('#periodTitle').append(`${w.y}년 ${w.m}월 ${w.k}주`,
    h('small', null, `${w.start.getMonth() + 1}.${pad(w.start.getDate())} (${DOW[w.start.getDay()]}) – ${w.end.getMonth() + 1}.${pad(w.end.getDate())} (${DOW[w.end.getDay()]})`));
  const refreshRate = () => setRate(weekStats(w), '이번 주');
  refreshRate();

  const wk = state.data.weeks[w.key] || {};
  const goals = h('div', { class: 'card' }, h('div', { class: 'panel-h' }, '이번 주 목표'));
  const goalRows = h('div', { class: 'goal-rows' });
  goals.append(goalRows);
  for (let g = 0; g < WEEK_GOALS; g++) {
    const inp = h('input', { class: 'plain', value: (wk.goals || [])[g] || '', placeholder: g === 0 ? '이번 주에 꼭 이룰 것' : '' });
    let typed = false;
    inp.addEventListener('focus', () => { typed = false; });
    inp.addEventListener('input', () => {
      if (!typed) { snapshot(); typed = true; }
      const ww = state.data.weeks[w.key] || (state.data.weeks[w.key] = {});
      (ww.goals || (ww.goals = []))[g] = inp.value;
      markDirty();
    });
    goalRows.append(h('div', { class: 'goal-row' }, h('span', { class: 'num' }, g + 1), inp));
  }
  const memo = h('textarea', { class: 'memo-area', placeholder: '주간 메모' });
  memo.value = wk.memo || '';
  let mt = false;
  memo.addEventListener('focus', () => { mt = false; });
  memo.addEventListener('input', () => {
    if (!mt) { snapshot(); mt = true; }
    (state.data.weeks[w.key] || (state.data.weeks[w.key] = {})).memo = memo.value;
    markDirty();
  });
  const memoCard = h('div', { class: 'card' }, h('div', { class: 'panel-h' }, '주간 메모'), memo);

  const grid = h('div', { class: 'week-grid' });
  for (let i = 0; i < 7; i++) grid.append(dayCard(addDays(w.start, i), refreshRate));
  // 다음 주로 넘길 일
  const nid = 'n:' + w.key;
  const next = h('div', { class: 'card next-card' },
    h('div', { class: 'card-h' }, h('div', { class: 'dhead' }, h('span', { class: 'dlabel' }, '다음 주로 넘길 일 · 메모'))),
    h('div', { class: 'cols simple', style: 'grid-template-columns:30px 1fr' }, h('span', null, '상태'), h('span', { style: 'text-align:left;padding-left:6px' }, '내용  ·  → 누르면 다음 주 첫날로 복사')));
  const rows = h('div', { class: 'rows' });
  for (let r = 0; r < LEN.n; r++) rows.append(itemRow(nid, r, 'simple'));
  next.append(rows);
  grid.append(next);

  return h('section', { class: 'week' }, h('div', { class: 'week-top' }, goals, memoCard), grid);
}

function dayHead(d, extraTags) {
  const k = dkey(d);
  const hol = holidayOf(k);
  return h('div', { class: 'dhead' },
    h('span', { class: 'dnum' }, d.getDate()),
    h('span', { class: 'dlabel' }, `${d.getMonth() + 1}월 · ${DOW[d.getDay()]}요일`),
    hol ? h('span', { class: 'tag hol', title: hol }, hol) : null,
    sameDay(d, today()) ? h('span', { class: 'tag today' }, '오늘') : null,
    extraTags || null);
}

function dayCard(d, onAnyChange) {
  const k = dkey(d);
  const id = 'd:' + k;
  const count = h('span', { class: 'count' });
  const refresh = () => {
    const s = dayStats(k);
    count.textContent = s.total ? `${s.done}/${s.total}` : '';
    count.classList.toggle('all', s.total > 0 && s.done === s.total);
    onAnyChange && onAnyChange();
  };
  const card = h('div', { class: `card day ${dayTone(d)} ${sameDay(d, today()) ? 'today' : ''}` },
    h('div', { class: 'card-h' }, dayHead(d), count),
    h('div', { class: 'cols' }, h('span', null, '상태'), h('span', null, '시간'), h('span', null, '할일 · 일정'), h('span', null, '구분')));
  const rows = h('div', { class: 'rows' });
  for (let r = 0; r < LEN.d; r++) rows.append(itemRow(id, r, 'day', refresh));
  card.append(rows);
  refresh();
  return card;
}

/* month */
function monthGridStart(first) { return weekStartOf(first); }

function renderMonth() {
  const c = state.cursor;
  const first = new Date(c.getFullYear(), c.getMonth(), 1);
  const mk = mkey(first);
  $('#periodTitle').innerHTML = '';
  $('#periodTitle').append(`${first.getFullYear()}년 ${first.getMonth() + 1}월`, h('small', null, '날짜를 누르면 그날 할일을 편집해요'));

  const start = monthGridStart(first);
  const nextFirst = addMonths(first, 1);
  const weeks = Math.ceil(((nextFirst - start) / 86400000) / 7);
  const ws = state.data.settings.weekStart;

  const dowRow = h('div', { class: 'dow-row' });
  for (let i = 0; i < 7; i++) {
    const di = (ws + i) % 7;
    dowRow.append(h('div', { class: di === 0 ? 'sun' : di === 6 ? 'sat' : '' }, DOW[di]));
  }
  const grid = h('div', { class: 'cal-grid' });
  grid.style.gridTemplateRows = `repeat(${weeks}, minmax(0, 1fr))`;
  const dayArrs = [];
  for (let i = 0; i < weeks * 7; i++) {
    const d = addDays(start, i);
    const inMonth = d.getMonth() === first.getMonth();
    if (inMonth) dayArrs.push(state.data.days[dkey(d)]);
    grid.append(monthCell(d, inMonth));
  }
  const cal = h('div', { class: 'card cal' }, dowRow, grid);

  // side
  const mo = state.data.months[mk] || {};
  const goals = h('div', { class: 'card' }, h('div', { class: 'panel-h' }, '이달의 목표'));
  const mGoalRows = h('div', { class: 'goal-rows' });
  goals.append(mGoalRows);
  for (let g = 0; g < MONTH_GOALS; g++) {
    const inp = h('input', { class: 'plain', value: (mo.goals || [])[g] || '', placeholder: g === 0 ? '이번 달 목표' : '' });
    let typed = false;
    inp.addEventListener('focus', () => { typed = false; });
    inp.addEventListener('input', () => {
      if (!typed) { snapshot(); typed = true; }
      const m = state.data.months[mk] || (state.data.months[mk] = {});
      (m.goals || (m.goals = []))[g] = inp.value;
      markDirty();
    });
    mGoalRows.append(h('div', { class: 'goal-row' }, h('span', { class: 'num' }, g + 1), inp));
  }
  const tid = 'm:' + mk;
  const badge = h('span', { class: 'count' });
  const refresh = () => {
    const s = statsOf([state.data.months[mk] && state.data.months[mk].todos]);
    badge.textContent = s.total ? `${s.done}/${s.total}` : '';
    badge.classList.toggle('all', s.total > 0 && s.done === s.total);
    setRate(statsOf([...dayArrs, state.data.months[mk] && state.data.months[mk].todos]), '이번 달');
  };
  const todo = h('div', { class: 'card todo-card' },
    h('div', { class: 'panel-h' }, h('span', null, '월간 할일'), badge),
    h('div', { class: 'cols todo' }, h('span', null, '상태'), h('span', { style: 'padding-left:6px' }, '할일'), h('span', { style: 'text-align:center' }, '마감'), h('span', { style: 'text-align:center' }, '구분')));
  const rows = h('div', { class: 'rows todo-rows' });
  todo.append(rows);
  refresh();

  const section = h('section', { class: 'month' }, cal, h('aside', { class: 'side' }, goals, todo));
  section._mount = () => {
    // 화면 높이에 맞춰 할일 줄 수 결정 (입력된 줄보다 적어지지는 않음)
    const arr = getStore(tid, false);
    let last = -1;
    arr.forEach((x, i) => { if (!isEmpty(x)) last = i; });
    const fit = Math.max(8, Math.floor(rows.clientHeight / TODO_ROW_H));
    const n = Math.min(LEN.m, Math.max(fit, last + 2));
    rows.style.gridTemplateRows = n <= fit ? `repeat(${n}, minmax(26px, 1fr))` : '';
    for (let r = 0; r < n; r++) rows.append(itemRow(tid, r, 'todo', refresh));
    fitMonthCells(grid);
  };
  return section;
}

function monthCell(d, inMonth) {
  const k = dkey(d);
  const hol = holidayOf(k);
  const items = (state.data.days[k] || []).filter((x) => !isEmpty(x));
  const cell = h('div', { class: `cell ${inMonth ? '' : 'out'} ${dayTone(d)} ${sameDay(d, today()) ? 'today' : ''}`, onclick: () => openDrawer(k) },
    h('div', { class: 'chead' }, h('span', { class: 'cnum' }, d.getDate()), hol ? h('span', { class: 'chol', title: hol }, hol) : null),
    h('span', { class: 'cadd', html: ICONS.plus }));
  for (const it of items) {
    cell.append(h('div', { class: `cline ${it.s ? 's-' + it.s : ''}`, title: `${it.t ? it.t + ' ' : ''}${it.x || ''}` },
      h('span', { class: 'dot', dataset: it.c ? { cat: it.c } : null }),
      it.t ? h('span', { class: 'ctime' }, it.t) : null,
      h('span', { class: 'ctext' }, it.x || '')));
  }
  cell.append(h('div', { class: 'cmore', hidden: true }));
  return cell;
}

// 칸 높이에 들어가는 만큼만 보여 주고 나머지는 "+N개 더"
function fitMonthCells(grid) {
  for (const cell of $$('.cell', grid)) {
    const lines = $$('.cline', cell);
    const more = $('.cmore', cell);
    if (!lines.length) continue;
    lines.forEach((l) => { l.hidden = false; });
    more.hidden = true;
    const head = $('.chead', cell);
    const avail = cell.clientHeight - head.offsetHeight - 14;
    const lh = lines[0].offsetHeight || 19;
    let cap = Math.max(1, Math.floor(avail / lh));
    if (lines.length > cap) {
      cap = Math.max(1, cap - 1);
      lines.forEach((l, i) => { l.hidden = i >= cap; });
      more.textContent = `+${lines.length - cap}개 더`;
      more.hidden = false;
    }
  }
}

/* drawer (월간에서 하루 편집) */
function openDrawer(k) {
  state.drawerKey = k;
  const dr = $('#drawer');
  const d = parseKey(k);
  const id = 'd:' + k;
  dr.innerHTML = '';
  const count = h('span', { class: 'count' });
  let rerenderTimer = null;
  const refresh = () => {
    const s = dayStats(k);
    count.textContent = s.total ? `${s.done}/${s.total}` : '';
    count.classList.toggle('all', s.total > 0 && s.done === s.total);
    clearTimeout(rerenderTimer);
    if (!initial) rerenderTimer = setTimeout(() => { if (state.view === 'month') render(false); }, 250);
  };
  let initial = true;
  const card = h('div', { class: `card day ${dayTone(d)} ${sameDay(d, today()) ? 'today' : ''}` },
    h('div', { class: 'cols' }, h('span', null, '상태'), h('span', null, '시간'), h('span', null, '할일 · 일정'), h('span', null, '구분')));
  const rows = h('div', { class: 'rows' });
  for (let r = 0; r < LEN.d; r++) rows.append(itemRow(id, r, 'day', refresh));
  card.append(rows);
  const panel = h('div', { class: `drawer-panel ${dayTone(d)}` },
    h('div', { class: 'drawer-head' },
      dayHead(d), count,
      h('button', { class: 'icon-btn', html: ICONS.left, title: '전날', onclick: () => openDrawer(dkey(addDays(d, -1))) }),
      h('button', { class: 'icon-btn', html: ICONS.right, title: '다음날', onclick: () => openDrawer(dkey(addDays(d, 1))) }),
      h('button', { class: 'pill-btn', onclick: () => { closeDrawer(); state.cursor = d; state.view = 'week'; render(); } }, '주간에서 보기'),
      h('button', { class: 'icon-btn', html: ICONS.close, title: '닫기 (Esc)', onclick: closeDrawer })),
    h('div', { class: 'drawer-body' }, card,
      h('p', { class: 'help', style: 'margin:14px 4px 0' }, '상태 동그라미를 누르면 완료 → 이월(다음 날로 복사) → 취소 순으로 바뀌어요. Enter로 다음 줄, Ctrl+Enter로 상태 전환.')));
  dr.append(panel);
  dr.hidden = false;
  dr.onclick = (e) => { if (e.target === dr) closeDrawer(); };
  refresh();
  initial = false;
  setTimeout(() => { const first = rows.querySelector('.row .tx'); const empty = $$('.row', rows).find((r) => !$('.tx', r).value); (empty ? $('.tx', empty) : first).focus(); }, 60);
}
function closeDrawer() {
  if ($('#drawer').hidden) return;
  $('#drawer').hidden = true;
  state.drawerKey = null;
  if (state.view === 'month') render(false);
}

/* year */
function renderYear() {
  const y = state.cursor.getFullYear();
  $('#periodTitle').innerHTML = '';
  $('#periodTitle').append(`${y}년`, h('small', null, '월 이름 → 월간 · 날짜 → 주간'));
  const ws = state.data.settings.weekStart;
  const t = today();
  const curWeek = weekInfo(t);
  const wrap = h('section', { class: 'year' });
  const yearArrs = [];
  for (let m = 0; m < 12; m++) {
    const first = new Date(y, m, 1);
    const mk = mkey(first);
    const arrs = [];
    const card = h('div', { class: `card mini ${y === t.getFullYear() && m === t.getMonth() ? 'cur' : ''}` });
    const grid = h('div', { class: 'mini-grid' });
    for (let i = 0; i < 7; i++) { const di = (ws + i) % 7; grid.append(h('div', { class: `md ${di === 0 ? 'sun' : di === 6 ? 'sat' : ''}` }, DOW[di])); }
    const lead = (first.getDay() - ws + 7) % 7;
    for (let i = 0; i < lead; i++) grid.append(h('span', { class: 'mday blank' }));
    const days = new Date(y, m + 1, 0).getDate();
    for (let dd = 1; dd <= days; dd++) {
      const d = new Date(y, m, dd);
      const k = dkey(d);
      arrs.push(state.data.days[k]);
      const s = dayStats(k);
      const hol = holidayOf(k);
      const inWeek = d >= curWeek.start && d <= curWeek.end;
      const btn = h('button', {
        class: `mday ${dayTone(d)} ${sameDay(d, t) ? 'today' : inWeek ? 'inweek' : ''}`,
        title: [hol, s.total ? `할일 ${s.total}개 · 완료 ${s.done}` : ''].filter(Boolean).join('\n') || null,
        onclick: () => { state.cursor = d; state.view = 'week'; render(); },
      }, dd);
      if (s.total) btn.append(h('span', { class: `bar ${s.done === s.total ? 'alldone' : s.total >= 6 ? 'l3' : s.total >= 3 ? 'l2' : ''}` }));
      grid.append(btn);
    }
    for (let i = lead + days; i < 42; i++) grid.append(h('span', { class: 'mday blank' }));
    const mt = state.data.months[mk] && state.data.months[mk].todos;
    arrs.push(mt);
    yearArrs.push(...arrs);
    const st = statsOf(arrs);
    card.append(
      h('div', { class: 'mini-h' },
        h('button', { onclick: () => { state.cursor = first; state.view = 'month'; render(); } }, `${m + 1}월`),
        st.total ? h('span', { class: `count ${st.done === st.total ? 'all' : ''}` }, `${Math.round(st.done / st.total * 100)}%`) : null),
      grid,
      h('div', { class: 'mini-foot' }, h('span', null, '할일 ', h('b', null, st.total)), h('span', null, '완료 ', h('b', null, st.done))));
    wrap.append(card);
  }
  setRate(statsOf(yearArrs), `${y}년`);
  return wrap;
}

/* ───────────────────────── navigation ───────────────────────── */
function go(dir) {
  const c = state.cursor;
  if (state.view === 'week') state.cursor = addDays(c, 7 * dir);
  else if (state.view === 'month') state.cursor = new Date(c.getFullYear(), c.getMonth() + dir, Math.min(c.getDate(), 28));
  else state.cursor = new Date(c.getFullYear() + dir, c.getMonth(), Math.min(c.getDate(), 28));
  render();
}

/* ───────────────────────── scale ───────────────────────── */
// 넓은 화면(예: 2560×1440 100%)에서는 글자·칸을 함께 키워 균형을 맞춤
function applyScale() {
  const pref = state.data.settings.scale || 'auto';
  const z = pref === 'auto' ? Math.min(1.25, Math.max(1, window.innerWidth / 2130)) : Number(pref) || 1;
  const cur = Number(document.documentElement.style.zoom || 1);
  if (Math.abs(cur - z) > 0.001) document.documentElement.style.zoom = String(Math.round(z * 1000) / 1000);
}

/* ───────────────────────── theme ───────────────────────── */
const mq = window.matchMedia('(prefers-color-scheme: dark)');
function applyTheme() {
  const t = state.data.settings.theme;
  const dark = t === 'dark' || (t === 'system' && mq.matches);
  document.documentElement.dataset.theme = dark ? 'dark' : 'light';
}
mq.addEventListener && mq.addEventListener('change', applyTheme);

/* ───────────────────────── settings ───────────────────────── */
async function openSettings() {
  const s = state.data.settings;
  const choice = (opts, cur, onPick) => {
    const box = h('div', { class: 'choice' });
    for (const [v, label] of opts) {
      box.append(h('button', { class: v === cur ? 'on' : '', onclick: (e) => { $$('button', box).forEach((b) => b.classList.remove('on')); e.currentTarget.classList.add('on'); onPick(v); } }, label));
    }
    return box;
  };
  const body = h('div');
  body.append(h('div', { class: 'set-row' },
    h('div', { class: 'lbl' }, h('b', null, '한 주의 시작'), h('small', null, '주간·월간·연간 화면 모두 바뀌어요')),
    choice([[0, '일요일'], [1, '월요일']], s.weekStart, (v) => { s.weekStart = v; markDirty(); render(); })));
  body.append(h('div', { class: 'set-row' },
    h('div', { class: 'lbl' }, h('b', null, '화면 배율'), h('small', null, '자동: 넓은 모니터에서 글자와 칸을 함께 키워요')),
    choice([['auto', '자동'], ['1', '100%'], ['1.1', '110%'], ['1.2', '120%'], ['1.3', '130%']], s.scale || 'auto', (v) => { s.scale = v; markDirty(); render(false); })));
  body.append(h('div', { class: 'set-row' },
    h('div', { class: 'lbl' }, h('b', null, '테마')),
    choice([['system', '시스템'], ['light', '라이트'], ['dark', '다크']], s.theme, (v) => { s.theme = v; markDirty(); applyTheme(); })));

  if (invoke) {
    const sw = h('button', { class: 'switch' });
    try { if (await invoke('get_autostart')) sw.classList.add('on'); } catch (e) { /* ignore */ }
    sw.onclick = async () => {
      const on = !sw.classList.contains('on');
      try { await invoke('set_autostart', { enabled: on }); sw.classList.toggle('on', on); toast(on ? '윈도우 시작 시 자동 실행' : '자동 실행 해제'); }
      catch (e) { toast('설정 실패: ' + e); }
    };
    body.append(h('div', { class: 'set-row' },
      h('div', { class: 'lbl' }, h('b', null, '윈도우 시작 시 자동 실행'), h('small', null, '트레이에 조용히 띄워 둡니다')), sw));
    let dir = '';
    try { dir = await invoke('data_dir'); } catch (e) { /* ignore */ }
    body.append(h('div', { class: 'set-row' },
      h('div', { class: 'lbl' }, h('b', null, '데이터 저장 위치'), h('small', null, dir || '문서\\Planner'), h('small', { style: 'display:block' }, '하루 한 번 backups 폴더에 자동 백업 (14일 보관)')),
      h('button', { class: 'pill-btn', onclick: () => invoke('open_data_dir').catch((e) => toast(String(e))) }, '폴더 열기')));
  }

  // 사용자 휴일
  const list = h('div', { class: 'hol-list' });
  const paintHol = () => {
    list.innerHTML = '';
    const entries = Object.entries(s.customHolidays || {}).sort();
    if (!entries.length) list.append(h('small', { style: 'color:var(--ink-3)' }, '추가한 휴일이 없어요'));
    for (const [k, name] of entries) {
      list.append(h('div', { class: 'hol-item' }, h('span', null, k), h('span', null, name),
        h('button', { class: 'icon-btn', style: 'width:26px;height:26px', html: ICONS.close, onclick: () => { snapshot(); delete s.customHolidays[k]; markDirty(); paintHol(); render(); } })));
    }
  };
  paintHol();
  const dIn = h('input', { class: 'field', placeholder: '2027-03-15', style: 'width:120px' });
  const nIn = h('input', { class: 'field', placeholder: '이름 (예: 창립기념일)', style: 'flex:1' });
  const addHol = () => {
    const k = parseDue(dIn.value, mkey(state.cursor));
    if (!k || !nIn.value.trim()) return toast('날짜와 이름을 적어 주세요');
    snapshot();
    (s.customHolidays || (s.customHolidays = {}))[k] = nIn.value.trim();
    dIn.value = ''; nIn.value = '';
    markDirty(); paintHol(); render();
  };
  nIn.addEventListener('keydown', (e) => { if (e.key === 'Enter') addHol(); });
  body.append(h('div', { class: 'set-row', style: 'display:block' },
    h('div', { class: 'lbl' }, h('b', null, '사용자 휴일'), h('small', null, '회사 휴무일 등 — 빨간 날로 표시돼요')),
    list, h('div', { class: 'hol-add' }, dIn, nIn, h('button', { class: 'pill-btn primary', onclick: addHol }, '추가'))));

  // 백업
  const fileIn = h('input', { type: 'file', accept: '.json,application/json', style: 'display:none' });
  fileIn.onchange = async () => {
    const f = fileIn.files[0];
    if (!f) return;
    try {
      const d = normalize(JSON.parse(await f.text()));
      if (!(await confirmBox('불러온 백업으로 현재 데이터를 바꿀까요?'))) return;
      snapshot(); state.data = d; markDirty(); closeModal(); render(); toast('백업을 불러왔어요');
    } catch (e) { toast('백업 파일을 읽지 못했어요'); }
  };
  body.append(h('div', { class: 'set-row' },
    h('div', { class: 'lbl' }, h('b', null, '백업'), h('small', null, 'JSON 파일로 내보내거나 불러오기')),
    h('div', { style: 'display:flex;gap:6px' },
      h('button', { class: 'pill-btn', onclick: exportJson }, '내보내기'),
      h('button', { class: 'pill-btn', onclick: () => fileIn.click() }, '불러오기'), fileIn)));

  body.append(h('div', { class: 'set-row', style: 'display:block' },
    h('div', { class: 'lbl' }, h('b', null, '단축키')),
    h('table', { class: 'help' },
      h('tr', null, h('td', null, h('span', { class: 'kbd' }, 'Ctrl+Alt+P')), h('td', null, '어디서든 플래너 열기/숨기기 (앱)')),
      h('tr', null, h('td', null, h('span', { class: 'kbd' }, 'Enter'), ' ', h('span', { class: 'kbd' }, '↑↓')), h('td', null, '줄 이동')),
      h('tr', null, h('td', null, h('span', { class: 'kbd' }, 'Ctrl+Enter')), h('td', null, '상태 전환 (완료 → 이월 → 취소)')),
      h('tr', null, h('td', null, h('span', { class: 'kbd' }, 'Alt+←/→'), ' ', h('span', { class: 'kbd' }, 'T')), h('td', null, '이전/다음 기간 · 오늘로')),
      h('tr', null, h('td', null, h('span', { class: 'kbd' }, '1'), ' ', h('span', { class: 'kbd' }, '2'), ' ', h('span', { class: 'kbd' }, '3')), h('td', null, '주간 · 월간 · 연간')),
      h('tr', null, h('td', null, h('span', { class: 'kbd' }, 'Ctrl+Z'), ' ', h('span', { class: 'kbd' }, 'Ctrl+Y')), h('td', null, '되돌리기 · 다시 실행 (입력칸 밖에서)')))));
  openModal('설정', body);
}
function confirmBox(msg) {
  return new Promise((resolve) => {
    const done = (v) => { closeModal(); resolve(v); };
    openModal('확인', h('p', { class: 'help', style: 'font-size:14px' }, msg), [
      h('button', { class: 'pill-btn', onclick: () => done(false) }, '취소'),
      h('button', { class: 'pill-btn primary', onclick: () => done(true) }, '확인'),
    ]);
  });
}

/* ───────────────────────── export ───────────────────────── */
function openExport() {
  const t = state.cursor;
  const w = weekInfo(t);
  const ranges = [
    ['이번 주', `${w.m}월 ${w.k}주`, w.start, w.end],
    ['이번 달', `${t.getFullYear()}년 ${t.getMonth() + 1}월`, new Date(t.getFullYear(), t.getMonth(), 1), new Date(t.getFullYear(), t.getMonth() + 1, 0)],
    ['올해', `${t.getFullYear()}년 전체`, new Date(t.getFullYear(), 0, 1), new Date(t.getFullYear(), 11, 31)],
    ['전체', '기록된 모든 데이터', null, null],
  ];
  const grid = h('div', { class: 'range-grid' });
  for (const [title, sub, a, b] of ranges) {
    grid.append(h('button', { onclick: () => { closeModal(); doExport(title, a, b); } }, h('b', null, title), h('small', null, sub)));
  }
  openModal('엑셀로 내보내기', h('div', null, grid,
    h('p', { class: 'help', style: 'margin:14px 2px 4px' }, invoke ? '문서\\Planner\\exports 폴더에 .xlsx 파일로 저장하고 폴더를 열어 드려요.' : '브라우저 미리보기에서는 CSV 파일로 내려받아요 (엑셀에서 열림).')));
}

function buildExport(a, b) {
  const D = state.data;
  const inRange = (k) => (!a || k >= dkey(a)) && (!b || k <= dkey(b));
  const sheets = [];
  // 1) 할일·일정
  const rows = [];
  for (const k of Object.keys(D.days).sort()) {
    if (!inRange(k)) continue;
    const d = parseKey(k);
    for (const it of D.days[k] || []) {
      if (isEmpty(it)) continue;
      rows.push([k, DOW[d.getDay()], holidayOf(k), STATUS_LABEL[it.s || ''], it.t || '', it.x || '', it.c || '', it.from ? `${it.from}에서 이월` : '']);
    }
  }
  sheets.push({ name: '할일·일정', columns: ['날짜', '요일', '공휴일', '상태', '시간', '할일·일정', '구분', '비고'], widths: [12, 6, 12, 9, 8, 48, 8, 16], rows });
  // 2) 주간 목표·메모
  const wrows = [];
  for (const k of Object.keys(D.weeks).sort()) {
    const wi = weekInfo(parseKey(k));
    if (!inRange(dkey(wi.end)) && !inRange(dkey(wi.start))) continue;
    const wk = D.weeks[k];
    const next = (wk.next || []).filter((x) => !isEmpty(x)).map((x) => (x.s ? STATUS_LABEL[x.s] + ' ' : '') + x.x).join('\n');
    wrows.push([`${wi.y}년 ${wi.m}월 ${wi.k}주`, `${dkey(wi.start)} ~ ${dkey(wi.end)}`, ...[0, 1, 2].map((g) => (wk.goals || [])[g] || ''), wk.memo || '', next]);
  }
  sheets.push({ name: '주간 목표·메모', columns: ['주차', '기간', '목표 1', '목표 2', '목표 3', '주간 메모', '다음 주로 넘길 일'], widths: [16, 24, 26, 26, 26, 40, 40], rows: wrows });
  // 3) 월간
  const mrows = [];
  for (const k of Object.keys(D.months).sort()) {
    const first = parseKey(k + '-01');
    const last = new Date(first.getFullYear(), first.getMonth() + 1, 0);
    if (a && dkey(last) < dkey(a)) continue;
    if (b && dkey(first) > dkey(b)) continue;
    const mo = D.months[k];
    const label = `${first.getFullYear()}년 ${first.getMonth() + 1}월`;
    (mo.goals || []).forEach((g, i) => { if (g) mrows.push([label, `목표 ${i + 1}`, '', g, '', '']); });
    for (const it of mo.todos || []) if (!isEmpty(it)) mrows.push([label, '할일', STATUS_LABEL[it.s || ''], it.x || '', it.d || '', it.c || '']);
  }
  sheets.push({ name: '월간', columns: ['월', '종류', '상태', '내용', '마감일', '구분'], widths: [14, 9, 9, 48, 12, 8], rows: mrows });
  return sheets;
}

async function doExport(title, a, b) {
  if (window.PLANNER_PREVIEW) return toast('미리보기에서는 내보내기가 동작하지 않아요 — 설치된 앱에서 사용해 주세요');
  await persistNow();
  const sheets = buildExport(a, b);
  const total = sheets.reduce((n, s) => n + s.rows.length, 0);
  if (!total) return toast('내보낼 내용이 없어요');
  const stamp = dkey(today()).replace(/-/g, '');
  const fname = `플래너_${title}_${stamp}`;
  if (invoke) {
    try {
      await invoke('export_xlsx', { payload: JSON.stringify({ sheets }), name: fname });
      toast('엑셀 파일을 저장했어요');
    } catch (e) { toast('내보내기 실패: ' + e); }
  } else {
    const s = sheets[0];
    const esc = (v) => `"${String(v).replace(/"/g, '""')}"`;
    const csv = '﻿' + [s.columns, ...s.rows].map((r) => r.map(esc).join(',')).join('\r\n');
    download(fname + '.csv', csv, 'text/csv;charset=utf-8');
  }
}
async function exportJson() {
  if (window.PLANNER_PREVIEW) return toast('미리보기에서는 백업 내보내기가 동작하지 않아요');
  await persistNow();
  const json = JSON.stringify(compact(state.data), null, 1);
  const name = `planner-backup-${dkey(today())}.json`;
  if (invoke) {
    try { await invoke('export_json', { json, name }); toast('백업 파일을 저장했어요'); } catch (e) { toast('실패: ' + e); }
  } else download(name, json, 'application/json');
}
function download(name, text, type) {
  const a = h('a', { href: URL.createObjectURL(new Blob([text], { type })), download: name });
  document.body.append(a); a.click(); a.remove();
}

/* ───────────────────────── global events ───────────────────────── */
function bindChrome() {
  $$('#viewSeg button').forEach((b) => b.addEventListener('click', () => { state.view = b.dataset.view; render(); }));
  $('#prevBtn').onclick = () => go(-1);
  $('#nextBtn').onclick = () => go(1);
  $('#todayBtn').onclick = () => { state.cursor = today(); render(); };
  $('#themeBtn').onclick = () => {
    const s = state.data.settings;
    const dark = document.documentElement.dataset.theme === 'dark';
    s.theme = dark ? 'light' : 'dark';
    markDirty(); applyTheme();
  };
  $('#settingsBtn').onclick = openSettings;
  $('#exportBtn').onclick = openExport;
  let resizeTimer = null;
  window.addEventListener('resize', () => {
    applyScale();
    updateSeg();
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => {
      const typing = document.activeElement && document.activeElement.closest('#view input, #view textarea');
      if ((state.view === 'month' || state.data.settings.scale === 'auto') && !typing) render(false);
    }, 160);
  });

  document.addEventListener('keydown', (e) => {
    const typing = e.target.closest('input, textarea');
    if (e.key === 'Escape') {
      if (!$('#popover').hidden) return closePopover();
      if (!$('#modal').hidden) return closeModal();
      if (!$('#drawer').hidden) return closeDrawer();
    }
    if (e.altKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight')) { e.preventDefault(); return go(e.key === 'ArrowLeft' ? -1 : 1); }
    if (typing || !$('#modal').hidden) return;
    const k = e.key.toLowerCase();
    if ((e.ctrlKey || e.metaKey) && k === 'z' && !e.shiftKey) { e.preventDefault(); return undo(); }
    if ((e.ctrlKey || e.metaKey) && (k === 'y' || (k === 'z' && e.shiftKey))) { e.preventDefault(); return redo(); }
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    if (k === 't') { state.cursor = today(); render(); }
    else if (k === '1' || k === '2' || k === '3') { state.view = ['week', 'month', 'year'][Number(k) - 1]; render(); }
  });

  // 자정이 지나면 오늘 표시 갱신 (트레이에 오래 띄워 두는 경우)
  let lastDay = dkey(today());
  setInterval(() => {
    const nowKey = dkey(today());
    if (nowKey !== lastDay) {
      lastDay = nowKey;
      if (!document.activeElement || !document.activeElement.closest('input, textarea')) render();
    }
  }, 60000);

  window.addEventListener('beforeunload', () => { if (state.saveTimer) persistNow(); });
  document.addEventListener('visibilitychange', () => { if (document.hidden && state.saveTimer) persistNow(); });

  if (TAURI && TAURI.event) {
    TAURI.event.listen('request-quit', async () => { await persistNow(); invoke('quit_app'); });
    TAURI.event.listen('window-shown', () => {
      const nowKey = dkey(today());
      if (nowKey !== lastDay) { lastDay = nowKey; state.cursor = today(); render(); }
    });
  }
}

(async function init() {
  state.data = await loadData();
  applyScale();
  if (window.PLANNER_PREVIEW && window.PLANNER_SEED && !Object.keys(state.data.days).length) { window.PLANNER_SEED(state.data); markDirty(); setTimeout(() => toast('미리보기용 예시 데이터가 들어 있어요'), 600); }
  applyTheme();
  bindChrome();
  render();
  document.fonts && document.fonts.ready.then(updateSeg);
})();
