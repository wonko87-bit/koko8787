/* ============================================================
   OfficeBridge – Main Logic
   MSAL 인증 → 자연어 파싱 → Microsoft Graph API 호출
   ============================================================ */

'use strict';

// ──────────────────────────────────────────────
// 설정 (localStorage에서 불러옴)
// ──────────────────────────────────────────────
let CONFIG = {
  clientId:        '',
  tenantId:        '',
  teamsWebhookUrl: '',
};

function loadConfig() {
  try {
    const saved = JSON.parse(localStorage.getItem('ob_config') || '{}');
    Object.assign(CONFIG, saved);
  } catch (_) {}
}

function saveConfig() {
  localStorage.setItem('ob_config', JSON.stringify(CONFIG));
}


// ──────────────────────────────────────────────
// MSAL – Azure AD 인증
// ──────────────────────────────────────────────
let msalInstance  = null;
let currentAccount = null;

const GRAPH_SCOPES = [
  'User.Read',
  'Tasks.ReadWrite',
  'Calendars.ReadWrite',
  'Mail.ReadWrite',
];

function initMsal() {
  if (!CONFIG.clientId || !CONFIG.tenantId) return;

  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId:    CONFIG.clientId,
      authority:   `https://login.microsoftonline.com/${CONFIG.tenantId}`,
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: { cacheLocation: 'localStorage' },
  });

  // 페이지 로드 시 기존 계정 확인
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    currentAccount = accounts[0];
    updateAuthUI();
  }
}

async function signIn() {
  if (!msalInstance) {
    showResult('먼저 설정에서 클라이언트 ID와 테넌트 ID를 입력해주세요.', 'error');
    openSettings();
    return;
  }
  try {
    const result = await msalInstance.loginPopup({ scopes: GRAPH_SCOPES });
    currentAccount = result.account;
    updateAuthUI();
  } catch (e) {
    showResult(`로그인 실패: ${e.message}`, 'error');
  }
}

async function signOut() {
  if (!msalInstance || !currentAccount) return;
  await msalInstance.logoutPopup({ account: currentAccount });
  currentAccount = null;
  updateAuthUI();
}

async function getToken() {
  if (!msalInstance || !currentAccount) return null;
  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: GRAPH_SCOPES,
      account: currentAccount,
    });
    return result.accessToken;
  } catch (_) {
    // silent 실패 → popup으로 재시도
    try {
      const result = await msalInstance.acquireTokenPopup({ scopes: GRAPH_SCOPES });
      currentAccount = result.account;
      return result.accessToken;
    } catch (e) {
      showResult(`토큰 획득 실패: ${e.message}`, 'error');
      return null;
    }
  }
}

function updateAuthUI() {
  const authBtn  = document.getElementById('authBtn');
  const userName = document.getElementById('userName');
  if (currentAccount) {
    authBtn.textContent  = '로그아웃';
    userName.textContent = currentAccount.name || currentAccount.username;
  } else {
    authBtn.textContent  = '로그인';
    userName.textContent = '';
  }
}


// ──────────────────────────────────────────────
// 자연어 파서
// ──────────────────────────────────────────────

const KOR_DAYS = { '월': 1, '화': 2, '수': 3, '목': 4, '금': 5, '토': 6, '일': 0 };
const ENG_DAYS = { 'mon': 1, 'tue': 2, 'wed': 3, 'thu': 4, 'fri': 5, 'sat': 6, 'sun': 0 };

function parseInput(raw) {
  let text = raw.trim();

  // 타겟 감지 (해시태그 우선, 없으면 키워드)
  let target = 'task';
  if      (/#teams|#팀즈/i.test(text) || /^팀즈[:\s]/i.test(text))   target = 'teams';
  else if (/#mail|#메일/i.test(text)   || /^메일[:\s]/i.test(text))   target = 'mail';
  else if (/#meeting|#회의|#미팅/i.test(text) ||
           /^(회의|미팅)[:\s]/i.test(text))                           target = 'meeting';
  else if (/#task|#할일/i.test(text)   || /^할일[:\s]/i.test(text))   target = 'task';

  // 해시태그 제거
  text = text.replace(/#\S+/g, '').trim();

  // 날짜/시간 추출
  const { startDt, endDt, remainder: afterDate } = extractDateTime(text);

  // 수신자 추출 (메일/회의용)
  const { recipients, remainder: afterRecip } = extractRecipients(afterDate);

  // 지속 시간 추출 (회의용, 남은 텍스트에서)
  const durationMin = extractDuration(afterRecip);

  // 최종 제목/본문
  const lines = afterRecip.split('\n');
  const title  = lines[0].trim();
  const body   = lines.slice(1).join('\n').trim();

  return { target, title, body, startDt, endDt, durationMin, recipients };
}


function extractDateTime(text) {
  const now = new Date();
  let startDt = null;
  let endDt   = null;
  let remainder = text;

  // 날짜 패턴 후보
  const datePatterns = [
    // 오늘 / 내일 / 모레
    { re: /오늘/,         fn: () => new Date(now.getFullYear(), now.getMonth(), now.getDate()) },
    { re: /내일/,         fn: () => new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1) },
    { re: /모레/,         fn: () => new Date(now.getFullYear(), now.getMonth(), now.getDate() + 2) },
    // 다음주
    { re: /다음\s*주/,   fn: () => {
        const d = new Date(now);
        d.setDate(d.getDate() + (8 - d.getDay()) % 7 + 1);
        return d;
      }
    },
    // 이번주 / 다음주 + 요일  (이번주 목, 다음주 금 등)
    { re: /(?:이번\s*주\s*)?([월화수목금토일])요일/,
      fn: (m) => {
        const target = KOR_DAYS[m[1]];
        const d = new Date(now);
        let diff = target - d.getDay();
        if (diff <= 0) diff += 7;
        d.setDate(d.getDate() + diff);
        return d;
      }
    },
    { re: /다음\s*주\s*([월화수목금토일])요일/,
      fn: (m) => {
        const target = KOR_DAYS[m[1]];
        const d = new Date(now);
        let diff = target - d.getDay() + 7;
        d.setDate(d.getDate() + diff);
        return d;
      }
    },
    // M월 D일
    { re: /(\d{1,2})월\s*(\d{1,2})일/,
      fn: (m) => new Date(now.getFullYear(), parseInt(m[1]) - 1, parseInt(m[2]))
    },
    // YYYY-MM-DD 또는 YYYY/MM/DD
    { re: /(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/,
      fn: (m) => new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]))
    },
  ];

  let baseDate = null;
  for (const { re, fn } of datePatterns) {
    const m = remainder.match(re);
    if (m) {
      baseDate  = fn(m);
      remainder = remainder.replace(m[0], '').trim();
      break;
    }
  }

  // 시간 패턴
  const timePatterns = [
    // 시간 범위: 오전/오후 HH시 ~ HH시  또는  HH:MM ~ HH:MM
    { re: /(오전|오후)?\s*(\d{1,2})시\s*[-~]\s*(오전|오후)?\s*(\d{1,2})시/,
      fn: (m, base) => {
        const s = toHour(parseInt(m[2]), m[1]);
        const e = toHour(parseInt(m[4]), m[3] || m[1]);
        return { start: setTime(base, s, 0), end: setTime(base, e, 0) };
      }
    },
    { re: /(\d{1,2}):(\d{2})\s*[-~]\s*(\d{1,2}):(\d{2})/,
      fn: (m, base) => ({
        start: setTime(base, parseInt(m[1]), parseInt(m[2])),
        end:   setTime(base, parseInt(m[3]), parseInt(m[4])),
      })
    },
    // 단일 시간: 오전/오후 HH시 MM분  또는  HH:MM
    { re: /(오전|오후)\s*(\d{1,2})시(?:\s*(\d{1,2})분)?/,
      fn: (m, base) => {
        const h = toHour(parseInt(m[2]), m[1]);
        const min = parseInt(m[3] || '0');
        return { start: setTime(base, h, min), end: null };
      }
    },
    { re: /(\d{1,2})시(?:\s*(\d{1,2})분)?/,
      fn: (m, base) => ({
        start: setTime(base, parseInt(m[1]), parseInt(m[2] || '0')),
        end: null,
      })
    },
    { re: /(\d{1,2}):(\d{2})/,
      fn: (m, base) => ({
        start: setTime(base, parseInt(m[1]), parseInt(m[2])),
        end: null,
      })
    },
  ];

  if (baseDate) {
    for (const { re, fn } of timePatterns) {
      const m = remainder.match(re);
      if (m) {
        const { start, end } = fn(m, baseDate);
        startDt   = start;
        endDt     = end;
        remainder = remainder.replace(m[0], '').trim();
        break;
      }
    }
    if (!startDt) startDt = baseDate;
  }

  remainder = remainder.replace(/\s{2,}/g, ' ').trim();
  return { startDt, endDt, remainder };
}

function toHour(h, ampm) {
  if (!ampm) return h;
  if (ampm === '오후' && h < 12) return h + 12;
  if (ampm === '오전' && h === 12) return 0;
  return h;
}

function setTime(base, h, m) {
  const d = new Date(base);
  d.setHours(h, m, 0, 0);
  return d;
}

function extractDuration(text) {
  // "1시간", "30분", "1시간 30분"
  const m = text.match(/(\d+)\s*시간(?:\s*(\d+)\s*분)?|(\d+)\s*분/);
  if (!m) return 60;
  if (m[3]) return parseInt(m[3]);
  return parseInt(m[1]) * 60 + parseInt(m[2] || '0');
}

function extractRecipients(text) {
  // "홍길동에게", "홍길동한테", "user@domain.com"
  const recipients = [];
  let remainder = text;

  const emailRe = /[\w.+-]+@[\w-]+\.[a-z]{2,}/gi;
  const emails  = remainder.match(emailRe) || [];
  emails.forEach(e => {
    recipients.push({ address: e, name: '' });
    remainder = remainder.replace(e, '');
  });

  const nameRe = /([가-힣a-zA-Z\s]{2,6})(?:에게|한테)/g;
  let m;
  while ((m = nameRe.exec(remainder)) !== null) {
    recipients.push({ address: '', name: m[1].trim() });
    remainder = remainder.replace(m[0], '');
  }

  return { recipients, remainder: remainder.trim() };
}

function fmtDt(dt) {
  // "YYYY-MM-DDTHH:MM:SS" (Graph API 형식)
  if (!dt) return null;
  const pad = n => String(n).padStart(2, '0');
  return `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())}T${pad(dt.getHours())}:${pad(dt.getMinutes())}:00`;
}

function displayDt(dt) {
  if (!dt) return '';
  const days = ['일', '월', '화', '수', '목', '금', '토'];
  const pad  = n => String(n).padStart(2, '0');
  return `${dt.getMonth() + 1}/${dt.getDate()}(${days[dt.getDay()]}) ${pad(dt.getHours())}:${pad(dt.getMinutes())}`;
}


// ──────────────────────────────────────────────
// Microsoft Graph API
// ──────────────────────────────────────────────
const GRAPH = 'https://graph.microsoft.com/v1.0';
const TZ    = Intl.DateTimeFormat().resolvedOptions().timeZone || 'Asia/Seoul';

async function graphRequest(token, method, path, body) {
  const resp = await fetch(`${GRAPH}${path}`, {
    method,
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type':  'application/json',
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err.error?.message || `HTTP ${resp.status}`);
  }
  return resp.status === 204 ? null : resp.json();
}

async function getDefaultTaskListId(token) {
  const data = await graphRequest(token, 'GET', '/me/todo/lists');
  const list  = (data.value || []).find(l => l.wellknownListName === 'defaultList')
             || data.value?.[0];
  if (!list) throw new Error('To Do 기본 목록을 찾을 수 없습니다.');
  return list.id;
}

async function createTask(token, parsed) {
  const listId = await getDefaultTaskListId(token);

  const taskBody = { title: parsed.title || '(제목 없음)' };
  if (parsed.startDt) {
    taskBody.dueDateTime = { dateTime: fmtDt(parsed.startDt), timeZone: TZ };
  }
  if (parsed.body) {
    taskBody.body = { content: parsed.body, contentType: 'text' };
  }

  await graphRequest(token, 'POST', `/me/todo/lists/${listId}/tasks`, taskBody);

  return `✅ <b>To Do 할 일 등록 완료</b><br/>
    제목: ${parsed.title}<br/>
    ${parsed.startDt ? `마감일: ${displayDt(parsed.startDt)}` : ''}`;
}

async function createEvent(token, parsed) {
  const start = parsed.startDt || new Date();
  const durationMs = (parsed.durationMin || 60) * 60 * 1000;
  const end   = parsed.endDt || new Date(start.getTime() + durationMs);

  const event = {
    subject: parsed.title || '(제목 없음)',
    body:    { contentType: 'HTML', content: parsed.body || '' },
    start:   { dateTime: fmtDt(start), timeZone: TZ },
    end:     { dateTime: fmtDt(end),   timeZone: TZ },
    allowNewTimeProposals: true,
  };

  if (parsed.recipients.length > 0) {
    event.attendees = parsed.recipients
      .filter(r => r.address)
      .map(r => ({
        emailAddress: { address: r.address, name: r.name || r.address },
        type: 'required',
      }));
  }

  const created = await graphRequest(token, 'POST', '/me/events', event);

  return `✅ <b>캘린더 일정 등록 완료</b><br/>
    제목: ${parsed.title}<br/>
    시작: ${displayDt(start)}<br/>
    종료: ${displayDt(end)}<br/>
    ${created?.webLink ? `<a href="${created.webLink}" target="_blank">Outlook에서 보기 →</a>` : ''}`;
}

async function createDraftEmail(token, parsed) {
  const msg = {
    subject: parsed.title || '(제목 없음)',
    body:    { contentType: 'HTML', content: parsed.body || '' },
    toRecipients: parsed.recipients
      .filter(r => r.address)
      .map(r => ({ emailAddress: { address: r.address, name: r.name || r.address } })),
  };

  await graphRequest(token, 'POST', '/me/messages', msg);

  const rcptStr = parsed.recipients.map(r => r.name || r.address).join(', ') || '(수신자 미지정)';
  return `✅ <b>메일 초안 저장 완료</b><br/>
    제목: ${parsed.title}<br/>
    수신: ${rcptStr}<br/>
    <small>초안 폴더에 저장되었습니다. Outlook에서 확인 후 발송하세요.</small>`;
}


// ──────────────────────────────────────────────
// Teams Incoming Webhook
// ──────────────────────────────────────────────
async function sendTeamsMessage(parsed) {
  const url = CONFIG.teamsWebhookUrl;
  if (!url) {
    throw new Error('Teams 웹훅 URL이 설정되지 않았습니다. 설정 패널에서 입력해주세요.');
  }

  const payload = {
    '@type':    'MessageCard',
    '@context': 'http://schema.org/extensions',
    themeColor: '0F6CBD',
    summary:    parsed.title || '(OfficeBridge 메시지)',
    sections: [{
      activityTitle:    `📢 ${parsed.title || '(제목 없음)'}`,
      activitySubtitle: `OfficeBridge • ${new Date().toLocaleString('ko-KR')}`,
      activityText:     parsed.body || '',
      markdown: true,
    }],
  };

  const resp = await fetch(url, {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify(payload),
  });

  if (!resp.ok) throw new Error(`Teams 전송 실패: HTTP ${resp.status}`);

  return `✅ <b>Teams 메시지 전송 완료</b><br/>내용: ${parsed.title}`;
}


// ──────────────────────────────────────────────
// Outlook 컨텍스트 (현재 메일 가져오기)
// ──────────────────────────────────────────────
function tryLoadFromMailItem() {
  try {
    const item = Office.context.mailbox?.item;
    if (!item) return;

    const mailCtx  = document.getElementById('mailContext');
    const mailText = document.getElementById('mailContextText');
    const loadBtn  = document.getElementById('loadMailBtn');

    const subject  = item.subject || '';
    if (subject) {
      mailCtx.classList.remove('hidden');
      mailText.textContent = `📧 ${subject}`;
    }

    loadBtn.addEventListener('click', () => {
      const nl = document.getElementById('nlInput');
      const fromStr = item.from?.displayName
        ? `${item.from.displayName}에게 Re: ${subject}`
        : `Re: ${subject}`;
      nl.value = fromStr + ' #mail';
      nl.focus();
    });
  } catch (_) {}
}


// ──────────────────────────────────────────────
// 결과 / UI 헬퍼
// ──────────────────────────────────────────────
function showResult(html, type = 'success') {
  const panel = document.getElementById('resultPanel');
  panel.innerHTML  = html;
  panel.className  = `result-panel ${type}`;
  panel.classList.remove('hidden');
  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

const TARGET_LABELS = { task: '할일', meeting: '회의', mail: '메일', teams: 'Teams' };

function addHistory(entry) {
  const list    = document.getElementById('historyList');
  const history = JSON.parse(localStorage.getItem('ob_history') || '[]');
  history.unshift(entry);
  if (history.length > 20) history.pop();
  localStorage.setItem('ob_history', JSON.stringify(history));
  renderHistory(history);
}

function renderHistory(history) {
  const list = document.getElementById('historyList');
  list.innerHTML = '';
  history.forEach(entry => {
    const li   = document.createElement('li');
    li.className = 'history-item';
    const time   = new Date(entry.time);
    const hhmm   = `${String(time.getHours()).padStart(2,'0')}:${String(time.getMinutes()).padStart(2,'0')}`;
    li.innerHTML = `
      <span class="history-tag tag-${entry.target}">${TARGET_LABELS[entry.target] || entry.target}</span>
      <span class="history-text">${escHtml(entry.title)}</span>
      <span class="history-time">${hhmm}</span>
    `;
    li.addEventListener('click', () => {
      document.getElementById('nlInput').value = entry.raw;
    });
    list.appendChild(li);
  });
}

function escHtml(str) {
  return (str || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function openSettings() {
  document.getElementById('settingsPanel').classList.remove('hidden');
}

function loadSettings() {
  loadConfig();
  document.getElementById('cfgClientId').value = CONFIG.clientId;
  document.getElementById('cfgTenantId').value = CONFIG.tenantId;
  document.getElementById('cfgWebhook').value  = CONFIG.teamsWebhookUrl;
  document.getElementById('redirectUriHint').textContent =
    window.location.origin + window.location.pathname;

  const saved = JSON.parse(localStorage.getItem('ob_history') || '[]');
  renderHistory(saved);
}


// ──────────────────────────────────────────────
// 메인 핸들러
// ──────────────────────────────────────────────
async function handleSubmit() {
  const nlInput   = document.getElementById('nlInput');
  const submitBtn = document.getElementById('submitBtn');
  const raw = nlInput.value.trim();
  if (!raw) return;

  submitBtn.disabled   = true;
  submitBtn.textContent = '처리 중...';

  try {
    const parsed = parseInput(raw);

    let resultHtml;

    if (parsed.target === 'teams') {
      resultHtml = await sendTeamsMessage(parsed);
    } else {
      const token = await getToken();
      if (!token) {
        showResult('먼저 로그인이 필요합니다.', 'error');
        return;
      }
      switch (parsed.target) {
        case 'task':    resultHtml = await createTask(token, parsed);        break;
        case 'meeting': resultHtml = await createEvent(token, parsed);       break;
        case 'mail':    resultHtml = await createDraftEmail(token, parsed);  break;
        default:        resultHtml = await createTask(token, parsed);
      }
    }

    showResult(resultHtml, 'success');
    addHistory({ raw, title: parsed.title, target: parsed.target, time: new Date().toISOString() });
    nlInput.value = '';

  } catch (e) {
    showResult(`❌ ${escHtml(e.message)}`, 'error');
  } finally {
    submitBtn.disabled   = false;
    submitBtn.textContent = '실행 ↵';
  }
}


// ──────────────────────────────────────────────
// Office.onReady – 진입점
// ──────────────────────────────────────────────
Office.onReady(() => {
  loadSettings();
  initMsal();
  tryLoadFromMailItem();

  // 이벤트 바인딩
  document.getElementById('submitBtn').addEventListener('click', handleSubmit);

  document.getElementById('nlInput').addEventListener('keydown', e => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) handleSubmit();
  });

  document.getElementById('authBtn').addEventListener('click', () => {
    currentAccount ? signOut() : signIn();
  });

  document.getElementById('settingsBtn').addEventListener('click', () => {
    document.getElementById('settingsPanel').classList.toggle('hidden');
  });

  document.getElementById('saveSettingsBtn').addEventListener('click', () => {
    CONFIG.clientId        = document.getElementById('cfgClientId').value.trim();
    CONFIG.tenantId        = document.getElementById('cfgTenantId').value.trim();
    CONFIG.teamsWebhookUrl = document.getElementById('cfgWebhook').value.trim();
    saveConfig();
    initMsal();
    document.getElementById('settingsPanel').classList.add('hidden');
    showResult('✅ 설정이 저장되었습니다. 다시 로그인해주세요.', 'success');
  });

  document.getElementById('clearHistoryBtn').addEventListener('click', () => {
    localStorage.removeItem('ob_history');
    renderHistory([]);
  });
});
