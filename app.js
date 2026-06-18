/**
 * Trip Conflict Checker — app.js  v1.0（Outlook 個人版）
 * Microsoft Graph API + MSAL.js 2.x（Browser）
 *
 * ── 填入你的金鑰 ──────────────────────────────
 * 步驟：Azure Portal → Azure Active Directory → 應用程式註冊
 * 1. 新增註冊 → 取得 Application (client) ID → 填入 CLIENT_ID
 * 2. 目錄 (tenant) ID → 填入 TENANT_ID
 *    （若是個人 Microsoft 帳號，TENANT_ID 填 'common'）
 *    （若是組織帳號 Microsoft 365，填實際 Tenant ID 或 'organizations'）
 *
 * 版本紀錄：
 * v1.0  個人單機版。僅讀寫 c1552@csd.org.tw 的 Outlook 日曆。
 *        共用日曆功能保留擴充介面，預設停用。
 * ─────────────────────────────────────────────
 */

const CONFIG = {
  CLIENT_ID:  'd742d26c-286a-4bcf-bbfa-ce9e670b8743',  // trip-checker-outlook Application ID
  TENANT_ID:  'cf8303ef-32f0-47cc-bee0-bb6092fcc083',  // 財團法人中衛發展中心 Tenant ID
  ACCOUNT:    'c1552@csd.org.tw',

  SCOPES: [
    'Calendars.ReadWrite',             // 讀取與寫入日曆
    'User.Read',                       // 讀取使用者基本資料
  ],

  GRAPH_ENDPOINT: 'https://graph.microsoft.com/v1.0',

  /* ── 擴充預留：日後啟用共用日曆時填入並設為 true ── */
  SHARED_CALENDAR_ENABLED: false,
  SHARED_CALENDAR_ID:      '',
};

/* ── 不可移動關鍵字（與 Google 版相同） ── */
const DEFAULT_LOCKED_KEYWORDS = [
  '董事長','副總','總經理','執行長','局長','處長','院長','主委',
  '評審','簡報','開幕','結案','簽約','考試','答辯','典禮',
  '課程','授課','演講','工作坊','培訓',
];

const DEFAULT_MOVABLE_KEYWORDS = ['待確認','暫定','TBD','tbd','草稿','draft','提醒'];

/* ════════════════════════════════════════
   MSAL 初始化
   ════════════════════════════════════════ */
let msalInstance = null;
let state = {
  isSignedIn:     false,
  account:        null,
  accessToken:    null,
  lockedKeywords:  [...DEFAULT_LOCKED_KEYWORDS],
  movableKeywords: [...DEFAULT_MOVABLE_KEYWORDS],
  tripInfo:       null,
  events:         [],
  classified:     { conflict: [], gray: [], movable: [], freeSlots: [] },
  grayUnlocked:   new Set(),
  currentModal:   null,
};

function initMSAL() {
  const msalConfig = {
    auth: {
      clientId:    CONFIG.CLIENT_ID,
      authority:   `https://login.microsoftonline.com/${CONFIG.TENANT_ID}`,
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: {
      cacheLocation:        'localStorage',
      storeAuthStateInCookie: false,
    },
  };

  msalInstance = new msal.PublicClientApplication(msalConfig);

  // 處理登入後的 redirect 回傳
  msalInstance.handleRedirectPromise().then(resp => {
    if (resp && resp.account) {
      state.account     = resp.account;
      state.accessToken = resp.accessToken;
      state.isSignedIn  = true;
      renderAuthArea();
      showSection('input');
    } else {
      // 檢查是否已有快取帳號
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        state.account    = accounts[0];
        state.isSignedIn = true;
        acquireTokenSilently().then(() => {
          renderAuthArea();
          showSection('input');
        });
      }
    }
  }).catch(err => {
    console.error('MSAL redirect error:', err);
    showToast('登入處理失敗，請重新嘗試', 'error');
  });
}

async function acquireTokenSilently() {
  try {
    const resp = await msalInstance.acquireTokenSilent({
      scopes:  CONFIG.SCOPES,
      account: state.account,
    });
    state.accessToken = resp.accessToken;
    return resp.accessToken;
  } catch (err) {
    // silent 失敗 → 改用互動式登入
    if (err instanceof msal.InteractionRequiredAuthError) {
      await loginRedirect();
    } else {
      throw err;
    }
  }
}

async function loginRedirect() {
  await msalInstance.loginRedirect({
    scopes:      CONFIG.SCOPES,
    loginHint:   CONFIG.ACCOUNT,  // 預填帳號
    prompt:      'select_account',
  });
}

function signOut() {
  if (!confirm('確定要登出？')) return;
  msalInstance.logoutRedirect({
    account:              state.account,
    postLogoutRedirectUri: window.location.href,
  });
}

/* ════════════════════════════════════════
   GRAPH API 通用請求
   ════════════════════════════════════════ */
async function graphRequest(endpoint, method = 'GET', body = null) {
  // 每次請求前確保 token 有效
  await acquireTokenSilently();

  const opts = {
    method,
    headers: {
      'Authorization': `Bearer ${state.accessToken}`,
      'Content-Type':  'application/json',
    },
  };
  if (body) opts.body = JSON.stringify(body);

  const resp = await fetch(`${CONFIG.GRAPH_ENDPOINT}${endpoint}`, opts);

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err?.error?.message || `HTTP ${resp.status}`);
  }

  // 204 No Content（PATCH 成功）不需要 parse
  if (resp.status === 204) return null;
  return resp.json();
}

/* ════════════════════════════════════════
   INIT
   ════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
  loadKeywordsFromStorage();
  renderKeywordTags();
  bindUIEvents();
  setDefaultDatetimes();

  // 等待 MSAL library 完全載入（最多等 10 秒）
  waitForMSAL(0);
});

function waitForMSAL(attempts) {
  if (typeof msal !== 'undefined' && msal.PublicClientApplication) {
    initMSAL();
  } else if (attempts < 50) {
    setTimeout(() => waitForMSAL(attempts + 1), 200);
  } else {
    showToast('Microsoft 登入服務載入逾時，請重新整理頁面後再試', 'error');
  }
}

/* ════════════════════════════════════════
   AUTH UI
   ════════════════════════════════════════ */
function renderAuthArea() {
  const area = document.getElementById('authArea');
  const name = state.account?.name || state.account?.username || CONFIG.ACCOUNT;
  const initial = name[0].toUpperCase();
  area.innerHTML = `
    <div class="user-chip" id="userChip" title="點擊登出">
      <div class="user-avatar-placeholder">${initial}</div>
      <span class="user-name">${name}</span>
      <span class="user-logout">登出</span>
    </div>`;
  document.getElementById('userChip').addEventListener('click', signOut);
}

/* ════════════════════════════════════════
   UI EVENTS
   ════════════════════════════════════════ */
function bindUIEvents() {
  document.getElementById('btnLogin').addEventListener('click', loginRedirect);
  document.getElementById('btnLoginWelcome').addEventListener('click', loginRedirect);
  document.getElementById('btnAnalyze').addEventListener('click', startAnalysis);
  document.getElementById('btnBackToInput').addEventListener('click', () => showSection('input'));

  document.getElementById('btnToggleKeywords').addEventListener('click', () => {
    const body = document.getElementById('keywordBody');
    const btn  = document.getElementById('btnToggleKeywords');
    const open = body.style.display === 'none';
    body.style.display = open ? '' : 'none';
    btn.textContent    = open ? '收起設定' : '展開設定';
  });

  document.getElementById('btnAddKeyword').addEventListener('click', addKeyword);
  document.getElementById('newKeyword').addEventListener('keydown', e => {
    if (e.key === 'Enter') addKeyword();
  });

  document.getElementById('btnCopyText').addEventListener('click', copyTextSummary);
  document.getElementById('btnExportPDF').addEventListener('click', () => window.print());

  document.getElementById('modalClose').addEventListener('click',  closeModal);
  document.getElementById('modalCancel').addEventListener('click', closeModal);
  document.getElementById('modalConfirm').addEventListener('click', confirmWriteToCalendar);
  document.getElementById('modalOverlay').addEventListener('click', e => {
    if (e.target === document.getElementById('modalOverlay')) closeModal();
  });
}

/* ════════════════════════════════════════
   DEFAULT DATETIMES
   ════════════════════════════════════════ */
function setDefaultDatetimes() {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  tomorrow.setHours(8, 0, 0, 0);
  const dayAfter = new Date(tomorrow);
  dayAfter.setDate(dayAfter.getDate() + 2);
  dayAfter.setHours(18, 0, 0, 0);
  document.getElementById('tripStart').value = toDatetimeLocal(tomorrow);
  document.getElementById('tripEnd').value   = toDatetimeLocal(dayAfter);
}

/* ════════════════════════════════════════
   KEYWORDS
   ════════════════════════════════════════ */
function loadKeywordsFromStorage() {
  try {
    const saved = localStorage.getItem('trip_checker_outlook_keywords');
    if (saved) state.lockedKeywords = JSON.parse(saved);
  } catch(e) {}
}

function saveKeywordsToStorage() {
  localStorage.setItem('trip_checker_outlook_keywords', JSON.stringify(state.lockedKeywords));
}

function addKeyword() {
  const input = document.getElementById('newKeyword');
  const kw = input.value.trim();
  if (!kw || state.lockedKeywords.includes(kw)) { input.value = ''; return; }
  state.lockedKeywords.push(kw);
  saveKeywordsToStorage();
  renderKeywordTags();
  input.value = '';
}

function removeKeyword(kw) {
  state.lockedKeywords = state.lockedKeywords.filter(k => k !== kw);
  saveKeywordsToStorage();
  renderKeywordTags();
}

function renderKeywordTags() {
  const container = document.getElementById('keywordTags');
  container.innerHTML = state.lockedKeywords.map(kw => `
    <div class="keyword-tag">
      <span>${kw}</span>
      <span class="tag-remove" data-kw="${kw}" title="移除">✕</span>
    </div>`).join('');
  container.querySelectorAll('.tag-remove').forEach(el => {
    el.addEventListener('click', () => removeKeyword(el.dataset.kw));
  });
}

/* ════════════════════════════════════════
   ANALYSIS ENTRY
   ════════════════════════════════════════ */
async function startAnalysis() {
  const location  = document.getElementById('tripLocation').value.trim();
  const purpose   = document.getElementById('tripPurpose').value.trim();
  const startStr  = document.getElementById('tripStart').value;
  const endStr    = document.getElementById('tripEnd').value;
  const bufBefore = parseInt(document.getElementById('bufferBefore').value) || 0;
  const bufAfter  = parseInt(document.getElementById('bufferAfter').value) || 0;

  if (!location || !startStr || !endStr) {
    showToast('請填寫出差地點、出發與返回時間', 'error'); return;
  }

  const tripStart = new Date(startStr);
  const tripEnd   = new Date(endStr);
  if (tripEnd <= tripStart) {
    showToast('返回時間必須晚於出發時間', 'error'); return;
  }

  state.tripInfo = { location, purpose, tripStart, tripEnd, bufBefore, bufAfter };

  const queryStart = new Date(tripStart.getTime() - bufBefore * 3600000);
  const queryEnd   = new Date(tripEnd.getTime()   + bufAfter  * 3600000);

  await fetchAndClassifyEvents(queryStart, queryEnd);
}

/* ════════════════════════════════════════
   FETCH OUTLOOK CALENDAR EVENTS（Graph API）
   ════════════════════════════════════════ */
async function fetchAndClassifyEvents(queryStart, queryEnd) {
  showLoading('讀取 Outlook 日曆...');
  try {
    // Graph API：讀取主日曆事件
    // calendarView 會展開重複行程，比 events 更準確
    const startIso = queryStart.toISOString();
    const endIso   = queryEnd.toISOString();

    const params = new URLSearchParams({
      startDateTime: startIso,
      endDateTime:   endIso,
      '$top':        '250',
      '$select':     'id,subject,start,end,location,attendees,showAs,isAllDay,recurrence,sensitivity',
      '$orderby':    'start/dateTime',
    });

    const data = await graphRequest(`/me/calendarView?${params}`);
    state.events = (data.value || []).filter(e => e.sensitivity !== 'private' || true);
    // 注意：private 行程仍讀取（因為是本人帳號），但顯示時標記

    // 補充：對每個行程加上 JS Date 物件，方便後續比對
    state.events.forEach(ev => {
      ev._evStart = new Date(ev.start?.dateTime
        ? ev.start.dateTime + (ev.start.timeZone === 'UTC' ? 'Z' : '')
        : ev.start?.date);
      ev._evEnd = new Date(ev.end?.dateTime
        ? ev.end.dateTime + (ev.end.timeZone === 'UTC' ? 'Z' : '')
        : ev.end?.date);
    });

    classifyEvents();
    buildFreeSlots(queryStart, queryEnd);
    renderResults();
    showSection('result');
  } catch(e) {
    console.error(e);
    showToast('讀取日曆失敗：' + (e.message || '請確認已登入且授權日曆存取'), 'error');
  } finally {
    hideLoading();
  }
}

/* ════════════════════════════════════════
   CLASSIFY EVENTS
   ════════════════════════════════════════ */
function classifyEvents() {
  const { tripStart, tripEnd, bufBefore, bufAfter } = state.tripInfo;
  const effectiveStart = new Date(tripStart.getTime() - bufBefore * 3600000);
  const effectiveEnd   = new Date(tripEnd.getTime()   + bufAfter  * 3600000);

  state.classified = { conflict: [], gray: [], movable: [], freeSlots: [] };
  state.grayUnlocked.clear();

  for (const ev of state.events) {
    const evStart = ev._evStart;
    const evEnd   = ev._evEnd;
    const overlaps = evStart < effectiveEnd && evEnd > effectiveStart;
    if (!overlaps) continue;

    const classification = classifyEvent(ev);
    ev._classification = classification;
    ev._conflictLevel  = getConflictLevel(evStart, evEnd, tripStart, tripEnd, bufBefore, bufAfter);

    if (classification === 'locked')   state.classified.conflict.push(ev);
    else if (classification === 'gray')    state.classified.gray.push(ev);
    else if (classification === 'movable') state.classified.movable.push(ev);
  }
}

function classifyEvent(ev) {
  const title     = ev.subject || '';
  // Graph API：showAs = 'free' | 'tentative' | 'busy' | 'oof' | 'workingElsewhere' | 'unknown'
  const showAs    = ev.showAs || 'busy';
  const attendees = ev.attendees || [];

  // 可移動條件
  if (showAs === 'free') return 'movable';
  if (showAs === 'tentative') return 'movable';
  if (isKeywordMatch(title, state.movableKeywords)) return 'movable';
  if (attendees.length === 0 && !isKeywordMatch(title, state.lockedKeywords)) return 'movable';

  // 不可移動條件
  if (isKeywordMatch(title, state.lockedKeywords)) return 'locked';
  const hasAcceptedExternal = attendees.some(a =>
    a.type !== 'required' || // organizer 以外
    (a.emailAddress?.address !== CONFIG.ACCOUNT && a.status?.response === 'accepted')
  );
  if (hasAcceptedExternal) return 'locked';

  // 灰色地帶
  return 'gray';
}

function isKeywordMatch(title, keywords) {
  return keywords.some(kw => title.toLowerCase().includes(kw.toLowerCase()));
}

function getConflictLevel(evStart, evEnd, tripStart, tripEnd) {
  const fullConflict = evStart < tripEnd && evEnd > tripStart;
  if (!fullConflict) return 'buffer';
  if (evStart >= tripStart && evEnd <= tripEnd) return 'full';
  return 'partial';
}

/* ════════════════════════════════════════
   FREE SLOTS
   ════════════════════════════════════════ */
function buildFreeSlots(queryStart, queryEnd) {
  const { tripStart, tripEnd } = state.tripInfo;
  const allEvents = state.events.filter(e => e._classification !== 'movable');
  const slots = [];

  let cur = new Date(tripStart);
  cur.setHours(0, 0, 0, 0);
  const end = new Date(tripEnd);
  end.setHours(23, 59, 59, 999);

  while (cur <= end) {
    const dayStart = new Date(cur); dayStart.setHours(8,  0, 0, 0);
    const dayEnd   = new Date(cur); dayEnd.setHours(18, 0, 0, 0);

    const eventsThisDay = allEvents.filter(e =>
      e._evStart < dayEnd && e._evEnd > dayStart
    );

    if (eventsThisDay.length === 0) {
      slots.push({ date: new Date(cur), quality: 'full', label: '完整空白天' });
    } else {
      const noon = new Date(cur); noon.setHours(12, 0, 0, 0);
      const morningBusy   = eventsThisDay.some(e => e._evStart < noon);
      const afternoonBusy = eventsThisDay.some(e => e._evEnd   > noon);
      if (!morningBusy)   slots.push({ date: new Date(cur), quality: 'morning',   label: '上午空檔' });
      if (!afternoonBusy) slots.push({ date: new Date(cur), quality: 'afternoon', label: '下午空檔' });
    }

    cur.setDate(cur.getDate() + 1);
  }

  const order = { full:0, morning:1, afternoon:2 };
  state.classified.freeSlots = slots.sort((a,b) => order[a.quality] - order[b.quality]);
}

/* ════════════════════════════════════════
   RENDER RESULTS
   ════════════════════════════════════════ */
function renderResults() {
  renderTripSummary();
  renderTimeline();
  renderConflictList();
  renderGrayList();
  renderMovableList();
  renderFreeList();
}

function renderTripSummary() {
  const { location, purpose, tripStart, tripEnd } = state.tripInfo;
  const days = Math.ceil((tripEnd - tripStart) / 86400000);
  document.getElementById('tripSummary').innerHTML = `
    <div class="summary-item">
      <span class="summary-label">出差地點</span>
      <span class="summary-value">📍 ${location}</span>
    </div>
    <div class="summary-item">
      <span class="summary-label">出發</span>
      <span class="summary-value">${formatDatetime(tripStart)}</span>
    </div>
    <div class="summary-item">
      <span class="summary-label">返回</span>
      <span class="summary-value">${formatDatetime(tripEnd)}</span>
    </div>
    <div class="summary-item">
      <span class="summary-label">天數</span>
      <span class="summary-value">${days} 天</span>
    </div>
    ${purpose ? `<div class="summary-item"><span class="summary-label">目的</span><span class="summary-value">${purpose}</span></div>` : ''}
  `;
}

function renderTimeline() {
  const { tripStart, tripEnd } = state.tripInfo;
  const container = document.getElementById('timelineContainer');
  container.innerHTML = '';

  let cur = new Date(tripStart);
  cur.setHours(0, 0, 0, 0);
  const end = new Date(tripEnd);
  end.setHours(23, 59, 59, 999);

  while (cur <= end) {
    const dayKey    = cur.toDateString();
    const dayEvents = state.events.filter(e => e._evStart?.toDateString() === dayKey);

    const dayEl = document.createElement('div');
    dayEl.className = 'timeline-day';
    dayEl.innerHTML = `
      <div class="timeline-day-header is-trip">
        ${cur.getMonth()+1}/${cur.getDate()}<br>
        <small>${['日','一','二','三','四','五','六'][cur.getDay()]}</small>
      </div>
      <div class="timeline-events">
        ${dayEvents.length === 0
          ? `<div class="timeline-event ev-free">空檔</div>`
          : dayEvents.map(e => {
              const cls = e._classification === 'locked'   ? 'ev-conflict'
                        : e._classification === 'movable'  ? 'ev-movable'
                        : 'ev-gray';
              const title = e.subject || '(無標題)';
              return `<div class="timeline-event ${cls}" title="${title}">${title}</div>`;
            }).join('')}
      </div>`;
    container.appendChild(dayEl);
    cur.setDate(cur.getDate() + 1);
  }
}

function renderConflictList() {
  const el = document.getElementById('conflictList');
  const section = document.getElementById('conflictSection');
  const items = state.classified.conflict;
  if (items.length === 0) { section.style.display = 'none'; return; }
  section.style.display = '';
  el.innerHTML = items.map(ev => {
    const badge = ev._conflictLevel === 'full'    ? '<span class="event-badge badge-conflict">完全衝突</span>'
                : ev._conflictLevel === 'partial'  ? '<span class="event-badge badge-partial">部分重疊</span>'
                : '<span class="event-badge badge-partial">緩衝時間內</span>';
    return `
      <div class="event-card conflict">
        <div class="event-info">
          <div class="event-title">🔴 ${ev.subject || '(無標題)'}</div>
          <div class="event-time">${formatEventTime(ev)}</div>
        </div>
        ${badge}
      </div>`;
  }).join('');
}

function renderGrayList() {
  const el = document.getElementById('grayList');
  const section = document.getElementById('graySection');
  const items = state.classified.gray;
  if (items.length === 0) { section.style.display = 'none'; return; }
  section.style.display = '';
  renderGrayItems(el, items);
}

function renderGrayItems(el, items) {
  el.innerHTML = items.map(ev => {
    const unlocked = state.grayUnlocked.has(ev.id);
    return `
      <div class="event-card gray" data-id="${ev.id}">
        <div class="event-info">
          <div class="event-title">🔒 ${ev.subject || '(無標題)'}</div>
          <div class="event-time">${formatEventTime(ev)}</div>
        </div>
        <div class="event-actions">
          ${unlocked
            ? `<span class="event-badge" style="background:var(--green);color:#fff">已解鎖</span>
               <button class="btn btn-sm btn-success" onclick="openWriteModal('${ev.id}')">修改</button>`
            : `<span class="event-badge badge-locked">不可移動</span>
               <button class="btn btn-sm btn-ghost" onclick="unlockGray('${ev.id}')">解鎖</button>`}
        </div>
      </div>`;
  }).join('');
}

function unlockGray(id) {
  state.grayUnlocked.add(id);
  renderGrayItems(document.getElementById('grayList'), state.classified.gray);
}

function renderMovableList() {
  const el = document.getElementById('movableList');
  const section = document.getElementById('movableSection');
  const items = state.classified.movable;
  if (items.length === 0) { section.style.display = 'none'; return; }
  section.style.display = '';
  el.innerHTML = items.map(ev => `
    <div class="event-card movable">
      <div class="event-info">
        <div class="event-title">🟡 ${ev.subject || '(無標題)'}</div>
        <div class="event-time">${formatEventTime(ev)}</div>
      </div>
      <div class="event-actions">
        <span class="event-badge badge-movable">可移動</span>
        <button class="btn btn-sm btn-secondary" onclick="openWriteModal('${ev.id}')">修改</button>
      </div>
    </div>`).join('');
}

function renderFreeList() {
  const el = document.getElementById('freeList');
  const section = document.getElementById('freeSection');
  const items = state.classified.freeSlots;
  if (items.length === 0) {
    section.style.display = '';
    el.innerHTML = `<div class="empty-state">出差區間內無完整空檔，建議重新選擇出差日期。</div>`;
    return;
  }
  section.style.display = '';
  const qualityIcon = { full:'🟢', morning:'🔵', afternoon:'🔵' };
  el.innerHTML = items.map((slot, i) => `
    <div class="event-card free-slot">
      <div class="event-info">
        <div class="event-title">${qualityIcon[slot.quality]} ${formatDate(slot.date)}（${slot.label}）</div>
        <div class="event-time">${slot.quality === 'full' ? '08:00 – 18:00 完整可用' : slot.quality === 'morning' ? '08:00 – 12:00' : '13:00 – 18:00'}</div>
      </div>
      <div class="event-actions">
        <button class="btn btn-sm btn-success" onclick="openInsertModal(${i})">填入出差行程</button>
      </div>
    </div>`).join('');
}

/* ════════════════════════════════════════
   MODAL：修改行程 / 填入行程
   ════════════════════════════════════════ */
function openWriteModal(evId) {
  const ev = state.events.find(e => e.id === evId);
  if (!ev) return;

  document.getElementById('modalTitle').textContent       = '修改行程';
  document.getElementById('modalEventTitle').value        = ev.subject || '';
  document.getElementById('modalEventStart').value        = toDatetimeLocal(ev._evStart);
  document.getElementById('modalEventEnd').value          = toDatetimeLocal(ev._evEnd);
  document.getElementById('modalEventLocation').value     = ev.location?.displayName || '';
  document.getElementById('modalNote').innerHTML          =
    `⚠️ 此操作將覆蓋您 Outlook 日曆（${CONFIG.ACCOUNT}）中的原始行程。請確認後再執行。`;

  state.currentModal = { type: 'update', eventId: evId };
  document.getElementById('modalOverlay').style.display = 'flex';
}

function openInsertModal(slotIndex) {
  const slot = state.classified.freeSlots[slotIndex];
  const { location, purpose } = state.tripInfo;

  const slotStart = new Date(slot.date);
  slotStart.setHours(slot.quality === 'afternoon' ? 13 : 8, 0, 0, 0);
  const slotEnd = new Date(slot.date);
  slotEnd.setHours(slot.quality === 'morning' ? 12 : 18, 0, 0, 0);

  document.getElementById('modalTitle').textContent       = '填入出差行程';
  document.getElementById('modalEventTitle').value        = `【出差】${location}${purpose ? '─'+purpose : ''}`;
  document.getElementById('modalEventStart').value        = toDatetimeLocal(slotStart);
  document.getElementById('modalEventEnd').value          = toDatetimeLocal(slotEnd);
  document.getElementById('modalEventLocation').value     = location;
  document.getElementById('modalNote').innerHTML          =
    `✅ 此行程將寫入您的 Outlook 日曆（${CONFIG.ACCOUNT}）。`;

  state.currentModal = { type: 'insert', slotIndex };
  document.getElementById('modalOverlay').style.display = 'flex';
}

function closeModal() {
  document.getElementById('modalOverlay').style.display = 'none';
  state.currentModal = null;
}

async function confirmWriteToCalendar() {
  const m = state.currentModal;
  if (!m) return;

  const title     = document.getElementById('modalEventTitle').value.trim();
  const startStr  = document.getElementById('modalEventStart').value;
  const endStr    = document.getElementById('modalEventEnd').value;
  const location  = document.getElementById('modalEventLocation').value.trim();

  if (!title || !startStr || !endStr) {
    showToast('請填寫完整資訊', 'error'); return;
  }
  const startDt = new Date(startStr);
  const endDt   = new Date(endStr);
  if (endDt <= startDt) { showToast('結束時間必須晚於開始時間', 'error'); return; }

  // Graph API 行程資源格式
  const resource = {
    subject: title,
    start:   { dateTime: startDt.toISOString(), timeZone: 'Asia/Taipei' },
    end:     { dateTime: endDt.toISOString(),   timeZone: 'Asia/Taipei' },
    location: { displayName: location },
    body: {
      contentType: 'text',
      content: `由出差衝突檢查器（Outlook 版）建立 · ${new Date().toLocaleString('zh-TW')}`,
    },
  };

  closeModal();
  showLoading('寫入 Outlook 日曆...');

  try {
    if (m.type === 'update') {
      // PATCH 更新既有行程
      await graphRequest(`/me/events/${m.eventId}`, 'PATCH', resource);
      showToast('✅ 行程已更新至 Outlook 日曆', 'success');
    } else {
      // POST 新增行程
      await graphRequest('/me/events', 'POST', resource);
      showToast('✅ 出差行程已寫入 Outlook 日曆', 'success');
    }

    // ── 擴充預留：日後啟用共用日曆時，取消以下註解 ──
    // if (CONFIG.SHARED_CALENDAR_ENABLED) {
    //   await insertToSharedCalendar(resource);
    // }

    await startAnalysis();
  } catch(e) {
    console.error(e);
    showToast('寫入失敗：' + (e.message || '請確認日曆權限'), 'error');
  } finally {
    hideLoading();
  }
}

/* ── 擴充預留：共用日曆寫入（v1.0 停用）── */
async function insertToSharedCalendar(resource) {
  if (!CONFIG.SHARED_CALENDAR_ENABLED || !CONFIG.SHARED_CALENDAR_ID) return;
  await graphRequest(`/groups/${CONFIG.SHARED_CALENDAR_ID}/events`, 'POST', resource);
}

/* ════════════════════════════════════════
   EXPORT
   ════════════════════════════════════════ */
function copyTextSummary() {
  const { location, purpose, tripStart, tripEnd } = state.tripInfo;
  const lines = [
    '═══ 出差衝突檢查報告（Outlook 版）═══',
    `出差地點：${location}`,
    purpose ? `出差目的：${purpose}` : '',
    `出發時間：${formatDatetime(tripStart)}`,
    `返回時間：${formatDatetime(tripEnd)}`,
    '',
    `▸ 不可移動衝突（${state.classified.conflict.length} 筆）`,
    ...state.classified.conflict.map(e => `  · ${e.subject || '(無標題)'} ${formatEventTime(e)}`),
    '',
    `▸ 待確認行程（${state.classified.gray.length} 筆）`,
    ...state.classified.gray.map(e => `  · ${e.subject || '(無標題)'} ${formatEventTime(e)}`),
    '',
    `▸ 建議空檔`,
    ...state.classified.freeSlots.map(s => `  · ${formatDate(s.date)} ${s.label}`),
    '',
    `報告產生時間：${new Date().toLocaleString('zh-TW')}`,
  ].filter(Boolean).join('\n');

  navigator.clipboard.writeText(lines)
    .then(() => showToast('✅ 摘要已複製至剪貼簿'))
    .catch(() => showToast('複製失敗，請手動選取', 'error'));
}

/* ════════════════════════════════════════
   UTILITIES
   ════════════════════════════════════════ */
function showSection(name) {
  document.getElementById('sectionWelcome').style.display = name === 'welcome' ? '' : 'none';
  document.getElementById('sectionInput').style.display   = name === 'input'   ? '' : 'none';
  document.getElementById('sectionResult').style.display  = name === 'result'  ? '' : 'none';
  if (name === 'result') document.getElementById('sectionResult').scrollIntoView({ behavior: 'smooth' });
}

function showLoading(text = '處理中...') {
  document.getElementById('loadingText').textContent = text;
  document.getElementById('loadingOverlay').style.display = 'flex';
}
function hideLoading() {
  document.getElementById('loadingOverlay').style.display = 'none';
}

function showToast(msg, type = '') {
  const t = document.createElement('div');
  t.className = 'toast' + (type ? ' ' + type : '');
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 3500);
}

function toDatetimeLocal(date) {
  if (!date) return '';
  const d = new Date(date);
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function formatDate(date) {
  return new Date(date).toLocaleDateString('zh-TW', { month:'numeric', day:'numeric', weekday:'short' });
}

function formatDatetime(date) {
  return new Date(date).toLocaleString('zh-TW', { month:'numeric', day:'numeric', weekday:'short', hour:'2-digit', minute:'2-digit' });
}

function formatEventTime(ev) {
  if (!ev._evStart) return '';
  if (ev.isAllDay) return `全天 · ${formatDate(ev._evStart)}`;
  return `${formatDatetime(ev._evStart)} – ${ev._evEnd.toLocaleTimeString('zh-TW', { hour:'2-digit', minute:'2-digit' })}`;
}

/* ── Service Worker 註冊（PWA） ── */
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW 註冊失敗', e));
}
