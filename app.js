/**
 * Trip Conflict Checker — app.js  v2.0（Outlook ICS 版）
 *
 * 版本紀錄：
 * v1.0  MSAL + Graph API，需管理員同意，無法使用。
 * v1.1  降權 Calendars.Read，仍被組織政策攔截。
 * v2.0  改用 ICS 訂閱連結，完全繞過 OAuth，不需任何登入。
 *        只讀取忙碌時段（不含行程詳細資料），符合組織資安規範。
 */

const CONFIG = {
  // Outlook 行事曆 ICS 訂閱連結
  // 來源：Outlook 網頁版 → 設定 → 行事曆 → 共用行事曆 → 發佈行事曆 → ICS 連結
  // 權限：可以檢視我忙碌的時間（不含行程詳細資料）
  ICS_URL: 'https://outlook.office365.com/owa/calendar/b20c984f351b476188f50db2d8f9b125@csd.org.tw/209834b43630471ebb654af7fd0fce7516556205859713450498/calendar.ics',

  // CORS Proxy（ICS 需要透過 proxy 讀取，避免瀏覽器跨來源限制）
  CORS_PROXY: 'https://corsproxy.io/?',
};

/* ════════════════════════════════════════
   STATE
   ════════════════════════════════════════ */
let state = {
  tripInfo:   null,
  busySlots:  [],   // 忙碌時段陣列 {start, end}
  freeSlots:  [],   // 空檔陣列
};

/* ════════════════════════════════════════
   INIT
   ════════════════════════════════════════ */
document.addEventListener('DOMContentLoaded', () => {
  bindUIEvents();
  setDefaultDatetimes();
});

/* ════════════════════════════════════════
   UI EVENTS
   ════════════════════════════════════════ */
function bindUIEvents() {
  document.getElementById('btnAnalyze').addEventListener('click', startAnalysis);
  document.getElementById('btnBackToInput').addEventListener('click', () => {
    document.getElementById('sectionResult').style.display = 'none';
    document.getElementById('sectionInput').style.display = '';
  });
  document.getElementById('btnCopyText').addEventListener('click', copyTextSummary);
  document.getElementById('btnExportPDF').addEventListener('click', () => window.print());
}

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

  await fetchAndAnalyzeICS();
}

/* ════════════════════════════════════════
   FETCH ICS
   ════════════════════════════════════════ */
async function fetchAndAnalyzeICS() {
  showLoading('讀取 Outlook 行事曆...');
  try {
    const url = CONFIG.CORS_PROXY + encodeURIComponent(CONFIG.ICS_URL);
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const icsText = await resp.text();
    parseICS(icsText);
    buildFreeSlots();
    renderResults();
    document.getElementById('sectionInput').style.display = 'none';
    document.getElementById('sectionResult').style.display = '';
    document.getElementById('sectionResult').scrollIntoView({ behavior: 'smooth' });
  } catch(e) {
    console.error(e);
    showToast('讀取行事曆失敗：' + (e.message || '請確認網路連線'), 'error');
  } finally {
    hideLoading();
  }
}

/* ════════════════════════════════════════
   ICS PARSER
   ════════════════════════════════════════ */
function parseICS(text) {
  const { tripStart, tripEnd, bufBefore, bufAfter } = state.tripInfo;
  const effectiveStart = new Date(tripStart.getTime() - bufBefore * 3600000);
  const effectiveEnd   = new Date(tripEnd.getTime()   + bufAfter  * 3600000);

  state.busySlots = [];

  // 分割各個 VEVENT
  const events = text.split('BEGIN:VEVENT').slice(1);

  events.forEach(ev => {
    // 解析 DTSTART 與 DTEND
    const startMatch = ev.match(/DTSTART(?:;[^:]*)?:([^\r\n]+)/);
    const endMatch   = ev.match(/DTEND(?:;[^:]*)?:([^\r\n]+)/);
    if (!startMatch || !endMatch) return;

    const evStart = parseICSDate(startMatch[1].trim());
    const evEnd   = parseICSDate(endMatch[1].trim());
    if (!evStart || !evEnd) return;

    // 過濾出在出差區間（含緩衝）內的行程
    if (evStart < effectiveEnd && evEnd > effectiveStart) {
      state.busySlots.push({ start: evStart, end: evEnd });
    }
  });

  // 依時間排序
  state.busySlots.sort((a, b) => a.start - b.start);

  // 合併重疊時段
  state.busySlots = mergeBusySlots(state.busySlots);
}

function parseICSDate(str) {
  try {
    // 格式：20240617T090000Z 或 20240617T090000 或 20240617
    if (str.length === 8) {
      // 全天行程 YYYYMMDD
      return new Date(
        parseInt(str.slice(0,4)),
        parseInt(str.slice(4,6)) - 1,
        parseInt(str.slice(6,8))
      );
    }
    // 標準格式
    const y  = str.slice(0,4);
    const mo = str.slice(4,6);
    const d  = str.slice(6,8);
    const h  = str.slice(9,11)  || '00';
    const mi = str.slice(11,13) || '00';
    const s  = str.slice(13,15) || '00';
    const isUTC = str.endsWith('Z');
    const iso = `${y}-${mo}-${d}T${h}:${mi}:${s}${isUTC ? 'Z' : '+08:00'}`;
    return new Date(iso);
  } catch(e) {
    return null;
  }
}

function mergeBusySlots(slots) {
  if (slots.length === 0) return [];
  const merged = [{ ...slots[0] }];
  for (let i = 1; i < slots.length; i++) {
    const last = merged[merged.length - 1];
    if (slots[i].start <= last.end) {
      last.end = new Date(Math.max(last.end, slots[i].end));
    } else {
      merged.push({ ...slots[i] });
    }
  }
  return merged;
}

/* ════════════════════════════════════════
   FREE SLOTS
   ════════════════════════════════════════ */
function buildFreeSlots() {
  const { tripStart, tripEnd } = state.tripInfo;
  state.freeSlots = [];

  let cur = new Date(tripStart);
  cur.setHours(0, 0, 0, 0);
  const end = new Date(tripEnd);
  end.setHours(23, 59, 59, 999);

  while (cur <= end) {
    const dayStart = new Date(cur); dayStart.setHours(8,  0, 0, 0);
    const dayEnd   = new Date(cur); dayEnd.setHours(18, 0, 0, 0);
    const noon     = new Date(cur); noon.setHours(12, 0, 0, 0);

    const busyThisDay = state.busySlots.filter(s =>
      s.start < dayEnd && s.end > dayStart
    );

    if (busyThisDay.length === 0) {
      state.freeSlots.push({ date: new Date(cur), quality: 'full', label: '完整空白天' });
    } else {
      const morningBusy   = busyThisDay.some(s => s.start < noon);
      const afternoonBusy = busyThisDay.some(s => s.end   > noon);
      if (!morningBusy)   state.freeSlots.push({ date: new Date(cur), quality: 'morning',   label: '上午空檔' });
      if (!afternoonBusy) state.freeSlots.push({ date: new Date(cur), quality: 'afternoon', label: '下午空檔' });
    }

    cur.setDate(cur.getDate() + 1);
  }

  const order = { full:0, morning:1, afternoon:2 };
  state.freeSlots.sort((a,b) => order[a.quality] - order[b.quality]);
}

/* ════════════════════════════════════════
   RENDER RESULTS
   ════════════════════════════════════════ */
function renderResults() {
  renderTripSummary();
  renderTimeline();
  renderConflictList();
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
    <div class="summary-item">
      <span class="summary-label">忙碌時段</span>
      <span class="summary-value">${state.busySlots.length} 筆</span>
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
    const dayStart = new Date(cur); dayStart.setHours(0,  0, 0, 0);
    const dayEnd   = new Date(cur); dayEnd.setHours(23, 59, 59, 999);
    const busyThisDay = state.busySlots.filter(s => s.start < dayEnd && s.end > dayStart);

    const dayEl = document.createElement('div');
    dayEl.className = 'timeline-day';
    dayEl.innerHTML = `
      <div class="timeline-day-header is-trip">
        ${cur.getMonth()+1}/${cur.getDate()}<br>
        <small>${['日','一','二','三','四','五','六'][cur.getDay()]}</small>
      </div>
      <div class="timeline-events">
        ${busyThisDay.length === 0
          ? `<div class="timeline-event ev-free">空檔</div>`
          : busyThisDay.map(s => {
              const startStr = s.start.toLocaleTimeString('zh-TW', { hour:'2-digit', minute:'2-digit' });
              const endStr   = s.end.toLocaleTimeString('zh-TW',   { hour:'2-digit', minute:'2-digit' });
              return `<div class="timeline-event ev-conflict" title="${startStr}–${endStr}">忙碌</div>`;
            }).join('')}
      </div>`;
    container.appendChild(dayEl);
    cur.setDate(cur.getDate() + 1);
  }
}

function renderConflictList() {
  const el = document.getElementById('conflictList');
  const section = document.getElementById('conflictSection');

  if (state.busySlots.length === 0) {
    section.style.display = '';
    el.innerHTML = `<div class="empty-state">🎉 出差區間內無忙碌時段，時間完全空白！</div>`;
    return;
  }

  section.style.display = '';
  el.innerHTML = state.busySlots.map(slot => `
    <div class="event-card conflict">
      <div class="event-info">
        <div class="event-title">🔴 忙碌時段</div>
        <div class="event-time">${formatDatetime(slot.start)} – ${slot.end.toLocaleTimeString('zh-TW', { hour:'2-digit', minute:'2-digit' })}</div>
      </div>
      <span class="event-badge badge-conflict">有行程</span>
    </div>`).join('');
}

function renderFreeList() {
  const el = document.getElementById('freeList');
  const section = document.getElementById('freeSection');

  if (state.freeSlots.length === 0) {
    section.style.display = '';
    el.innerHTML = `<div class="empty-state">出差區間內無完整空檔，建議重新選擇出差日期。</div>`;
    return;
  }

  section.style.display = '';
  const qualityIcon = { full:'🟢', morning:'🔵', afternoon:'🔵' };
  el.innerHTML = state.freeSlots.map(slot => `
    <div class="event-card free-slot">
      <div class="event-info">
        <div class="event-title">${qualityIcon[slot.quality]} ${formatDate(slot.date)}（${slot.label}）</div>
        <div class="event-time">${slot.quality === 'full' ? '08:00 – 18:00 完整可用' : slot.quality === 'morning' ? '08:00 – 12:00' : '13:00 – 18:00'}</div>
      </div>
    </div>`).join('');
}

/* ════════════════════════════════════════
   EXPORT
   ════════════════════════════════════════ */
function copyTextSummary() {
  const { location, purpose, tripStart, tripEnd } = state.tripInfo;
  const lines = [
    '═══ 出差衝突檢查報告（Outlook ICS 版）═══',
    `出差地點：${location}`,
    purpose ? `出差目的：${purpose}` : '',
    `出發時間：${formatDatetime(tripStart)}`,
    `返回時間：${formatDatetime(tripEnd)}`,
    '',
    `▸ 忙碌時段（${state.busySlots.length} 筆）`,
    ...state.busySlots.map(s =>
      `  · ${formatDatetime(s.start)} – ${s.end.toLocaleTimeString('zh-TW', { hour:'2-digit', minute:'2-digit' })}`
    ),
    '',
    `▸ 建議空檔`,
    ...state.freeSlots.map(s => `  · ${formatDate(s.date)} ${s.label}`),
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

/* ── Service Worker ── */
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW 註冊失敗', e));
}
