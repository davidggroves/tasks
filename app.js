// ── CONFIG ──────────────────────────────────────────────────────────
const CLIENT_ID = '469662960124-l8ssq4psn55oupe0t6ouu2laeiol1abv.apps.googleusercontent.com';
const APP_VERSION = '2026.08.03.1';
const SPREADSHEET_ID = '16J873aq698SxJsFgiWcNJOubn3R3z_5_J8NMlrsuIsA';
const TAXONOMY_SHEET_ID = '1oQM1alY_nyVpk8LcHsFmEMpEk-jy0b18vf506_ubj9s';
const BOARD_SHEET_ID = '1LPMQEO9DCQ7-CAKtCIDkWTOFZ3mK6e1qBcDfa3lggQQ';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';
const RANGE = 'A:I'; // Task ID | Task | Activity Tag | Due Date | Priority | Notes | Done | Today | Quick Note
const LONG_PRESS_MS = 550;
const TAXONOMY_RANGE = 'A:E'; // Code | Label | Parent | Domain | Notes
const BOARD_RANGE = 'A:J'; // Code | Name | Domain | Stage | Priority | NextUp | Notes | MilestoneDate | Updated | Show

let activeTab = 'tasks'; // 'tasks' | 'board'
let domainFirst = 'Work'; // 'Work' | 'Personal' — which section renders first
let sortMode = 'priority'; // 'category' | 'priority'
let boardMode = 'tracked'; // 'tracked' | 'all'

// ── STATE ───────────────────────────────────────────────────────────
let tokenClient = null;
let accessToken = null;
let tasks = []; // [{row, id, task, tag, due, priority, notes, done}]
let taxonomy = {}; // Code -> Label lookup
let taxonomyParent = {}; // Code -> Parent code lookup (for resolving top-level category)
let board = []; // [{code, name, domain, stage, priority, nextUp, notes, milestoneDate, updated, show}]

// ── STORAGE HELPERS (offline cache + pending write queue) ─────────
const CACHE_KEY = 'tasks_cache';
const TAXONOMY_CACHE_KEY = 'taxonomy_cache';
const QUEUE_KEY = 'tasks_pending_writes';
const TOKEN_KEY = 'tasks_token_cache';

function saveCache(taskList) {
  localStorage.setItem(CACHE_KEY, JSON.stringify(taskList));
}
function loadCache() {
  const raw = localStorage.getItem(CACHE_KEY);
  return raw ? JSON.parse(raw) : [];
}
function saveTaxonomyCache(map) {
  localStorage.setItem(TAXONOMY_CACHE_KEY, JSON.stringify(map));
}
function loadTaxonomyCache() {
  const raw = localStorage.getItem(TAXONOMY_CACHE_KEY);
  return raw ? JSON.parse(raw) : {};
}
function queueWrite(row, column, value) {
  const q = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]');
  q.push({ row, column, value, ts: Date.now() });
  localStorage.setItem(QUEUE_KEY, JSON.stringify(q));
}
function getQueue() {
  return JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]');
}
function clearQueue() {
  localStorage.setItem(QUEUE_KEY, '[]');
}
function saveToken(token, expiresInSec) {
  const expiry = Date.now() + expiresInSec * 1000 - 60000; // 60s safety buffer
  localStorage.setItem(TOKEN_KEY, JSON.stringify({ token, expiry }));
}
function loadValidToken() {
  const raw = localStorage.getItem(TOKEN_KEY);
  if (!raw) return null;
  const { token, expiry } = JSON.parse(raw);
  return Date.now() < expiry ? token : null;
}

// ── GOOGLE AUTH ─────────────────────────────────────────────────────
function initAuth() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPES,
    callback: (resp) => {
      if (resp.error) {
        setStatus('Sign-in failed: ' + resp.error);
        return;
      }
      accessToken = resp.access_token;
      saveToken(accessToken, resp.expires_in);
      setStatus('Connected.');
      document.getElementById('connect-btn').style.display = 'none';
      Promise.all([fetchTaxonomy(), fetchBoard()]).then(fetchTasks);
    }
  });

  const cached = loadValidToken();
  if (cached) {
    accessToken = cached;
    document.getElementById('connect-btn').style.display = 'none';
    setStatus('Connected (cached).');
    if (navigator.onLine) Promise.all([fetchTaxonomy(), fetchBoard()]).then(fetchTasks);
  } else {
    const cachedTax = loadTaxonomyCache();
    taxonomy = cachedTax.labels || {};
    taxonomyParent = cachedTax.parents || {};
  }
}

function connect() {
  tokenClient.requestAccessToken({ prompt: '' });
}

// ── DATA FETCH ──────────────────────────────────────────────────────
async function fetchTaxonomy() {
  if (!accessToken || !navigator.onLine) { taxonomy = loadTaxonomyCache(); return; }
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${TAXONOMY_SHEET_ID}/values/${TAXONOMY_RANGE}`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    const rows = (data.values || []).slice(1); // drop header
    const map = {};
    const parentMap = {};
    rows.forEach(r => {
      if (!r[0]) return;
      const code = r[0].trim();
      map[code] = r[1] || code;
      parentMap[code] = (r[2] || '').trim();
    });
    taxonomy = map;
    taxonomyParent = parentMap;
    saveTaxonomyCache({ labels: map, parents: parentMap });
  } catch (e) {
    const cached = loadTaxonomyCache();
    taxonomy = cached.labels || {};
    taxonomyParent = cached.parents || {};
  }
}

// Turns "W.1 · C.AFE" or "W.1 / W.3" into "W.1 (CCDRs) · C.AFE (Africa East)" —
// unrecognized fragments (like free-text "KTCCA") pass through unchanged.
// Climbs the taxonomy parent chain to find the top-level category under Work/Personal
// (e.g. W.12.1 -> W.12 "WBG Admin"; W.1 -> itself "CCDRs"). Falls back to the raw code
// if it's not a recognized taxonomy entry (e.g. free-text like "KTCCA").
function getCategoryLabel(tag) {
  if (!tag) return 'Uncategorized';
  const primary = tag.split(/[·/]/).map(t => t.trim()).find(t => /^[WP]/.test(t)) || tag.trim();
  let current = primary;
  let guard = 0;
  while (taxonomyParent[current] && taxonomyParent[current] !== 'W' && taxonomyParent[current] !== 'P' && guard < 10) {
    current = taxonomyParent[current];
    guard++;
  }
  return taxonomy[current] || current;
}

function labelTag(tag) {
  if (!tag) return '';
  return tag.split(/[·/]/).map(part => {
    const code = part.trim();
    const label = taxonomy[code];
    return label ? `${code} (${label})` : code;
  }).join(' · ');
}

async function fetchBoard() {
  if (!accessToken || !navigator.onLine) return;
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${BOARD_SHEET_ID}/values/${BOARD_RANGE}`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    const rows = (data.values || []).slice(1);
    board = rows.map(r => ({
      code: r[0] || '',
      name: r[1] || '',
      domain: r[2] || '',
      stage: r[3] || '',
      priority: r[4] || 'Medium',
      nextUp: r[5] || '',
      notes: r[6] || '',
      milestoneDate: r[7] || '',
      updated: r[8] || '',
      show: (r[9] || '').toString().toUpperCase() !== 'FALSE' // default true unless explicitly FALSE
    })).filter(p => p.code);
    localStorage.setItem('board_cache', JSON.stringify(board));
  } catch (e) {
    const raw = localStorage.getItem('board_cache');
    board = raw ? JSON.parse(raw) : [];
  }
}

async function fetchTasks() {
  if (!accessToken || !navigator.onLine) return;
  setStatus('Syncing…');
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${RANGE}`;
    const res = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    const rows = (data.values || []).slice(1); // drop header
    tasks = rows.map((r, i) => ({
      row: i + 2, // actual sheet row number
      id: r[0] || '',
      task: r[1] || '',
      tag: r[2] || '',
      due: r[3] || '',
      priority: r[4] || 'Medium',
      notes: r[5] || '',
      done: (r[6] || '').toString().toUpperCase() === 'TRUE',
      todayFlag: (r[7] || '').toString().trim().toUpperCase(), // 'TRUE' | 'FALSE' | '' (never set)
      note: r[8] || ''
    })).filter(t => t.id);
    saveCache(tasks);
    setStatus('Synced ' + new Date().toLocaleTimeString());
    render();
    flushQueue();
  } catch (e) {
    setStatus('Sync failed, showing cached data: ' + e.message);
    tasks = loadCache();
    render();
  }
}

// ── WRITE (checkbox / today toggle / quick note) ─────────────────────
// `value` may be a boolean (Done/Today columns, written as TRUE/FALSE) or a
// free-text string (Quick Note column, written as-is).
async function writeColumn(task, column, field, value) {
  task[field] = value;
  saveCache(tasks); // optimistic local update
  render();

  const cellValue = (typeof value === 'boolean') ? (value ? 'TRUE' : 'FALSE') : value;

  if (accessToken && navigator.onLine) {
    try {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${column}${task.row}?valueInputOption=RAW`;
      const res = await fetch(url, {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: [[cellValue]] })
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      setStatus('Saved.');
    } catch (e) {
      queueWrite(task.row, column, cellValue);
      setStatus('Offline or write failed — queued for later sync.');
    }
  } else {
    queueWrite(task.row, column, cellValue);
    setStatus('Offline — change queued, will sync when connected.');
  }
}

function toggleDone(task, checked) {
  return writeColumn(task, 'G', 'done', checked);
}

function toggleTodayShared(task, checked) {
  return writeColumn(task, 'H', 'todayFlag', checked ? 'TRUE' : 'FALSE');
}

// Appends a timestamped quick note to column I, preserving whatever's already there
// (including anything Grover itself wrote there via a Grover Tasks push).
function saveQuickNote(task, text) {
  const trimmed = (text || '').trim();
  if (!trimmed) return Promise.resolve();
  const stamp = todayStr();
  const existing = (task.note || '').trim();
  const entry = '[' + stamp + '] ' + trimmed;
  const updated = existing ? existing + '\n' + entry : entry;
  return writeColumn(task, 'I', 'note', updated);
}

async function flushQueue() {
  const q = getQueue();
  if (!q.length || !accessToken || !navigator.onLine) return;
  setStatus('Syncing ' + q.length + ' queued change(s)…');
  for (const item of q) {
    try {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${item.column}${item.row}?valueInputOption=RAW`;
      await fetch(url, {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: [[item.value]] })
      });
    } catch (e) {
      setStatus('Some queued changes failed to sync — will retry next time.');
      return; // stop, keep queue for retry
    }
  }
  clearQueue();
  setStatus('All queued changes synced.');
}

// ── QUICK NOTE MODAL ──────────────────────────────────────────────────
let noteModalTask = null;

function openNoteModal(task) {
  noteModalTask = task;
  document.getElementById('note-modal-title').textContent = task.task;
  document.getElementById('note-modal-existing').textContent = task.note || '';
  document.getElementById('note-modal-existing').style.display = task.note ? 'block' : 'none';
  document.getElementById('note-input').value = '';
  document.getElementById('note-modal').style.display = 'flex';
  setTimeout(() => document.getElementById('note-input').focus(), 50);
}

function closeNoteModal() {
  document.getElementById('note-modal').style.display = 'none';
  noteModalTask = null;
}

// ── RENDER ──────────────────────────────────────────────────────────
function setStatus(msg) {
  document.getElementById('status').textContent = msg;
}

function todayStr() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

// Parses either "YYYY-MM-DD" or "M/D/YYYY" (both appear in the sheet) into a
// sortable "YYYY-MM-DD" string. Returns '' if unparseable/empty.
function normalizeDue(due) {
  if (!due) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(due)) return due;
  const m = due.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) {
    const [, mo, da, yr] = m;
    return yr + '-' + mo.padStart(2, '0') + '-' + da.padStart(2, '0');
  }
  return '';
}

function isOverdue(t) {
  if (t.done) return false;
  const norm = normalizeDue(t.due);
  if (!norm) return false;
  return norm < todayStr();
}

function render() {
  if (activeTab === 'board') { renderBoard(); return; }

  const list = tasks.length ? tasks : loadCache();
  const today = todayStr();
  const container = document.getElementById('tasks');
  container.innerHTML = '';

  const todayItems = list.filter(t => isToday(t));
  const workItems = list.filter(t => t.tag.trim().startsWith('W') && !isToday(t));
  const personalItems = list.filter(t => t.tag.trim().startsWith('P') && !isToday(t));

  function isToday(t) {
    if (t.todayFlag === 'FALSE') return false; // explicit dismissal always wins
    if (t.todayFlag === 'TRUE') return true;   // explicit flag always wins
    return t.due === today;                    // no explicit assertion yet — auto-surface if due today
  }

  function taskRow(t) {
    const row = document.createElement('label');
    row.className = 'item' + (t.done ? ' done' : '');
    row.innerHTML = `
      <input type="checkbox" ${t.done ? 'checked' : ''} />
      <div class="content">
        <div class="label">${t.task}${t.note ? ' <span class="note-flag" title="Has a note">📝</span>' : ''}</div>
        <div class="proj">${labelTag(t.tag)} · <span class="priority-tag">${t.priority || 'Medium'}</span></div>
      </div>
      <div class="due-col${isOverdue(t) ? ' overdue' : ''}">${t.due || ''}</div>
      <div class="today-btn ${isToday(t) ? 'active' : ''}" title="Toggle Today"></div>`;
    row.querySelector('input').addEventListener('change', (e) => toggleDone(t, e.target.checked));
    row.querySelector('.today-btn').addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleTodayShared(t, !isToday(t));
    });

    // Long-press (touch or mouse-hold) opens the quick-note screen instead of
    // toggling Done. We suppress the label's synthetic click on release so the
    // checkbox doesn't also fire.
    let pressTimer = null;
    let longPressed = false;
    const startPress = () => {
      longPressed = false;
      pressTimer = setTimeout(() => {
        longPressed = true;
        if (navigator.vibrate) navigator.vibrate(12);
        openNoteModal(t);
      }, LONG_PRESS_MS);
    };
    const cancelPress = () => clearTimeout(pressTimer);
    row.addEventListener('touchstart', startPress, { passive: true });
    row.addEventListener('touchend', cancelPress);
    row.addEventListener('touchmove', cancelPress);
    row.addEventListener('mousedown', startPress);
    row.addEventListener('mouseup', cancelPress);
    row.addEventListener('mouseleave', cancelPress);
    row.addEventListener('click', (e) => {
      if (longPressed) {
        e.preventDefault(); // stop the label's synthetic checkbox click
        longPressed = false;
      }
    });
    return row;
  }

  function subhead(text) {
    const h = document.createElement('div');
    h.className = 'subhead';
    h.textContent = text;
    return h;
  }

  function domainHead(text) {
    const h = document.createElement('div');
    h.className = 'group-head';
    h.textContent = text;
    return h;
  }

  function byDateAscending(items) {
    return [...items].sort((a, b) => {
      const da = normalizeDue(a.due), db = normalizeDue(b.due);
      if (!da && !db) return 0;
      if (!da) return 1;
      if (!db) return -1;
      return da.localeCompare(db);
    });
  }

  function byPriorityThenDate(items) {
    const order = { High: 0, Medium: 1, Low: 2 };
    return byDateAscending(items).sort((a, b) => (order[a.priority] ?? 1) - (order[b.priority] ?? 1));
  }

  function renderGroupedItems(items) {
    if (sortMode === 'category') {
      const groups = {};
      items.forEach(t => {
        const cat = getCategoryLabel(t.tag);
        (groups[cat] = groups[cat] || []).push(t);
      });
      Object.keys(groups).sort().forEach(cat => {
        container.appendChild(subhead(cat));
        byPriorityThenDate(groups[cat]).forEach(t => container.appendChild(taskRow(t)));
      });
    } else { // priority
      const tiers = ['High', 'Medium', 'Low'];
      tiers.forEach(tier => {
        const inTier = items.filter(t => (t.priority || 'Medium') === tier);
        if (!inTier.length) return;
        container.appendChild(subhead(tier));
        byDateAscending(inTier).forEach(t => container.appendChild(taskRow(t)));
      });
    }
  }

  function renderDomainSection(title, items) {
    if (!items.length) return;
    container.appendChild(domainHead(title));
    renderGroupedItems(items);
  }

  if (todayItems.length) {
    container.appendChild(domainHead('Today'));
    renderGroupedItems(todayItems); // now respects the By Category / By Priority toggle, same as Work/Personal
  }

  if (domainFirst === 'Personal') {
    renderDomainSection('Personal', personalItems);
    renderDomainSection('Work', workItems);
  } else {
    renderDomainSection('Work', workItems);
    renderDomainSection('Personal', personalItems);
  }

  if (!list.length) {
    container.innerHTML = '<p class="empty">No tasks loaded yet. Connect to sync.</p>';
  }
}

// ── BOARD VIEW ──────────────────────────────────────────────────────
function yearProgressPct(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr);
  const start = new Date(d.getFullYear(), 0, 1);
  const end = new Date(d.getFullYear(), 11, 31);
  return Math.round(((d - start) / (end - start)) * 100);
}

function getTopLevelCodes(domainRoot) {
  return Object.keys(taxonomyParent).filter(code => taxonomyParent[code] === domainRoot).sort();
}

function renderBoard() {
  const container = document.getElementById('tasks');
  container.innerHTML = '';
  const list = board.length ? board : JSON.parse(localStorage.getItem('board_cache') || '[]');
  const boardByCode = {};
  list.forEach(p => { boardByCode[p.code] = p; });

  function boardCard(p) {
    const card = document.createElement('div');
    card.className = 'board-card priority-' + (p.priority || 'Medium').toLowerCase()
      + (p.stage === 'Untracked' ? ' untracked' : '');
    const pct = yearProgressPct(p.milestoneDate);
    card.innerHTML = `
      <div class="board-head">
        <span class="board-code">${p.code}${taxonomy[p.code.split(' ')[0]] ? ' · ' + taxonomy[p.code.split(' ')[0]] : ''}</span>
        <span class="board-stage">${p.stage}</span>
      </div>
      <div class="board-name">${p.name}</div>
      ${p.nextUp ? `<div class="board-next">${p.nextUp}</div>` : ''}
      ${pct !== null ? `<div class="board-track"><div class="board-marker" style="left:${pct}%"></div></div><div class="board-date">${p.milestoneDate}</div>` : ''}
      ${p.notes ? `<div class="board-notes">${p.notes}</div>` : ''}`;
    return card;
  }

  function section(title, domainCode) {
    let items;
    if (boardMode === 'all') {
      const codes = getTopLevelCodes(domainCode);
      items = codes.map(code => boardByCode[code] || {
        code, name: taxonomy[code] || code, domain: domainCode, stage: 'Untracked',
        priority: 'Low', nextUp: '', notes: '', milestoneDate: '', updated: '', show: true
      });
    } else {
      items = list.filter(p => p.domain === domainCode && p.show);
    }
    if (!items.length) return;

    const h = document.createElement('div');
    h.className = 'group-head';
    h.textContent = title;
    container.appendChild(h);

    const stageOrder = { Milestone: 0, Active: 1, Scoping: 2, Waiting: 3, 'Wrapping up': 4, Monitor: 5, 'On-hold': 6, Untracked: 7 };
    items.sort((a, b) =>
      (stageOrder[a.stage] ?? 5) - (stageOrder[b.stage] ?? 5) ||
      a.code.localeCompare(b.code, undefined, { numeric: true })
    );
    items.forEach(p => container.appendChild(boardCard(p)));
  }

  section('Work', 'W');
  section('Personal', 'P');

  const anyContent = boardMode === 'all'
    ? (getTopLevelCodes('W').length || getTopLevelCodes('P').length)
    : list.some(p => p.show);
  if (!anyContent) {
    container.innerHTML = '<p class="empty">No board data loaded yet.</p>';
  }
}

// ── INIT ────────────────────────────────────────────────────────────
window.addEventListener('load', () => {
  document.title = 'Tasks (' + APP_VERSION + ')';
  const versionTag = document.createElement('div');
  versionTag.id = 'version-tag';
  versionTag.textContent = 'App code: ' + APP_VERSION;
  document.body.insertBefore(versionTag, document.getElementById('status'));

  // Weekend default: Personal shows first on Sat/Sun unless user toggles it during the session.
  const dayOfWeek = new Date().getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    domainFirst = 'Personal';
    document.getElementById('sort-domain').textContent = 'Personal first';
  }

  tasks = loadCache();
  const cachedTax = loadTaxonomyCache();
  taxonomy = cachedTax.labels || {};
  taxonomyParent = cachedTax.parents || {};
  board = JSON.parse(localStorage.getItem('board_cache') || '[]');
  render();
  initAuth();
  document.getElementById('connect-btn').addEventListener('click', connect);

  document.getElementById('tab-tasks').addEventListener('click', () => setTab('tasks'));
  document.getElementById('tab-board').addEventListener('click', () => setTab('board'));
  document.getElementById('sort-domain').addEventListener('click', () => {
    domainFirst = domainFirst === 'Work' ? 'Personal' : 'Work';
    document.getElementById('sort-domain').textContent = domainFirst + ' first';
    render();
  });

  ['category', 'priority'].forEach(mode => {
    document.getElementById('sort-' + mode).addEventListener('click', () => {
      sortMode = mode;
      ['category', 'priority'].forEach(m =>
        document.getElementById('sort-' + m).classList.toggle('active', m === mode)
      );
      render();
    });
  });

  ['tracked', 'all'].forEach(mode => {
    document.getElementById('board-' + mode).addEventListener('click', () => {
      boardMode = mode;
      ['tracked', 'all'].forEach(m =>
        document.getElementById('board-' + m).classList.toggle('active', m === mode)
      );
      render();
    });
  });

  document.getElementById('note-cancel').addEventListener('click', closeNoteModal);
  document.getElementById('note-modal').addEventListener('click', (e) => {
    if (e.target.id === 'note-modal') closeNoteModal(); // tap outside the box
  });
  document.getElementById('note-save').addEventListener('click', () => {
    const text = document.getElementById('note-input').value;
    const task = noteModalTask;
    closeNoteModal();
    if (task) saveQuickNote(task, text);
  });
});

function setTab(tab) {
  activeTab = tab;
  document.getElementById('tab-tasks').classList.toggle('active', tab === 'tasks');
  document.getElementById('tab-board').classList.toggle('active', tab === 'board');
  document.getElementById('sort-controls').style.display = tab === 'tasks' ? 'flex' : 'none';
  document.getElementById('board-controls').style.display = tab === 'board' ? 'flex' : 'none';
  render();
}

window.addEventListener('online', () => {
  setStatus('Back online, syncing…');
  fetchTasks();
});
window.addEventListener('offline', () => setStatus('Offline — showing cached data.'));

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./service-worker.js');
    let reloaded = false;
    navigator.serviceWorker.addEventListener('controllerchange', () => {
      if (reloaded) return;
      reloaded = true;
      window.location.reload();
    });
  });
}
