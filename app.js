// ── CONFIG ──────────────────────────────────────────────────────────
const CLIENT_ID = '469662960124-l8ssq4psn55oupe0t6ouu2laeiol1abv.apps.googleusercontent.com';
const SPREADSHEET_ID = '16J873aq698SxJsFgiWcNJOubn3R3z_5_J8NMlrsuIsA';
const TAXONOMY_SHEET_ID = '1oQM1alY_nyVpk8LcHsFmEMpEk-jy0b18vf506_ubj9s';
const BOARD_SHEET_ID = '1LPMQEO9DCQ7-CAKtCIDkWTOFZ3mK6e1qBcDfa3lggQQ';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';
const RANGE = 'A:G'; // Task ID | Task | Activity Tag | Due Date | Priority | Notes | Done
const TAXONOMY_RANGE = 'A:E'; // Code | Label | Parent | Domain | Notes
const BOARD_RANGE = 'A:J'; // Code | Name | Domain | Stage | Priority | NextUp | Notes | MilestoneDate | Updated | Show

let activeTab = 'tasks'; // 'tasks' | 'board'
let domainFirst = 'Work'; // 'Work' | 'Personal' — which section renders first
let sortMode = 'priority'; // 'priority' | 'urgency'

// ── STATE ───────────────────────────────────────────────────────────
let tokenClient = null;
let accessToken = null;
let tasks = []; // [{row, id, task, tag, due, priority, notes, done}]
let taxonomy = {}; // Code -> Label lookup
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
function queueWrite(row, value) {
  const q = JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]');
  q.push({ row, value, ts: Date.now() });
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
    taxonomy = loadTaxonomyCache(); // still show cached labels while offline/signed-out
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
    rows.forEach(r => { if (r[0]) map[r[0].trim()] = r[1] || r[0]; });
    taxonomy = map;
    saveTaxonomyCache(map);
  } catch (e) {
    taxonomy = loadTaxonomyCache(); // fall back silently, labels just won't show
  }
}

// Turns "W.1 · C.AFE" or "W.1 / W.3" into "W.1 (CCDRs) · C.AFE (Africa East)" —
// unrecognized fragments (like free-text "KTCCA") pass through unchanged.
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
      done: (r[6] || '').toString().toUpperCase() === 'TRUE'
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

// ── WRITE (checkbox toggle) ────────────────────────────────────────
async function toggleDone(task, checked) {
  task.done = checked;
  saveCache(tasks); // optimistic local update
  render();

  if (accessToken && navigator.onLine) {
    try {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/G${task.row}?valueInputOption=RAW`;
      const res = await fetch(url, {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: [[checked ? 'TRUE' : 'FALSE']] })
      });
      if (!res.ok) throw new Error('HTTP ' + res.status);
      setStatus('Saved.');
    } catch (e) {
      queueWrite(task.row, checked ? 'TRUE' : 'FALSE');
      setStatus('Offline or write failed — queued for later sync.');
    }
  } else {
    queueWrite(task.row, checked ? 'TRUE' : 'FALSE');
    setStatus('Offline — change queued, will sync when connected.');
  }
}

async function flushQueue() {
  const q = getQueue();
  if (!q.length || !accessToken || !navigator.onLine) return;
  setStatus('Syncing ' + q.length + ' queued change(s)…');
  for (const item of q) {
    try {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/G${item.row}?valueInputOption=RAW`;
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

// ── RENDER ──────────────────────────────────────────────────────────
function setStatus(msg) {
  document.getElementById('status').textContent = msg;
}

function todayStr() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function render() {
  if (activeTab === 'board') { renderBoard(); return; }

  const list = tasks.length ? tasks : loadCache();
  const today = todayStr();
  const container = document.getElementById('tasks');
  container.innerHTML = '';

  const todayItems = list.filter(t => t.due === today && !t.done);
  const workItems = list.filter(t => t.tag.trim().startsWith('W') && t.due !== today && !t.done);
  const personalItems = list.filter(t => t.tag.trim().startsWith('P') && t.due !== today && !t.done);

  function sortItems(items) {
    if (sortMode === 'urgency') {
      // Overdue/soonest due date first; blank due dates sink to the bottom.
      return items.sort((a, b) => {
        if (!a.due && !b.due) return 0;
        if (!a.due) return 1;
        if (!b.due) return -1;
        return a.due.localeCompare(b.due);
      });
    }
    const order = { High: 0, Medium: 1, Low: 2 };
    return items.sort((a, b) => (order[a.priority] ?? 1) - (order[b.priority] ?? 1));
  }

  function section(title, items) {
    if (!items.length) return;
    const h = document.createElement('div');
    h.className = 'group-head';
    h.textContent = title;
    container.appendChild(h);

    sortItems(items).forEach(t => {
      const row = document.createElement('label');
      row.className = 'item' + (t.done ? ' done' : '');
      row.innerHTML = `
        <input type="checkbox" ${t.done ? 'checked' : ''} />
        <div class="content">
          <div class="label">${t.task}</div>
          <div class="proj">${labelTag(t.tag)}${t.due ? ' · due ' + t.due : ''}</div>
        </div>`;
      row.querySelector('input').addEventListener('change', (e) => toggleDone(t, e.target.checked));
      container.appendChild(row);
    });
  }

  section('Today', todayItems);
  if (domainFirst === 'Personal') {
    section('Personal', personalItems);
    section('Work', workItems);
  } else {
    section('Work', workItems);
    section('Personal', personalItems);
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

function renderBoard() {
  const container = document.getElementById('tasks');
  container.innerHTML = '';
  const list = board.length ? board : JSON.parse(localStorage.getItem('board_cache') || '[]');
  const visible = list.filter(p => p.show);

  function section(title, domainCode) {
    const items = visible.filter(p => p.domain === domainCode);
    if (!items.length) return;
    const h = document.createElement('div');
    h.className = 'group-head';
    h.textContent = title;
    container.appendChild(h);

    const stageOrder = { Milestone: 0, Active: 1, Scoping: 2, Waiting: 3, 'Wrapping up': 4, Monitor: 5, 'On-hold': 6 };
    items.sort((a, b) => (stageOrder[a.stage] ?? 5) - (stageOrder[b.stage] ?? 5));

    items.forEach(p => {
      const card = document.createElement('div');
      card.className = 'board-card priority-' + (p.priority || 'Medium').toLowerCase();
      const pct = yearProgressPct(p.milestoneDate);
      card.innerHTML = `
        <div class="board-head">
          <span class="board-code">${p.code}${taxonomy[p.code.split(' ')[0]] ? ' · ' + taxonomy[p.code.split(' ')[0]] : ''}</span>
          <span class="board-stage">${p.stage}</span>
        </div>
        <div class="board-name">${p.name}</div>
        <div class="board-next">${p.nextUp}</div>
        ${pct !== null ? `<div class="board-track"><div class="board-marker" style="left:${pct}%"></div></div><div class="board-date">${p.milestoneDate}</div>` : ''}
        <div class="board-notes">${p.notes}</div>`;
      container.appendChild(card);
    });
  }

  section('Work', 'W');
  section('Personal', 'P');

  if (!visible.length) {
    container.innerHTML = '<p class="empty">No board data loaded yet.</p>';
  }
}

// ── INIT ────────────────────────────────────────────────────────────
window.addEventListener('load', () => {
  tasks = loadCache();
  taxonomy = loadTaxonomyCache();
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
  document.getElementById('sort-mode').addEventListener('click', () => {
    sortMode = sortMode === 'priority' ? 'urgency' : 'priority';
    document.getElementById('sort-mode').textContent = sortMode === 'priority' ? 'By priority' : 'By urgency';
    render();
  });
});

function setTab(tab) {
  activeTab = tab;
  document.getElementById('tab-tasks').classList.toggle('active', tab === 'tasks');
  document.getElementById('tab-board').classList.toggle('active', tab === 'board');
  document.getElementById('sort-controls').style.display = tab === 'tasks' ? 'flex' : 'none';
  render();
}

window.addEventListener('online', () => {
  setStatus('Back online, syncing…');
  fetchTasks();
});
window.addEventListener('offline', () => setStatus('Offline — showing cached data.'));

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('./service-worker.js'));
}
