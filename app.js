// ── CONFIG ──────────────────────────────────────────────────────────
const CLIENT_ID = '469662960124-l8ssq4psn55oupe0t6ouu2laeiol1abv.apps.googleusercontent.com';
const SPREADSHEET_ID = '16J873aq698SxJsFgiWcNJOubn3R3z_5_J8NMlrsuIsA';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';
const RANGE = 'A:G'; // Task ID | Task | Activity Tag | Due Date | Priority | Notes | Done

// ── STATE ───────────────────────────────────────────────────────────
let tokenClient = null;
let accessToken = null;
let tasks = []; // [{row, id, task, tag, due, priority, notes, done}]

// ── STORAGE HELPERS (offline cache + pending write queue) ─────────
const CACHE_KEY = 'tasks_cache';
const QUEUE_KEY = 'tasks_pending_writes';
const TOKEN_KEY = 'tasks_token_cache';

function saveCache(taskList) {
  localStorage.setItem(CACHE_KEY, JSON.stringify(taskList));
}
function loadCache() {
  const raw = localStorage.getItem(CACHE_KEY);
  return raw ? JSON.parse(raw) : [];
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
      fetchTasks();
    }
  });

  const cached = loadValidToken();
  if (cached) {
    accessToken = cached;
    document.getElementById('connect-btn').style.display = 'none';
    setStatus('Connected (cached).');
    if (navigator.onLine) fetchTasks();
  }
}

function connect() {
  tokenClient.requestAccessToken({ prompt: '' });
}

// ── DATA FETCH ──────────────────────────────────────────────────────
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
  const list = tasks.length ? tasks : loadCache();
  const today = todayStr();
  const container = document.getElementById('tasks');
  container.innerHTML = '';

  const todayItems = list.filter(t => t.due === today && !t.done);
  const workItems = list.filter(t => t.tag.trim().startsWith('W') && t.due !== today);
  const personalItems = list.filter(t => t.tag.trim().startsWith('P') && t.due !== today);

  function section(title, items) {
    if (!items.length) return;
    const h = document.createElement('div');
    h.className = 'group-head';
    h.textContent = title;
    container.appendChild(h);

    const order = { High: 0, Medium: 1, Low: 2 };
    items.sort((a, b) => (order[a.priority] ?? 1) - (order[b.priority] ?? 1) || (a.done - b.done));

    items.forEach(t => {
      const row = document.createElement('label');
      row.className = 'item' + (t.done ? ' done' : '');
      row.innerHTML = `
        <input type="checkbox" ${t.done ? 'checked' : ''} />
        <div class="content">
          <div class="label">${t.task}</div>
          <div class="proj">${t.tag}${t.due ? ' · due ' + t.due : ''}</div>
        </div>`;
      row.querySelector('input').addEventListener('change', (e) => toggleDone(t, e.target.checked));
      container.appendChild(row);
    });
  }

  section('Today', todayItems);
  section('Work', workItems);
  section('Personal', personalItems);

  if (!list.length) {
    container.innerHTML = '<p class="empty">No tasks loaded yet. Connect to sync.</p>';
  }
}

// ── INIT ────────────────────────────────────────────────────────────
window.addEventListener('load', () => {
  tasks = loadCache();
  render();
  initAuth();
  document.getElementById('connect-btn').addEventListener('click', connect);
});

window.addEventListener('online', () => {
  setStatus('Back online, syncing…');
  fetchTasks();
});
window.addEventListener('offline', () => setStatus('Offline — showing cached data.'));

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('./service-worker.js'));
}
