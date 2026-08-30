// ── CONFIG ──────────────────────────────────────────────────────────
const CLIENT_ID = '469662960124-l8ssq4psn55oupe0t6ouu2laeiol1abv.apps.googleusercontent.com';
const APP_VERSION = '2026.08.30.2';
const SPREADSHEET_ID = '16J873aq698SxJsFgiWcNJOubn3R3z_5_J8NMlrsuIsA';
const TAXONOMY_SHEET_ID = '1oQM1alY_nyVpk8LcHsFmEMpEk-jy0b18vf506_ubj9s';
const BOARD_SHEET_ID = '1LPMQEO9DCQ7-CAKtCIDkWTOFZ3mK6e1qBcDfa3lggQQ';
const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';
const RANGE = 'A:J'; // Task ID | Task | Activity Tag | Due Date | Priority | Notes | Done | Today | Quick Note | Recurrence
const COMPLETIONS_RANGE = "'Recurring Completions'!A:B"; // Task ID | Date
const LONG_PRESS_MS = 550;
const TAXONOMY_RANGE = 'A:E'; // Code | Label | Parent | Domain | Notes
const BOARD_RANGE = 'A:J'; // Code | Name | Domain | Stage | Priority | NextUp | Notes | MilestoneDate | Updated | Show

let activeTab = 'tasks'; // 'tasks' | 'board'
let domainFirst = 'Work'; // 'Work' | 'Personal'
let sortMode = 'priority'; // 'category' | 'priority'
let boardMode = 'tracked'; // 'tracked' | 'all'

// ── STATE ───────────────────────────────────────────────────────────
let tokenClient = null;
let accessToken = null;
let tasks = [];
let recurringCompletions = [];
let taxonomy = {};
let taxonomyParent = {};
let board = [];

// ── STORAGE HELPERS ─────────────────────────────────────────────────
const CACHE_KEY = 'tasks_cache';
const TAXONOMY_CACHE_KEY = 'taxonomy_cache';
const QUEUE_KEY = 'tasks_pending_writes';
const TOKEN_KEY = 'tasks_token_cache';
const COMPLETIONS_CACHE_KEY = 'recurring_completions_cache';

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

function saveCompletionsCache(items) {
  localStorage.setItem(COMPLETIONS_CACHE_KEY, JSON.stringify(items));
}

function loadCompletionsCache() {
  const raw = localStorage.getItem(COMPLETIONS_CACHE_KEY);
  return raw ? JSON.parse(raw) : [];
}

function getQueue() {
  return JSON.parse(localStorage.getItem(QUEUE_KEY) || '[]');
}

function setQueue(q) {
  localStorage.setItem(QUEUE_KEY, JSON.stringify(q));
}

function queueOperation(op) {
  const q = getQueue();
  q.push({ ...op, ts: Date.now() });
  setQueue(q);
}

function queueWrite(row, column, value) {
  queueOperation({ type: 'cell', row, column, value });
}

function clearQueue() {
  setQueue([]);
}

function removeQueuedCompletionAdd(taskId, date) {
  const q = getQueue();
  const idx = q.findIndex(item =>
    item.type === 'completion-add' &&
    item.taskId === taskId &&
    item.date === date
  );

  if (idx === -1) return false;

  q.splice(idx, 1);
  setQueue(q);
  return true;
}

function saveToken(token, expiresInSec) {
  const expiry = Date.now() + expiresInSec * 1000 - 60000;
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

      Promise.all([
        fetchTaxonomy(),
        fetchBoard(),
        fetchRecurringCompletions()
      ]).then(fetchTasks);
    }
  });

  const cached = loadValidToken();

  if (cached) {
    accessToken = cached;
    document.getElementById('connect-btn').style.display = 'none';
    setStatus('Connected (cached).');

    if (navigator.onLine) {
      Promise.all([
        fetchTaxonomy(),
        fetchBoard(),
        fetchRecurringCompletions()
      ]).then(fetchTasks);
    }
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
  if (!accessToken || !navigator.onLine) {
    const cached = loadTaxonomyCache();
    taxonomy = cached.labels || {};
    taxonomyParent = cached.parents || {};
    return;
  }

  try {
    const url =
      `https://sheets.googleapis.com/v4/spreadsheets/${TAXONOMY_SHEET_ID}/values/${TAXONOMY_RANGE}`;

    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!res.ok) throw new Error('HTTP ' + res.status);

    const data = await res.json();
    const rows = (data.values || []).slice(1);

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

    saveTaxonomyCache({
      labels: map,
      parents: parentMap
    });
  } catch (e) {
    const cached = loadTaxonomyCache();
    taxonomy = cached.labels || {};
    taxonomyParent = cached.parents || {};
  }
}

function getCategoryLabel(tag) {
  if (!tag) return 'Uncategorized';

  const primary =
    tag.split(/[·/]/)
      .map(t => t.trim())
      .find(t => /^[WP]/.test(t)) || tag.trim();

  let current = primary;
  let guard = 0;

  while (
    taxonomyParent[current] &&
    taxonomyParent[current] !== 'W' &&
    taxonomyParent[current] !== 'P' &&
    guard < 10
  ) {
    current = taxonomyParent[current];
    guard++;
  }

  return taxonomy[current] || current;
}

function labelTag(tag) {
  if (!tag) return '';

  return tag
    .split(/[·/]/)
    .map(part => {
      const code = part.trim();
      const label = taxonomy[code];
      return label ? `${code} (${label})` : code;
    })
    .join(' · ');
}

async function fetchBoard() {
  if (!accessToken || !navigator.onLine) return;

  try {
    const url =
      `https://sheets.googleapis.com/v4/spreadsheets/${BOARD_SHEET_ID}/values/${BOARD_RANGE}`;

    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!res.ok) throw new Error('HTTP ' + res.status);

    const data = await res.json();
    const rows = (data.values || []).slice(1);

    board = rows
      .map(r => ({
        code: r[0] || '',
        name: r[1] || '',
        domain: r[2] || '',
        stage: r[3] || '',
        priority: r[4] || 'Medium',
        nextUp: r[5] || '',
        notes: r[6] || '',
        milestoneDate: r[7] || '',
        updated: r[8] || '',
        show: (r[9] || '').toString().toUpperCase() !== 'FALSE'
      }))
      .filter(p => p.code);

    localStorage.setItem('board_cache', JSON.stringify(board));
  } catch (e) {
    const raw = localStorage.getItem('board_cache');
    board = raw ? JSON.parse(raw) : [];
  }
}

async function fetchRecurringCompletions() {
  if (!accessToken || !navigator.onLine) {
    recurringCompletions = loadCompletionsCache();
    return;
  }

  try {
    const range = encodeURIComponent(COMPLETIONS_RANGE);

    const url =
      `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}`;

    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!res.ok) throw new Error('HTTP ' + res.status);

    const data = await res.json();
    const rows = (data.values || []).slice(1);

    recurringCompletions = rows
      .map((r, i) => ({
        row: i + 2,
        taskId: r[0] || '',
        date: r[1] || ''
      }))
      .filter(c => c.taskId && c.date);

    saveCompletionsCache(recurringCompletions);
  } catch (e) {
    recurringCompletions = loadCompletionsCache();
  }
}

async function fetchTasks() {
  if (!accessToken || !navigator.onLine) return;

  setStatus('Syncing…');

  try {
    const url =
      `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${RANGE}`;

    const res = await fetch(url, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });

    if (!res.ok) throw new Error('HTTP ' + res.status);

    const data = await res.json();
    const rows = (data.values || []).slice(1);

    tasks = rows
      .map((r, i) => ({
        row: i + 2,
        id: r[0] || '',
        task: r[1] || '',
        tag: r[2] || '',
        due: r[3] || '',
        priority: r[4] || 'Medium',
        notes: r[5] || '',
        done: (r[6] || '').toString().toUpperCase() === 'TRUE',
        todayFlag: (r[7] || '').toString().trim().toUpperCase(),
        note: r[8] || '',
        recurrence: (r[9] || '').trim()
      }))
      .filter(t => t.id);

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

// ── WRITE ───────────────────────────────────────────────────────────
async function writeColumn(task, column, field, value) {
  task[field] = value;

  saveCache(tasks);
  render();

  const cellValue =
    typeof value === 'boolean'
      ? (value ? 'TRUE' : 'FALSE')
      : value;

  if (accessToken && navigator.onLine) {
    try {
      const url =
        `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}` +
        `/values/${column}${task.row}?valueInputOption=RAW`;

      const res = await fetch(url, {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          values: [[cellValue]]
        })
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
  return writeColumn(
    task,
    'H',
    'todayFlag',
    checked ? 'TRUE' : 'FALSE'
  );
}

// ── RECURRENCE ──────────────────────────────────────────────────────
function isRecurring(task) {
  const r = (task.recurrence || '').trim();

  return (
    r === 'Daily' ||
    r === 'Weekdays' ||
    /^[2-7]x\/week$/.test(r)
  );
}

function recurrenceTarget(task) {
  const m = (task.recurrence || '').match(/^([2-7])x\/week$/);
  return m ? Number(m[1]) : null;
}

function completionForToday(task) {
  const today = todayStr();

  return recurringCompletions.find(
    c => c.taskId === task.id && c.date === today
  ) || null;
}

function hasCompletionToday(task) {
  return Boolean(completionForToday(task));
}

function weekBounds() {
  const d = new Date();

  const day = d.getDay();
  const deltaToMonday = day === 0 ? -6 : 1 - day;

  const monday = new Date(
    d.getFullYear(),
    d.getMonth(),
    d.getDate() + deltaToMonday
  );

  const sunday = new Date(
    monday.getFullYear(),
    monday.getMonth(),
    monday.getDate() + 6
  );

  const fmt = x =>
    x.getFullYear() +
    '-' +
    String(x.getMonth() + 1).padStart(2, '0') +
    '-' +
    String(x.getDate()).padStart(2, '0');

  return {
    start: fmt(monday),
    end: fmt(sunday)
  };
}

function weeklyCompletionCount(task) {
  const { start, end } = weekBounds();

  return recurringCompletions.filter(c =>
    c.taskId === task.id &&
    c.date >= start &&
    c.date <= end
  ).length;
}

function recurrenceAppliesToday(task) {
  if (!isRecurring(task) || task.done) return false;

  const r = task.recurrence;

  if (r === 'Daily') return true;

  if (r === 'Weekdays') {
    const day = new Date().getDay();
    return day >= 1 && day <= 5;
  }

  const target = recurrenceTarget(task);
  if (!target) return false;

  return (
    weeklyCompletionCount(task) < target ||
    hasCompletionToday(task)
  );
}

function recurrenceDisplay(task) {
  const r = task.recurrence || '';

  if (r === 'Daily' || r === 'Weekdays') {
    return r;
  }

  const target = recurrenceTarget(task);
  if (!target) return '';

  const count = weeklyCompletionCount(task);

  return `${target}×/week · ${count}/${target}`;
}

async function appendRecurringCompletion(task) {
  if (!isRecurring(task) || hasCompletionToday(task)) return;

  const date = todayStr();

  const localRecord = {
    row: null,
    taskId: task.id,
    date
  };

  recurringCompletions.push(localRecord);
  saveCompletionsCache(recurringCompletions);
  render();

  if (accessToken && navigator.onLine) {
    try {
      const range = encodeURIComponent(COMPLETIONS_RANGE);

      const url =
        `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}` +
        `/values/${range}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;

      const res = await fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          values: [[task.id, date]]
        })
      });

      if (!res.ok) throw new Error('HTTP ' + res.status);

      const data = await res.json();
      const updatedRange = data?.updates?.updatedRange || '';

      const rowMatch =
        updatedRange.match(/!A(\d+):B\d+$/);

      if (rowMatch) {
        localRecord.row = Number(rowMatch[1]);
      }

      saveCompletionsCache(recurringCompletions);
      setStatus('Recurring completion saved.');
    } catch (e) {
      queueOperation({
        type: 'completion-add',
        taskId: task.id,
        date
      });

      setStatus(
        'Offline or write failed — recurring completion queued.'
      );
    }
  } else {
    queueOperation({
      type: 'completion-add',
      taskId: task.id,
      date
    });

    setStatus('Offline — recurring completion queued.');
  }
}

async function removeRecurringCompletion(task) {
  const record = completionForToday(task);
  if (!record) return;

  recurringCompletions =
    recurringCompletions.filter(c => c !== record);

  saveCompletionsCache(recurringCompletions);
  render();

  if (
    !record.row &&
    removeQueuedCompletionAdd(task.id, record.date)
  ) {
    setStatus('Recurring completion undone.');
    return;
  }

  if (!record.row) {
    if (accessToken && navigator.onLine) {
      await fetchRecurringCompletions();

      const refreshed = completionForToday(task);

      if (refreshed) {
        recurringCompletions =
          recurringCompletions.filter(c => c !== refreshed);

        saveCompletionsCache(recurringCompletions);

        return clearRecurringCompletionRow(refreshed.row);
      }
    }

    setStatus('Could not undo yet — sync and try again.');
    return;
  }

  return clearRecurringCompletionRow(record.row);
}

async function clearRecurringCompletionRow(rowNumber) {
  if (accessToken && navigator.onLine) {
    try {
      const range = encodeURIComponent(
        `'Recurring Completions'!A${rowNumber}:B${rowNumber}`
      );

      const url =
        `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}` +
        `/values/${range}:clear`;

      const res = await fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: '{}'
      });

      if (!res.ok) throw new Error('HTTP ' + res.status);

      setStatus('Recurring completion undone.');
    } catch (e) {
      queueOperation({
        type: 'completion-clear',
        row: rowNumber
      });

      setStatus('Undo queued for later sync.');
    }
  } else {
    queueOperation({
      type: 'completion-clear',
      row: rowNumber
    });

    setStatus('Offline — undo queued.');
  }
}

function toggleTaskCompletion(task, checked) {
  if (!isRecurring(task)) {
    return toggleDone(task, checked);
  }

  return checked
    ? appendRecurringCompletion(task)
    : removeRecurringCompletion(task);
}

// ── QUICK NOTES ─────────────────────────────────────────────────────
function saveQuickNote(task, text) {
  const trimmed = (text || '').trim();
  if (!trimmed) return Promise.resolve();

  const stamp = todayStr();
  const existing = (task.note || '').trim();

  const entry = '[' + stamp + '] ' + trimmed;
  const updated =
    existing
      ? existing + '\n' + entry
      : entry;

  return writeColumn(task, 'I', 'note', updated);
}

// ── OFFLINE QUEUE ───────────────────────────────────────────────────
async function flushQueue() {
  const q = getQueue();

  if (
    !q.length ||
    !accessToken ||
    !navigator.onLine
  ) {
    return;
  }

  setStatus(
    'Syncing ' + q.length + ' queued change(s)…'
  );

  for (let i = 0; i < q.length; i++) {
    const item = q[i];

    try {
      if (!item.type || item.type === 'cell') {
        const url =
          `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}` +
          `/values/${item.column}${item.row}?valueInputOption=RAW`;

        const res = await fetch(url, {
          method: 'PUT',
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            values: [[item.value]]
          })
        });

        if (!res.ok) throw new Error('HTTP ' + res.status);
      } else if (item.type === 'completion-add') {
        const range = encodeURIComponent(COMPLETIONS_RANGE);

        const url =
          `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}` +
          `/values/${range}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;

        const res = await fetch(url, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          body: JSON.stringify({
            values: [[item.taskId, item.date]]
          })
        });

        if (!res.ok) throw new Error('HTTP ' + res.status);
      } else if (item.type === 'completion-clear') {
        const range = encodeURIComponent(
          `'Recurring Completions'!A${item.row}:B${item.row}`
        );

        const url =
          `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}` +
          `/values/${range}:clear`;

        const res = await fetch(url, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          },
          body: '{}'
        });

        if (!res.ok) throw new Error('HTTP ' + res.status);
      }
    } catch (e) {
      setQueue(q.slice(i));

      setStatus(
        'Some queued changes failed to sync — will retry next time.'
      );

      return;
    }
  }

  clearQueue();

  await fetchRecurringCompletions();
  render();

  setStatus('All queued changes synced.');
}

// ── QUICK NOTE MODAL ────────────────────────────────────────────────
let noteModalTask = null;

function openNoteModal(task) {
  noteModalTask = task;

  document.getElementById(
    'note-modal-title'
  ).textContent = task.task;

  document.getElementById(
    'note-modal-existing'
  ).textContent = task.note || '';

  document.getElementById(
    'note-modal-existing'
  ).style.display =
    task.note ? 'block' : 'none';

  document.getElementById(
    'note-input'
  ).value = '';

  document.getElementById(
    'note-modal'
  ).style.display = 'flex';

  setTimeout(
    () => document.getElementById('note-input').focus(),
    50
  );
}

function closeNoteModal() {
  document.getElementById(
    'note-modal'
  ).style.display = 'none';

  noteModalTask = null;
}

// ── RENDER HELPERS ──────────────────────────────────────────────────
function setStatus(msg) {
  document.getElementById('status').textContent = msg;
}

function todayStr() {
  const d = new Date();

  return (
    d.getFullYear() +
    '-' +
    String(d.getMonth() + 1).padStart(2, '0') +
    '-' +
    String(d.getDate()).padStart(2, '0')
  );
}

function normalizeDue(due) {
  if (!due) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(due)) {
    return due;
  }

  const m =
    due.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);

  if (m) {
    const [, mo, da, yr] = m;

    return (
      yr +
      '-' +
      mo.padStart(2, '0') +
      '-' +
      da.padStart(2, '0')
    );
  }

  return '';
}

function isOverdue(t) {
  if (t.done || isRecurring(t)) return false;

  const norm = normalizeDue(t.due);
  if (!norm) return false;

  return norm < todayStr();
}

// ── TASK VIEW ───────────────────────────────────────────────────────
function render() {
  if (activeTab === 'board') {
    renderBoard();
    return;
  }

  const list =
    tasks.length
      ? tasks
      : loadCache();

  const container =
    document.getElementById('tasks');

  container.innerHTML = '';

  function isToday(t) {
    if (recurrenceAppliesToday(t)) return true;

    return t.todayFlag === 'TRUE';
  }

  const todayItems =
    list.filter(t => isToday(t));

  const workItems =
    list.filter(t =>
      !isRecurring(t) &&
      t.tag.trim().startsWith('W') &&
      !isToday(t)
    );

  const personalItems =
    list.filter(t =>
      !isRecurring(t) &&
      t.tag.trim().startsWith('P') &&
      !isToday(t)
    );

  function taskRow(t) {
    const row =
      document.createElement('label');

    const checked =
      isRecurring(t)
        ? hasCompletionToday(t)
        : t.done;

    row.className =
      'item' + (checked ? ' done' : '');

    row.innerHTML = `
      <input type="checkbox" ${checked ? 'checked' : ''} />
      <div class="content">
        <div class="label">${t.task}${t.note ? ' <span class="note-flag" title="Has a note">📝</span>' : ''}</div>
        <div class="proj">
          ${labelTag(t.tag)} ·
          <span class="priority-tag">${t.priority || 'Medium'}</span>
          ${isRecurring(t) ? ` · <span class="recurrence-tag">${recurrenceDisplay(t)}</span>` : ''}
        </div>
      </div>
      <div class="due-col${isOverdue(t) ? ' overdue' : ''}">
        ${t.due || ''}
      </div>
      ${isRecurring(t) ? '' : `<div class="today-btn ${isToday(t) ? 'active' : ''}" title="Toggle Today"></div>`}
    `;

    row.querySelector('input')
      .addEventListener(
        'change',
        e => toggleTaskCompletion(
          t,
          e.target.checked
        )
      );

    const todayButton =
      row.querySelector('.today-btn');

    if (todayButton) {
      todayButton.addEventListener(
        'click',
        e => {
          e.preventDefault();
          e.stopPropagation();

          toggleTodayShared(
            t,
            !isToday(t)
          );
        }
      );
    }

    let pressTimer = null;
    let longPressed = false;

    const startPress = () => {
      longPressed = false;

      pressTimer = setTimeout(() => {
        longPressed = true;

        if (navigator.vibrate) {
          navigator.vibrate(12);
        }

        openNoteModal(t);
      }, LONG_PRESS_MS);
    };

    const cancelPress = () =>
      clearTimeout(pressTimer);

    row.addEventListener(
      'touchstart',
      startPress,
      { passive: true }
    );

    row.addEventListener(
      'touchend',
      cancelPress
    );

    row.addEventListener(
      'touchmove',
      cancelPress
    );

    row.addEventListener(
      'mousedown',
      startPress
    );

    row.addEventListener(
      'mouseup',
      cancelPress
    );

    row.addEventListener(
      'mouseleave',
      cancelPress
    );

    row.addEventListener(
      'click',
      e => {
        if (longPressed) {
          e.preventDefault();
          longPressed = false;
        }
      }
    );

    return row;
  }

  function subhead(text) {
    const h =
      document.createElement('div');

    h.className = 'subhead';
    h.textContent = text;

    return h;
  }

  function domainHead(text) {
    const h =
      document.createElement('div');

    h.className = 'group-head';
    h.textContent = text;

    return h;
  }

  function byDateAscending(items) {
    return [...items]
      .sort((a, b) => {
        const da = normalizeDue(a.due);
        const db = normalizeDue(b.due);

        if (!da && !db) return 0;
        if (!da) return 1;
        if (!db) return -1;

        return da.localeCompare(db);
      });
  }

  function byPriorityThenDate(items) {
    const order = {
      High: 0,
      Medium: 1,
      Low: 2
    };

    return byDateAscending(items)
      .sort(
        (a, b) =>
          (order[a.priority] ?? 1) -
          (order[b.priority] ?? 1)
      );
  }

  function renderGroupedItems(items) {
    if (sortMode === 'category') {
      const groups = {};

      items.forEach(t => {
        const cat =
          getCategoryLabel(t.tag);

        (
          groups[cat] =
          groups[cat] || []
        ).push(t);
      });

      Object.keys(groups)
        .sort()
        .forEach(cat => {
          container.appendChild(
            subhead(cat)
          );

          byPriorityThenDate(
            groups[cat]
          ).forEach(t =>
            container.appendChild(
              taskRow(t)
            )
          );
        });
    } else {
      const tiers = [
        'High',
        'Medium',
        'Low'
      ];

      tiers.forEach(tier => {
        const inTier =
          items.filter(
            t =>
              (t.priority || 'Medium') === tier
          );

        if (!inTier.length) return;

        container.appendChild(
          subhead(tier)
        );

        byDateAscending(inTier)
          .forEach(t =>
            container.appendChild(
              taskRow(t)
            )
          );
      });
    }
  }

  function renderDomainSection(
    title,
    items
  ) {
    if (!items.length) return;

    container.appendChild(
      domainHead(title)
    );

    renderGroupedItems(items);
  }

  if (todayItems.length) {
    container.appendChild(
      domainHead('Today')
    );

    renderGroupedItems(todayItems);
  }

  if (domainFirst === 'Personal') {
    renderDomainSection(
      'Personal',
      personalItems
    );

    renderDomainSection(
      'Work',
      workItems
    );
  } else {
    renderDomainSection(
      'Work',
      workItems
    );

    renderDomainSection(
      'Personal',
      personalItems
    );
  }

  if (!list.length) {
    container.innerHTML =
      '<p class="empty">No tasks loaded yet. Connect to sync.</p>';
  }
}

// ── BOARD VIEW ──────────────────────────────────────────────────────
function yearProgressPct(dateStr) {
  if (!dateStr) return null;

  const d = new Date(dateStr);
  const start =
    new Date(d.getFullYear(), 0, 1);
  const end =
    new Date(d.getFullYear(), 11, 31);

  return Math.round(
    ((d - start) / (end - start)) * 100
  );
}

function getTopLevelCodes(domainRoot) {
  return Object.keys(taxonomyParent)
    .filter(
      code =>
        taxonomyParent[code] === domainRoot
    )
    .sort();
}

function renderBoard() {
  const container =
    document.getElementById('tasks');

  container.innerHTML = '';

  const list =
    board.length
      ? board
      : JSON.parse(
          localStorage.getItem('board_cache') ||
          '[]'
        );

  const boardByCode = {};

  list.forEach(p => {
    boardByCode[p.code] = p;
  });

  function boardCard(p) {
    const card =
      document.createElement('div');

    card.className =
      'board-card priority-' +
      (p.priority || 'Medium').toLowerCase() +
      (p.stage === 'Untracked'
        ? ' untracked'
        : '');

    const pct =
      yearProgressPct(
        p.milestoneDate
      );

    card.innerHTML = `
      <div class="board-head">
        <span class="board-code">
          ${p.code}${taxonomy[p.code.split(' ')[0]] ? ' · ' + taxonomy[p.code.split(' ')[0]] : ''}
        </span>
        <span class="board-stage">${p.stage}</span>
      </div>
      <div class="board-name">${p.name}</div>
      ${p.nextUp ? `<div class="board-next">${p.nextUp}</div>` : ''}
      ${pct !== null ? `<div class="board-track"><div class="board-marker" style="left:${pct}%"></div></div><div class="board-date">${p.milestoneDate}</div>` : ''}
      ${p.notes ? `<div class="board-notes">${p.notes}</div>` : ''}
    `;

    return card;
  }

  function section(
    title,
    domainCode
  ) {
    let items;

    if (boardMode === 'all') {
      const codes =
        getTopLevelCodes(
          domainCode
        );

      items = codes.map(
        code =>
          boardByCode[code] || {
            code,
            name:
              taxonomy[code] || code,
            domain: domainCode,
            stage: 'Untracked',
            priority: 'Low',
            nextUp: '',
            notes: '',
            milestoneDate: '',
            updated: '',
            show: true
          }
      );
    } else {
      items =
        list.filter(
          p =>
            p.domain === domainCode &&
            p.show
        );
    }

    if (!items.length) return;

    const h =
      document.createElement('div');

    h.className = 'group-head';
    h.textContent = title;

    container.appendChild(h);

    const stageOrder = {
      Milestone: 0,
      Active: 1,
      Scoping: 2,
      Waiting: 3,
      'Wrapping up': 4,
      Monitor: 5,
      'On-hold': 6,
      Untracked: 7
    };

    items.sort(
      (a, b) =>
        (stageOrder[a.stage] ?? 5) -
          (stageOrder[b.stage] ?? 5) ||
        a.code.localeCompare(
          b.code,
          undefined,
          { numeric: true }
        )
    );

    items.forEach(p =>
      container.appendChild(
        boardCard(p)
      )
    );
  }

  section('Work', 'W');
  section('Personal', 'P');

  const anyContent =
    boardMode === 'all'
      ? (
          getTopLevelCodes('W').length ||
          getTopLevelCodes('P').length
        )
      : list.some(p => p.show);

  if (!anyContent) {
    container.innerHTML =
      '<p class="empty">No board data loaded yet.</p>';
  }
}

// ── INIT ────────────────────────────────────────────────────────────
window.addEventListener(
  'load',
  () => {
    document.title =
      'Tasks (' + APP_VERSION + ')';

    const versionTag =
      document.createElement('div');

    versionTag.id = 'version-tag';
    versionTag.textContent =
      'App code: ' + APP_VERSION;

    document.body.insertBefore(
      versionTag,
      document.getElementById('status')
    );

    const dayOfWeek =
      new Date().getDay();

    if (
      dayOfWeek === 0 ||
      dayOfWeek === 6
    ) {
      domainFirst = 'Personal';

      document.getElementById(
        'sort-domain'
      ).textContent =
        'Personal first';
    }

    tasks = loadCache();
    recurringCompletions =
      loadCompletionsCache();

    const cachedTax =
      loadTaxonomyCache();

    taxonomy =
      cachedTax.labels || {};

    taxonomyParent =
      cachedTax.parents || {};

    board = JSON.parse(
      localStorage.getItem(
        'board_cache'
      ) || '[]'
    );

    render();
    initAuth();

    document.getElementById(
      'connect-btn'
    ).addEventListener(
      'click',
      connect
    );

    document.getElementById(
      'tab-tasks'
    ).addEventListener(
      'click',
      () => setTab('tasks')
    );

    document.getElementById(
      'tab-board'
    ).addEventListener(
      'click',
      () => setTab('board')
    );

    document.getElementById(
      'sort-domain'
    ).addEventListener(
      'click',
      () => {
        domainFirst =
          domainFirst === 'Work'
            ? 'Personal'
            : 'Work';

        document.getElementById(
          'sort-domain'
        ).textContent =
          domainFirst + ' first';

        render();
      }
    );

    ['category', 'priority']
      .forEach(mode => {
        document.getElementById(
          'sort-' + mode
        ).addEventListener(
          'click',
          () => {
            sortMode = mode;

            ['category', 'priority']
              .forEach(m =>
                document.getElementById(
                  'sort-' + m
                ).classList.toggle(
                  'active',
                  m === mode
                )
              );

            render();
          }
        );
      });

    ['tracked', 'all']
      .forEach(mode => {
        document.getElementById(
          'board-' + mode
        ).addEventListener(
          'click',
          () => {
            boardMode = mode;

            ['tracked', 'all']
              .forEach(m =>
                document.getElementById(
                  'board-' + m
                ).classList.toggle(
                  'active',
                  m === mode
                )
              );

            render();
          }
        );
      });

    document.getElementById(
      'note-cancel'
    ).addEventListener(
      'click',
      closeNoteModal
    );

    document.getElementById(
      'note-modal'
    ).addEventListener(
      'click',
      e => {
        if (
          e.target.id ===
          'note-modal'
        ) {
          closeNoteModal();
        }
      }
    );

    document.getElementById(
      'note-save'
    ).addEventListener(
      'click',
      () => {
        const text =
          document.getElementById(
            'note-input'
          ).value;

        const task =
          noteModalTask;

        closeNoteModal();

        if (task) {
          saveQuickNote(
            task,
            text
          );
        }
      }
    );
  }
);

function setTab(tab) {
  activeTab = tab;

  document.getElementById(
    'tab-tasks'
  ).classList.toggle(
    'active',
    tab === 'tasks'
  );

  document.getElementById(
    'tab-board'
  ).classList.toggle(
    'active',
    tab === 'board'
  );

  document.getElementById(
    'sort-controls'
  ).style.display =
    tab === 'tasks'
      ? 'flex'
      : 'none';

  document.getElementById(
    'board-controls'
  ).style.display =
    tab === 'board'
      ? 'flex'
      : 'none';

  render();
}

window.addEventListener(
  'online',
  async () => {
    setStatus(
      'Back online, syncing…'
    );

    await fetchRecurringCompletions();
    fetchTasks();
  }
);

window.addEventListener(
  'offline',
  () =>
    setStatus(
      'Offline — showing cached data.'
    )
);

if ('serviceWorker' in navigator) {
  window.addEventListener(
    'load',
    () => {
      navigator.serviceWorker.register(
        './service-worker.js'
      );

      let reloaded = false;

      navigator.serviceWorker.addEventListener(
        'controllerchange',
        () => {
          if (reloaded) return;

          reloaded = true;
          window.location.reload();
        }
      );
    }
  );
}
