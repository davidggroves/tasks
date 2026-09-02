/* Tasks app — shared application core
 * Personalization belongs in window.TASK_APP_CONFIG, loaded before this file.
 */
(() => {
  'use strict';

  const CFG = window.TASK_APP_CONFIG;
  if (!CFG) throw new Error('TASK_APP_CONFIG must be loaded before app.js');

  const APP_VERSION = CFG.appVersion || 'dev';
  const CLIENT_ID = CFG.clientId || '';
  const SPREADSHEET_ID = CFG.sheets?.tasks || '';
  const TAXONOMY_SHEET_ID = CFG.sheets?.taxonomy || '';
  const BOARD_SHEET_ID = CFG.sheets?.board || '';
  const TASK_RANGE = CFG.ranges?.tasks || 'A:J';
  const TAXONOMY_RANGE = CFG.ranges?.taxonomy || 'A:E';
  const BOARD_RANGE = CFG.ranges?.board || 'A:J';
  const RECURRING_RANGE = CFG.ranges?.recurring || "'Recurring Completions'!A:B";
  const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';
  const PREFIX = (CFG.storagePrefix || 'tasks_app').replace(/[^a-z0-9_-]/gi, '_').toLowerCase();

  const CACHE_KEY = `${PREFIX}_tasks_cache`;
  const TAXONOMY_CACHE_KEY = `${PREFIX}_taxonomy_cache`;
  const BOARD_CACHE_KEY = `${PREFIX}_board_cache`;
  const RECURRING_CACHE_KEY = `${PREFIX}_recurring_completions_cache`;
  const QUEUE_KEY = `${PREFIX}_tasks_pending_writes`;
  const TOKEN_KEY = `${PREFIX}_tasks_token_cache`;

  let activeTab = 'tasks';
  let domainFirst = 'Work';
  let sortMode = 'priority';
  let boardMode = 'tracked';
  let tokenClient = null;
  let accessToken = null;
  let tasks = [];
  let taxonomy = {};
  let taxonomyParent = {};
  let board = [];
  let recurringCompletions = [];
  let detailTask = null;

  function escapeHtml(s) {
    return String(s || '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[ch]));
  }

  function injectShell() {
    document.documentElement.style.setProperty('--app-accent', CFG.themeColor || '#0d8c72');
    document.title = `${CFG.appName || 'Tasks'} (${APP_VERSION})`;
    document.head.insertAdjacentHTML('beforeend', `
      <meta name="theme-color" content="${escapeHtml(CFG.themeColor || '#0d8c72')}">
      <meta name="apple-mobile-web-app-capable" content="yes">
      <meta name="apple-mobile-web-app-status-bar-style" content="default">
      <link rel="manifest" href="./manifest.json">
      <link rel="apple-touch-icon" href="./icon-192.png">
      <style>
        *{box-sizing:border-box;margin:0;padding:0} body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f5f4f0;color:#1a1a18;padding:1.25rem 1rem 3rem;max-width:640px;margin:0 auto}
        h1{font-size:18px;font-weight:600;margin-bottom:2px}.app-header{display:flex;align-items:center;gap:12px;margin-bottom:2px}.app-image{width:52px;height:52px;object-fit:cover;border-radius:10px;flex-shrink:0}
        #status{font-size:12px;color:#999;margin-bottom:1rem;min-height:16px}#version-tag{font-size:11px;color:#bbb;margin-bottom:4px}#connect-btn{background:var(--app-accent);color:#fff;border:none;border-radius:8px;padding:10px 16px;font-size:14px;font-weight:500;cursor:pointer;margin-bottom:1rem}
        .tab-bar{display:flex;gap:6px;margin-bottom:.75rem}.tab-bar button{font-size:13px;padding:6px 14px;border:1px solid #d0cfc8;border-radius:20px;background:#fff;cursor:pointer;color:#555;font-family:inherit}.tab-bar button.active{background:#1a1a18;color:#fff;border-color:#1a1a18}
        #sort-controls,#board-controls{display:flex;gap:6px;margin-bottom:1rem;flex-wrap:wrap}#sort-controls button,#board-controls button{font-size:12px;padding:5px 12px;border:1px solid #d0cfc8;border-radius:16px;background:#f0f0ee;cursor:pointer;color:#666;font-family:inherit}#sort-controls button.active,#board-controls button.active{background:var(--app-accent);color:#fff;border-color:var(--app-accent)}
        .group-head{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.08em;color:#888;margin:1.2rem 0 .4rem;padding-bottom:4px;border-bottom:1.5px solid #d0cfc8}.subhead{font-size:12px;font-weight:600;color:#666;margin:.6rem 0 .15rem 4px}
        .item{display:flex;align-items:flex-start;gap:10px;padding:10px 4px;border-bottom:.5px solid #e8e6df;cursor:pointer}.item input[type=checkbox]{width:18px;height:18px;margin-top:2px;flex-shrink:0;accent-color:var(--app-accent)}.content{flex:1;min-width:0}.label{font-size:14px;font-weight:500}.item.done .label{color:#aaa;text-decoration:line-through}.proj{font-size:11.5px;color:#999;margin-top:2px}.priority-tag{font-weight:600}.due-col{margin-left:auto;font-size:11.5px;color:#999;white-space:nowrap;padding-top:2px;flex-shrink:0}.due-col.overdue{color:#d32f2f;font-weight:700}.today-btn{width:16px;height:16px;border-radius:50%;border:1.5px solid #ccc;margin-top:2px;flex-shrink:0;cursor:pointer;background:transparent}.today-btn.active{background:var(--app-accent);border-color:var(--app-accent)}.note-flag{font-size:12px;margin-left:2px}.empty{font-size:13px;color:#999;padding:2rem 0;text-align:center}
        .board-card{border:1px solid #e0dfd8;border-radius:8px;padding:10px 12px;margin-bottom:8px;border-left:3px solid #ccc}.board-card.priority-high{border-left-color:#c0392b}.board-card.priority-medium{border-left-color:#e67e22}.board-card.priority-low{border-left-color:#ccc}.board-card.untracked{opacity:.6;border-left-color:#ddd}.board-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:3px}.board-code{font-size:10.5px;color:#999}.board-stage{font-size:9px;font-weight:700;text-transform:uppercase;color:#888;background:#f0f0ee;padding:2px 6px;border-radius:3px}.board-name{font-size:14px;font-weight:600;margin-bottom:3px}.board-next{font-size:12.5px;color:#444;margin-bottom:4px}.board-track{position:relative;height:4px;background:#eee;border-radius:2px;margin:6px 0 2px}.board-marker{position:absolute;top:-3px;width:10px;height:10px;border-radius:50%;background:var(--app-accent);transform:translateX(-50%)}.board-date{font-size:10.5px;color:#999;margin-bottom:4px}.board-notes{font-size:11.5px;color:#888;line-height:1.4}
        #task-modal{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);align-items:flex-end;justify-content:center;z-index:50}.modal-box{background:#fff;width:100%;max-width:640px;border-radius:14px 14px 0 0;padding:1rem 1rem 1.25rem;max-height:82%;overflow-y:auto}.modal-title{font-size:16px;font-weight:650;margin-bottom:8px}.detail-meta{font-size:12px;color:#777;margin-bottom:10px}.detail-section{margin-top:10px}.detail-label{font-size:10px;text-transform:uppercase;letter-spacing:.06em;color:#999;font-weight:700;margin-bottom:4px}.detail-text{font-size:13.5px;line-height:1.45;white-space:pre-wrap;background:#f5f4f0;border-radius:8px;padding:9px 10px}.detail-empty{color:#aaa;font-style:italic}.detail-actions{display:flex;justify-content:flex-end;gap:8px;margin-top:12px;flex-wrap:wrap}.detail-actions button{font-size:14px;padding:8px 14px;border-radius:8px;border:none;cursor:pointer;font-family:inherit}.secondary{background:#f0f0ee;color:#555}.primary{background:var(--app-accent);color:#fff;font-weight:500}#task-modal-notes{width:100%;font-family:inherit;font-size:14px;border:1px solid #d0cfc8;border-radius:8px;padding:10px;resize:vertical;min-height:110px;background:#fff;color:inherit}
        @media(prefers-color-scheme:dark){body{background:#1c1c1a;color:#e8e6df}.group-head{color:#888;border-bottom-color:#333}.item{border-bottom-color:#2e2e2b}.label{color:#e8e6df}.item.done .label{color:#666}.proj,.subhead{color:#888}.tab-bar button,#sort-controls button,#board-controls button{background:#252523;border-color:#333;color:#999}.tab-bar button.active{background:#e8e6df;color:#1a1a18;border-color:#e8e6df}.board-card{border-color:#333}.board-stage{background:#252523;color:#999}.board-name{color:#e8e6df}.board-next{color:#ccc}.modal-box{background:#1c1c1a}.detail-text,#task-modal-notes{background:#252523;color:#e8e6df;border-color:#333}.secondary{background:#252523;color:#aaa}}
      </style>`);

    document.body.innerHTML = `
      <div class="app-header"><img src="${escapeHtml(CFG.icon || './icon.jpg')}" class="app-image" alt="App"><h1>${escapeHtml(CFG.appName || 'Tasks')}</h1></div>
      <div id="version-tag">App code: ${escapeHtml(APP_VERSION)}</div>
      <div id="status">Loading…</div><button id="connect-btn">Connect to Google</button>
      <div class="tab-bar"><button id="tab-tasks" class="active">Tasks</button><button id="tab-board">Board</button></div>
      <div id="sort-controls"><button id="sort-domain">Work first</button><button id="sort-category">By Category</button><button id="sort-priority" class="active">By Priority</button></div>
      <div id="board-controls" style="display:none"><button id="board-tracked" class="active">Tracked Only</button><button id="board-all">All Categories</button></div>
      <div id="tasks"></div>
      <div id="task-modal"><div class="modal-box"><div class="modal-title" id="task-modal-title"></div><div class="detail-meta" id="task-modal-meta"></div><div class="detail-section"><div class="detail-label">Notes</div><textarea id="task-modal-notes" placeholder="Add task details…"></textarea></div><div class="detail-actions"><button id="task-notes-save" class="primary">Save notes</button><button id="task-modal-close" class="secondary">Close</button></div></div></div>`;
  }

  function saveCache(v){ localStorage.setItem(CACHE_KEY, JSON.stringify(v)); }
  function loadCache(){ const raw=localStorage.getItem(CACHE_KEY); return raw?JSON.parse(raw):[]; }
  function saveTaxonomyCache(v){ localStorage.setItem(TAXONOMY_CACHE_KEY, JSON.stringify(v)); }
  function loadTaxonomyCache(){ const raw=localStorage.getItem(TAXONOMY_CACHE_KEY); return raw?JSON.parse(raw):{}; }
  function queueWrite(row,column,value){ const q=JSON.parse(localStorage.getItem(QUEUE_KEY)||'[]'); q.push({row,column,value,ts:Date.now()}); localStorage.setItem(QUEUE_KEY,JSON.stringify(q)); }
  function getQueue(){ return JSON.parse(localStorage.getItem(QUEUE_KEY)||'[]'); }
  function clearQueue(){ localStorage.setItem(QUEUE_KEY,'[]'); }
  function saveToken(token,expiresInSec){ localStorage.setItem(TOKEN_KEY,JSON.stringify({token,expiry:Date.now()+expiresInSec*1000-60000})); }
  function loadValidToken(){ const raw=localStorage.getItem(TOKEN_KEY); if(!raw)return null; const {token,expiry}=JSON.parse(raw); return Date.now()<expiry?token:null; }
  function setStatus(msg){ document.getElementById('status').textContent=msg; }

  function initAuth(){
    if(!CLIENT_ID || CLIENT_ID.includes('REPLACE_')){ setStatus('OAuth client ID is not configured.'); return; }
    tokenClient=google.accounts.oauth2.initTokenClient({client_id:CLIENT_ID,scope:SCOPES,callback:(resp)=>{if(resp.error){setStatus('Sign-in failed: '+resp.error);return;}accessToken=resp.access_token;saveToken(accessToken,resp.expires_in);setStatus('Connected.');document.getElementById('connect-btn').style.display='none';Promise.all([fetchTaxonomy(),fetchBoard(),fetchRecurringCompletions()]).then(fetchTasks);}});
    const cached=loadValidToken();
    if(cached){accessToken=cached;document.getElementById('connect-btn').style.display='none';setStatus('Connected (cached).');if(navigator.onLine)Promise.all([fetchTaxonomy(),fetchBoard(),fetchRecurringCompletions()]).then(fetchTasks);}else{const c=loadTaxonomyCache();taxonomy=c.labels||{};taxonomyParent=c.parents||{};}
  }
  function connect(){ if(tokenClient) tokenClient.requestAccessToken({prompt:''}); }

  async function fetchTaxonomy(){
    if(!accessToken||!navigator.onLine){const c=loadTaxonomyCache();taxonomy=c.labels||{};taxonomyParent=c.parents||{};return;}
    try{const res=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${TAXONOMY_SHEET_ID}/values/${TAXONOMY_RANGE}`,{headers:{Authorization:`Bearer ${accessToken}`}});if(!res.ok)throw new Error('HTTP '+res.status);const data=await res.json();const map={},parents={};(data.values||[]).slice(1).forEach(r=>{if(!r[0])return;const code=r[0].trim();map[code]=r[1]||code;parents[code]=(r[2]||'').trim();});taxonomy=map;taxonomyParent=parents;saveTaxonomyCache({labels:map,parents});}catch(e){const c=loadTaxonomyCache();taxonomy=c.labels||{};taxonomyParent=c.parents||{};}
  }
  function getCategoryLabel(tag){if(!tag)return'Uncategorized';const primary=tag.split(/[·/]/).map(t=>t.trim()).find(t=>/^[WP]/.test(t))||tag.trim();let current=primary,guard=0;while(taxonomyParent[current]&&taxonomyParent[current]!=='W'&&taxonomyParent[current]!=='P'&&guard<10){current=taxonomyParent[current];guard++;}return taxonomy[current]||current;}
  function labelTag(tag){if(!tag)return'';return tag.split(/[·/]/).map(part=>{const code=part.trim();return taxonomy[code]?`${code} (${taxonomy[code]})`:code;}).join(' · ');}

  async function fetchBoard(){
    if(!BOARD_SHEET_ID||!accessToken||!navigator.onLine)return;
    try{const res=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${BOARD_SHEET_ID}/values/${BOARD_RANGE}`,{headers:{Authorization:`Bearer ${accessToken}`}});if(!res.ok)throw new Error('HTTP '+res.status);const data=await res.json();board=(data.values||[]).slice(1).map(r=>({code:r[0]||'',name:r[1]||'',domain:r[2]||'',stage:r[3]||'',priority:r[4]||'Medium',nextUp:r[5]||'',notes:r[6]||'',milestoneDate:r[7]||'',updated:r[8]||'',show:(r[9]||'').toString().toUpperCase()!=='FALSE'})).filter(p=>p.code);localStorage.setItem(BOARD_CACHE_KEY,JSON.stringify(board));}catch(e){const raw=localStorage.getItem(BOARD_CACHE_KEY);board=raw?JSON.parse(raw):[];}
  }

  function loadRecurringCache(){try{return JSON.parse(localStorage.getItem(RECURRING_CACHE_KEY)||'[]');}catch(_){return[];}}
  function saveRecurringCache(){localStorage.setItem(RECURRING_CACHE_KEY,JSON.stringify(recurringCompletions));}
  async function fetchRecurringCompletions(){
    if(!accessToken||!navigator.onLine){recurringCompletions=loadRecurringCache();return;}
    try{
      const res=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(RECURRING_RANGE)}`,{headers:{Authorization:`Bearer ${accessToken}`}});
      if(!res.ok)throw new Error('HTTP '+res.status);
      const data=await res.json();
      recurringCompletions=(data.values||[]).slice(1).map((r,i)=>({row:i+2,id:r[0]||'',date:r[1]||''})).filter(x=>x.id&&x.date);
      saveRecurringCache();
    }catch(_){recurringCompletions=loadRecurringCache();}
  }

  function localDateStr(d=new Date()){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}
  function weekBounds(){const now=new Date();now.setHours(12,0,0,0);const day=(now.getDay()+6)%7;const start=new Date(now);start.setDate(now.getDate()-day);const end=new Date(start);end.setDate(start.getDate()+6);return [localDateStr(start),localDateStr(end)];}
  function recurringQuota(task){const m=String(task.recurrence||'').match(/^([2-7])x\/week$/i);return m?Number(m[1]):null;}
  function recurringCountThisWeek(task){const [start,end]=weekBounds();return recurringCompletions.filter(x=>x.id===task.id&&x.date>=start&&x.date<=end).length;}
  function recurringCompletedToday(task){const today=localDateStr();return recurringCompletions.some(x=>x.id===task.id&&x.date===today);}
  function recurringProgress(task){const q=recurringQuota(task);return q?`${recurringCountThisWeek(task)}/${q}`:'';}
  function recurringVisible(task){
    if(!task.recurrence)return true;
    if(task.done)return false;
    const todayDone=recurringCompletedToday(task);
    const rule=String(task.recurrence).trim();
    if(/^daily$/i.test(rule))return true;
    if(/^weekdays$/i.test(rule)){const d=new Date().getDay();return d>=1&&d<=5;}
    const quota=recurringQuota(task);
    if(quota)return todayDone||recurringCountThisWeek(task)<quota;
    return true;
  }
  async function setRecurringCompletion(task,complete){
    if(!accessToken||!navigator.onLine){setStatus('Connect to Google to update recurring completion.');render();return;}
    const today=localDateStr();
    try{
      if(complete){
        if(recurringCompletedToday(task)){render();return;}
        const appendRange=encodeURIComponent("'Recurring Completions'!A:B");
        const res=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${appendRange}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,{method:'POST',headers:{Authorization:`Bearer ${accessToken}`,'Content-Type':'application/json'},body:JSON.stringify({values:[[task.id,today]]})});
        if(!res.ok){let detail='HTTP '+res.status;try{const b=await res.json();detail=b?.error?.message||detail;}catch(_){}throw new Error(detail);}
        const data=await res.json();
        const updated=data?.updates?.updatedRange||'';const m=updated.match(/!A(\d+):/);const row=m?Number(m[1]):null;
        recurringCompletions.push({row,id:task.id,date:today});
      }else{
        const hit=recurringCompletions.find(x=>x.id===task.id&&x.date===today);
        if(!hit){render();return;}
        if(!hit.row)await fetchRecurringCompletions();
        const target=recurringCompletions.find(x=>x.id===task.id&&x.date===today);
        if(!target?.row)throw new Error('Could not locate today’s completion row');
        const clearRange=encodeURIComponent(`'Recurring Completions'!A${target.row}:B${target.row}`);
        const res=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${clearRange}:clear`,{method:'POST',headers:{Authorization:`Bearer ${accessToken}`,'Content-Type':'application/json'},body:'{}'});
        if(!res.ok){let detail='HTTP '+res.status;try{const b=await res.json();detail=b?.error?.message||detail;}catch(_){}throw new Error(detail);}
        recurringCompletions=recurringCompletions.filter(x=>!(x.id===task.id&&x.date===today));
      }
      saveRecurringCache();render();setStatus(complete?'Recurring completion recorded.':'Recurring completion undone.');
    }catch(e){setStatus('Recurring update failed: '+e.message);await fetchRecurringCompletions();render();}
  }

  async function fetchTasks(){
    if(!accessToken||!navigator.onLine)return;setStatus('Syncing…');
    try{const res=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${TASK_RANGE}`,{headers:{Authorization:`Bearer ${accessToken}`}});if(!res.ok)throw new Error('HTTP '+res.status);const data=await res.json();tasks=(data.values||[]).slice(1).map((r,i)=>({row:i+2,id:r[0]||'',task:r[1]||'',tag:r[2]||'',due:r[3]||'',priority:r[4]||'Medium',notes:r[5]||'',done:(r[6]||'').toString().toUpperCase()==='TRUE',todayFlag:(r[7]||'').toString().trim().toUpperCase(),note:r[8]||'',recurrence:r[9]||''})).filter(t=>t.id);saveCache(tasks);setStatus('Synced '+new Date().toLocaleTimeString());render();flushQueue();}catch(e){setStatus('Sync failed, showing cached data: '+e.message);tasks=loadCache();render();}
  }

  async function writeColumn(task,column,field,value){task[field]=value;saveCache(tasks);render();const cellValue=typeof value==='boolean'?(value?'TRUE':'FALSE'):value;if(accessToken&&navigator.onLine){try{const res=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${column}${task.row}?valueInputOption=RAW`,{method:'PUT',headers:{Authorization:`Bearer ${accessToken}`,'Content-Type':'application/json'},body:JSON.stringify({values:[[cellValue]]})});if(!res.ok)throw new Error('HTTP '+res.status);setStatus('Saved.');}catch(e){queueWrite(task.row,column,cellValue);setStatus('Offline or write failed — queued for later sync.');}}else{queueWrite(task.row,column,cellValue);setStatus('Offline — change queued, will sync when connected.');}}
  const toggleDone=(task,v)=>task.recurrence?setRecurringCompletion(task,v):writeColumn(task,'G','done',v);
  const toggleTodayShared=(task,v)=>writeColumn(task,'H','todayFlag',v?'TRUE':'FALSE');
  function todayStr(){return localDateStr();}
  async function saveTaskNotes(task,text){
    const button=document.getElementById('task-notes-save');
    const original=button.textContent;
    const value=text||'';
    button.disabled=true;
    button.textContent='Saving…';
    try{
      if(!accessToken||!navigator.onLine)throw new Error('Not connected to Google');
      const range=`F${task.row}`;
      const res=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(range)}?valueInputOption=RAW`,{method:'PUT',headers:{Authorization:`Bearer ${accessToken}`,'Content-Type':'application/json'},body:JSON.stringify({values:[[value]]})});
      if(!res.ok){let detail='HTTP '+res.status;try{const body=await res.json();detail=body?.error?.message||detail;}catch(_){}throw new Error(detail);}
      task.notes=value;
      saveCache(tasks);
      render();
      setStatus('Notes saved.');
      button.textContent='Saved';
      setTimeout(()=>{if(detailTask===task){button.disabled=false;button.textContent='Save notes';}},900);
      return true;
    }catch(e){
      setStatus('Notes save failed: '+e.message);
      button.disabled=false;
      button.textContent='Try again';
      return false;
    }
  }

  async function flushQueue(){const q=getQueue();if(!q.length||!accessToken||!navigator.onLine)return;setStatus(`Syncing ${q.length} queued change(s)…`);for(const item of q){try{const res=await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${item.column}${item.row}?valueInputOption=RAW`,{method:'PUT',headers:{Authorization:`Bearer ${accessToken}`,'Content-Type':'application/json'},body:JSON.stringify({values:[[item.value]]})});if(!res.ok)throw new Error('HTTP '+res.status);}catch(e){setStatus('Some queued changes failed to sync — will retry next time.');return;}}clearQueue();setStatus('All queued changes synced.');}

  function normalizeDue(due){if(!due)return'';if(/^\d{4}-\d{2}-\d{2}$/.test(due))return due;const m=due.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);return m?`${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`:'';}
  function isOverdue(t){if(t.recurrence)return false;const norm=normalizeDue(t.due);return !t.done&&!!norm&&norm<todayStr();}
  function isToday(t){if(t.todayFlag==='FALSE')return false;if(t.todayFlag==='TRUE')return true;return CFG.dueDateImpliesToday ? normalizeDue(t.due)===todayStr() : false;}

  function openTaskModal(task){detailTask=task;document.getElementById('task-modal-title').textContent=task.task;document.getElementById('task-modal-meta').textContent=[labelTag(task.tag),task.priority,task.due?`Due ${task.due}`:'',task.recurrence?`Recurring ${task.recurrence}`:''].filter(Boolean).join(' · ');document.getElementById('task-modal-notes').value=task.notes||'';document.getElementById('task-modal').style.display='flex';}
  function closeTaskModal(){document.getElementById('task-modal').style.display='none';detailTask=null;}

  function render(){if(activeTab==='board'){renderBoard();return;}const base=tasks.length?tasks:loadCache();const list=base.filter(recurringVisible);const container=document.getElementById('tasks');container.innerHTML='';const todayItems=list.filter(isToday),workItems=list.filter(t=>t.tag.trim().startsWith('W')&&!isToday(t)),personalItems=list.filter(t=>t.tag.trim().startsWith('P')&&!isToday(t)),otherItems=list.filter(t=>!t.tag.trim().startsWith('W')&&!t.tag.trim().startsWith('P')&&!isToday(t));
    function taskRow(t){const recurringDone=t.recurrence&&recurringCompletedToday(t);const checked=t.recurrence?recurringDone:t.done;const progress=t.recurrence?recurringProgress(t):'';const row=document.createElement('div');row.className='item'+(checked?' done':'');row.innerHTML=`<input type="checkbox" ${checked?'checked':''}/><div class="content"><div class="label">${escapeHtml(t.task)}${t.notes?' <span class="note-flag" title="Has notes">📝</span>':''}</div><div class="proj">${escapeHtml(labelTag(t.tag))} · <span class="priority-tag">${escapeHtml(t.priority||'Medium')}</span>${progress?` · <span class="recurring-progress">${escapeHtml(progress)}</span>`:''}</div></div><div class="due-col${(!t.recurrence&&isOverdue(t))?' overdue':''}">${escapeHtml(t.recurrence?'':(t.due||''))}</div><button type="button" class="today-btn ${isToday(t)?'active':''}" aria-label="Toggle Today"></button>`;
      row.querySelector('input').addEventListener('click',e=>e.stopPropagation());row.querySelector('input').addEventListener('change',e=>toggleDone(t,e.target.checked));row.querySelector('.today-btn').addEventListener('click',e=>{e.preventDefault();e.stopPropagation();toggleTodayShared(t,!isToday(t));});row.querySelector('.content').addEventListener('click',e=>{e.preventDefault();e.stopPropagation();openTaskModal(t);});row.addEventListener('click',()=>openTaskModal(t));return row;}
    const subhead=t=>{const h=document.createElement('div');h.className='subhead';h.textContent=t;return h;};const domainHead=t=>{const h=document.createElement('div');h.className='group-head';h.textContent=t;return h;};
    const byDate=items=>[...items].sort((a,b)=>{if(a.done!==b.done)return a.done?1:-1;const da=normalizeDue(a.due),db=normalizeDue(b.due);if(!da&&!db)return 0;if(!da)return 1;if(!db)return-1;return da.localeCompare(db);});const byPriority=items=>{const order={High:0,Medium:1,Low:2};return [...items].sort((a,b)=>{if(a.done!==b.done)return a.done?1:-1;const p=(order[a.priority]??1)-(order[b.priority]??1);if(p)return p;const da=normalizeDue(a.due),db=normalizeDue(b.due);if(!da&&!db)return 0;if(!da)return 1;if(!db)return-1;return da.localeCompare(db);});};
    function renderGrouped(items){if(sortMode==='category'){const groups={};items.forEach(t=>(groups[getCategoryLabel(t.tag)]??=[]).push(t));Object.keys(groups).sort().forEach(cat=>{container.appendChild(subhead(cat));byPriority(groups[cat]).forEach(t=>container.appendChild(taskRow(t)));});}else{['High','Medium','Low'].forEach(tier=>{const x=items.filter(t=>(t.priority||'Medium')===tier);if(x.length){container.appendChild(subhead(tier));byDate(x).forEach(t=>container.appendChild(taskRow(t)));}});}}
    function section(title,items){if(!items.length)return;container.appendChild(domainHead(title));renderGrouped(items);}
    function todaySection(items){if(!items.length)return;container.appendChild(domainHead('Today'));const work=items.filter(t=>t.tag.trim().startsWith('W')),personal=items.filter(t=>t.tag.trim().startsWith('P')),other=items.filter(t=>!t.tag.trim().startsWith('W')&&!t.tag.trim().startsWith('P'));const ordered=domainFirst==='Personal'?[['Personal',personal],['Work',work]]:[['Work',work],['Personal',personal]];ordered.forEach(([label,group])=>{if(!group.length)return;container.appendChild(subhead(label));renderGrouped(group);});if(other.length){container.appendChild(subhead('Other'));renderGrouped(other);}}
    todaySection(todayItems);if(domainFirst==='Personal'){section('Personal',personalItems);section('Work',workItems);}else{section('Work',workItems);section('Personal',personalItems);}section('Other',otherItems);if(!list.length)container.innerHTML='<p class="empty">No tasks loaded yet. Connect to sync.</p>';
  }

  function yearProgressPct(dateStr){if(!dateStr)return null;const d=new Date(dateStr+(dateStr.length===10?'T12:00:00':''));if(Number.isNaN(d.getTime()))return null;const start=new Date(d.getFullYear(),0,1),end=new Date(d.getFullYear(),11,31);return Math.round(((d-start)/(end-start))*100);}
  function getTopLevelCodes(root){return Object.keys(taxonomyParent).filter(code=>taxonomyParent[code]===root).sort();}
  function renderBoard(){const container=document.getElementById('tasks');container.innerHTML='';const list=board.length?board:JSON.parse(localStorage.getItem(BOARD_CACHE_KEY)||'[]');function card(p){const el=document.createElement('div');el.className='board-card priority-'+(p.priority||'Medium').toLowerCase()+(p.stage==='Untracked'?' untracked':'');const pct=yearProgressPct(p.milestoneDate),codeLabel=taxonomy[p.code]?' · '+taxonomy[p.code]:'';el.innerHTML=`<div class="board-head"><span class="board-code">${escapeHtml(p.code+codeLabel)}</span><span class="board-stage">${escapeHtml(p.stage)}</span></div><div class="board-name">${escapeHtml(p.name)}</div>${p.nextUp?`<div class="board-next">${escapeHtml(p.nextUp)}</div>`:''}${pct!==null?`<div class="board-track"><div class="board-marker" style="left:${pct}%"></div></div><div class="board-date">${escapeHtml(p.milestoneDate)}</div>`:''}${p.notes?`<div class="board-notes">${escapeHtml(p.notes)}</div>`:''}`;return el;}function section(title,root){let items;if(boardMode==='all'){const tracked={};list.forEach(p=>{if(!tracked[p.code])tracked[p.code]=p;});items=getTopLevelCodes(root).map(code=>tracked[code]||{code,name:taxonomy[code]||code,domain:root,stage:'Untracked',priority:'Low',nextUp:'',notes:'',milestoneDate:'',updated:'',show:true});}else items=list.filter(p=>p.domain===root&&p.show);if(!items.length)return;const h=document.createElement('div');h.className='group-head';h.textContent=title;container.appendChild(h);const order={Milestone:0,Active:1,Scoping:2,Waiting:3,'Wrapping up':4,Monitor:5,'On-hold':6,Untracked:7};items.sort((a,b)=>(order[a.stage]??5)-(order[b.stage]??5)||a.code.localeCompare(b.code,undefined,{numeric:true}));items.forEach(p=>container.appendChild(card(p)));}section('Work','W');section('Personal','P');if(!(boardMode==='all'?(getTopLevelCodes('W').length||getTopLevelCodes('P').length):list.some(p=>p.show)))container.innerHTML='<p class="empty">No board data loaded yet.</p>';}

  function setTab(tab){activeTab=tab;document.getElementById('tab-tasks').classList.toggle('active',tab==='tasks');document.getElementById('tab-board').classList.toggle('active',tab==='board');document.getElementById('sort-controls').style.display=tab==='tasks'?'flex':'none';document.getElementById('board-controls').style.display=tab==='board'?'flex':'none';render();}

  function bindEvents(){document.getElementById('connect-btn').addEventListener('click',connect);document.getElementById('tab-tasks').addEventListener('click',()=>setTab('tasks'));document.getElementById('tab-board').addEventListener('click',()=>setTab('board'));document.getElementById('sort-domain').addEventListener('click',()=>{domainFirst=domainFirst==='Work'?'Personal':'Work';document.getElementById('sort-domain').textContent=domainFirst+' first';render();});['category','priority'].forEach(mode=>document.getElementById('sort-'+mode).addEventListener('click',()=>{sortMode=mode;['category','priority'].forEach(m=>document.getElementById('sort-'+m).classList.toggle('active',m===mode));render();}));['tracked','all'].forEach(mode=>document.getElementById('board-'+mode).addEventListener('click',()=>{boardMode=mode;['tracked','all'].forEach(m=>document.getElementById('board-'+m).classList.toggle('active',m===mode));render();}));document.getElementById('task-modal-close').addEventListener('click',closeTaskModal);document.getElementById('task-modal').addEventListener('click',e=>{if(e.target.id==='task-modal')closeTaskModal();});document.getElementById('task-notes-save').addEventListener('click',async()=>{const task=detailTask;if(!task)return;const text=document.getElementById('task-modal-notes').value;const ok=await saveTaskNotes(task,text);if(ok&&detailTask===task)document.getElementById('task-modal-notes').value=task.notes||'';});}

  function boot(){injectShell();if(new Date().getDay()===0||new Date().getDay()===6){domainFirst='Personal';document.getElementById('sort-domain').textContent='Personal first';}tasks=loadCache();const c=loadTaxonomyCache();taxonomy=c.labels||{};taxonomyParent=c.parents||{};board=JSON.parse(localStorage.getItem(BOARD_CACHE_KEY)||'[]');render();bindEvents();if(typeof google==='undefined'||!google.accounts?.oauth2)setStatus('Google sign-in library did not load. Refresh when online.');else initAuth();}

  window.addEventListener('online',()=>{setStatus('Back online, syncing…');if(accessToken)Promise.all([fetchTaxonomy(),fetchBoard(),fetchRecurringCompletions()]).then(fetchTasks);});window.addEventListener('offline',()=>setStatus('Offline — showing cached data.'));
  if('serviceWorker'in navigator)window.addEventListener('load',()=>navigator.serviceWorker.register('./service-worker.js'));
  window.addEventListener('load',boot);
})();
