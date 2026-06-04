/* =============================================
   HOLOS Equipment Check — app.js  v2.0
   Two-way Google Sheets sync via Apps Script
   ============================================= */

'use strict';

// ── Constants ────────────────────────────────
const STORAGE_KEY    = 'holos_equipment_log';
const PW_KEY         = 'holos_admin_pw';
const SCRIPT_URL_KEY = 'holos_script_url';
const EQ_CACHE_KEY   = 'holos_equipment_cache';
const DEFAULT_PW     = 'holos2024';

// ── Storage helpers ───────────────────────────
const loadLog      = ()  => { try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || []; } catch { return []; } };
const saveLog      = (d) => localStorage.setItem(STORAGE_KEY, JSON.stringify(d));
const getPassword  = ()  => localStorage.getItem(PW_KEY) || DEFAULT_PW;
const setPassword  = (p) => localStorage.setItem(PW_KEY, p);
const getScriptUrl = ()  => localStorage.getItem(SCRIPT_URL_KEY) || '';
const setScriptUrl = (u) => localStorage.setItem(SCRIPT_URL_KEY, u);
const getEqCache   = ()  => { try { return JSON.parse(localStorage.getItem(EQ_CACHE_KEY)) || []; } catch { return []; } };
const saveEqCache  = (d) => localStorage.setItem(EQ_CACHE_KEY, JSON.stringify(d));

// ── State ─────────────────────────────────────
let log          = loadLog();
let equipmentList = getEqCache();   // [{name, category}]
let eqQueue      = [];
let cameraStream = null;
let scanInterval = null;
let pwMode       = 'clear';

// ── DOM refs ──────────────────────────────────
const $ = (id) => document.getElementById(id);

const btnMenu        = $('btn-menu');
const btnCloseMenu   = $('btn-close-menu');
const overlay        = $('overlay');
const sideMenu       = $('side-menu');
const syncBadge      = $('sync-badge');
const syncLabel      = syncBadge.querySelector('.sync-label');

const navTracker     = $('nav-tracker');
const navLog         = $('nav-log');
const navExport      = $('nav-export');
const navSettings    = $('nav-settings');
const navChangePass  = $('nav-change-pass');
const navClear       = $('nav-clear');

const pageTracker    = $('page-tracker');
const pageLog        = $('page-log');
const pageSettings   = $('page-settings');

const btnScan        = $('btn-scan');
const eqPickerWrap   = $('eq-picker-wrap');
const eqPickerGrid   = $('eq-picker-grid');
const pickerSearch   = $('picker-search');
const entryFields    = $('entry-fields');
const equipmentId    = $('equipment-id');
const btnAddEq       = $('btn-add-eq');
const eqQueueWrap    = $('eq-queue-wrap');
const eqQueueEl      = $('eq-queue');
const queueCount     = $('queue-count');
const operatorName   = $('operator-name');
const remarks        = $('remarks');
const btnIn          = $('btn-in');
const btnOut         = $('btn-out');
const toast          = $('toast');

const logBody        = $('log-body');
const logEmpty       = $('log-empty');
const logLoading     = $('log-loading');
const logCount       = $('log-count');
const logSearch      = $('log-search');
const logFilter      = $('log-filter');
const btnRefreshLog  = $('btn-refresh-log');

const scriptUrlInput = $('script-url');
const btnTestConn    = $('btn-test-conn');
const connResult     = $('conn-result');
const btnSaveSettings = $('btn-save-settings');

const cameraModal    = $('camera-modal');
const cameraFeed     = $('camera-feed');
const cameraCanvas   = $('camera-canvas');
const btnCloseCamera = $('btn-close-camera');
const btnManual      = $('btn-manual-entry');

const passwordModal  = $('password-modal');
const pwModalTitle   = $('pw-modal-title');
const pwModalHint    = $('pw-modal-hint');
const pwInput        = $('pw-input');
const btnPwToggle    = $('btn-pw-toggle');
const pwError        = $('pw-error');
const newPwFields    = $('new-pw-fields');
const newPwInput     = $('new-pw-input');
const confirmPwInput = $('confirm-pw-input');
const pwMatchError   = $('pw-match-error');
const btnPwSubmit    = $('btn-pw-submit');
const btnClosePw     = $('btn-close-pw');

// ── Utilities ─────────────────────────────────
function formatDateTime(iso) {
  const d = new Date(iso);
  return d.toLocaleDateString('en-PH', { month: 'short', day: 'numeric', year: 'numeric' })
    + ' ' + d.toLocaleTimeString('en-PH', { hour: '2-digit', minute: '2-digit' });
}
function escHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
let toastTimer;
function showToast(msg, type = '') {
  toast.textContent = msg;
  toast.className = 'toast show ' + type;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { toast.className = 'toast'; }, 2800);
}

// ── Sync badge ───────────────────────────────
function setSyncBadge(state, label) {
  syncBadge.className = 'sync-badge sync-' + state;
  syncLabel.textContent = label;
}

// ── Google Sheets API ─────────────────────────
async function sheetsGet(action) {
  const url = getScriptUrl();
  if (!url) throw new Error('No script URL configured');
  const res = await fetch(`${url}?action=${action}`, { redirect: 'follow' });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

async function sheetsPost(entries) {
  const url = getScriptUrl();
  if (!url) throw new Error('No script URL configured');
  // Apps Script POST — send as text/plain to avoid CORS preflight
  const res = await fetch(url, {
    method:   'POST',
    redirect: 'follow',
    headers:  { 'Content-Type': 'text/plain' },
    body:     JSON.stringify({ entries })
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

// ── Load equipment list from Sheets ──────────
async function loadEquipmentList() {
  if (!getScriptUrl()) {
    equipmentList = getEqCache();
    renderPicker(equipmentList);
    return;
  }
  try {
    setSyncBadge('busy', 'Loading…');
    const data = await sheetsGet('getEquipment');
    if (data.error) throw new Error(data.error);
    equipmentList = data.equipment || [];
    saveEqCache(equipmentList);
    renderPicker(equipmentList);
    setSyncBadge('ok', 'Connected');
  } catch (err) {
    console.warn('Equipment list fetch failed:', err);
    equipmentList = getEqCache();
    renderPicker(equipmentList);
    setSyncBadge('err', 'Offline');
  }
}

// ── Push entries to Sheets ───────────────────
async function pushToSheets(entries) {
  if (!getScriptUrl()) return;
  setSyncBadge('busy', 'Syncing…');
  try {
    await sheetsPost(entries);
    setSyncBadge('ok', 'Synced ✓');
    setTimeout(() => setSyncBadge('ok', 'Connected'), 2000);
  } catch (err) {
    console.warn('Sheets push failed:', err);
    setSyncBadge('err', 'Sync failed');
    showToast('⚠️  Could not sync to Sheets — saved locally');
  }
}

// ── Fetch live log from Sheets ───────────────
async function fetchLiveLog() {
  if (!getScriptUrl()) return null;
  logLoading.classList.remove('hidden');
  logEmpty.style.display = 'none';
  logBody.innerHTML = '';
  btnRefreshLog.classList.add('spinning');

  try {
    const data = await sheetsGet('getLog');
    if (data.error) throw new Error(data.error);
    // Merge: Sheets log wins (it's the source of truth)
    log = data.log.map(e => ({
      id:        e.id,
      equipment: e.equipment,
      status:    e.status,
      operator:  e.operator,
      notes:     e.notes || '',
      timestamp: e.timestamp
    }));
    saveLog(log);
    setSyncBadge('ok', 'Connected');
    return log;
  } catch (err) {
    console.warn('Log fetch failed:', err);
    setSyncBadge('err', 'Offline');
    showToast('⚠️  Could not reach Sheets — showing local log');
    return null;
  } finally {
    logLoading.classList.add('hidden');
    btnRefreshLog.classList.remove('spinning');
  }
}

// ── Equipment Picker ─────────────────────────
function renderPicker(list, query = '') {
  const q = query.toLowerCase();
  const filtered = list.filter(e => !q || e.name.toLowerCase().includes(q) || e.category.toLowerCase().includes(q));

  if (list.length === 0) {
    eqPickerWrap.classList.add('hidden');
    return;
  }
  eqPickerWrap.classList.remove('hidden');

  if (filtered.length === 0) {
    eqPickerGrid.innerHTML = '<p class="picker-empty">No equipment found</p>';
    return;
  }

  eqPickerGrid.innerHTML = filtered.map(e => `
    <button class="picker-chip" data-name="${escHtml(e.name)}" aria-label="Add ${escHtml(e.name)}">
      <span class="picker-chip__name">${escHtml(e.name)}</span>
      ${e.category ? `<span class="picker-chip__cat">${escHtml(e.category)}</span>` : ''}
    </button>
  `).join('');

  eqPickerGrid.querySelectorAll('.picker-chip').forEach(btn => {
    btn.addEventListener('click', () => {
      const name = btn.dataset.name;
      addToQueue(name);
      btn.classList.add('selected');
      setTimeout(() => btn.classList.remove('selected'), 600);
      revealFields();
    });
  });
}

pickerSearch.addEventListener('input', () => renderPicker(equipmentList, pickerSearch.value));

// ── Menu ──────────────────────────────────────
function openMenu()  { sideMenu.classList.add('open'); overlay.classList.add('open'); sideMenu.removeAttribute('aria-hidden'); }
function closeMenu() { sideMenu.classList.remove('open'); overlay.classList.remove('open'); sideMenu.setAttribute('aria-hidden', 'true'); }
btnMenu.addEventListener('click', openMenu);
btnCloseMenu.addEventListener('click', closeMenu);
overlay.addEventListener('click', () => { closeMenu(); closePwModal(); });

// ── Navigation ────────────────────────────────
const allPages = [pageTracker, pageLog, pageSettings];

function showPage(pageId) {
  allPages.forEach(p => { p.classList.remove('active'); });
  document.querySelectorAll('.menu-link').forEach(l => l.classList.remove('active'));

  if (pageId === 'tracker') {
    pageTracker.classList.add('active'); navTracker.classList.add('active');
  } else if (pageId === 'log') {
    pageLog.classList.add('active'); navLog.classList.add('active');
    fetchLiveLog().then(() => renderLog());
  } else if (pageId === 'settings') {
    pageSettings.classList.add('active'); navSettings.classList.add('active');
    scriptUrlInput.value = getScriptUrl();
    connResult.className = 'conn-result hidden';
  }
  closeMenu();
}

navTracker.addEventListener('click',    e => { e.preventDefault(); showPage('tracker'); });
navLog.addEventListener('click',        e => { e.preventDefault(); showPage('log'); });
navExport.addEventListener('click',     e => { e.preventDefault(); exportCSV(); closeMenu(); });
navSettings.addEventListener('click',  e => { e.preventDefault(); showPage('settings'); });
navChangePass.addEventListener('click', e => { e.preventDefault(); openPwModal('change'); closeMenu(); });
navClear.addEventListener('click',      e => { e.preventDefault(); openPwModal('clear');  closeMenu(); });

btnRefreshLog.addEventListener('click', () => fetchLiveLog().then(() => renderLog()));

// ── Settings ──────────────────────────────────
btnTestConn.addEventListener('click', async () => {
  const url = scriptUrlInput.value.trim();
  if (!url) { showConnResult('⚠️  Please enter a URL first', false); return; }
  btnTestConn.textContent = '⏳ Testing…'; btnTestConn.disabled = true;
  try {
    const res = await fetch(`${url}?action=ping`, { redirect: 'follow' });
    const data = await res.json();
    if (data.ok) showConnResult('✅ Connection successful!', true);
    else         showConnResult('❌ Script responded but returned an error.', false);
  } catch (e) {
    showConnResult('❌ Could not reach the script. Check the URL and deployment settings.', false);
  } finally {
    btnTestConn.textContent = '🔌 Test Connection'; btnTestConn.disabled = false;
  }
});

function showConnResult(msg, ok) {
  connResult.textContent  = msg;
  connResult.className    = 'conn-result ' + (ok ? 'ok' : 'err');
}

btnSaveSettings.addEventListener('click', () => {
  const url = scriptUrlInput.value.trim();
  setScriptUrl(url);
  showToast('✅ Settings saved');
  loadEquipmentList();
  showPage('tracker');
});

// ── Reveal fields ─────────────────────────────
function revealFields() {
  entryFields.classList.remove('hidden');
}

// ── Scan button ───────────────────────────────
btnScan.addEventListener('click', () => {
  revealFields();
  if (typeof jsQR === 'undefined') { showToast('📝 Enter equipment details below'); return; }
  openCamera();
});

// ── Camera / QR ──────────────────────────────
async function openCamera() {
  cameraModal.classList.remove('hidden');
  try {
    cameraStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } }
    });
    cameraFeed.srcObject = cameraStream;
    await cameraFeed.play();
    startQRScan();
  } catch {
    closeCamera();
    showToast('📷 Camera unavailable — enter manually');
  }
}

function startQRScan() {
  const ctx = cameraCanvas.getContext('2d', { willReadFrequently: true });
  scanInterval = setInterval(() => {
    if (cameraFeed.readyState !== cameraFeed.HAVE_ENOUGH_DATA) return;
    cameraCanvas.width  = cameraFeed.videoWidth;
    cameraCanvas.height = cameraFeed.videoHeight;
    ctx.drawImage(cameraFeed, 0, 0);
    const imgData = ctx.getImageData(0, 0, cameraCanvas.width, cameraCanvas.height);
    const code = jsQR(imgData.data, imgData.width, imgData.height, { inversionAttempts: 'dontInvert' });
    if (code && code.data) handleQRResult(code.data);
  }, 200);
}

function handleQRResult(data) {
  closeCamera();
  addToQueue(data.trim());
  showToast(`✅ Scanned: ${data.trim()} — added to queue`);
}

function closeCamera() {
  clearInterval(scanInterval); scanInterval = null;
  if (cameraStream) { cameraStream.getTracks().forEach(t => t.stop()); cameraStream = null; }
  cameraFeed.srcObject = null;
  cameraModal.classList.add('hidden');
}

btnCloseCamera.addEventListener('click', closeCamera);
btnManual.addEventListener('click', () => { closeCamera(); equipmentId.focus(); });

// ── Equipment Queue ───────────────────────────
function addToQueue(name) {
  name = name.trim();
  if (!name) return;
  if (eqQueue.includes(name)) { showToast(`⚠️  "${name}" already in queue`); return; }
  eqQueue.push(name);
  renderQueue();
  equipmentId.value = '';
  equipmentId.focus();
}

function removeFromQueue(i) { eqQueue.splice(i, 1); renderQueue(); }

function renderQueue() {
  if (eqQueue.length === 0) { eqQueueWrap.classList.add('hidden'); return; }
  eqQueueWrap.classList.remove('hidden');
  queueCount.textContent = `${eqQueue.length} item${eqQueue.length > 1 ? 's' : ''}`;
  eqQueueEl.innerHTML = eqQueue.map((name, i) => `
    <span class="eq-chip">
      <span class="eq-chip__name" title="${escHtml(name)}">${escHtml(name)}</span>
      <button class="eq-chip__remove" data-i="${i}" aria-label="Remove ${escHtml(name)}">✕</button>
    </span>
  `).join('');
  eqQueueEl.querySelectorAll('.eq-chip__remove').forEach(btn =>
    btn.addEventListener('click', () => removeFromQueue(Number(btn.dataset.i)))
  );
}

btnAddEq.addEventListener('click', () => addToQueue(equipmentId.value));
equipmentId.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); addToQueue(equipmentId.value); } });

// ── IN / OUT ──────────────────────────────────
function logEntry(status) {
  if (entryFields.classList.contains('hidden')) {
    revealFields(); showToast(`📝 Add equipment then press ${status}`); return;
  }
  const raw = equipmentId.value.trim();
  if (raw) addToQueue(raw);
  if (eqQueue.length === 0) { equipmentId.focus(); showToast('⚠️  Add at least one equipment to the queue'); return; }
  const op = operatorName.value.trim();
  if (!op) { operatorName.focus(); showToast('⚠️  Please enter the operator name'); return; }

  const now    = new Date().toISOString();
  const note   = remarks.value.trim();
  const count  = eqQueue.length;

  const newEntries = eqQueue.map(name => ({
    id:        Date.now() + Math.random(),
    equipment: name, status, operator: op,
    notes:     note, timestamp: now
  }));

  log.unshift(...newEntries);
  saveLog(log);

  // Push to Sheets asynchronously
  pushToSheets(newEntries);

  const emoji = status === 'IN' ? '✅' : '📤';
  const label = count === 1 ? eqQueue[0] : `${count} items`;
  showToast(`${emoji}  ${label} marked ${status}`, status === 'IN' ? 'green' : 'red');

  const btn = status === 'IN' ? btnIn : btnOut;
  btn.style.transform = 'scale(0.9)';
  setTimeout(() => { btn.style.transform = ''; }, 160);

  eqQueue = [];
  renderQueue();
  equipmentId.value = '';
}

btnIn.addEventListener('click',  () => logEntry('IN'));
btnOut.addEventListener('click', () => logEntry('OUT'));
[operatorName, remarks].forEach(inp => inp.addEventListener('keydown', e => { if (e.key === 'Enter') btnIn.click(); }));

// ── Render Log ────────────────────────────────
function renderLog() {
  const query  = logSearch.value.toLowerCase();
  const filter = logFilter.value;
  const filtered = log.filter(e => {
    const mf = filter === 'all' || e.status === filter;
    const mq = !query || e.equipment.toLowerCase().includes(query) || e.operator.toLowerCase().includes(query) || (e.notes && e.notes.toLowerCase().includes(query));
    return mf && mq;
  });
  logCount.textContent = `${log.length} entr${log.length === 1 ? 'y' : 'ies'}`;
  if (filtered.length === 0) { logBody.innerHTML = ''; logEmpty.style.display = 'block'; return; }
  logEmpty.style.display = 'none';
  logBody.innerHTML = filtered.map((e, i) => `
    <tr>
      <td style="color:var(--text-muted);font-size:.75rem;">${filtered.length - i}</td>
      <td style="font-weight:700;">${escHtml(e.equipment)}</td>
      <td><span class="badge ${e.status.toLowerCase()}">${e.status}</span></td>
      <td>${escHtml(e.operator)}</td>
      <td style="color:var(--text-muted);font-size:.8rem;white-space:nowrap;">${formatDateTime(e.timestamp)}</td>
      <td style="color:var(--text-muted);font-size:.8rem;">${escHtml(e.notes || '—')}</td>
    </tr>
  `).join('');
}

logSearch.addEventListener('input', renderLog);
logFilter.addEventListener('change', renderLog);

// ── Export CSV ────────────────────────────────
function exportCSV() {
  if (log.length === 0) { showToast('ℹ️  No entries to export'); return; }
  const headers = ['#', 'Equipment', 'Status', 'Operator', 'Date & Time', 'Notes'];
  const rows = log.map((e, i) => [
    log.length - i,
    `"${e.equipment.replace(/"/g,'""')}"`,
    e.status,
    `"${e.operator.replace(/"/g,'""')}"`,
    formatDateTime(e.timestamp),
    `"${(e.notes||'').replace(/"/g,'""')}"`
  ].join(','));
  const csv  = [headers.join(','), ...rows].join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url; a.download = `holos-equipment-log-${new Date().toISOString().slice(0,10)}.csv`;
  a.click(); URL.revokeObjectURL(url);
  showToast('⬇️  CSV exported!');
}

// ── Password Modal ────────────────────────────
function openPwModal(mode) {
  pwMode = mode;
  pwInput.value = ''; newPwInput.value = ''; confirmPwInput.value = '';
  pwError.classList.add('hidden'); pwMatchError.classList.add('hidden');
  if (mode === 'clear') {
    pwModalTitle.textContent = '🔒 Admin Password Required';
    pwModalHint.textContent  = 'Enter the admin password to clear all logs.';
    newPwFields.classList.add('hidden'); btnPwSubmit.textContent = 'Clear All Logs';
  } else {
    pwModalTitle.textContent = '🔑 Change Password';
    pwModalHint.textContent  = 'Enter your current password, then set a new one.';
    newPwFields.classList.remove('hidden'); btnPwSubmit.textContent = 'Update Password';
  }
  passwordModal.classList.remove('hidden');
  setTimeout(() => pwInput.focus(), 100);
}

function closePwModal() { passwordModal.classList.add('hidden'); pwInput.value = ''; }

btnClosePw.addEventListener('click', closePwModal);

btnPwToggle.addEventListener('click', () => {
  const show = pwInput.type === 'password';
  pwInput.type = show ? 'text' : 'password';
  btnPwToggle.textContent = show ? '🙈' : '👁';
});

btnPwSubmit.addEventListener('click', () => {
  pwError.classList.add('hidden'); pwMatchError.classList.add('hidden');
  if (pwInput.value !== getPassword()) { pwError.classList.remove('hidden'); pwInput.focus(); return; }
  if (pwMode === 'clear') {
    log = []; saveLog(log); renderLog(); closePwModal(); showToast('🗑️  All logs cleared');
  } else {
    const np = newPwInput.value, cp = confirmPwInput.value;
    if (!np || np !== cp) { pwMatchError.classList.remove('hidden'); newPwInput.focus(); return; }
    setPassword(np); closePwModal(); showToast('🔑 Password updated');
  }
});
[pwInput, newPwInput, confirmPwInput].forEach(inp => inp.addEventListener('keydown', e => { if (e.key === 'Enter') btnPwSubmit.click(); }));

// ── Init ──────────────────────────────────────
showPage('tracker');
loadEquipmentList();
