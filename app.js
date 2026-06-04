/* =============================================
   HOLOS Equipment Check — app.js  v2.0
   Two-way Google Sheets sync via Apps Script
   ============================================= */

'use strict';

// ── Constants ────────────────────────────────
const STORAGE_KEY        = 'holos_equipment_log';
const PW_KEY             = 'holos_admin_pw';
const SCRIPT_URL_KEY     = 'holos_script_url';
const DEFAULT_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby6vl87SRH-CtrFW1LSNGhb3bpHaoB9ucxUeNarSsGt2YEabIT7m9vNQ_w82ZVUJdLG/exec';
const EQ_CACHE_KEY       = 'holos_equipment_cache';
const USER_KEY           = 'holos_current_user';
const USERS_CACHE_KEY    = 'holos_users_cache';
const DEFAULT_PW         = 'holos2024';

// ── Storage helpers ───────────────────────────
const loadLog         = ()  => { try { return JSON.parse(localStorage.getItem(STORAGE_KEY)) || []; } catch { return []; } };
const saveLog         = (d) => localStorage.setItem(STORAGE_KEY, JSON.stringify(d));
const getPassword     = ()  => localStorage.getItem(PW_KEY) || DEFAULT_PW;
const setPassword     = (p) => localStorage.setItem(PW_KEY, p);
const getScriptUrl    = ()  => localStorage.getItem(SCRIPT_URL_KEY) || DEFAULT_SCRIPT_URL;
const setScriptUrl    = (u) => localStorage.setItem(SCRIPT_URL_KEY, u);
const getEqCache      = ()  => { try { return JSON.parse(localStorage.getItem(EQ_CACHE_KEY)) || []; } catch { return []; } };
const saveEqCache     = (d) => localStorage.setItem(EQ_CACHE_KEY, JSON.stringify(d));
const getCurrentUser  = ()  => { try { return JSON.parse(localStorage.getItem(USER_KEY)); } catch { return null; } };
const saveCurrentUser = (u) => localStorage.setItem(USER_KEY, JSON.stringify(u));
const getUsersCache   = ()  => { try { return JSON.parse(localStorage.getItem(USERS_CACHE_KEY)) || []; } catch { return []; } };
const saveUsersCache  = (u) => localStorage.setItem(USERS_CACHE_KEY, JSON.stringify(u));

// ── State ─────────────────────────────────────
let log           = loadLog();
let equipmentList = getEqCache();
let eqQueue       = [];
let cameraStream  = null;
let scanInterval  = null;
let pwMode        = 'clear';
let currentUser   = getCurrentUser();
let usersList     = getUsersCache();

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
const navLogControl  = $('nav-log-control');

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
const remarks        = $('remarks');
const btnIn          = $('btn-in');
const btnOut         = $('btn-out');
const toast          = $('toast');

// User selector
const userCard          = $('user-card');
const userAvatar        = $('user-avatar');
const userCardName      = $('user-card-name');
const userCardContact   = $('user-card-contact');
const btnSwitchUser     = $('btn-switch-user');
const btnSelectUser     = $('btn-select-user');

// User modal
const userModal         = $('user-modal');
const btnCloseUserModal = $('btn-close-user-modal');
const userChips         = $('user-chips');
const userChipsEmpty    = $('user-chips-empty');
const regUserForm       = $('reg-user-form');
const regUserName       = $('reg-user-name');
const regUserContact    = $('reg-user-contact');
const regUserError      = $('reg-user-error');
const btnSubmitRegUser  = $('btn-submit-reg-user');
const btnToggleRegUser  = $('btn-toggle-reg-user');

const logBody        = $('log-body');
const logEmpty       = $('log-empty');
const logLoading     = $('log-loading');
const logCount       = $('log-count');
const logSearch      = $('log-search');
const logFilter      = $('log-filter');
const btnRefreshLog  = $('btn-refresh-log');

const scriptUrlInput  = $('script-url');
const btnTestConn     = $('btn-test-conn');
const connResult      = $('conn-result');
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

// ── Checked-out map (name → { operator }) ────
function getCheckedOutMap() {
  const seen = new Set();
  const map  = new Map();
  for (const entry of log) {
    if (!seen.has(entry.equipment)) {
      seen.add(entry.equipment);
      if (entry.status === 'OUT')
        map.set(entry.equipment, { operator: entry.operator, timestamp: entry.timestamp });
    }
  }
  return map;
}
function getCheckedOutEquipment() { return new Set(getCheckedOutMap().keys()); }

// ── Utilities ──────────────────────────────────
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

// ── Sync badge ────────────────────────────────
function setSyncBadge(state, label) {
  syncBadge.className = 'sync-badge sync-' + state;
  syncLabel.textContent = label;
}

// ── User display ──────────────────────────────
function initials(name) {
  return name.trim().split(/\s+/).map(w => w[0]).join('').toUpperCase().slice(0, 2);
}

function updateUserDisplay() {
  if (currentUser) {
    userCard.classList.remove('hidden');
    btnSelectUser.classList.add('hidden');
    userAvatar.textContent      = initials(currentUser.name);
    userCardName.textContent    = currentUser.name;
    userCardContact.textContent = currentUser.contact;
  } else {
    userCard.classList.add('hidden');
    btnSelectUser.classList.remove('hidden');
  }
}

// ── User Modal ────────────────────────────────
function openUserModal() {
  regUserForm.classList.add('hidden');
  btnToggleRegUser.textContent = '＋ Register New User';
  regUserName.value = ''; regUserContact.value = '';
  regUserError.classList.add('hidden');
  renderUserChips();
  userModal.classList.remove('hidden');
}
function closeUserModal() { userModal.classList.add('hidden'); }

function renderUserChips() {
  if (usersList.length === 0) {
    userChips.innerHTML = '';
    userChipsEmpty.classList.remove('hidden');
    return;
  }
  userChipsEmpty.classList.add('hidden');
  userChips.innerHTML = usersList.map((u, i) => `
    <button class="user-chip" data-i="${i}" aria-label="Select ${escHtml(u.name)}">
      <span class="user-chip__avatar">${escHtml(initials(u.name))}</span>
      <span class="user-chip__info">
        <span class="user-chip__name">${escHtml(u.name)}</span>
        <span class="user-chip__contact">${escHtml(u.contact)}</span>
      </span>
    </button>
  `).join('');
  userChips.querySelectorAll('.user-chip').forEach(btn => {
    btn.addEventListener('click', () => {
      currentUser = usersList[Number(btn.dataset.i)];
      saveCurrentUser(currentUser);
      updateUserDisplay();
      closeUserModal();
      revealFields();
    });
  });
}

btnSelectUser.addEventListener('click', () => { openUserModal(); revealFields(); });
btnSwitchUser.addEventListener('click', openUserModal);
btnCloseUserModal.addEventListener('click', closeUserModal);

btnToggleRegUser.addEventListener('click', () => {
  const hidden = regUserForm.classList.toggle('hidden');
  btnToggleRegUser.textContent = hidden ? '＋ Register New User' : '− Cancel';
  if (!hidden) setTimeout(() => regUserName.focus(), 100);
});

btnSubmitRegUser.addEventListener('click', async () => {
  const name    = regUserName.value.trim();
  const contact = regUserContact.value.trim();
  if (!name || !contact) { regUserError.classList.remove('hidden'); return; }
  regUserError.classList.add('hidden');

  const newUser = { name, contact };
  btnSubmitRegUser.textContent = 'Saving…'; btnSubmitRegUser.disabled = true;

  if (!usersList.find(u => u.name.toLowerCase() === name.toLowerCase())) {
    usersList.push(newUser);
    saveUsersCache(usersList);
  }

  registerUserToSheet(name, contact).catch(err => console.warn('User sheet sync failed:', err));

  currentUser = newUser;
  saveCurrentUser(currentUser);
  updateUserDisplay();
  closeUserModal();
  revealFields();

  btnSubmitRegUser.textContent = 'Register & Select'; btnSubmitRegUser.disabled = false;
});

[regUserName, regUserContact].forEach(inp => {
  inp.addEventListener('keydown', e => { if (e.key === 'Enter') btnSubmitRegUser.click(); });
});

async function loadUsers() {
  usersList = getUsersCache();
  if (!getScriptUrl()) return;
  try {
    const data = await sheetsGet('getUsers');
    if (data.users && Array.isArray(data.users)) {
      usersList = data.users;
      saveUsersCache(usersList);
    }
  } catch (e) { console.warn('Users fetch failed:', e); }
}

async function registerUserToSheet(name, contact) {
  const url = getScriptUrl();
  if (!url) return;
  await fetch(url, {
    method: 'POST', redirect: 'follow',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({ action: 'registerUser', name, contact })
  });
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
  const checkedOutMap = getCheckedOutMap();
  const filtered = list.filter(e => !q || e.name.toLowerCase().includes(q) || e.category.toLowerCase().includes(q));

  if (list.length === 0) { eqPickerWrap.classList.add('hidden'); return; }
  eqPickerWrap.classList.remove('hidden');

  if (filtered.length === 0) {
    eqPickerGrid.innerHTML = '<p class="picker-empty">No equipment found</p>';
    return;
  }

  eqPickerGrid.innerHTML = filtered.map(e => {
    const info  = checkedOutMap.get(e.name);
    const isOut = !!info;
    const byStr = isOut ? info.operator : '';
    return `
      <button class="picker-chip${isOut ? ' picker-chip--out' : ''}"
              data-name="${escHtml(e.name)}"
              aria-label="${isOut ? `Out with ${escHtml(byStr)}: ` : 'Add '}${escHtml(e.name)}">
        <span class="picker-chip__name">${escHtml(e.name)}</span>
        ${e.category ? `<span class="picker-chip__cat">${escHtml(e.category)}</span>` : ''}
        ${isOut ? `<span class="picker-chip__out-badge">OUT &middot; <span class="out-by">${escHtml(byStr)}</span></span>` : ''}
      </button>
    `;
  }).join('');

  eqPickerGrid.querySelectorAll('.picker-chip').forEach(btn => {
    btn.addEventListener('click', () => {
      const name = btn.dataset.name;
      const info = checkedOutMap.get(name);
      if (info) {
        if (currentUser && currentUser.name === info.operator) {
          addToQueue(name); revealFields();
        } else {
          showToast(`⛔ ${name} is out with ${info.operator}`, 'red');
        }
        return;
      }
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
  allPages.forEach(p => p.classList.remove('active'));
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
navSettings.addEventListener('click',   e => { e.preventDefault(); showPage('settings'); });
navChangePass.addEventListener('click', e => { e.preventDefault(); openPwModal('change'); closeMenu(); });
navClear.addEventListener('click',      e => { e.preventDefault(); openPwModal('clear');  closeMenu(); });
navLogControl.addEventListener('click', e => { e.preventDefault(); openPwModal('log-control'); closeMenu(); });

btnRefreshLog.addEventListener('click', () => fetchLiveLog().then(() => renderLog()));

// ── Settings ──────────────────────────────────
btnTestConn.addEventListener('click', async () => {
  const url = scriptUrlInput.value.trim();
  if (!url) { showConnResult('⚠️  Please enter a URL first', false); return; }
  btnTestConn.textContent = '⏳ Testing…'; btnTestConn.disabled = true;
  try {
    const res  = await fetch(`${url}?action=ping`, { redirect: 'follow' });
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
  connResult.textContent = msg;
  connResult.className   = 'conn-result ' + (ok ? 'ok' : 'err');
}

btnSaveSettings.addEventListener('click', () => {
  const url = scriptUrlInput.value.trim();
  setScriptUrl(url);
  showToast('✅ Settings saved');
  loadEquipmentList();
  showPage('tracker');
});

// ── Reveal fields ─────────────────────────────
function revealFields() { entryFields.classList.remove('hidden'); }

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

  const checkedOutMap = getCheckedOutMap();
  const info = checkedOutMap.get(name);
  if (info) {
    if (!currentUser || currentUser.name !== info.operator) {
      showToast(`⛔ ${name} is out with ${info.operator} — only they can return it`, 'red');
      return;
    }
  }

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

  if (!currentUser) {
    openUserModal();
    showToast('👤 Please select your name first');
    return;
  }

  const raw = equipmentId.value.trim();
  if (raw) addToQueue(raw);
  if (eqQueue.length === 0) { equipmentId.focus(); showToast('⚠️  Add at least one equipment to the queue'); return; }

  const checkedOutMap = getCheckedOutMap();

  if (status === 'IN') {
    for (const name of eqQueue) {
      const info = checkedOutMap.get(name);
      if (info && info.operator !== currentUser.name) {
        showToast(`⛔ ${name} was taken by ${info.operator} — only they can return it`, 'red');
        return;
      }
    }
  }

  const now   = new Date().toISOString();
  const note  = remarks.value.trim();
  const count = eqQueue.length;

  const newEntries = eqQueue.map(name => ({
    id:        Date.now() + Math.random(),
    equipment: name, status,
    operator:  currentUser.name,
    notes:     note, timestamp: now
  }));

  log.unshift(...newEntries);
  saveLog(log);
  pushToSheets(newEntries);
  renderPicker(equipmentList, pickerSearch.value);

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
[remarks].forEach(inp => inp.addEventListener('keydown', e => { if (e.key === 'Enter') btnIn.click(); }));

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
  } else if (mode === 'log-control') {
    pwModalTitle.textContent = '🔐 Log Control';
    pwModalHint.textContent  = 'Enter the admin password to manage equipment logs.';
    newPwFields.classList.add('hidden'); btnPwSubmit.textContent = 'Open Log Control';
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
  } else if (pwMode === 'log-control') {
    closePwModal(); openLogControlModal();
  } else {
    const np = newPwInput.value, cp = confirmPwInput.value;
    if (!np || np !== cp) { pwMatchError.classList.remove('hidden'); newPwInput.focus(); return; }
    setPassword(np); closePwModal(); showToast('🔑 Password updated');
  }
});
[pwInput, newPwInput, confirmPwInput].forEach(inp => inp.addEventListener('keydown', e => { if (e.key === 'Enter') btnPwSubmit.click(); }));

// ── Log Control Modal ──────────────────────────────
const logControlModal = $('log-control-modal');
const logControlList  = $('log-control-list');
const logControlEmpty = $('log-control-empty');
const btnCloseLC      = $('btn-close-log-control');

function openLogControlModal() {
  renderLogControlModal();
  logControlModal.classList.remove('hidden');
}
function closeLogControlModal() { logControlModal.classList.add('hidden'); }
btnCloseLC.addEventListener('click', closeLogControlModal);
overlay.addEventListener('click', () => closeLogControlModal());

function renderLogControlModal() {
  const checkedOutMap = getCheckedOutMap();
  if (checkedOutMap.size === 0) {
    logControlList.innerHTML = '';
    logControlEmpty.classList.remove('hidden');
    return;
  }
  logControlEmpty.classList.add('hidden');
  logControlList.innerHTML = [...checkedOutMap.entries()].map(([name, info]) => `
    <div class="lc-item" id="lc-${escHtml(name)}">
      <div class="lc-info">
        <span class="lc-name">${escHtml(name)}</span>
        <span class="lc-meta">Out with <strong>${escHtml(info.operator)}</strong> &middot; ${formatDateTime(info.timestamp)}</span>
      </div>
      <button class="lc-btn" data-name="${escHtml(name)}">&#x21a9; Force Return</button>
    </div>
  `).join('');

  logControlList.querySelectorAll('.lc-btn').forEach(btn => {
    btn.addEventListener('click', () => forceReturnEquipment(btn.dataset.name, btn));
  });
}

function forceReturnEquipment(name, btn) {
  btn.textContent = 'Returning…'; btn.disabled = true;
  const entry = {
    id:        Date.now() + Math.random(),
    equipment: name,
    status:    'IN',
    operator:  currentUser ? currentUser.name : 'Admin',
    notes:     'Force returned via Log Control',
    timestamp: new Date().toISOString()
  };
  log.unshift(entry);
  saveLog(log);
  pushToSheets([entry]);
  renderPicker(equipmentList, pickerSearch.value);
  showToast(`↩ ${name} force-returned to IN`);
  // Remove row from modal
  const row = document.getElementById(`lc-${name}`);
  if (row) row.remove();
  // If none left, show empty state
  if (logControlList.children.length === 0) {
    logControlEmpty.classList.remove('hidden');
  }
}

// ── Init ──────────────────────────────────────
showPage('tracker');
loadEquipmentList();
loadUsers();
updateUserDisplay();
