// ═══════════════════════════════════════════════════════════════
//  AUTO-SAVE — writes to a shared CSV file in the planner folder
//
//  Flow:
//    1. User clicks "Set Master CSV" once → picks / creates a .csv file
//    2. The FileSystemFileHandle is stored in IndexedDB so it survives
//       page reloads (browser remembers the permission for the origin).
//    3. Every 30 s (and 3 s after each change) the current state is
//       written silently to that file — no dialogs, no downloads.
//    4. On startup, if a handle exists the user sees a restore-from-CSV
//       banner offering to reload the shared file.
//
//  All users point at the same .csv on a shared drive → true shared save.
// ═══════════════════════════════════════════════════════════════

const _AUTOSAVE_INTERVAL_MS = 30000;
const _IDB_NAME   = 'diaglo_planner';
const _IDB_STORE  = 'handles';
const _IDB_KEY    = 'masterCSV';

let _autoSaveTimer      = null;
let _masterFileHandle   = null;   // FileSystemFileHandle | null
let _changesSinceLastSave = false;

// ── IndexedDB helpers ─────────────────────────────────────────

function _openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(_IDB_NAME, 1);
    req.onupgradeneeded = (e) => e.target.result.createObjectStore(_IDB_STORE);
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror   = ()  => reject(req.error);
  });
}

async function _storeHandle(handle) {
  const db = await _openDB();
  return new Promise((resolve, reject) => {
    const tx  = db.transaction(_IDB_STORE, 'readwrite');
    const req = tx.objectStore(_IDB_STORE).put(handle, _IDB_KEY);
    req.onsuccess = resolve;
    req.onerror   = () => reject(req.error);
  });
}

async function _loadHandle() {
  try {
    const db = await _openDB();
    return new Promise((resolve) => {
      const tx  = db.transaction(_IDB_STORE, 'readonly');
      const req = tx.objectStore(_IDB_STORE).get(_IDB_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror   = () => resolve(null);
    });
  } catch (e) { return null; }
}

async function _clearHandle() {
  try {
    const db = await _openDB();
    return new Promise((resolve) => {
      const tx = db.transaction(_IDB_STORE, 'readwrite');
      tx.objectStore(_IDB_STORE).delete(_IDB_KEY);
      tx.oncomplete = resolve;
    });
  } catch (e) {}
}

// ── Core write ────────────────────────────────────────────────

async function _writeToHandle(handle) {
  const csv      = generateCSVString();
  const writable = await handle.createWritable();
  await writable.write(csv);
  await writable.close();
}

async function _doAutoSave() {
  if (!_changesSinceLastSave || !_masterFileHandle) return;
  try {
    // Verify we still have write permission
    const perm = await _masterFileHandle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted') {
      const req = await _masterFileHandle.requestPermission({ mode: 'readwrite' });
      if (req !== 'granted') { _updateAutoSaveUI(); return; }
    }
    await _writeToHandle(_masterFileHandle);
    _changesSinceLastSave = false;
    _updateAutoSaveUI();
  } catch (e) {
    console.warn('Auto-save write failed:', e);
    const s = document.getElementById('autosave-status');
    if (s) s.textContent = 'Save failed';
  }
}

// ── markChanged — called by pushUndo / CSV import ─────────────

function markChanged() {
  _changesSinceLastSave = true;
  clearTimeout(markChanged._debounce);
  markChanged._debounce = setTimeout(() => {
    if (_masterFileHandle) _doAutoSave();
  }, 3000);
}
markChanged._debounce = null;

// ── Set / clear master file ───────────────────────────────────

async function setMasterCSVFile() {
  if (!window.showSaveFilePicker) {
    alert('Your browser does not support the File System Access API.\nUse Chrome or Edge.');
    return;
  }
  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: 'planner_master.csv',
      types: [{ description: 'CSV Files', accept: { 'text/csv': ['.csv'] } }],
    });
    _masterFileHandle = handle;
    await _storeHandle(handle);
    // Write immediately so the file is created with current data
    _changesSinceLastSave = true;
    await _doAutoSave();
    _startTimer();
    _updateAutoSaveUI();
    dismissRestoreBanner();
  } catch (e) {
    if (e.name !== 'AbortError') console.error(e);
  }
}

async function clearMasterCSVFile() {
  _masterFileHandle = null;
  _stopTimer();
  await _clearHandle();
  _updateAutoSaveUI();
}

function _startTimer() {
  if (_autoSaveTimer) return;
  _autoSaveTimer = setInterval(_doAutoSave, _AUTOSAVE_INTERVAL_MS);
}

function _stopTimer() {
  clearInterval(_autoSaveTimer);
  _autoSaveTimer = null;
}

// ── UI ────────────────────────────────────────────────────────

function _updateAutoSaveUI() {
  const btn       = document.getElementById('autosave-btn');
  const statusEl  = document.getElementById('autosave-status');
  const label     = document.getElementById('autosave-label');
  const active    = !!_masterFileHandle;

  if (label) label.textContent = active ? 'Auto-save: ON' : 'Auto-save: OFF';
  if (btn) {
    btn.style.color      = active ? '#10b981' : '';
    btn.style.fontWeight = active ? '700'     : '';
    btn.title = active
      ? 'Click to turn off auto-save'
      : 'Click to set the shared master CSV file for auto-save';
  }

  if (statusEl) {
    if (!active) { statusEl.textContent = ''; return; }
    const name = _masterFileHandle.name || 'master.csv';
    statusEl.textContent = _changesSinceLastSave ? '● ' + name : '✓ ' + name;
  }
}

// ── Restore banner ────────────────────────────────────────────

function _showRestoreBanner(handle) {
  const banner = document.getElementById('autosave-restore-banner');
  if (!banner) return;
  banner.innerHTML =
    '<span style="flex:1">📂 Master CSV found: <strong>' + handle.name + '</strong> — load it now?</span>' +
    '<button class="success" onclick="restoreFromMasterCSV()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Load</button>' +
    '<button class="outline" onclick="dismissRestoreBanner()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Skip</button>';
  banner.style.display = 'flex';
}

async function restoreFromMasterCSV() {
  if (!_masterFileHandle) return;
  try {
    const perm = await _masterFileHandle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted') {
      const req = await _masterFileHandle.requestPermission({ mode: 'readwrite' });
      if (req !== 'granted') return;
    }
    const file = await _masterFileHandle.getFile();
    const text = await file.text();
    // Reuse existing CSV import logic
    _importCSVText(text);
    dismissRestoreBanner();
    _startTimer();
    _updateAutoSaveUI();
  } catch (e) {
    console.error('Restore from master CSV failed:', e);
  }
}

function dismissRestoreBanner() {
  const banner = document.getElementById('autosave-restore-banner');
  if (banner) banner.style.display = 'none';
}

// ── Button click handler (replaces toggleAutoSave) ────────────

function toggleAutoSave() {
  if (_masterFileHandle) {
    // Already active — turn it off
    clearMasterCSVFile();
  } else {
    setMasterCSVFile();
  }
}

// ── Init (called after first render) ─────────────────────────

async function initAutoSave() {
  if (!window.showSaveFilePicker) {
    // Browser doesn't support FSA — hide the button silently
    const btn = document.getElementById('autosave-btn');
    if (btn) btn.style.display = 'none';
    return;
  }

  const handle = await _loadHandle();
  if (!handle) { _updateAutoSaveUI(); return; }

  // Verify the stored handle is still queryable
  try {
    await handle.queryPermission({ mode: 'readwrite' });
    _masterFileHandle = handle;
    _startTimer();
    _updateAutoSaveUI();
    // Offer to reload from the master file on startup
    _showRestoreBanner(handle);
  } catch (e) {
    // Handle is stale (file deleted/moved) — clear it
    await _clearHandle();
    _updateAutoSaveUI();
  }
}
