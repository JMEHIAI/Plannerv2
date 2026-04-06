// ═══════════════════════════════════════════════════════════════
//  SHARED MASTER CSV + SINGLE-EDITOR LOCK
//
//  Goals:
//    1. Keep the planner browser-only and easy to share.
//    2. Allow only one active editor at a time.
//    3. Let everyone else load the latest file in read-only mode.
//    4. Recover automatically from crashes via stale-lock timeout.
// ═══════════════════════════════════════════════════════════════

const _AUTOSAVE_INTERVAL_MS = 30000;
const _EDITOR_LOCK_TTL_MS = 90000;
const _IDB_NAME = "diaglo_planner";
const _IDB_STORE = "handles";
const _IDB_KEY = "masterCSV";
const _EDITOR_NAME_KEY = "diaglo_planner_editor_name";

let _autoSaveTimer = null;
let _masterFileHandle = null;
let _changesSinceLastSave = false;
let _hasEditLock = false;
let _isReadOnlyMaster = false;
let _editorLockInfo = null;
let _editorIdentity = "";
let _editorSessionId =
  "editor_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8);

// ── IndexedDB helpers ─────────────────────────────────────────

function _openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(_IDB_NAME, 1);
    req.onupgradeneeded = (e) => e.target.result.createObjectStore(_IDB_STORE);
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = () => reject(req.error);
  });
}

async function _storeHandle(handle) {
  const db = await _openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(_IDB_STORE, "readwrite");
    const req = tx.objectStore(_IDB_STORE).put(handle, _IDB_KEY);
    req.onsuccess = resolve;
    req.onerror = () => reject(req.error);
  });
}

async function _loadHandle() {
  try {
    const db = await _openDB();
    return new Promise((resolve) => {
      const tx = db.transaction(_IDB_STORE, "readonly");
      const req = tx.objectStore(_IDB_STORE).get(_IDB_KEY);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => resolve(null);
    });
  } catch (e) {
    return null;
  }
}

async function _clearHandle() {
  try {
    const db = await _openDB();
    return new Promise((resolve) => {
      const tx = db.transaction(_IDB_STORE, "readwrite");
      tx.objectStore(_IDB_STORE).delete(_IDB_KEY);
      tx.oncomplete = resolve;
    });
  } catch (e) {}
}

// ── Lock helpers ──────────────────────────────────────────────

function _getEditorIdentity() {
  if (_editorIdentity) return _editorIdentity;
  _editorIdentity = (localStorage.getItem(_EDITOR_NAME_KEY) || "").trim();
  return _editorIdentity;
}

function _ensureEditorIdentity() {
  const existing = _getEditorIdentity();
  if (existing) return existing;

  const entered = (window.prompt("Enter your name or initials for the shared edit lock:", "") || "").trim();
  if (!entered) return null;

  _editorIdentity = entered;
  localStorage.setItem(_EDITOR_NAME_KEY, entered);
  return entered;
}

function _plannerHasContent() {
  return (
    items.length > 0 ||
    alarms.length > 0 ||
    people.length > 0 ||
    holidays.length > 0 ||
    links.length > 0 ||
    customTemplates.length > 0
  );
}

function _formatLockOwner(lockInfo) {
  return (lockInfo && lockInfo.editorName) || "another editor";
}

function _isLockActive(lockInfo) {
  if (!lockInfo || !lockInfo.sessionId) return false;
  const expiresAt = Date.parse(lockInfo.expiresAt || lockInfo.heartbeatAt || "");
  return !isNaN(expiresAt) && expiresAt > Date.now();
}

function _refreshOwnLockInfo() {
  const editorName = _getEditorIdentity() || "Unknown";
  const now = new Date();

  if (!_editorLockInfo || _editorLockInfo.sessionId !== _editorSessionId) {
    _editorLockInfo = {
      sessionId: _editorSessionId,
      editorName,
      acquiredAt: now.toISOString(),
    };
  }

  _editorLockInfo.editorName = editorName;
  _editorLockInfo.heartbeatAt = now.toISOString();
  _editorLockInfo.expiresAt = new Date(now.getTime() + _EDITOR_LOCK_TTL_MS).toISOString();

  return _editorLockInfo;
}

function _buildEditorLockCSVLine() {
  if (!_hasEditLock) return "";
  const lockInfo = _refreshOwnLockInfo();
  return 'EditorLock,"' + JSON.stringify(lockInfo).replace(/"/g, '""') + '"';
}

function _parseEditorLockFromText(text) {
  if (!text) return null;
  const lines = text.split(String.fromCharCode(10));
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) continue;
    if (line.startsWith("EditorLock")) {
      const values = _parseCSVLine(line);
      if (!values[1]) return null;
      try {
        return JSON.parse(values[1]);
      } catch (e) {
        return null;
      }
    }
    if (line.startsWith("Id,Type,Name")) break;
  }
  return null;
}

async function _readMasterText(handle) {
  const file = await handle.getFile();
  return file.text();
}

async function _readMasterSnapshot(handle) {
  const text = await _readMasterText(handle);
  return {
    text,
    lockInfo: _parseEditorLockFromText(text),
  };
}

async function _ensureMasterPermission() {
  if (!_masterFileHandle) return false;
  const perm = await _masterFileHandle.queryPermission({ mode: "readwrite" });
  if (perm === "granted") return true;
  const req = await _masterFileHandle.requestPermission({ mode: "readwrite" });
  return req === "granted";
}

function _setDisconnectedState() {
  _hasEditLock = false;
  _isReadOnlyMaster = false;
  _editorLockInfo = null;
  _masterFileHandle = null;
  _stopTimer();
  _updateAutoSaveUI();
}

function _setReadOnlyState(lockInfo) {
  _hasEditLock = false;
  _isReadOnlyMaster = !!lockInfo;
  _editorLockInfo = lockInfo || null;
  _stopTimer();
  _updateAutoSaveUI();
}

function _setEditingState() {
  _hasEditLock = true;
  _isReadOnlyMaster = false;
  _startTimer();
  _updateAutoSaveUI();
}

// ── Core write ────────────────────────────────────────────────

async function _writeToHandle(handle, options) {
  const csv = generateCSVString({
    includeEditorLock: !!(options && options.includeEditorLock),
  });
  const writable = await handle.createWritable();
  await writable.write(csv);
  await writable.close();
}

async function _doAutoSave() {
  if (!_masterFileHandle || !_hasEditLock) return;
  try {
    const hasPermission = await _ensureMasterPermission();
    if (!hasPermission) {
      _updateAutoSaveUI();
      return;
    }

    await _writeToHandle(_masterFileHandle, { includeEditorLock: true });
    _changesSinceLastSave = false;
    _updateAutoSaveUI();
  } catch (e) {
    console.warn("Auto-save write failed:", e);
    const s = document.getElementById("autosave-status");
    if (s) s.textContent = "Save failed";
  }
}

async function _releaseEditLock(clearStoredHandle) {
  if (_masterFileHandle && _hasEditLock) {
    try {
      const hasPermission = await _ensureMasterPermission();
      if (hasPermission) {
        await _writeToHandle(_masterFileHandle, { includeEditorLock: false });
      }
    } catch (e) {
      console.warn("Could not release edit lock cleanly:", e);
    }
  }

  _hasEditLock = false;
  _isReadOnlyMaster = false;
  _editorLockInfo = null;
  _stopTimer();

  if (clearStoredHandle) {
    _masterFileHandle = null;
    await _clearHandle();
  }

  _updateAutoSaveUI();
}

// ── markChanged — called by pushUndo / CSV import ─────────────

function markChanged() {
  _changesSinceLastSave = true;
  clearTimeout(markChanged._debounce);
  markChanged._debounce = setTimeout(() => {
    if (_masterFileHandle && _hasEditLock) _doAutoSave();
  }, 3000);
}
markChanged._debounce = null;

// ── Banner helpers ────────────────────────────────────────────

function _showMasterBanner() {
  const banner = document.getElementById("autosave-restore-banner");
  if (!banner || !_masterFileHandle) return;

  if (_isReadOnlyMaster && _isLockActive(_editorLockInfo)) {
    banner.innerHTML =
      '<span style="flex:1">📂 Shared planner: <strong>' +
      _masterFileHandle.name +
      "</strong> — read-only right now, locked by <strong>" +
      _formatLockOwner(_editorLockInfo).replace(/</g, "&lt;") +
      "</strong>.</span>" +
      '<button class="success" onclick="restoreFromMasterCSV()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Load Latest</button>' +
      '<button class="outline" onclick="startEditingMasterCSV()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Try Edit</button>' +
      '<button class="outline" onclick="dismissRestoreBanner()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Hide</button>';
    banner.style.display = "flex";
    return;
  }

  banner.innerHTML =
    '<span style="flex:1">📂 Shared planner ready: <strong>' +
    _masterFileHandle.name +
    "</strong>.</span>" +
    '<button class="success" onclick="restoreFromMasterCSV()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Load</button>' +
    '<button class="outline" onclick="startEditingMasterCSV()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Start Editing</button>' +
    '<button class="outline" onclick="dismissRestoreBanner()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Hide</button>';
  banner.style.display = "flex";
}

function dismissRestoreBanner() {
  const banner = document.getElementById("autosave-restore-banner");
  if (banner) banner.style.display = "none";
}

// ── Master file / edit-lock control ───────────────────────────

async function _importMasterText(text) {
  if (text && text.trim()) _importCSVText(text);
}

async function setMasterCSVFile() {
  if (!window.showSaveFilePicker) {
    alert("Your browser does not support the File System Access API.\nUse Chrome or Edge.");
    return;
  }

  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: "planner_master.csv",
      types: [{ description: "CSV Files", accept: { "text/csv": [".csv"] } }],
    });

    _masterFileHandle = handle;
    await _storeHandle(handle);
    dismissRestoreBanner();
    await startEditingMasterCSV();
  } catch (e) {
    if (e.name !== "AbortError") console.error(e);
  }
}

async function startEditingMasterCSV() {
  if (!_masterFileHandle) {
    await setMasterCSVFile();
    return;
  }

  const editorName = _ensureEditorIdentity();
  if (!editorName) return;

  try {
    const hasPermission = await _ensureMasterPermission();
    if (!hasPermission) return;

    const snapshot = await _readMasterSnapshot(_masterFileHandle);
    const activeLock = _isLockActive(snapshot.lockInfo) ? snapshot.lockInfo : null;

    if (activeLock && activeLock.sessionId !== _editorSessionId) {
      _setReadOnlyState(activeLock);
      if (snapshot.text && snapshot.text.trim()) {
        await _importMasterText(snapshot.text);
      }
      _showMasterBanner();
      alert("Shared planner is currently being edited by " + _formatLockOwner(activeLock) + ".");
      return;
    }

    if (snapshot.text && snapshot.text.trim()) {
      const shouldLoadShared =
        !_plannerHasContent() ||
        confirm(
          "Load the selected shared planner before taking the edit lock?\n\n" +
            "OK = load the shared file first\n" +
            "Cancel = keep the current planner and overwrite the shared file after locking",
        );
      if (shouldLoadShared) {
        await _importMasterText(snapshot.text);
      }
    }

    _editorLockInfo = {
      sessionId: _editorSessionId,
      editorName,
      acquiredAt:
        snapshot.lockInfo && snapshot.lockInfo.sessionId === _editorSessionId
          ? snapshot.lockInfo.acquiredAt || new Date().toISOString()
          : new Date().toISOString(),
    };

    _setEditingState();
    _changesSinceLastSave = true;
    await _doAutoSave();

    const verifySnapshot = await _readMasterSnapshot(_masterFileHandle);
    if (!verifySnapshot.lockInfo || verifySnapshot.lockInfo.sessionId !== _editorSessionId) {
      _setReadOnlyState(verifySnapshot.lockInfo);
      _showMasterBanner();
      alert("Could not secure the edit lock. Another editor took it first.");
      return;
    }

    dismissRestoreBanner();
  } catch (e) {
    console.error("Could not start editing shared planner:", e);
  }
}

async function clearMasterCSVFile() {
  await _releaseEditLock(true);
  dismissRestoreBanner();
}

async function restoreFromMasterCSV() {
  if (!_masterFileHandle) return;

  try {
    const text = await _readMasterText(_masterFileHandle);
    await _importMasterText(text);

    const activeLock = _isLockActive(_parseEditorLockFromText(text))
      ? _parseEditorLockFromText(text)
      : null;
    if (!_hasEditLock) {
      _setReadOnlyState(activeLock);
      _showMasterBanner();
    } else {
      dismissRestoreBanner();
    }
    _updateAutoSaveUI();
  } catch (e) {
    console.error("Restore from master CSV failed:", e);
  }
}

// ── Timer / UI ────────────────────────────────────────────────

function _startTimer() {
  if (_autoSaveTimer) return;
  _autoSaveTimer = setInterval(_doAutoSave, _AUTOSAVE_INTERVAL_MS);
}

function _stopTimer() {
  clearInterval(_autoSaveTimer);
  _autoSaveTimer = null;
}

function _updateAutoSaveUI() {
  const btn = document.getElementById("autosave-btn");
  const statusEl = document.getElementById("autosave-status");
  const label = document.getElementById("autosave-label");

  if (label) {
    if (_hasEditLock) label.textContent = "Edit Lock: ON";
    else if (_masterFileHandle && _isReadOnlyMaster) label.textContent = "Read-only";
    else if (_masterFileHandle) label.textContent = "Shared File";
    else label.textContent = "Edit Lock: OFF";
  }

  if (btn) {
    if (_hasEditLock) {
      btn.style.color = "#10b981";
      btn.style.fontWeight = "700";
      btn.title = "Click to release the edit lock and stop shared auto-save";
    } else if (_masterFileHandle && _isReadOnlyMaster) {
      btn.style.color = "#b45309";
      btn.style.fontWeight = "700";
      btn.title = "Shared file is read-only right now. Click to try taking the edit lock";
    } else if (_masterFileHandle) {
      btn.style.color = "#2563eb";
      btn.style.fontWeight = "700";
      btn.title = "Shared file is connected. Click to take the edit lock";
    } else {
      btn.style.color = "";
      btn.style.fontWeight = "";
      btn.title = "Pick a shared CSV file for single-editor auto-save";
    }
  }

  if (!statusEl) return;
  if (!_masterFileHandle) {
    statusEl.textContent = "";
    return;
  }

  if (_hasEditLock) {
    const name = _masterFileHandle.name || "master.csv";
    statusEl.textContent = (_changesSinceLastSave ? "● " : "🔒 ") + name;
    return;
  }

  if (_isReadOnlyMaster && _isLockActive(_editorLockInfo)) {
    statusEl.textContent = "👀 " + _formatLockOwner(_editorLockInfo);
    return;
  }

  statusEl.textContent = "✓ " + (_masterFileHandle.name || "master.csv");
}

// ── Button click handler ──────────────────────────────────────

function toggleAutoSave() {
  if (_hasEditLock) {
    clearMasterCSVFile();
  } else if (_masterFileHandle) {
    startEditingMasterCSV();
  } else {
    setMasterCSVFile();
  }
}

// ── Init ──────────────────────────────────────────────────────

async function initAutoSave() {
  if (!window.showSaveFilePicker) {
    const btn = document.getElementById("autosave-btn");
    if (btn) btn.style.display = "none";
    return;
  }

  const handle = await _loadHandle();
  if (!handle) {
    _updateAutoSaveUI();
    return;
  }

  try {
    await handle.queryPermission({ mode: "readwrite" });
    _masterFileHandle = handle;
    const snapshot = await _readMasterSnapshot(handle);
    const activeLock = _isLockActive(snapshot.lockInfo) ? snapshot.lockInfo : null;
    _setReadOnlyState(activeLock);
    if (snapshot.text && snapshot.text.trim()) _showMasterBanner();
  } catch (e) {
    await _clearHandle();
    _setDisconnectedState();
  }
}

window.addEventListener("pagehide", () => {
  if (_hasEditLock) {
    _changesSinceLastSave = true;
    _doAutoSave();
  }
});
