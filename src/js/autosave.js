// ═══════════════════════════════════════════════════════════════
//  SHARED MASTER CSV — last-write-wins, no locks
//
//  Flow:
//    1. On load: check IndexedDB for a stored file handle.
//       If found and permission already granted → auto-load + auto-save silently.
//       If found but permission expired → show banner to re-grant (one click).
//       If not found → show banner to pick the file (one click).
//    2. Everyone edits freely. Changes auto-save 3 s after last edit + every 30 s.
//    3. Every 60 s the file is re-read; if someone else saved, a toast appears
//       and content reloads silently.
// ═══════════════════════════════════════════════════════════════

const _AUTOSAVE_INTERVAL_MS = 30000;
const _POLL_INTERVAL_MS     = 60000;
const _IDB_NAME             = "diaglo_planner";
const _IDB_STORE            = "handles";
const _IDB_KEY              = "masterCSV";
const _EDITOR_NAME_KEY      = "diaglo_planner_editor_name";
const _AUTO_MASTER_FILENAME = "planner_master.csv";

let _autoSaveTimer    = null;
let _pollTimer        = null;
let _masterFileHandle = null;
let _canWrite         = false;
let _changesSinceLastSave = false;
let _lastSavedAt      = 0;
let _lastSavedByStr   = "";   // "Name @ HH:MM" from last read
let _editorName       = "";

// ── IndexedDB — persist handle across page loads ──────────────

function _openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(_IDB_NAME, 1);
    req.onupgradeneeded = (e) => e.target.result.createObjectStore(_IDB_STORE);
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = () => reject(req.error);
  });
}
async function _storeHandle(handle) {
  try {
    const db = await _openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(_IDB_STORE, "readwrite");
      const req = tx.objectStore(_IDB_STORE).put(handle, _IDB_KEY);
      req.onsuccess = resolve; req.onerror = () => reject(req.error);
    });
  } catch (e) {}
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
  } catch (e) { return null; }
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

// ── Editor name ───────────────────────────────────────────────

function _getEditorName() {
  if (_editorName) return _editorName;
  _editorName = (localStorage.getItem(_EDITOR_NAME_KEY) || "").trim();
  return _editorName;
}
function _ensureEditorName() {
  const n = _getEditorName();
  if (n) return n;
  const entered = (window.prompt("Your name or initials (shown when you save):", "") || "").trim();
  _editorName = entered || "User";
  localStorage.setItem(_EDITOR_NAME_KEY, _editorName);
  return _editorName;
}

// ── "Saved by" metadata embedded in CSV ──────────────────────
// Re-uses the EditorLock row key for backwards compat — no lock semantics.

function _buildSavedByCSVLine() {
  const info = { editorName: _getEditorName() || "?", savedAt: new Date().toISOString() };
  return 'EditorLock,"' + JSON.stringify(info).replace(/"/g, '""') + '"';
}
function _parseSavedByFromText(text) {
  if (!text) return null;
  const lines = text.split("\n");
  for (let i = 0; i < Math.min(lines.length, 20); i++) {
    const line = lines[i].trim();
    if (line.startsWith("EditorLock")) {
      try { return JSON.parse(_parseCSVLine(line)[1]); } catch (e) { return null; }
    }
    if (line.startsWith("Id,Type,Name")) break;
  }
  return null;
}
function _formatSavedBy(info) {
  if (!info || !info.editorName) return "";
  if (!info.savedAt) return info.editorName;
  const d = new Date(info.savedAt);
  return info.editorName + " @ " + String(d.getHours()).padStart(2,"0") + ":" + String(d.getMinutes()).padStart(2,"0");
}

// ── File I/O ──────────────────────────────────────────────────

async function _checkWritePermission() {
  if (!_masterFileHandle) return false;
  try {
    const perm = await _masterFileHandle.queryPermission({ mode: "readwrite" });
    if (perm === "granted") { _canWrite = true; return true; }
    return false;
  } catch (e) { return false; }
}
async function _requestWritePermission() {
  if (!_masterFileHandle) return false;
  try {
    const perm = await _masterFileHandle.requestPermission({ mode: "readwrite" });
    _canWrite = perm === "granted";
    return _canWrite;
  } catch (e) { _canWrite = false; return false; }
}

async function _readFileText(handle) {
  try { const f = await handle.getFile(); return await f.text(); } catch (e) { return null; }
}
async function _writeFile(handle) {
  const csv = generateCSVString({ includeEditorLock: true });
  const writable = await handle.createWritable();
  await writable.write(csv);
  await writable.close();
}

function _plannerHasContent() {
  return items.length > 0 || alarms.length > 0 || people.length > 0 ||
    holidays.length > 0 || links.length > 0 || customTemplates.length > 0;
}
async function _importMasterText(text) {
  if (text && text.trim()) _importCSVText(text);
}

// ── Auto-save & polling ───────────────────────────────────────

async function _doAutoSave() {
  if (!_masterFileHandle || !_canWrite) return;
  try {
    await _writeFile(_masterFileHandle);
    _changesSinceLastSave = false;
    _lastSavedAt = Date.now();
    _updateAutoSaveUI();
  } catch (e) {
    console.warn("Auto-save failed:", e);
    _showStatus("⚠ Save failed", "#ef4444");
  }
}

async function _pollForChanges() {
  if (!_masterFileHandle) return;
  const text = await _readFileText(_masterFileHandle);
  if (!text || !text.trim()) return;
  const info = _parseSavedByFromText(text);
  if (!info || !info.savedAt) return;
  const savedAt = Date.parse(info.savedAt);
  if (savedAt <= _lastSavedAt + 5000) return;     // our own save
  const myName = _getEditorName();
  if (info.editorName === myName) return;          // same user, different tab
  _showToast("Updated by " + _formatSavedBy(info));
  await _importMasterText(text);
  _lastSavedAt = savedAt;
}

function _startTimers() {
  if (!_autoSaveTimer) _autoSaveTimer = setInterval(_doAutoSave, _AUTOSAVE_INTERVAL_MS);
  if (!_pollTimer)     _pollTimer     = setInterval(_pollForChanges, _POLL_INTERVAL_MS);
}
function _stopTimers() {
  clearInterval(_autoSaveTimer); _autoSaveTimer = null;
  clearInterval(_pollTimer);     _pollTimer     = null;
}

// markChanged — called by pushUndo / CSV import ───────────────

function markChanged() {
  _changesSinceLastSave = true;
  _showStatus("● Unsaved", "#f59e0b");
  clearTimeout(markChanged._debounce);
  markChanged._debounce = setTimeout(() => {
    if (_masterFileHandle && _canWrite) _doAutoSave();
  }, 3000);
}
markChanged._debounce = null;

// ── Toast ────────────────────────────────────────────────────

function _showToast(msg) {
  let t = document.getElementById("_as_toast");
  if (!t) {
    t = document.createElement("div");
    t.id = "_as_toast";
    t.style.cssText = "position:fixed;bottom:24px;right:24px;background:#1e293b;color:#fff;" +
      "padding:10px 18px;border-radius:8px;font-size:13px;z-index:99999;" +
      "box-shadow:0 4px 16px rgba(0,0,0,.3);transition:opacity .4s;pointer-events:none;";
    document.body.appendChild(t);
  }
  t.textContent = msg; t.style.opacity = "1";
  clearTimeout(_showToast._t);
  _showToast._t = setTimeout(() => { t.style.opacity = "0"; }, 3500);
}
_showToast._t = null;

// ── Status bar ───────────────────────────────────────────────

function _showStatus(text, color) {
  const el = document.getElementById("autosave-status");
  if (el) { el.textContent = text; el.style.color = color || "#64748b"; }
}
function _updateAutoSaveUI() {
  const btn   = document.getElementById("autosave-btn");
  const label = document.getElementById("autosave-label");
  if (label) label.textContent = _masterFileHandle ? "Auto-save" : "Open File";
  if (btn) {
    btn.title      = _masterFileHandle ? "Switch to a different shared file" : "Open the shared planner CSV file";
    btn.style.color      = _masterFileHandle && _canWrite ? "#10b981" : "";
    btn.style.fontWeight = _masterFileHandle && _canWrite ? "700" : "";
  }
  if (!_masterFileHandle) { _showStatus("No file", "#94a3b8"); return; }
  if (_changesSinceLastSave) { _showStatus("● " + _masterFileHandle.name, "#f59e0b"); return; }
  if (_lastSavedAt > 0) {
    const d = new Date(_lastSavedAt);
    _showStatus("✓ Saved " + String(d.getHours()).padStart(2,"0") + ":" + String(d.getMinutes()).padStart(2,"0"), "#10b981");
  } else {
    _showStatus("✓ " + _masterFileHandle.name, "#64748b");
  }
}

// ── Banner ───────────────────────────────────────────────────

function _showBanner(html) {
  const b = document.getElementById("autosave-restore-banner");
  if (b) { b.innerHTML = html; b.style.display = "flex"; }
}
function _hideBanner() {
  const b = document.getElementById("autosave-restore-banner");
  if (b) b.style.display = "none";
}
function dismissRestoreBanner() { _hideBanner(); }

function _showOpenFileBanner() {
  _showBanner(
    '<span style="flex:1">📂 Open <strong>' + _AUTO_MASTER_FILENAME +
    '</strong> to load and auto-save the shared planner.</span>' +
    '<button class="success" onclick="openMasterCSVFile()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Open File</button>' +
    '<button class="outline" onclick="_hideBanner()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Dismiss</button>'
  );
}
function _showPermissionBanner(name) {
  _showBanner(
    '<span style="flex:1">📂 <strong>' + (name || _AUTO_MASTER_FILENAME) +
    '</strong> — click <b>Grant access</b> to re-enable auto-save.</span>' +
    '<button class="success" onclick="_grantAndStart()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Grant access</button>' +
    '<button class="outline" onclick="_hideBanner()" style="padding:4px 14px;font-size:13px;flex-shrink:0;">Dismiss</button>'
  );
}

// ── Grant permission after banner click ───────────────────────

async function _grantAndStart() {
  _hideBanner();
  if (!_masterFileHandle) { await openMasterCSVFile(); return; }
  const ok = await _requestWritePermission();
  if (!ok) { _showToast("Permission denied — edits won't be auto-saved."); return; }
  _ensureEditorName();
  // Load latest content then start saving
  const text = await _readFileText(_masterFileHandle);
  if (text && text.trim()) await _importMasterText(text);
  _changesSinceLastSave = true;
  await _doAutoSave();
  _startTimers();
  _updateAutoSaveUI();
}

// ── Public: open file picker ──────────────────────────────────

async function openMasterCSVFile() {
  if (!window.showOpenFilePicker) {
    alert("Your browser does not support the File System Access API.\nUse Chrome or Edge.");
    return;
  }
  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{ description: "CSV Files", accept: { "text/csv": [".csv"] } }],
      startIn: "documents",
    });
    _masterFileHandle = handle;
    await _storeHandle(handle);
    _ensureEditorName();

    // Load existing content
    const text = await _readFileText(handle);
    if (text && text.trim()) await _importMasterText(text);

    // Request write permission (may auto-grant since user just picked it)
    _canWrite = await _checkWritePermission();
    if (!_canWrite) _canWrite = await _requestWritePermission();

    _changesSinceLastSave = true;
    if (_canWrite) await _doAutoSave();
    _startTimers();
    _hideBanner();
    _updateAutoSaveUI();
    if (!_canWrite) _showToast("Loaded read-only — could not get write permission.");
  } catch (e) {
    if (e.name !== "AbortError") console.error("Could not open file:", e);
  }
}

// toggleAutoSave — toolbar button
async function toggleAutoSave() {
  await openMasterCSVFile();
}

// ── Init ──────────────────────────────────────────────────────

async function initAutoSave() {
  // Browsers without File System Access API: nothing we can do for auto-save
  if (!window.showSaveFilePicker) {
    const btn = document.getElementById("autosave-btn");
    if (btn) btn.style.display = "none";
    _updateAutoSaveUI();
    _showOpenFileBanner();
    return;
  }

  const handle = await _loadHandle();

  if (!handle) {
    // First time ever — show open-file banner
    _updateAutoSaveUI();
    _showOpenFileBanner();
    return;
  }

  // We have a stored handle — check permission
  _masterFileHandle = handle;
  const granted = await _checkWritePermission();

  if (granted) {
    // Best case: silently auto-load and start auto-saving
    _canWrite = true;
    const text = await _readFileText(handle);
    if (text && text.trim()) {
      await _importMasterText(text);
      const info = _parseSavedByFromText(text);
      if (info) {
        _lastSeenSavedBy = _formatSavedBy(info);
        _lastSavedAt = info.savedAt ? Date.parse(info.savedAt) : 0;
        const myName = _getEditorName();
        if (info.editorName && myName && info.editorName !== myName) {
          _showToast("Loaded • Last saved by " + _lastSeenSavedBy);
        }
      }
    }
    _ensureEditorName();
    _changesSinceLastSave = true;
    await _doAutoSave();
    _startTimers();
    _hideBanner();
  } else {
    // Permission expired — need one click to re-grant
    // Try to show file content from a fresh file read (may fail without permission)
    // but at least show something useful in the banner
    _showPermissionBanner(handle.name);
    // Try read-only load (queryPermission for "read" only)
    try {
      const readPerm = await handle.queryPermission({ mode: "read" });
      if (readPerm === "granted") {
        const text = await _readFileText(handle);
        if (text && text.trim()) await _importMasterText(text);
      }
    } catch (e) {}
  }

  _updateAutoSaveUI();
}

window.addEventListener("pagehide", () => {
  if (_masterFileHandle && _canWrite && _changesSinceLastSave) _doAutoSave();
});
