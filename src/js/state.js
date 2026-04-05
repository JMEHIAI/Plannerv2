// ═══════════════════════════════════════════════════════════════
//  UNDO SYSTEM
// ═══════════════════════════════════════════════════════════════

let _undoStack = [];
const _UNDO_MAX = 30;

/** Call at the start of any function that mutates items/links/nextId. */
function pushUndo() {
  _undoStack.push({
    items: JSON.parse(JSON.stringify(items)),
    links: JSON.parse(JSON.stringify(links)),
    nextId: nextId,
  });
  if (_undoStack.length > _UNDO_MAX) _undoStack.shift();
  _updateUndoButton();
  // Notify auto-save that the planner state has changed
  if (typeof markChanged === 'function') markChanged();
}

function undo() {
  if (_undoStack.length === 0) return;
  const snapshot = _undoStack.pop();
  items = snapshot.items;
  links = snapshot.links;
  nextId = snapshot.nextId;
  // Invalidate render caches that depend on data
  _monthSpansCacheKey = ""; // force re-check (years may have changed)
  _dayColMetaCacheKey = "";
  _cachedHeaderHeights = null;
  _updateUndoButton();
  render();
}

function _updateUndoButton() {
  const btn = document.getElementById("undo-btn");
  if (btn) btn.disabled = _undoStack.length === 0;
}

// ═══════════════════════════════════════════════════════════════
//  GLOBAL STATE
// ═══════════════════════════════════════════════════════════════

let items = [];
let people = [];
let customTemplates = [];

// Load templates from localStorage on boot
try {
  const saved = localStorage.getItem("diaglo_activity_templates");
  if (saved) customTemplates = JSON.parse(saved);
} catch (e) {
  console.warn("Could not load templates", e);
}

let nextId = 1;
let nextPersonId = 1;
let years = ["2025", "2026", "2027", "2028"];
let hiddenYears = new Set();
let alarms = [];
let holidays = [];
let nextAlarmId = 1;
let currentAlarmItemId = null;
let highlightedWeek = null;
let zoomMode = "weeks";
let showSettings = false;
let showComments = false;

// Dependency Links state
let links = [];
let nextLinkId = 1;
let linkMode = false;
let showLinks = true;
let linkSource = null;
let commentWidth = 200;
let nameWidthBase = 350;
let filters = { name: [], type: [] };
let activeFilterMenu = null;
let filterMenuPosition = { top: 0, left: 0 };

// Block Comment State
let draggingCommentId = null;
let resizingCommentId = null;
let commentDragStartX = 0;
let commentDragStartY = 0;
let commentInitialX = 0;
let commentInitialY = 0;
let commentInitialW = 0;
let commentInitialH = 0;

// Drag/Resize State
let _interactionRafId = null;
let _dragItemsMap = new Map(); // O(1) item lookup during drag
let draggingId = null;
let dragStartX = 0;
let initialStartWeeks = {};
let draggingItems = [];
let finalWeeksMoved = 0;
let resizingId = null;
let resizeType = null;
let resizeStartX = 0;
let initialStartWeek = 0;
let initialDuration = 0;
let _resizeEl = null;       // DOM element being resized (visual-only during resize)
let _resizeInitialBarW = 0; // offsetWidth at resize start
let _resizeInitialBarML = 0; // marginLeft at resize start
let _resizeNewStart = 0;    // pending startWeek (committed on stopResize)
let _resizeNewDuration = 0; // pending duration (committed on stopResize)
let draggedRowId = null;
let globalLocked = false;

// Pan State
let isPanning = false;
let panStartX = 0, panStartY = 0, panScrollLeft = 0, panScrollTop = 0;

// Column Resize State
let colResizing = null;
let colResizeStartX = 0;
let colResizeStartWidth = 0;

// Assignee & Templates UI State
let activeAssigneeMenu = null;
let composerState = [];
let editingMasterId = null;
let _pendingCsvContent = "";

// Predefined milestones
const predefinedMilestones = [
  "HU-INT (Powertrain Intention)",
  "HU-PreC (Powertrain PreConcept)",
  "HU-SP (Start Of Project)",
  "HU-Concept (Concept Freeze)",
  "HU-PC (Pre-Contract)",
  "HU-CO (Contract)",
  "HU-VC (Agreement to Build Off Tool Unit)",
  "HU-PT (Agreement to Build Plant Trial)",
  "HU-PPC (Product & Process Certification)",
  "HU-SOP (Start Of Production)",
];

// ═══════════════════════════════════════════════════════════════
//  SHARED UTILITIES (consolidated from duplicates)
// ═══════════════════════════════════════════════════════════════

// Shared pad function (was defined 6+ times in original)
function pad2(n) {
  return n.toString().padStart(2, "0");
}

// Shared date formatter (was inline in 3+ places)
function formatDateStr(d) {
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
}

// Shared getSubtreeIds (was defined 4 times identically in original)
function getSubtreeIds(rootId) {
  const ids = new Set([rootId]);
  let changed = true;
  while (changed) {
    changed = false;
    items.forEach((i) => {
      if (i.parentId && ids.has(i.parentId) && !ids.has(i.id)) {
        ids.add(i.id);
        changed = true;
      }
    });
  }
  return ids;
}

// Shared month-end check (was computed 5+ times inline in original)
function isMonthEndWeek(yearIndex, relWeek, dynamicMonthSpans) {
  let acc = 0;
  for (const m of dynamicMonthSpans[yearIndex]) {
    acc += m.weeks;
    if (relWeek === acc) return true;
    if (acc > relWeek) return false;
  }
  return false;
}

// Returns the predefined color for a milestone based on its short name
function getMilestoneColor(name) {
  const short = name.replace(/\s*\(.*/, "").trim();
  if (short === "HU-CO" || short === "HU-PPC") return "#ef4444";
  if (short === "HU-SP" || short === "HU-SOP") return "#10b981";
  return "#f59e0b";
}

function getDescendantMilestones(parentId) {
  const result = [];
  function recurse(pid) {
    items
      .filter((i) => i.parentId === pid)
      .forEach((child) => {
        if (child.type === "milestone") result.push(child);
        recurse(child.id);
      });
  }
  recurse(parentId);
  return result;
}

function getShortYear(yearIndex) {
  const label = years[yearIndex] || "";
  const m4 = label.match(/\d{4}/);
  if (m4) return m4[0].slice(-2);
  const m2 = label.match(/\d{2}/);
  if (m2) return m2[0];
  return String(yearIndex + 1).padStart(2, "0");
}

function weekToYYWW(absWeek) {
  const yi = Math.floor((absWeek - 1) / 52);
  const ww = ((absWeek - 1) % 52) + 1;
  return getShortYear(yi) + String(ww).padStart(2, "0");
}

function formatYYWWD(absWeek) {
  if (!absWeek) return "";
  const baseWeek = Math.max(1, Math.floor(absWeek));
  const yrIndex = Math.floor((baseWeek - 1) / 52);
  const yrStr = years[yrIndex] || "2026";
  const yy = yrStr.slice(-2);
  const wkFull = ((baseWeek - 1) % 52) + 1;
  const wk = String(wkFull).padStart(2, "0");
  const offset = absWeek - Math.floor(absWeek);
  const d = Math.round(offset / 0.2) + 1;
  return yy + wk + (d > 1 ? "." + Math.min(5, d) : "");
}

function parseYYWWD(str) {
  if (!str) return null;
  const parts = str.split(".");
  const yyww = parts[0];
  const d = parts[1] ? parseInt(parts[1]) : 1;
  if (yyww.length < 3) return null;
  const yy = yyww.substring(0, 2);
  const ww = parseInt(yyww.substring(2));
  if (ww < 1 || ww > 52) return null;

  let yrIdx = years.findIndex((y) => y.endsWith(yy));

  // Auto-create missing years if needed
  if (yrIdx === -1) {
    const century = parseInt(yy) >= 50 ? 1900 : 2000;
    const targetYear = century + parseInt(yy);
    const firstYear = parseInt(years[0]) || 2025;
    const lastYear = parseInt(years[years.length - 1]) || 2028;

    if (targetYear < firstYear) {
      // Prepend years and shift all items
      while (parseInt(years[0]) > targetYear) {
        const fy = parseInt(years[0]);
        years.unshift(String(fy - 1));
        items.forEach(item => { item.startWeek += 52; });
      }
    } else if (targetYear > lastYear) {
      // Append years
      while (parseInt(years[years.length - 1]) < targetYear) {
        const ly = parseInt(years[years.length - 1]);
        years.push(String(ly + 1));
      }
    }
    if (typeof _updateYearCountLabel === 'function') _updateYearCountLabel();
    yrIdx = years.findIndex((y) => y.endsWith(yy));
    if (yrIdx === -1) return null;
  }

  const baseWeek = yrIdx * 52 + ww;
  const offset = (Math.max(1, Math.min(5, d)) - 1) * 0.2;
  return baseWeek + offset;
}

function isHighlightActive(cellWeek, highlightVal) {
  if (highlightVal === null) return false;
  if (typeof highlightVal === "string" && highlightVal.startsWith("W")) {
    const hW = parseInt(highlightVal.substring(1));
    return Math.floor(cellWeek) === hW;
  }
  return Math.abs(cellWeek - highlightVal) < 0.01;
}

// Returns true if this holiday applies to the given item assignees.
function holidayApplies(h, itemAssignees) {
  if (!h.people || h.people.length === 0) return true;
  if (!itemAssignees || itemAssignees.length === 0) return false;
  return h.people.some(pid => itemAssignees.includes(pid));
}

function getLastItemEndWeek() {
  const displayItems = getDisplayItems();
  if (displayItems.length === 0) return 1;
  const lastItem = displayItems[displayItems.length - 1];
  const raw = Math.min(years.length * 52, lastItem.startWeek + (lastItem.duration || 0));
  return Math.round(raw * 5) / 5;
}
