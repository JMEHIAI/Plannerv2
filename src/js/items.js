// ═══════════════════════════════════════════════════════════════
//  ITEM CRUD OPERATIONS
// ═══════════════════════════════════════════════════════════════

function snapNewItemStartWeek(startWeek, type) {
  const normalized = Math.round(Math.max(1, startWeek) * 5) / 5;
  if (type === "milestone") return normalized;
  return Math.max(1, Math.ceil(normalized - 0.0001));
}

function ensureTimelineCovers(maxWeekNeeded) {
  const needed = Math.max(1, Math.ceil(maxWeekNeeded - 0.0001));
  let extended = false;
  while (getTotalWeekCount() < needed) {
    const lastYr = parseInt(years[years.length - 1]) || new Date().getFullYear();
    years.push(String(lastYr + 1));
    extended = true;
  }
  if (extended) _updateYearCountLabel();
}

// Consolidated addSubItem - replaces 3 near-identical functions
function addSubItem(parentId, type, defaults) {
  pushUndo();
  const parent = items.find((i) => i.id === parentId);
  if (!parent) return;
  parent.isExpanded = true;
  const subItems = items.filter((i) => i.parentId === parentId);
  let startWeek = parent.startWeek + (parent.duration || 0);
  if (subItems.length > 0) {
    startWeek = subItems.reduce((maxEnd, subItem) => {
      let subDuration = subItem.duration || 0;
      if (subItem.type === "task" && holidays.length > 0 && typeof computeEffectiveDuration === "function") {
        subDuration = computeEffectiveDuration(subItem.startWeek, subDuration, subItem.assignees || []);
      } else if (subItem.type === "milestone" && subDuration === 0) {
        subDuration = 0.5;
      }
      return Math.max(maxEnd, subItem.startWeek + subDuration);
    }, parent.startWeek);
  }
  // Prefer starting on the current Monday (or next Monday if mid-week) over
  // placing the item at the tail of existing children, which may be far away.
  const todayWk = typeof _getSystemTodayWeek === "function" ? _getSystemTodayWeek() : null;
  if (todayWk !== null) {
    const thisMonday = Math.ceil(todayWk - 0.0001); // Monday on or after today
    if (thisMonday >= 1 && thisMonday <= getTotalWeekCount()) {
      startWeek = thisMonday;
    }
  }
  startWeek = snapNewItemStartWeek(Math.min(getTotalWeekCount(), startWeek), type);
  ensureTimelineCovers(startWeek + (defaults.duration || 0));

  items.push({
    id: nextId++,
    type: type,
    name: defaults.name,
    startWeek: startWeek,
    duration: defaults.duration || 0,
    parentId: parentId,
    isExpanded: defaults.isExpanded !== undefined ? defaults.isExpanded : true,
    color: defaults.color || parent.color || "#4f46e5",
  });
  render();
}

function addSubTask(parentId) {
  addSubItem(parentId, "task", { name: "New Sub-activity", duration: 2 });
}

function addSubProject(parentId) {
  addSubItem(parentId, "project", { name: "New Project", duration: 12.0, isExpanded: false, color: "#10b981" });
}

function addSubMilestone(parentId) {
  addSubItem(parentId, "milestone", { name: "New Milestone", duration: 0, color: "#f59e0b" });
}

function addFamily() {
  pushUndo();
  const startWeek = snapNewItemStartWeek(getLastItemEndWeek(), "family");
  ensureTimelineCovers(startWeek + 12.0);
  const parentId = nextId++;
  items.push({
    id: parentId,
    type: "family",
    name: "Family",
    startWeek: startWeek,
    duration: 12.0,
    isExpanded: true,
    color: "#6366f1",
  });

  const brotherId = nextId++;
  items.push({
    id: brotherId,
    type: "project",
    name: "New Project",
    startWeek: startWeek,
    duration: 12.0,
    isExpanded: false,
    color: "#10b981",
    parentId: parentId,
  });
  updateProjectName(brotherId, "Brother Project");

  const childId1 = nextId++;
  items.push({
    id: childId1,
    type: "project",
    name: "New Project",
    startWeek: startWeek,
    duration: 12.0,
    isExpanded: false,
    color: "#10b981",
    parentId: parentId,
  });
  updateProjectName(childId1, "Child Project");

  const childId2 = nextId++;
  items.push({
    id: childId2,
    type: "project",
    name: "New Project",
    startWeek: startWeek,
    duration: 12.0,
    isExpanded: false,
    color: "#10b981",
    parentId: parentId,
  });
  updateProjectName(childId2, "Child Project");

  render();
}

function addProject() {
  pushUndo();
  const startWeek = snapNewItemStartWeek(getLastItemEndWeek(), "project");
  ensureTimelineCovers(startWeek + 12.0);
  items.push({
    id: nextId++,
    type: "project",
    name: "New Project",
    startWeek: startWeek,
    duration: 12.0,
    isExpanded: true,
    color: "#10b981",
  });
  render();
}

function addTask() {
  pushUndo();
  const startWeek = snapNewItemStartWeek(getLastItemEndWeek(), "task");
  ensureTimelineCovers(startWeek + 2);
  items.push({
    id: nextId++,
    type: "task",
    name: "New Activity",
    startWeek: startWeek,
    duration: 2,
    isExpanded: true,
    color: "#4f46e5",
  });
  render();
}

function addMilestone() {
  pushUndo();
  const startWeek = snapNewItemStartWeek(getLastItemEndWeek(), "milestone");
  ensureTimelineCovers(startWeek);
  let group = items.find(i => i.type === "milestones-group" && !i.parentId);
  if (!group) {
    const groupId = nextId++;
    items.push({
      id: groupId,
      type: "milestones-group",
      name: "Milestones",
      startWeek: startWeek,
      duration: 0,
      isExpanded: true,
    });
    group = items.find(i => i.id === groupId);
  }
  items.push({
    id: nextId++,
    type: "milestone",
    name: "New Milestone",
    startWeek: startWeek,
    duration: 0,
    parentId: group.id,
    color: "#f59e0b",
  });
  render();
}

function toggleExpand(id) {
  const item = items.find((i) => i.id === id);
  if (item) {
    item.isExpanded = item.isExpanded === false ? true : false;
    render();
  }
}

function deleteItem(id) {
  pushUndo();
  const idsToDelete = getSubtreeIds(id);
  items = items.filter((i) => !idsToDelete.has(i.id));
  links = links.filter(
    (l) => !idsToDelete.has(l.fromId) && !idsToDelete.has(l.toId),
  );
  render();
}

function duplicateItem(id) {
  pushUndo();
  const itemToDuplicate = items.find((i) => i.id === id);
  if (!itemToDuplicate) return;

  function deepClone(oldId, newParentId) {
    const original = items.find((i) => i.id === oldId);
    const newId = nextId++;
    const copy = JSON.parse(JSON.stringify(original));
    copy.id = newId;
    copy.parentId = newParentId;
    copy.isLocked = true;
    if (oldId === id) {
      copy.name += " (Copy)";
    }
    items.push(copy);

    const children = items
      .filter((i) => i.parentId === oldId)
      .map((i) => i.id);
    children.forEach((childId) => deepClone(childId, newId));
  }

  deepClone(id, itemToDuplicate.parentId);
  render();
}

function updateName(id, name) {
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (item) item.name = name;
  render();
}

// Data-driven milestone configs (consolidated from 3 separate lists)
const PROJECT_MILESTONE_CONFIGS = {
  "Child Project": {
    milestones: [
      { name: "HU-SP (Start Of Project)", offset: 0 },
      { name: "HU-CO (Contract)", offset: 47 },
      { name: "HU-VC (Agreement to Build Off Tool Unit)", offset: 60 },
      { name: "HU-PT (Agreement to Build Plant Trial)", offset: 73 },
      { name: "HU-PPC (Product & Process Certification)", offset: 118 },
      { name: "HU-SOP (Start Of Production)", offset: 125 },
    ],
    activityCount: 3,
  },
  "Brother Project": {
    milestones: [
      { name: "HU-SP (Start Of Project)", offset: 0 },
      { name: "HU-PC (Pre-Contract)", offset: 46 },
      { name: "HU-CO (Contract)", offset: 88 },
      { name: "HU-VC (Agreement to Build Off Tool Unit)", offset: 103 },
      { name: "HU-PT (Agreement to Build Plant Trial)", offset: 118 },
      { name: "HU-PPC (Product & Process Certification)", offset: 139 },
      { name: "HU-SOP (Start Of Production)", offset: 169 },
    ],
    activityCount: 3,
  },
  "Mother Project": {
    milestones: [
      { name: "HU-INT (Powertrain Intention)", offset: 0 },
      { name: "HU-PreC (Powertrain PreConcept)", offset: 21 },
      { name: "HU-SP (Start Of Project)", offset: 42 },
      { name: "HU-PC (Pre-Contract)", offset: 104 },
      { name: "HU-CO (Contract)", offset: 151 },
      { name: "HU-VC (Agreement to Build Off Tool Unit)", offset: 166 },
      { name: "HU-PT (Agreement to Build Plant Trial)", offset: 181 },
      { name: "HU-PPC (Product & Process Certification)", offset: 209 },
      { name: "HU-SOP (Start Of Production)", offset: 238 },
    ],
    activityCount: 5,
  },
};

function updateProjectName(id, name) {
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (item) {
    if (name === "Custom Project") {
      item.name = "Custom Project Name";
    } else {
      item.name = name;
      const startWeek = item.startWeek;
      const config = PROJECT_MILESTONE_CONFIGS[name];

      if (config) {
        // Add milestones-group container, then milestones as its children
        const mgId = nextId++;
        items.push({
          id: mgId,
          type: "milestones-group",
          name: "Milestones",
          startWeek: startWeek,
          duration: 0,
          parentId: id,
          isExpanded: true,
        });
        config.milestones.forEach((mObj) => {
          items.push({
            id: nextId++,
            type: "milestone",
            name: mObj.name,
            startWeek: startWeek + mObj.offset,
            duration: 0,
            parentId: mgId,
            color: getMilestoneColor(mObj.name),
          });
        });

        // Add activities
        for (let i = 0; i < config.activityCount; i++) {
          items.push({
            id: nextId++,
            type: "task",
            name: "Activity " + (i + 1),
            startWeek: startWeek + i * 2,
            duration: 2,
            parentId: id,
            isExpanded: true,
            color: "#4f46e5",
          });
        }

        // Ensure enough years exist to fit all milestones and activities
        const lastMsOffset = config.milestones.reduce((max, m) => Math.max(max, m.offset), 0);
        const lastActEnd = startWeek + (config.activityCount - 1) * 2 + 2;
        const maxWeekNeeded = Math.max(startWeek + lastMsOffset, lastActEnd);
        ensureTimelineCovers(maxWeekNeeded);
      }
    }
    render();
  }
}

function updateMilestoneName(id, name) {
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (item) {
    if (name === "Custom Milestone") {
      item.name = "Custom Milestone Name";
    } else {
      item.name = name;
      item.color = getMilestoneColor(name);
    }
  }
  render();
}

// Internal: apply a template with optional overrides (called by config modal or directly)
function _applyTemplateCore(id, templateId, nameOverride, durationWeeksOverride, colorOverride) {
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (!item) return;

  if (templateId === "Custom Activity") {
    item.name = nameOverride || "Custom Activity Name";
    render();
    return;
  }

  const t = customTemplates.find((ct) => ct.id === templateId);
  if (!t) { render(); return; }

  item.name     = nameOverride          || t.name;
  item.duration = durationWeeksOverride || t.duration;
  item.color    = colorOverride         || t.color;
  item.isExpanded = true;
  let maxWeekNeeded = item.startWeek + item.duration;

  if (t.composition && t.composition.length > 0) {
    // Scale sub-item durations proportionally when duration was overridden
    const scale = durationWeeksOverride ? (durationWeeksOverride / (t.duration || 1)) : 1;
    let currentStartWeek = Math.round((item.startWeek || 1) * 25) / 25;

    t.composition.forEach((compItem) => {
      const subT = customTemplates.find((ct) => ct.id === compItem.templateId);
      if (subT) {
        const subDur = (subT.duration || 0.04) * scale;
        if (compItem.quantity > 1) {
          const wrapperId = nextId++;
          items.push({ id: wrapperId, type: "task", name: subT.name,
            duration: subDur * compItem.quantity, color: subT.color,
            startWeek: currentStartWeek, parentId: item.id, isExpanded: true });
          maxWeekNeeded = Math.max(maxWeekNeeded, currentStartWeek + subDur * compItem.quantity);
          for (let q = 0; q < compItem.quantity; q++) {
            items.push({ id: nextId++, type: "task", name: subT.name + " " + (q + 1),
              duration: subDur, color: subT.color, startWeek: currentStartWeek, parentId: wrapperId });
            maxWeekNeeded = Math.max(maxWeekNeeded, currentStartWeek + subDur);
            currentStartWeek += subDur;
          }
        } else {
          items.push({ id: nextId++, type: "task", name: subT.name,
            duration: subDur, color: subT.color,
            startWeek: currentStartWeek, parentId: item.id });
          maxWeekNeeded = Math.max(maxWeekNeeded, currentStartWeek + subDur);
          currentStartWeek += subDur;
        }
      }
    });
  }

  ensureTimelineCovers(maxWeekNeeded);
  render();
}

function applyActivityTemplate(id, templateId) {
  if (templateId === "Custom Activity") {
    _applyTemplateCore(id, templateId);
    return;
  }
  const t = customTemplates.find((ct) => ct.id === templateId);
  if (!t) return;

  // If template asks for parameters, show the config modal instead of applying directly
  if (t.askName || t.askDuration || t.askColor) {
    if (typeof showTemplateConfigModal === "function") showTemplateConfigModal(id, templateId);
    return;
  }

  _applyTemplateCore(id, templateId);
}

function updateDuration(id, duration) {
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (item) {
    const oldDuration = item.duration || 0;
    let parsedDuration = parseFloat(duration) || 0.04;
    if (zoomMode === "days") {
      parsedDuration = parsedDuration / 5.0;
    } else if (zoomMode === "months") {
      parsedDuration = parsedDuration * 4.0;
    }
    item.duration = Math.max(0.04, parsedDuration);
    const deltaEnd = parseFloat((item.duration - oldDuration).toFixed(2));
    if (deltaEnd !== 0) {
      propagateLinks(id, 0, deltaEnd, new Set());
    }
    render();
  }
}

function updateComment(id, comment) {
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (item) item.comment = comment;
  render();
}

// ── Custom Column Management ──────────────────────────────────
function addCustomColumn() {
  pushUndo();
  var name = "Column " + nextColId;
  customColumns.push({
    id: "col_" + Date.now() + "_" + nextColId,
    name: name,
    width: 150,
    visible: true,
  });
  nextColId++;
  if (typeof markChanged === "function") markChanged();
  render();
}

function removeCustomColumn(colId) {
  pushUndo();
  customColumns = customColumns.filter(function (c) { return c.id !== colId; });
  items.forEach(function (it) { if (it.customData) delete it.customData[colId]; });
  delete columnFilters[colId];
  if (typeof markChanged === "function") markChanged();
  render();
  if (typeof _renderColManager === "function") _renderColManager();
}

function renameCustomColumn(colId, newName) {
  pushUndo();
  var col = customColumns.find(function (c) { return c.id === colId; });
  if (col) col.name = newName;
  if (typeof markChanged === "function") markChanged();
  render();
}

function toggleCustomColumn(colId) {
  var col = customColumns.find(function (c) { return c.id === colId; });
  if (col) col.visible = !col.visible;
  if (typeof markChanged === "function") markChanged();
  render();
}

function updateCustomColumnCell(itemId, colId, value) {
  pushUndo();
  var item = items.find(function (i) { return i.id === itemId; });
  if (!item) return;
  if (!item.customData) item.customData = {};
  item.customData[colId] = value;
  if (typeof markChanged === "function") markChanged();
}

function moveCustomColumn(colId, direction) {
  var idx = customColumns.findIndex(function (c) { return c.id === colId; });
  if (idx < 0) return;
  var swapIdx = idx + direction;
  if (swapIdx < 0 || swapIdx >= customColumns.length) return;
  pushUndo();
  var tmp = customColumns[idx];
  customColumns[idx] = customColumns[swapIdx];
  customColumns[swapIdx] = tmp;
  if (typeof markChanged === "function") markChanged();
  render();
}

function updateColor(id, color) {
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (item) item.color = color;
  render();
}

function updateYear(index, value) {
  pushUndo();
  const newYearNum = parseInt(value);
  if (isNaN(newYearNum)) { render(); return; }
  years = years.map((_, yi) => String(newYearNum + yi - index));
  _updateYearCountLabel();
  render();
}

function updateCompletion(id, value) {
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (item) {
    item.completion = Math.min(100, Math.max(0, parseInt(value) || 0));
    render();
  }
}

function toggleSettings(id) {
  const item = items.find((i) => i.id === id);
  if (item) {
    item.isSettingsExpanded = !item.isSettingsExpanded;
    render();
  }
}

function setMilestoneWeek(id, val) {
  pushUndo();
  val = val.trim().replace(/\s/g, "");
  const parts = val.split(".");
  const yyww = parts[0];
  const d = parts[1] ? Math.max(1, Math.min(5, parseInt(parts[1], 10))) : 1;
  if (yyww.length !== 4) return;
  const inputYY = yyww.slice(0, 2);
  const inputWW = parseInt(yyww.slice(2), 10);
  if (isNaN(inputWW) || inputWW < 1) return;
  const offset = (d - 1) * 0.2;
  for (let yi = 0; yi < years.length; yi++) {
    if (getShortYear(yi) === inputYY) {
      const weeksInYear = getIsoWeeksInYear(years[yi]);
      if (inputWW > weeksInYear) return;
      const absWeek = getAbsWeekFromYearWeek(yi, inputWW) + offset;
      const item = items.find((i) => i.id === id);
      if (item) {
        const delta = absWeek - item.startWeek;
        item.startWeek = absWeek;
        if (delta !== 0) {
          propagateLinks(id, delta, delta, new Set());
        }
        render();
      }
      return;
    }
  }
}

function updateMilestoneDate(id, val) {
  const item = items.find((i) => i.id === id);
  if (!item) return;
  const plan = _planYYWWD(val.trim());
  if (!plan) return;
  const projectedStartWeek = item.startWeek + (plan.shiftWeeks || 0);
  const needsTimelineExpansion =
    (plan.prependYears && plan.prependYears.length > 0) ||
    (plan.appendYears && plan.appendYears.length > 0);
  const projectedDelta = parseFloat((plan.absWeek - projectedStartWeek).toFixed(1));
  if (!needsTimelineExpansion && projectedDelta === 0) return;

  pushUndo();
  _applyYYWWDPlan(plan);

  const updatedItem = items.find((i) => i.id === id);
  if (!updatedItem) {
    render();
    return;
  }
  const delta = parseFloat((plan.absWeek - updatedItem.startWeek).toFixed(1));
  if (delta === 0) {
    render();
    return;
  }

  updatedItem.startWeek = plan.absWeek;
  propagateLinks(id, delta, delta, new Set());
  render();
}

function toggleGlobalLock() {
  globalLocked = !globalLocked;
  items.forEach((i) => { i.isLocked = globalLocked; });
  const btn = document.getElementById('global-lock-btn');
  const lockLabel = document.getElementById('lock-label');
  if (lockLabel) lockLabel.textContent = globalLocked ? 'Lock: ON' : 'Lock: OFF';
  else if (btn) btn.textContent = globalLocked ? '🔒 Lock: ON' : '🔓 Lock: OFF';
  render();
}

function toggleLock(id, event) {
  const item = items.find((i) => i.id === id);
  if (!item) return;

  const isCurrentlyLocked = item.isLocked === true;
  const newLockedState = !isCurrentlyLocked;
  const isolateOnly = !!(event && (event.ctrlKey || event.metaKey));

  item.isLocked = newLockedState;

  if (isolateOnly) {
    render();
    return;
  }

  function cascadeDescendants(parentId) {
    items
      .filter((i) => i.parentId === parentId)
      .forEach((child) => {
        child.isLocked = newLockedState;
        cascadeDescendants(child.id);
      });
  }
  cascadeDescendants(id);

  if (!newLockedState) {
    const visitedLink = new Set([id]);
    function unlockLinked(itemId) {
      links.forEach((l) => {
        const otherId =
          l.fromId === itemId
            ? l.toId
            : l.toId === itemId
              ? l.fromId
              : null;
        if (otherId !== null && !visitedLink.has(otherId)) {
          visitedLink.add(otherId);
          const other = items.find((i) => i.id === otherId);
          if (other) {
            other.isLocked = false;
            cascadeDescendants(otherId);
          }
          unlockLinked(otherId);
        }
      });
    }
    unlockLinked(id);
  }

  render();
}

function addYear() {
  const last = years[years.length - 1];
  const nextYear = String(parseInt(last) + 1 || new Date().getFullYear() + years.length);
  years.push(nextYear);
  _updateYearCountLabel();
  render();
}

function prependYear() {
  pushUndo();
  const firstYear = parseInt(years[0]) || new Date().getFullYear();
  const newYear = String(firstYear - 1);
  years.unshift(newYear);
  const addedWeeks = getIsoWeeksInYear(newYear);
  items.forEach(item => { item.startWeek += addedWeeks; });
  _updateYearCountLabel();
  render();
}

function removeYear() {
  if (years.length <= 1) return;
  years.pop();
  hiddenYears.delete(years.length);
  _updateYearCountLabel();
  render();
}

function toggleYearVisibility(yi) {
  if (hiddenYears.has(yi)) {
    hiddenYears.delete(yi);
  } else {
    hiddenYears.add(yi);
  }
  render();
}

// ── Multi-select bulk actions ─────────────────────────────────

function deleteSelected() {
  if (selectedIds.size === 0) return;
  if (!confirm("Delete " + selectedIds.size + " selected item(s)? This cannot be undone.")) return;
  pushUndo();
  const allToDelete = new Set();
  selectedIds.forEach(id => getSubtreeIds(id).forEach(sid => allToDelete.add(sid)));
  items = items.filter(i => !allToDelete.has(i.id));
  links = links.filter(l => !allToDelete.has(l.fromId) && !allToDelete.has(l.toId));
  selectedIds.clear();
  _lastSelectedId = null;
  if (typeof _updateMultiselectBar === "function") _updateMultiselectBar();
  render();
}

function lockSelected(lock) {
  if (selectedIds.size === 0) return;
  pushUndo();
  items.forEach(i => { if (selectedIds.has(i.id)) i.isLocked = lock; });
  render();
}

function recolorSelected(color) {
  if (selectedIds.size === 0 || !color) return;
  pushUndo();
  items.forEach(i => {
    if (selectedIds.has(i.id) && i.type !== "milestones-group") i.color = color;
  });
  render();
}

function assignSelected(personId) {
  if (selectedIds.size === 0 || !personId) return;
  pushUndo();
  items.forEach(i => {
    if (selectedIds.has(i.id) && (i.type === "task" || i.type === "project" || i.type === "family")) {
      if (!i.assignees) i.assignees = [];
      if (!i.assignees.includes(personId)) i.assignees.push(personId);
    }
  });
  render();
}

function moveSelected(weekDelta) {
  if (selectedIds.size === 0 || !weekDelta) return;
  pushUndo();
  items.forEach(i => {
    if (selectedIds.has(i.id)) i.startWeek = Math.max(1, i.startWeek + weekDelta);
  });
  render();
}

function _updateYearCountLabel() {
  const el = document.getElementById("year-count-label");
  if (el) el.textContent = years.length + " yr" + (years.length === 1 ? "" : "s");
  const removeBtn = document.getElementById("remove-year-btn");
  if (removeBtn) removeBtn.disabled = years.length <= 1;
}
