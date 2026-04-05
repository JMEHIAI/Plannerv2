// ═══════════════════════════════════════════════════════════════
//  ITEM CRUD OPERATIONS
// ═══════════════════════════════════════════════════════════════

// Consolidated addSubItem - replaces 3 near-identical functions
function addSubItem(parentId, type, defaults) {
  pushUndo();
  const parent = items.find((i) => i.id === parentId);
  if (!parent) return;
  parent.isExpanded = true;
  const subItems = items.filter((i) => i.parentId === parentId);
  let startWeek = parent.startWeek;
  if (subItems.length > 0) {
    const lastSub = subItems[subItems.length - 1];
    startWeek = lastSub.startWeek + (lastSub.duration || 0);
  } else {
    startWeek = parent.startWeek + (parent.duration || 0);
  }
  startWeek = Math.round(Math.min(years.length * 52, startWeek) * 5) / 5;

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
  const startWeek = getLastItemEndWeek();
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
  const startWeek = getLastItemEndWeek();
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
  const startWeek = getLastItemEndWeek();
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
  const startWeek = getLastItemEndWeek();
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
        while (years.length * 52 < maxWeekNeeded) {
          const lastYr = parseInt(years[years.length - 1]) || 2025;
          years.push(String(lastYr + 1));
        }
        _updateYearCountLabel();
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

function applyActivityTemplate(id, templateId) {
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (item) {
    if (templateId === "Custom Activity") {
      item.name = "Custom Activity Name";
    } else {
      const t = customTemplates.find((ct) => ct.id === templateId);
      if (t) {
        item.name = t.name;
        item.duration = t.duration;
        item.color = t.color;
        item.isExpanded = true;

        if (t.composition && t.composition.length > 0) {
          let currentStartWeek = Math.round((item.startWeek || 1) * 25) / 25;

          t.composition.forEach((compItem) => {
            const subT = customTemplates.find((ct) => ct.id === compItem.templateId);
            if (subT) {
              const subDur = subT.duration || 0.04;
              if (compItem.quantity > 1) {
                const wrapperId = nextId++;
                items.push({
                  id: wrapperId,
                  type: "task",
                  name: subT.name,
                  duration: subDur * compItem.quantity,
                  color: subT.color,
                  startWeek: currentStartWeek,
                  parentId: item.id,
                  isExpanded: true,
                });
                for (let q = 0; q < compItem.quantity; q++) {
                  items.push({
                    id: nextId++,
                    type: "task",
                    name: subT.name + " " + (q + 1),
                    duration: subDur,
                    color: subT.color,
                    startWeek: currentStartWeek,
                    parentId: wrapperId,
                  });
                  currentStartWeek = currentStartWeek + subDur;
                }
              } else {
                items.push({
                  id: nextId++,
                  type: "task",
                  name: subT.name,
                  duration: subDur,
                  color: subT.color,
                  startWeek: currentStartWeek,
                  parentId: item.id,
                });
                currentStartWeek = currentStartWeek + subDur;
              }
            }
          });
        }
      }
    }
  }
  render();
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
  const item = items.find((i) => i.id === id);
  if (item) item.comment = comment;
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
  const oldYearNum = parseInt(years[index]);
  const yearDelta = newYearNum - oldYearNum;
  if (yearDelta !== 0) {
    const weekShift = -yearDelta * 52;
    items.forEach(item => {
      item.startWeek = Math.max(0.2, item.startWeek + weekShift);
    });
    years = years.map(y => String(parseInt(y) + yearDelta));
  }
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
  if (isNaN(inputWW) || inputWW < 1 || inputWW > 52) return;
  const offset = (d - 1) * 0.2;
  for (let yi = 0; yi < 3; yi++) {
    if (getShortYear(yi) === inputYY) {
      const absWeek = yi * 52 + inputWW + offset;
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
  pushUndo();
  const item = items.find((i) => i.id === id);
  if (!item) return;
  const newWeek = parseYYWWD(val.trim());
  if (newWeek !== null) {
    item.startWeek = newWeek;
    render();
  }
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

  item.isLocked = newLockedState;

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
  years.unshift(String(firstYear - 1));
  items.forEach(item => { item.startWeek += 52; });
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

function _updateYearCountLabel() {
  const el = document.getElementById("year-count-label");
  if (el) el.textContent = years.length + " yr" + (years.length === 1 ? "" : "s");
  const removeBtn = document.getElementById("remove-year-btn");
  if (removeBtn) removeBtn.disabled = years.length <= 1;
}
