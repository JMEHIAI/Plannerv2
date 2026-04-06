// ═══════════════════════════════════════════════════════════════
//  UI: TOOLBARS, MODALS, TEAM, TEMPLATES, ASSIGNEES
// ═══════════════════════════════════════════════════════════════

function toggleToolbars() {
  const wrapper = document.getElementById("toolbars-wrapper");
  const btnIcon = document.getElementById("toggle-toolbars-icon");
  const btnText = document.getElementById("toggle-toolbars-text");
  const isVisible = wrapper.style.display !== "none";

  if (isVisible) {
    wrapper.style.display = "none";
    btnIcon.textContent = "👁️";
    btnText.textContent = "Show Buttons";
    localStorage.setItem("planner_toolbars_visible", "false");
  } else {
    wrapper.style.display = "block";
    btnIcon.textContent = "👁️";
    btnText.textContent = "Hide Buttons";
    localStorage.setItem("planner_toolbars_visible", "true");
  }
}

function toggleManual() {
  const overlay = document.getElementById("manual-overlay");
  if (overlay.classList.contains("show")) {
    overlay.classList.remove("show");
  } else {
    overlay.classList.add("show");
  }
}

function toggleTeamModal() {
  const overlay = document.getElementById("team-modal-overlay");
  if (overlay.style.display === "flex") {
    overlay.style.display = "none";
  } else {
    overlay.style.display = "flex";
    renderTeamList();
  }
}

function renderTeamList() {
  const list = document.getElementById("team-list");
  if (people.length === 0) {
    list.innerHTML =
      '<div style="padding: 12px; color: #94a3b8; text-align: center; font-style: italic; font-size: 13px;">No team members added yet.</div>';
    return;
  }

  let html = "";
  people.forEach((p) => {
    html += `
      <div style="display: flex; align-items: center; justify-content: space-between; padding: 10px 12px; border-bottom: 1px solid #f1f5f9;">
        <div style="font-size: 14px; font-weight: 500; color: #1e293b;">${p.name.replace(/</g, "&lt;")}</div>
        <div style="display: flex; align-items: center; gap: 8px;">
          <button onclick="deletePerson('${p.id}')" style="margin-left: 8px; background: none; border: none; color: #ef4444; font-size: 16px; cursor: pointer; padding: 0 4px;" title="Remove">&times;</button>
        </div>
      </div>
    `;
  });
  list.innerHTML = html;
  renderWorkloadView();
}

function renderWorkloadView() {
  const container = document.getElementById("workload-view");
  if (!container) return;

  if (people.length === 0) {
    container.innerHTML =
      '<div style="text-align:center;color:#94a3b8;font-size:13px;padding:16px 0 4px;">Add team members to see workload.</div>';
    return;
  }

  const assignedItems = items.filter(
    (item) =>
      item.assignees &&
      item.assignees.length > 0 &&
      item.type !== "family" &&
      (item.duration || 0) > 0,
  );

  if (assignedItems.length === 0) {
    container.innerHTML =
      '<div style="border-top:1px solid #e2e8f0;padding-top:12px;margin-top:8px;"><h3 style="margin:0 0 8px;font-size:14px;color:#475569;font-weight:600;">📊 Team Capacity</h3><div style="text-align:center;color:#94a3b8;font-size:12px;font-style:italic;padding:8px 0;">No assignments yet. Enable ⚙️ Settings and use 👥 to assign people to tasks.</div></div>';
    return;
  }

  const activePeopleIds = new Set(people.map((p) => p.id));

  // Find week range of all assignments (1-based absolute weeks)
  let minWeek = Infinity,
    maxWeek = -Infinity;
  assignedItems.forEach((item) => {
    minWeek = Math.min(minWeek, item.startWeek);
    maxWeek = Math.max(maxWeek, item.startWeek + (item.duration || 1));
  });
  const spanWeeks = maxWeek - minWeek;
  const totalPeople = people.length;

  // --- Determine granularity ---
  const granularity =
    zoomMode === "days" ? (spanWeeks <= 8 ? "days" : "weeks") : "months";

  // Build slots array
  const slots = [];

  if (granularity === "days") {
    const dayNames = ["Mon", "Tue", "Wed", "Thu", "Fri"];
    for (let w = minWeek; w < maxWeek; w++) {
      const weekInfo = getYearWeekInfo(w);
      for (let d = 0; d < 5; d++) {
        slots.push({
          start: w + d / 5,
          end: w + (d + 1) / 5,
          label: dayNames[d],
          sublabel: d === 0 ? "W" + weekInfo.relWeek : "",
          yearIdx: weekInfo.yearIndex,
          year: weekInfo.yearLabel || "",
        });
      }
    }
  } else if (granularity === "weeks") {
    for (let w = minWeek; w < maxWeek; w++) {
      const weekInfo = getYearWeekInfo(w);
      slots.push({
        start: w,
        end: w + 1,
        label: "W" + weekInfo.relWeek,
        sublabel: weekInfo.relWeek === 1 ? weekInfo.yearLabel || "" : "",
        yearIdx: weekInfo.yearIndex,
        year: weekInfo.yearLabel || "",
      });
    }
  } else {
    for (let y = 0; y < years.length; y++) {
      let absWeek = getAbsWeekFromYearWeek(y, 1);
      getMonthWeekSpans(years[y]).forEach((m) => {
        const start = absWeek,
          end = absWeek + m.weeks;
        if (start < maxWeek && end > minWeek) {
          slots.push({
            start,
            end,
            label: m.name,
            sublabel: "",
            yearIdx: y,
            year: years[y] || "",
          });
        }
        absWeek += m.weeks;
      });
    }
  }

  // For each slot, compute which people are assigned
  const slotData = slots.map((slot) => {
    const assignedSet = new Set();
    const taskMap = {};
    const assignmentCount = {};
    assignedItems.forEach((item) => {
      const iStart = item.startWeek;
      const iEnd = item.startWeek + (item.duration || 1);
      const overlap =
        Math.min(iEnd, slot.end) - Math.max(iStart, slot.start);
      if (overlap > 0) {
        item.assignees.forEach((pid) => {
          if (!activePeopleIds.has(pid)) return;
          assignedSet.add(pid);
          assignmentCount[pid] = (assignmentCount[pid] || 0) + 1;
          if (!taskMap[pid]) taskMap[pid] = [];
          taskMap[pid].push(item.name || "Task");
        });
      }
    });
    const overloadedCount = Object.values(assignmentCount).filter((count) => count > 1).length;
    return { ...slot, assignedSet, taskMap, assignmentCount, overloadedCount };
  });

  // --- Build HTML ---
  const granLabel =
    granularity === "days"
      ? "day"
      : granularity === "weeks"
        ? "week"
        : "month";
  let html = `<div style="border-top:1px solid #e2e8f0; padding-top:16px; margin-top:8px;">`;
  html += `<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px;">
    <h3 style="margin:0; font-size:14px; color:#475569; font-weight:600;">📊 Team Capacity</h3>
    <div style="font-size:11px; color:#94a3b8;">${totalPeople} member${totalPeople !== 1 ? "s" : ""} · by ${granLabel}</div>
  </div>`;

  // Slot header labels (scrollable wrapper for many columns)
  const scrollable = slots.length > 20;
  if (scrollable) {
    html += `<div style="overflow-x:auto; padding-bottom:4px;">
      <div style="min-width:${slots.length * 22 + 110}px;">`;
  }

  html += `<div style="display:flex; margin-bottom:2px;">
    <div style="width:110px; flex-shrink:0;"></div>
    <div style="display:flex; flex:1; gap:1px;">`;

  slotData.forEach((s, si) => {
    const isNewYear = si > 0 && s.yearIdx !== slotData[si - 1].yearIdx;
    const borderL = isNewYear ? "2px solid #cbd5e1" : "none";
    html += `<div style="flex:1; text-align:center; font-size:8px; color:#94a3b8; border-left:${borderL}; white-space:nowrap; overflow:hidden;">
      ${s.sublabel ? `<div style="font-weight:700;color:#64748b;font-size:8px;">${s.sublabel}</div>` : ""}
      <div>${s.label}</div>
    </div>`;
  });
  html += `</div></div>`;

  // Person rows
  people.forEach((p) => {
    html += `<div style="display:flex; align-items:center; margin-bottom:3px; height:18px;">
      <div style="width:110px; flex-shrink:0; font-size:12px; color:#1e293b; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; padding-right:6px;" title="${p.name}">${p.name}</div>
      <div style="display:flex; flex:1; gap:1px; height:100%;">`;

    slotData.forEach((s, si) => {
      const assignmentCount = s.assignmentCount[p.id] || 0;
      const isAssigned = assignmentCount > 0;
      const isOverbooked = assignmentCount > 1;
      const taskNames = s.taskMap[p.id]
        ? [...new Set(s.taskMap[p.id])].join(", ")
        : "";
      const isNewYear = si > 0 && s.yearIdx !== slotData[si - 1].yearIdx;
      const borderL = isNewYear ? "2px solid #cbd5e1" : "none";
      const tooltip = isAssigned
        ? `${p.name}: ${taskNames} (${s.label} ${s.year})${isOverbooked ? ` • ${assignmentCount} overlapping assignments` : ""}`
        : `${p.name}: free (${s.label} ${s.year})`;
      const bgColor = isOverbooked ? "#ef4444" : isAssigned ? "#6366f1" : "#e2e8f0";
      html += `<div style="flex:1; background:${bgColor}; border-radius:1px; height:100%; border-left:${borderL}; opacity:${isAssigned ? "1" : "0.35"};" title="${tooltip}"></div>`;
    });

    html += `</div></div>`;
  });

  // Divider
  html += `<div style="height:1px; background:#f1f5f9; margin:6px 0;"></div>`;

  // Free slots summary bar
  html += `<div style="display:flex; align-items:center; height:22px;">
    <div style="width:110px; flex-shrink:0; font-size:11px; font-weight:600; color:#475569; padding-right:6px;">Free slots</div>
    <div style="display:flex; flex:1; gap:1px; height:100%;">`;

  slotData.forEach((s, si) => {
    const assignedCount = s.assignedSet.size;
    const freeCount = totalPeople - assignedCount;
    const isNewYear = si > 0 && s.yearIdx !== slotData[si - 1].yearIdx;
    const borderL = isNewYear ? "2px solid #cbd5e1" : "none";
    const color =
      s.overloadedCount > 0
        ? "#ef4444"
        : freeCount === 0
        ? "#ef4444"
        : assignedCount / totalPeople >= 0.7
          ? "#f59e0b"
          : "#22c55e";
    const tooltip = s.overloadedCount > 0
      ? `${freeCount} of ${totalPeople} free on ${s.label} ${s.year}; ${s.overloadedCount} overloaded`
      : `${freeCount} of ${totalPeople} free on ${s.label} ${s.year}`;
    html += `<div style="flex:1; background:${color}; border-radius:1px; height:100%; border-left:${borderL}; display:flex; align-items:center; justify-content:center;" title="${tooltip}">
      ${slots.length <= 40 ? `<span style="font-size:8px; font-weight:700; color:white; line-height:1;">${freeCount}</span>` : ""}
    </div>`;
  });

  html += `</div></div>`;

  // Legend
  html += `<div style="display:flex; gap:12px; font-size:10px; color:#94a3b8; margin-top:8px; flex-wrap:wrap;">
    <span style="display:flex;align-items:center;gap:4px;"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#6366f1;"></span>Assigned</span>
    <span style="display:flex;align-items:center;gap:4px;"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#22c55e;"></span>Free</span>
    <span style="display:flex;align-items:center;gap:4px;"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#f59e0b;"></span>≥70% booked</span>
    <span style="display:flex;align-items:center;gap:4px;"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;background:#ef4444;"></span>Full / overloaded</span>
  </div>`;

  if (scrollable) {
    html += `</div></div>`;
  }
  html += `</div>`;
  container.innerHTML = html;
}

// --- Assignee dropdown ---

// Close all open assignee dropdowns
function closeAllAssigneeDropdowns() {
  document
    .querySelectorAll(".assignee-dropdown-popup")
    .forEach((d) => d.remove());
  activeAssigneeMenu = null;
}

// --- Activity Templates Management ---
function toggleTemplatesModal() {
  const modal = document.getElementById("templates-modal-overlay");
  if (modal.style.display === "none") {
    modal.style.display = "flex";
    renderTemplatesList();
  } else {
    modal.style.display = "none";
  }
}

function saveTemplatesLocal() {
  try {
    localStorage.setItem(
      "diaglo_activity_templates",
      JSON.stringify(customTemplates),
    );
  } catch (e) {
    console.warn("Could not save templates", e);
  }
}

function addTemplate() {
  const nameInput = document.getElementById("new-template-name");
  const durInput = document.getElementById("new-template-duration");
  const colorInput = document.getElementById("new-template-color");

  const name = nameInput.value.trim();
  const durDays = parseFloat(durInput.value);
  const color = colorInput.value;

  if (!name) {
    alert("Please enter an activity name.");
    return;
  }
  if (isNaN(durDays) || durDays <= 0) {
    alert("Please enter a valid day length.");
    return;
  }

  // check duplicate name
  if (
    customTemplates.some(
      (t) => t.name.toLowerCase() === name.toLowerCase(),
    )
  ) {
    alert("A template with this name already exists.");
    return;
  }

  // Convert days to weeks internally
  const durWeeks = durDays / 5;

  customTemplates.push({
    id: "t_" + Date.now(),
    name,
    duration: durWeeks,
    color,
  });

  nameInput.value = "";
  durInput.value = "10";
  colorInput.value = "#4f46e5";

  saveTemplatesLocal();
  renderTemplatesList();
  render(); // refresh dropdowns in the grid
}

function deleteTemplate(id) {
  if (
    !confirm(
      "Remove this template? Existing activities will not be affected.",
    )
  )
    return;
  customTemplates = customTemplates.filter((t) => t.id !== id);
  saveTemplatesLocal();
  renderTemplatesList();
  render(); // refresh dropdowns in the grid
}

function renderTemplatesList() {
  const list = document.getElementById("templates-list");
  const searchEl = document.getElementById("template-search");
  let query = searchEl ? searchEl.value.toLowerCase() : "";

  const filteredTemplates = customTemplates.filter((t) =>
    t.name.toLowerCase().includes(query),
  );

  if (filteredTemplates.length === 0) {
    if (query) {
      list.innerHTML =
        '<div style="padding: 12px; color: #94a3b8; text-align: center; font-style: italic; font-size: 13px;">No templates found matching "' +
        query.replace(/</g, "&lt;") +
        '".</div>';
    } else {
      list.innerHTML =
        '<div style="padding: 12px; color: #94a3b8; text-align: center; font-style: italic; font-size: 13px;">No custom templates saved yet.</div>';
    }
    return;
  }

  let html = "";
  filteredTemplates.forEach((t) => {
    // Only allow basic templates to be dragged to avoid infinite loops
    const isMaster = t.composition && t.composition.length > 0;
    const dragAttrs = isMaster
      ? 'opacity: 0.6; cursor: pointer;" title="Click to edit this Master Template" onclick="loadMasterForEdit(\'' +
      t.id +
      "')\""
      : 'cursor: grab;" draggable="true" ondragstart="handleTemplateDragStart(event, \'' +
      t.id +
      '\')" title="Drag into Template Composer"';

    html += `
      <div style="display: flex; align-items: center; justify-content: space-between; padding: 10px 12px; border-bottom: 1px solid #f1f5f9; ${dragAttrs}">
        <div style="display: flex; align-items: center; gap: 8px;">
          <div style="width: 16px; height: 16px; border-radius: 4px; background: ${t.color}; border: 1px solid rgba(0,0,0,0.1);"></div>
          <div style="font-size: 14px; font-weight: 500; color: #1e293b;">${t.name.replace(/</g, "&lt;")} ${isMaster ? '<span style="font-size:10px;background:#fef3c7;padding:2px 4px;border-radius:4px;">Master</span>' : ""}</div>
        </div>
        <div style="display: flex; align-items: center; gap: 12px;">
          <span style="font-size: 12px; color: #64748b;">${Number((t.duration * 5).toFixed(1))} days</span>
          <button onclick="event.stopPropagation(); deleteTemplate('${t.id}')" style="background: none; border: none; color: #ef4444; font-size: 16px; cursor: pointer; padding: 0 4px;" title="Remove template">&times;</button>
        </div>
      </div>
    `;
  });
  list.innerHTML = html;
}

// --- Master Template Composer Logic ---

function loadMasterForEdit(id) {
  const t = customTemplates.find((ct) => ct.id === id);
  if (!t || !t.composition) return;

  composerState = JSON.parse(JSON.stringify(t.composition));
  document.getElementById("composer-master-name").value = t.name;
  editingMasterId = t.id;

  renderComposerUI();
}

function handleTemplateDragStart(event, id) {
  event.dataTransfer.setData("text/plain", id);
}

function allowComposeDrop(event) {
  event.preventDefault(); // allow drop
}

function handleComposeDrop(event) {
  event.preventDefault();
  const templateId = event.dataTransfer.getData("text/plain");
  if (!templateId) return;

  const t = customTemplates.find((ct) => ct.id === templateId);
  if (!t) return;

  // Add to composer
  composerState.push({ templateId: t.id, quantity: 1 });
  renderComposerUI();
}

function removeComposerItem(index) {
  composerState.splice(index, 1);
  renderComposerUI();
}

function updateComposerQuantity(index, qty) {
  const parsed = parseFloat(qty);
  if (!isNaN(parsed) && parsed > 0) {
    composerState[index].quantity = parsed;
  }
  renderComposerUI();
}

function renderComposerUI() {
  const list = document.getElementById("composer-list");
  const durLabel = document.getElementById("composer-total-duration");

  // Update button text based on edit mode
  const saveBtn = document.getElementById("composer-save-btn");
  const saveNewBtn = document.getElementById("composer-save-new-btn");
  if (saveBtn) {
    saveBtn.innerText = editingMasterId ? "Update Master" : "Save Master";
  }
  if (saveNewBtn) {
    saveNewBtn.style.display = editingMasterId ? "inline-block" : "none";
  }

  if (composerState.length === 0) {
    list.innerHTML =
      '<div style="color: #94a3b8; font-size: 13px; text-align: center; margin-top: 40px; font-style: italic;">Drag &amp; drop saved templates here</div>';
    durLabel.textContent = "0 days";
    return;
  }

  let html = "";
  let totalDur = 0;

  composerState.forEach((item, index) => {
    const t = customTemplates.find((ct) => ct.id === item.templateId);
    if (!t) return;

    const lineDur = t.duration * item.quantity;
    totalDur += lineDur;

    html += `
      <div style="display: flex; align-items: center; justify-content: space-between; background: white; border: 1px solid #cbd5e1; border-radius: 4px; padding: 8px;">
        <div style="display: flex; align-items: center; gap: 8px;">
          <div style="width:12px; height:12px; border-radius:2px; background:${t.color};"></div>
          <span style="font-size: 13px; font-weight: 500;">${t.name.replace(/</g, "&lt;")}</span>
        </div>
        <div style="display: flex; align-items: center; gap: 8px;">
          <span style="font-size: 11px; color:#64748b;">Qty</span>
          <input type="number" min="0.1" step="0.1" value="${item.quantity}" onchange="updateComposerQuantity(${index}, this.value)" style="width: 48px; padding: 4px; font-size: 12px; border: 1px solid #cbd5e1; border-radius: 4px;" title="Quantity">
          <span style="font-size: 12px; font-weight: bold; width: 44px; text-align: right; color:#334155;">${Number((lineDur * 5).toFixed(1))} days</span>
          <button onclick="removeComposerItem(${index})" style="background: none; border: none; color: #ef4444; font-size: 16px; cursor: pointer; padding: 0 4px;" title="Remove">&times;</button>
        </div>
      </div>
    `;
  });

  list.innerHTML = html;
  durLabel.textContent = Number((totalDur * 5).toFixed(1)) + " days";
}

function saveAsNewMasterTemplate() {
  if (!editingMasterId) return;

  // Wipe current tracking ID to trigger the standard Save fork path
  editingMasterId = null;

  const nameInput = document.getElementById("composer-master-name");
  let name = nameInput.value.trim();

  // If no name change was made, auto-append " (Copy)" to avoid duplicate alert collision
  if (
    name &&
    customTemplates.some(
      (t) => t.name.toLowerCase() === name.toLowerCase(),
    )
  ) {
    nameInput.value = name + " (Copy)";
  }

  saveMasterTemplate();
}

function saveMasterTemplate() {
  const nameInput = document.getElementById("composer-master-name");
  const name = nameInput.value.trim();

  if (!name) {
    alert("Please enter a name for the Master Template.");
    return;
  }
  if (composerState.length === 0) {
    alert("Please drag at least one template into the composer first.");
    return;
  }

  // Check for duplicate names, ignoring the currently edited template
  if (
    customTemplates.some(
      (t) =>
        t.name.toLowerCase() === name.toLowerCase() &&
        t.id !== editingMasterId,
    )
  ) {
    alert("A template with this name already exists.");
    return;
  }

  let totalDur = 0;
  let masterColor = "#10b981"; // Default green for master

  const compArray = composerState.map((item, index) => {
    const ct = customTemplates.find((t) => t.id === item.templateId);
    if (ct) {
      totalDur += ct.duration * item.quantity;
      if (index === 0) masterColor = ct.color; // Inherit color from first item
    }
    return { templateId: item.templateId, quantity: item.quantity };
  });

  if (editingMasterId) {
    // Update existing master
    const existingMaster = customTemplates.find(
      (t) => t.id === editingMasterId,
    );
    if (existingMaster) {
      existingMaster.name = name;
      existingMaster.duration = totalDur;
      existingMaster.color = masterColor;
      existingMaster.composition = compArray;
    } else {
      customTemplates.push({
        id: "t_m_" + Date.now(),
        name: name,
        duration: totalDur,
        color: masterColor,
        composition: compArray,
      });
    }
  } else {
    // Create new master
    customTemplates.push({
      id: "t_m_" + Date.now(),
      name: name,
      duration: totalDur,
      color: masterColor,
      composition: compArray,
    });
  }

  // Reset composer
  composerState = [];
  editingMasterId = null;
  nameInput.value = "";

  saveTemplatesLocal();
  renderTemplatesList();
  renderComposerUI();
  render(); // Refresh dropdowns in the grid
}

// Close assignee dropdowns when clicking outside
document.addEventListener("click", function (e) {
  if (
    !e.target.closest(".assignee-dropdown-popup") &&
    !e.target.closest(".assignee-select")
  ) {
    closeAllAssigneeDropdowns();
  }
});

function toggleAssigneeDropdown(itemId, event) {
  event.stopPropagation();

  if (activeAssigneeMenu === itemId) {
    closeAllAssigneeDropdowns();
    return;
  }

  closeAllAssigneeDropdowns();
  activeAssigneeMenu = itemId;

  const btn = event.currentTarget || event.target;
  const item = items.find((i) => i.id === itemId);
  if (!item) return;

  // Calculate position relative to button
  const rect = btn.getBoundingClientRect();

  const popup = document.createElement("div");
  popup.className = "assignee-dropdown-popup";
  popup.style.cssText =
    "position:fixed;" +
    "top:" +
    (rect.bottom + 4) +
    "px;" +
    "left:" +
    rect.left +
    "px;" +
    "background:white;" +
    "border:1px solid #e2e8f0;" +
    "border-radius:6px;" +
    "padding:10px 12px;" +
    "z-index:99999;" +
    "min-width:170px;" +
    "box-shadow:0 4px 16px rgba(0,0,0,0.18);";
  popup.onclick = (e) => e.stopPropagation();

  if (people.length === 0) {
    popup.innerHTML =
      '<div style="font-size:12px;color:#64748b;text-align:center;">No team members.<br>Add via <b>Team &amp; Workload</b>.</div>';
  } else {
    let html =
      '<div style="font-weight:600;font-size:12px;margin-bottom:8px;color:#1e293b;border-bottom:1px solid #f1f5f9;padding-bottom:6px;">Assign People:</div>';
    people.forEach((p) => {
      const isAssigned = item.assignees && item.assignees.includes(p.id);
      html += `<label style="display:flex;align-items:center;gap:8px;font-size:13px;margin-bottom:6px;cursor:pointer;">
        <input type="checkbox" ${isAssigned ? "checked" : ""} onchange="toggleAssignee(${itemId},'${p.id}',this.checked)">
        <span>${p.name.replace(/</g, "&lt;")}</span>
      </label>`;
    });
    popup.innerHTML = html;
  }

  document.body.appendChild(popup);
}

function toggleAssignee(itemId, personId, checked) {
  const item = items.find((i) => i.id === itemId);
  if (!item) return;
  if (!item.assignees) item.assignees = [];
  if (checked) {
    if (!item.assignees.includes(personId)) item.assignees.push(personId);
  } else {
    item.assignees = item.assignees.filter((id) => id !== personId);
  }
  closeAllAssigneeDropdowns();
  render();
}

function addPerson() {
  const nameInput = document.getElementById("new-person-name");
  const name = nameInput.value.trim();

  if (!name) {
    alert("Please enter a name.");
    return;
  }

  const id = "p_" + nextPersonId++;
  people.push({ id, name });

  nameInput.value = "";

  renderTeamList();
  render(); // Re-render planner to update assignee dropdowns
}

function deletePerson(id) {
  if (
    !confirm(
      "Remove this person? This will not remove them from existing task assignments.",
    )
  )
    return;
  people = people.filter((p) => p.id !== id);
  renderTeamList();
  render(); // Re-render planner
}

function toggleZoom() {
  if (zoomMode === "weeks") {
    zoomMode = "days";
  } else if (zoomMode === "days") {
    zoomMode = "months";
  } else {
    zoomMode = "weeks";
  }

  const zoomLabel = document.getElementById("zoom-label");
  const zoomLabelText = zoomMode === "days" ? "Zoom: Days" : zoomMode === "months" ? "Zoom: Months" : "Zoom: Weeks";
  if (zoomLabel) {
    zoomLabel.textContent = zoomLabelText;
  } else {
    document.getElementById("zoom-toggle-btn").innerHTML = "🔍 " + zoomLabelText;
  }
  // Invalidate header height cache so new zoom mode is measured fresh
  _cachedHeaderHeights = null;
  // Yield to the browser so the button label paints before the heavy render
  requestAnimationFrame(() => requestAnimationFrame(() => render()));
}

function toggleAllSettings() {
  showSettings = !showSettings;
  render();
}

function toggleComments() {
  showComments = !showComments;
  // Also show/hide all block comments (ctrl+click popups)
  items.forEach(item => {
    if (item.blockComment) {
      item.blockComment.isOpen = showComments;
    }
  });
  render();
}

function toggleHighlight(val) {
  if (highlightedWeek === val) {
    highlightedWeek = null;
  } else {
    highlightedWeek = val;
  }
  render();
}

function toggleFilterMenu(column, event) {
  if (activeFilterMenu === column) {
    activeFilterMenu = null;
  } else {
    activeFilterMenu = column;
    if (event && event.currentTarget && gridEl) {
      const triggerRect = event.currentTarget.getBoundingClientRect();
      const gridRect = gridEl.getBoundingClientRect();
      filterMenuPosition = {
        top: Math.max(8, triggerRect.bottom - gridRect.top + 6),
        left: Math.max(8, triggerRect.left - gridRect.left),
      };
    }
  }
  render();
  if (event) event.stopPropagation();
}

function handleFilterChange(column, value, event) {
  if (!filters[column]) filters[column] = [];
  const checked = event.target.checked;

  let valuesToChange = [value];

  if (column === "name") {
    // Find matching parent names and cascade the change to all of their descendants' names
    const affectedNames = new Set([value]);
    const matchingItems = items.filter(
      (i) => i.name === value || (i.name || "").trim() === value,
    );

    function addDescendantNames(parentId) {
      const children = items.filter((i) => i.parentId === parentId);
      children.forEach((child) => {
        if (child.name) affectedNames.add(child.name);
        addDescendantNames(child.id);
      });
    }

    matchingItems.forEach((item) => addDescendantNames(item.id));
    valuesToChange = Array.from(affectedNames);
  }

  valuesToChange.forEach((val) => {
    if (!checked) {
      // If unchecked, add to the filter list (hide this item)
      if (!filters[column].includes(val)) filters[column].push(val);
    } else {
      // If checked, remove from the filter list (show this item)
      filters[column] = filters[column].filter((v) => v !== val);
    }
  });

  render();
}

function clearFilter(column) {
  filters[column] = [];
  render();
}

function filterNameFilterOptions(query) {
  const menu = document.querySelector(".filter-menu.open");
  if (!menu) return;

  const normalized = (query || "").trim().toLowerCase();
  let visibleCount = 0;

  menu.querySelectorAll(".filter-name-option").forEach((option) => {
    const haystack = option.getAttribute("data-filter-search") || "";
    const isVisible = !normalized || haystack.includes(normalized);
    option.style.display = isVisible ? "flex" : "none";
    if (isVisible) visibleCount++;
  });

  const emptyState = menu.querySelector(".filter-menu-empty");
  if (emptyState) emptyState.style.display = visibleCount === 0 ? "block" : "none";
}

// Close filter menus when clicking outside
document.addEventListener("click", () => {
  if (activeFilterMenu) {
    activeFilterMenu = null;
    render();
  }
});

// Add beforeunload event to remind the user to save to CSV
window.addEventListener("beforeunload", function (e) {
  e.preventDefault();
  e.returnValue = "Please remember to save to CSV before leaving!";
});

// Initialize Toolbars Visibility
document.addEventListener("DOMContentLoaded", () => {
  const toolbarsVisible = localStorage.getItem("planner_toolbars_visible");
  if (toolbarsVisible === "false") {
    const wrapper = document.getElementById("toolbars-wrapper");
    const btnText = document.getElementById("toggle-toolbars-text");
    wrapper.style.display = "none";
    if (btnText) btnText.textContent = "Show Buttons";
  }
  const plannerContainer = document.querySelector(".planner-container");
  if (plannerContainer) plannerContainer.addEventListener("mousedown", startPan);
});
