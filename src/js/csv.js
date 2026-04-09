// ═══════════════════════════════════════════════════════════════
//  CSV EXPORT / IMPORT / MERGE
// ═══════════════════════════════════════════════════════════════

function _csvQuote(value) {
  return '"' + String(value == null ? "" : value).replace(/"/g, '""') + '"';
}

function _csvJson(value) {
  if (value == null) return '""';
  if (Array.isArray(value) && value.length === 0) return '""';
  if (
    typeof value === "object" &&
    !Array.isArray(value) &&
    Object.keys(value).length === 0
  ) {
    return '""';
  }
  return _csvQuote(JSON.stringify(value));
}

function _parseCsvJson(value) {
  if (!value || value === '""') return undefined;
  try {
    return JSON.parse(value);
  } catch (e) {
    return undefined;
  }
}

function _newTemplateCategoryId() {
  return "cat_" + Date.now() + "_" + Math.floor(Math.random() * 1000000);
}

function _newImportedTemplateId(prefix) {
  return prefix + Date.now() + "_" + Math.floor(Math.random() * 1000000);
}

function _parseTemplateCategoryRow(values) {
  if (values.length >= 3) {
    return {
      id: values[1] || "",
      name: values[2] || "",
    };
  }
  return {
    id: "",
    name: values[1] || "",
  };
}

function _parseTemplateRow(values) {
  const composition = _parseCsvJson(values[4]);
  return {
    oldId: values[9] || "",
    name: values[1] || "",
    duration: parseFloat(values[2]) || 0,
    color: values[3] || "#4f46e5",
    composition: Array.isArray(composition) ? composition : null,
    rawCategoryId: values[5] || "",
    askName: values[6] === "true",
    askDuration: values[7] === "true",
    askColor: values[8] === "true",
  };
}

function _mergeTemplateCategoriesFromCsv(rows) {
  const categoryIdMap = {};
  rows.forEach(function (row) {
    const name = (row.name || "").trim();
    if (!name) return;

    let existing = null;
    if (row.id) {
      existing = templateCategories.find(function (cat) {
        return cat.id === row.id;
      });
    }
    if (!existing) {
      existing = templateCategories.find(function (cat) {
        return (cat.name || "").toLowerCase() === name.toLowerCase();
      });
    }
    if (!existing) {
      let newId = row.id || _newTemplateCategoryId();
      if (templateCategories.some(function (cat) { return cat.id === newId; })) {
        newId = _newTemplateCategoryId();
      }
      existing = { id: newId, name: name };
      templateCategories.push(existing);
    }
    if (row.id) categoryIdMap[row.id] = existing.id;
  });
  return categoryIdMap;
}

function _mergeTemplatesFromCsv(rows, categoryIdMap) {
  const templateIdMap = {};
  const pendingComposition = [];

  rows.forEach(function (row) {
    const name = (row.name || "").trim();
    if (!name) return;

    const resolvedCategoryId = row.rawCategoryId
      ? (categoryIdMap[row.rawCategoryId] ||
        (templateCategories.find(function (cat) { return cat.id === row.rawCategoryId; }) || {}).id ||
        null)
      : null;

    let template = customTemplates.find(function (t) {
      return (t.name || "").toLowerCase() === name.toLowerCase();
    });
    let shouldApplyComposition = false;

    if (!template) {
      template = {
        id: _newImportedTemplateId(row.composition && row.composition.length > 0 ? "t_m_" : "t_"),
        name: name,
        duration: row.duration,
        color: row.color,
      };
      if (resolvedCategoryId) template.category = resolvedCategoryId;
      if (row.askName) template.askName = true;
      if (row.askDuration) template.askDuration = true;
      if (row.askColor) template.askColor = true;
      if (row.composition && row.composition.length > 0) {
        template.composition = JSON.parse(JSON.stringify(row.composition));
        shouldApplyComposition = true;
      }
      customTemplates.push(template);
    } else {
      if (resolvedCategoryId && !template.category) template.category = resolvedCategoryId;
      if (row.askName) template.askName = true;
      if (row.askDuration) template.askDuration = true;
      if (row.askColor) template.askColor = true;
      if ((!template.color || template.color === "#4f46e5") && row.color) template.color = row.color;
      if ((!template.duration || template.duration <= 0) && row.duration > 0) template.duration = row.duration;
      if (
        row.composition &&
        row.composition.length > 0 &&
        (!template.composition || template.composition.length === 0)
      ) {
        template.composition = JSON.parse(JSON.stringify(row.composition));
        shouldApplyComposition = true;
      }
    }

    if (row.oldId) templateIdMap[row.oldId] = template.id;
    if (shouldApplyComposition) {
      pendingComposition.push({
        template: template,
        composition: row.composition,
      });
    }
  });

  pendingComposition.forEach(function (entry) {
    if (!entry.composition || entry.composition.length === 0) return;
    entry.template.composition = entry.composition.map(function (part) {
      return {
        templateId: templateIdMap[part.templateId] || part.templateId,
        quantity: part.quantity,
      };
    });
  });
}

function _parseCustomColumnRow(values) {
  return {
    id: values[1] || "",
    name: values[2] || "Column",
    width: parseInt(values[3], 10) || 150,
    visible: values[4] !== "false",
  };
}

function _newCustomColumnId() {
  const id = "col_" + Date.now() + "_" + nextColId;
  nextColId++;
  return id;
}

function _mergeCustomColumnsFromCsv(rows) {
  const columnIdMap = {};
  rows.forEach(function (row) {
    const name = (row.name || "Column").trim() || "Column";
    let existing = null;
    if (row.id) {
      existing = customColumns.find(function (col) {
        return col.id === row.id && (col.name || "") === name;
      });
    }
    if (!existing) {
      existing = customColumns.find(function (col) {
        return (col.name || "").toLowerCase() === name.toLowerCase();
      });
    }
    if (!existing) {
      let newId = row.id || _newCustomColumnId();
      if (customColumns.some(function (col) { return col.id === newId; })) {
        newId = _newCustomColumnId();
      }
      existing = {
        id: newId,
        name: name,
        width: row.width || 150,
        visible: row.visible !== false,
      };
      customColumns.push(existing);
    }
    if (row.id) columnIdMap[row.id] = existing.id;
  });
  return columnIdMap;
}

/** Build and return the full CSV string for the current planner state. */
function generateCSVString(options) {
  options = options || {};
  const csv = ["Years," + years.join(",")];
  csv.push(
    "Globals," +
    zoomMode +
    "," +
    showSettings +
    "," +
    showComments +
    "," +
    commentWidth +
    "," +
    nameWidthBase +
    "," +
    (highlightedWeek || "") +
    "," +
    showLinks,
  );
  if (options.includeEditorLock && typeof _buildSavedByCSVLine === "function") {
    const savedByLine = _buildSavedByCSVLine();
    if (savedByLine) csv.push(savedByLine);
  }
  csv.push(
    "FiltersName," +
    filters.name
      .map((n) => '"' + n.replace(/"/g, '""') + '"')
      .join(","),
  );
  csv.push(
    "FiltersType," +
    filters.type
      .map((n) => '"' + n.replace(/"/g, '""') + '"')
      .join(","),
  );
  templateCategories.forEach(function (cat) {
    csv.push(
      "TemplateCategory," +
      _csvQuote(cat.id || "") +
      "," +
      _csvQuote(cat.name || ""),
    );
  });
  customTemplates.forEach((t) => {
    const safeComposition = _csvJson(t.composition);
    csv.push(
      "Template," +
      _csvQuote(t.name) +
      "," +
      t.duration +
      "," +
      _csvQuote(t.color || "") +
      "," +
      safeComposition +
      "," +
      _csvQuote(t.category || "") +
      "," +
      (t.askName === true) +
      "," +
      (t.askDuration === true) +
      "," +
      (t.askColor === true) +
      "," +
      _csvQuote(t.id || ""),
    );
  });
  alarms.forEach((a) => {
    csv.push(
      "Alarm," +
      a.id +
      "," +
      a.itemId +
      "," +
      a.date +
      "," +
      a.time +
      "," +
      a.duration +
      ',"' +
      a.title.replace(/"/g, '""') +
      '"',
    );
  });
  holidays.forEach((h) => {
    // Format: Holiday,date[,person1|person2|...]
    const peoplePart = h.people && h.people.length > 0 ? "," + h.people.join("|") : "";
    csv.push("Holiday," + h.date + peoplePart);
  });

  people.forEach((p) => {
    csv.push("Person," + p.id + ',"' + p.name.replace(/"/g, '""') + '"');
  });

  links.forEach((l) => {
    csv.push(
      "Link," +
      l.id +
      "," +
      l.fromId +
      "," +
      l.fromAnchor +
      "," +
      l.toId +
      "," +
      l.toAnchor,
    );
  });

  // Custom columns definitions
  customColumns.forEach((col) => {
    csv.push(
      "CustomCol," +
      '"' + col.id.replace(/"/g, '""') + '",' +
      '"' + col.name.replace(/"/g, '""') + '",' +
      col.width + "," +
      col.visible,
    );
  });

  // Column filters
  var colFilterKeys = Object.keys(columnFilters);
  colFilterKeys.forEach(function (colId) {
    var vals = columnFilters[colId];
    if (vals && vals.length > 0) {
      csv.push(
        "ColFilter," +
        '"' + colId.replace(/"/g, '""') + '",' +
        vals.map(function (v) { return '"' + v.replace(/"/g, '""') + '"'; }).join(","),
      );
    }
  });

  csv.push(
    "Id,Type,Name,StartWeek,Duration,ParentId,IsExpanded,Comment,Color,IsLocked,MilestoneRow,BlockCommentData,AssigneesData,Completion,CustomData",
  );
  items.forEach((item) => {
    const safeName = '"' + item.name.replace(/"/g, '""') + '"';
    const safeComment =
      '"' + (item.comment || "").replace(/"/g, '""') + '"';
    const safeBlockComment = item.blockComment
      ? '"' + JSON.stringify(item.blockComment).replace(/"/g, '""') + '"'
      : '""';
    const assigneesStr =
      item.assignees && item.assignees.length > 0
        ? '"' + item.assignees.join(";") + '"'
        : '""';
    const customDataStr = item.customData && Object.keys(item.customData).length > 0
      ? '"' + JSON.stringify(item.customData).replace(/"/g, '""') + '"'
      : '""';
    csv.push(
      item.id +
      "," +
      item.type +
      "," +
      safeName +
      "," +
      item.startWeek +
      "," +
      item.duration +
      "," +
      (item.parentId || "") +
      "," +
      (item.isExpanded !== false) +
      "," +
      safeComment +
      "," +
      (item.color || "") +
      "," +
      (item.isLocked === true) +
      "," +
      (item.milestoneRow || 0) +
      "," +
      safeBlockComment +
      "," +
      assigneesStr +
      "," +
      (item.completion || 0) +
      "," +
      customDataStr,
    );
  });

  return csv.join(String.fromCharCode(10));
}

function exportCSV() {
  const now = new Date();
  const dateStr =
    now.getFullYear() + "-" + pad2(now.getMonth() + 1) + "-" + pad2(now.getDate()) +
    "_" + pad2(now.getHours()) + "-" + pad2(now.getMinutes());

  const csvContent = generateCSVString();
  const defaultName = "project_plan_" + dateStr + ".csv";

  // Try native Save As dialog (Chrome/Edge File System Access API)
  if (window.showSaveFilePicker) {
    window.showSaveFilePicker({
      suggestedName: defaultName,
      types: [{ description: "CSV Files", accept: { "text/csv": [".csv"] } }],
    }).then(function(handle) {
      return handle.createWritable().then(function(writable) {
        return writable.write(csvContent).then(function() { return writable.close(); });
      });
    }).catch(function(err) {
      if (err.name !== "AbortError") console.error(err);
    });
    return;
  }

  // Fallback: show a small rename dialog then trigger download
  _pendingCsvContent = csvContent;
  const inp = document.getElementById("_csvFilenameInput");
  inp.value = defaultName.replace(/\.csv$/i, "");
  document.getElementById("_csvSaveModal").style.display = "flex";
  setTimeout(function() { inp.select(); }, 50);
}

function _confirmCsvSave() {
  let name = (document.getElementById("_csvFilenameInput").value || "project_plan").trim();
  if (!name.toLowerCase().endsWith(".csv")) name += ".csv";
  const blob = new Blob([_pendingCsvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
  document.getElementById("_csvSaveModal").style.display = "none";
}

// CSV helper: parse one CSV line respecting quoted fields
function _parseCSVLine(line) {
  const result = [];
  let cur = "",
    inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      if (inQ && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else inQ = !inQ;
    } else if (c === "," && !inQ) {
      result.push(cur);
      cur = "";
    } else cur += c;
  }
  result.push(cur);
  return result;
}

function _importCSVText(text) {
    const lines = text.split(String.fromCharCode(10));
    const newItems = [];
    const newAlarms = [];
    const newPeople = [];
    const newHolidays = [];
    const newLinks = [];
    const importedTemplateCategories = [];
    const importedTemplates = [];
    const newCustomColumns = [];
    const newColumnFilters = {};
    let tempId = 1;
    let tempAlarmId = 1;
    let tempPersonId = 1;

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;

      const values = _parseCSVLine(line);

      if (values[0] === "Years") {
        const imported = values.slice(1).filter(Boolean);
        years = imported.length > 0
          ? imported
          : ["Year 1", "Year 2", "Year 3"];
        continue;
      }
      if (values[0] === "Globals") {
        zoomMode = values[1] || "weeks";
        showSettings = values[2] === "true";
        showComments = values[3] === "true";
        commentWidth = parseInt(values[4]) || 200;
        nameWidthBase = parseInt(values[5]) || 350;
        if (values[6]) {
          if (values[6].startsWith("W")) highlightedWeek = values[6];
          else if (values[6].includes(".")) highlightedWeek = parseFloat(values[6]);
          else highlightedWeek = parseInt(values[6], 10);
          if (Number.isNaN(highlightedWeek)) highlightedWeek = null;
        } else {
          highlightedWeek = null;
        }
        showLinks = values[7] !== "false"; // default true if missing
        continue;
      }
      if (values[0] === "FiltersName") {
        filters.name = values.slice(1).filter((v) => v !== "");
        continue;
      }
      if (values[0] === "FiltersType") {
        filters.type = values.slice(1).filter((v) => v !== "");
        continue;
      }
      if (values[0] === "TemplateCategory") {
        importedTemplateCategories.push(_parseTemplateCategoryRow(values));
        continue;
      }
      if (values[0] === "Template") {
        importedTemplates.push(_parseTemplateRow(values));
        continue;
      }
      if (values[0] === "EditorLock") {
        continue;
      }
      if (values[0] === "Holiday") {
        // Support old format (bare date string) and new format (date + optional people)
        const _hDate = values[1];
        const _hPeople = values[2] ? values[2].split("|").filter(Boolean) : [];
        newHolidays.push({ date: _hDate, people: _hPeople });
        continue;
      }
      if (values[0] === "Person") {
        const pId = values[1];
        newPeople.push({
          id: pId,
          name: values[2],
        });
        const numId = parseInt(pId.replace("p_", ""), 10);
        if (!isNaN(numId) && numId >= tempPersonId)
          tempPersonId = numId + 1;
        continue;
      }
      if (values[0] === "Alarm") {
        const aId = parseInt(values[1]);
        newAlarms.push({
          id: aId,
          itemId:
            values[2] === "null" || !values[2]
              ? null
              : parseInt(values[2]),
          date: values[3],
          time: values[4],
          duration: parseInt(values[5]),
          title: values[6],
        });
        if (aId >= tempAlarmId) tempAlarmId = aId + 1;
        continue;
      }
      if (values[0] === "Link") {
        newLinks.push({
          id: parseInt(values[1]),
          fromId: parseInt(values[2]),
          fromAnchor: values[3],
          toId: parseInt(values[4]),
          toAnchor: values[5],
        });
        continue;
      }
      if (values[0] === "CustomCol") {
        newCustomColumns.push(_parseCustomColumnRow(values));
        continue;
      }
      if (values[0] === "ColFilter") {
        var cfId = values[1];
        var cfVals = values.slice(2).filter(function (v) { return v !== ""; });
        if (cfVals.length > 0) newColumnFilters[cfId] = cfVals;
        continue;
      }

      if (values[0] === "Id" || values[0] === "Type") continue;

      if (values.length >= 4) {
        if (
          values[0] === "task" ||
          values[0] === "milestone" ||
          values[0] === "project"
        ) {
          newItems.push({
            id: tempId++,
            type: values[0],
            name: values[1],
            startWeek: parseFloat(values[2]) || 1,
            duration: parseFloat(values[3]) || 0,
          });
        } else {
          const parsedId = parseInt(values[0]);
          let blockCommentData;
          try {
            blockCommentData = values[11] && values[11] !== '""'
              ? JSON.parse(values[11])
              : undefined;
          } catch (e) {
            blockCommentData = undefined;
          }
          let customDataParsed;
          try {
            customDataParsed = values[14] && values[14] !== '""' && values[14] !== ''
              ? JSON.parse(values[14])
              : undefined;
          } catch (e) {
            customDataParsed = undefined;
          }
          newItems.push({
            id: parsedId,
            type: values[1],
            name: values[2],
            startWeek: parseFloat(values[3]) || 1,
            duration: parseFloat(values[4]) || 0,
            parentId:
              values[5] && values[5] !== ""
                ? parseInt(values[5])
                : undefined,
            isExpanded: values[6] === "false" ? false : true,
            comment: values[7] || "",
            color: values[8] || "",
            isLocked: values[9] === "true",
            milestoneRow: values[10] ? parseInt(values[10]) : 0,
            blockComment: blockCommentData,
            assignees:
              values[12] && values[12] !== '""'
                ? values[12].split(";")
                : [],
            completion: values[13] ? parseInt(values[13]) : 0,
            customData: customDataParsed,
          });
          if (parsedId >= tempId) tempId = parsedId + 1;
        }
      }
    }

    if (newItems.length > 0) {
      items = newItems;
      alarms = newAlarms;
      people = newPeople;
      holidays = newHolidays;
      links = newLinks;
      const templateCategoryIdMap = _mergeTemplateCategoriesFromCsv(importedTemplateCategories);
      _mergeTemplatesFromCsv(importedTemplates, templateCategoryIdMap);
      nextLinkId =
        newLinks.length > 0
          ? Math.max(...newLinks.map((l) => l.id)) + 1
          : 1;
      nextId = tempId;
      nextAlarmId = tempAlarmId;
      nextPersonId = tempPersonId;
      if (newCustomColumns.length > 0) {
        customColumns = newCustomColumns;
        nextColId = newCustomColumns.length + 1;
      }
      columnFilters = newColumnFilters;
      if (typeof _saveCategoriesLocal === "function") _saveCategoriesLocal();
      if (typeof saveTemplatesLocal === "function") saveTemplatesLocal();
      if (typeof _syncCategoryDropdowns === "function") _syncCategoryDropdowns();
      if (typeof renderTemplatesList === "function") renderTemplatesList();
      render();
      if (typeof markChanged === 'function') markChanged();
    } else {
      alert("Invalid CSV format.");
    }
}

function importCSV(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (event) => _importCSVText(event.target.result);
  reader.readAsText(file);
  e.target.value = "";
}

function mergeCSV(e) {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (ev) {
    const text = ev.target.result;
    const lines = text
      .split("\n")
      .map((l) => l.trim())
      .filter((l) => l);

    // Find the start line for items, collecting alarms along the way
    let startLine = 0;
    const incomingAlarms = [];
    const incomingLinks = [];
    const incomingTemplateCategories = [];
    const incomingTemplates = [];
    const incomingCustomColumns = [];
    const oldToNewPersonIds = {};
    while (
      startLine < lines.length &&
      !lines[startLine].startsWith("Id,Type,Name")
    ) {
      const v = _parseCSVLine(lines[startLine]);
      if (v[0] === "Alarm") {
        incomingAlarms.push({
          oldItemId: v[2] === "null" || !v[2] ? null : parseInt(v[2]),
          date: v[3],
          time: v[4],
          duration: parseInt(v[5]),
          title: v[6],
        });
      } else if (v[0] === "Link") {
        incomingLinks.push({
          oldFromId: parseInt(v[2]),
          fromAnchor: v[3],
          oldToId: parseInt(v[4]),
          toAnchor: v[5],
        });
      } else if (v[0] === "TemplateCategory") {
        incomingTemplateCategories.push(_parseTemplateCategoryRow(v));
      } else if (v[0] === "Template") {
        incomingTemplates.push(_parseTemplateRow(v));
      } else if (v[0] === "CustomCol") {
        incomingCustomColumns.push(_parseCustomColumnRow(v));
      } else if (v[0] === "Holiday") {
        const _mhDate = v[1];
        const _mhPeople = v[2] ? v[2].split("|").filter(Boolean) : [];
        if (!holidays.some(h => h.date === _mhDate)) {
          holidays.push({ date: _mhDate, people: _mhPeople });
        }
      } else if (v[0] === "Person") {
        const oldId = v[1];
        const newId = "p_" + nextPersonId++;
        oldToNewPersonIds[oldId] = newId;
        people.push({
          id: newId,
          name: v[2],
        });
      }
      startLine++;
    }
    startLine++; // skip Id,Type,Name,... header

    const templateCategoryIdMap = _mergeTemplateCategoriesFromCsv(incomingTemplateCategories);
    _mergeTemplatesFromCsv(incomingTemplates, templateCategoryIdMap);
    const customColIdMap = _mergeCustomColumnsFromCsv(incomingCustomColumns);

    // First pass: build old-id → new-id map
    const idMap = {};
    for (let i = startLine; i < lines.length; i++) {
      const v = _parseCSVLine(lines[i]);
      if (!v[0]) continue;
      const oldId = parseInt(v[0]);
      if (!isNaN(oldId)) idMap[oldId] = nextId++;
    }

    // Second pass: create items with remapped IDs and parentIds
    const newItems = [];
    for (let i = startLine; i < lines.length; i++) {
      const v = _parseCSVLine(lines[i]);
      if (!v[0]) continue;
      const oldId = parseInt(v[0]);
      if (isNaN(oldId)) continue;
      const oldParentId = v[5] ? parseInt(v[5]) : null;

      let newAssignees = [];
      if (v[12] && v[12] !== '""') {
        const oldAssignees = v[12].split(";");
        oldAssignees.forEach((oldAssId) => {
          if (oldToNewPersonIds[oldAssId]) {
            newAssignees.push(oldToNewPersonIds[oldAssId]);
          } else {
            newAssignees.push(oldAssId);
          }
        });
      }

      let blockCommentData;
      try {
        blockCommentData = v[11] && v[11] !== '""' ? JSON.parse(v[11]) : undefined;
      } catch (e) {
        blockCommentData = undefined;
      }
      let customDataParsed = _parseCsvJson(v[14]);
      if (customDataParsed && typeof customDataParsed === "object") {
        const remappedCustomData = {};
        Object.keys(customDataParsed).forEach(function (oldKey) {
          const newKey = customColIdMap[oldKey] || oldKey;
          remappedCustomData[newKey] = customDataParsed[oldKey];
        });
        customDataParsed = remappedCustomData;
      } else {
        customDataParsed = undefined;
      }

      newItems.push({
        id: idMap[oldId],
        type: v[1] || "task",
        name: v[2] || "Unnamed",
        startWeek: parseFloat(v[3]) || 1,
        duration: parseFloat(v[4]) || 0,
        parentId:
          oldParentId && idMap[oldParentId]
            ? idMap[oldParentId]
            : undefined,
        isExpanded: v[6] !== "false",
        comment: v[7] || "",
        color: v[8] || "",
        isLocked: v[9] === "true",
        milestoneRow: v[10] ? parseInt(v[10]) : 0,
        blockComment: blockCommentData,
        assignees: newAssignees,
        completion: v[13] ? parseInt(v[13], 10) : 0,
        customData: customDataParsed,
      });
    }

    if (newItems.length === 0) {
      alert("No valid rows found in the merged CSV.");
      return;
    }

    // Remap holiday people IDs (holidays were parsed before person ID mapping was complete)
    holidays.forEach(h => {
      if (h.people && h.people.length > 0) {
        h.people = h.people.map(pid => oldToNewPersonIds[pid] || pid);
      }
    });

    // Push merged items
    items.push(...newItems);

    // Push merged alarms with remapped item references
    incomingAlarms.forEach((a) => {
      alarms.push({
        id: nextAlarmId++,
        itemId: a.oldItemId ? idMap[a.oldItemId] : null,
        date: a.date,
        time: a.time,
        duration: a.duration,
        title: a.title,
      });
    });

    incomingLinks.forEach((l) => {
      const fromId = idMap[l.oldFromId];
      const toId = idMap[l.oldToId];
      if (!fromId || !toId) return;
      links.push({
        id: nextLinkId++,
        fromId,
        fromAnchor: l.fromAnchor,
        toId,
        toAnchor: l.toAnchor,
      });
    });

    if (typeof _saveCategoriesLocal === "function") _saveCategoriesLocal();
    if (typeof saveTemplatesLocal === "function") saveTemplatesLocal();
    if (typeof _syncCategoryDropdowns === "function") _syncCategoryDropdowns();
    if (typeof renderTemplatesList === "function") renderTemplatesList();
    render();
    if (typeof markChanged === 'function') markChanged();
    let msg = "Merged " + newItems.length + " item rows";
    if (incomingAlarms.length)
      msg += " and " + incomingAlarms.length + " alarms";
    if (incomingLinks.length)
      msg += " and " + incomingLinks.length + " links";
    msg += ' from "' + file.name + '" into the planner.';
    alert(msg);
  };
  reader.readAsText(file);
  e.target.value = "";
}

function exportJPEG() {
  // Load html2canvas dynamically (cached after first load)
  function loadLib() {
    return new Promise(function (resolve, reject) {
      if (window.html2canvas) { resolve(); return; }
      var s = document.createElement("script");
      s.src = "https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js";
      s.onload = resolve;
      s.onerror = function () { reject(new Error("Could not load html2canvas")); };
      document.head.appendChild(s);
    });
  }

  var container = document.querySelector(".planner-container");
  var grid = document.getElementById("grid");
  if (!container || !grid) return;

  loadLib().then(function () {
    var exportBg = getExportBackground(container);
    var toolbarBtnIcon = document.getElementById("toggle-toolbars-icon");

    html2canvas(document.body, {
      scale: Math.max(window.devicePixelRatio || 1, 2),
      useCORS: true,
      backgroundColor: exportBg,
      width: window.innerWidth,
      height: window.innerHeight,
      windowWidth: window.innerWidth,
      windowHeight: window.innerHeight,
      scrollX: window.scrollX,
      scrollY: window.scrollY,
      onclone: function (clonedDoc) {
        clonedDoc.documentElement.style.backgroundColor = exportBg;
        clonedDoc.body.style.backgroundColor = exportBg;
        clonedDoc.body.style.zoom = "1";
        clonedDoc.body.style.transform = "scale(0.75)";
        clonedDoc.body.style.transformOrigin = "top left";
        clonedDoc.body.style.width = "133.3333vw";
        clonedDoc.body.style.height = "133.3333vh";
        clonedDoc.querySelectorAll("img").forEach(function (img) {
          img.style.visibility = "hidden";
        });
        var clonedWrapper = clonedDoc.getElementById("toolbars-wrapper");
        if (clonedWrapper) clonedWrapper.style.display = "none";
        var clonedBtnText = clonedDoc.getElementById("toggle-toolbars-text");
        if (clonedBtnText) clonedBtnText.textContent = "Show Buttons";
        var clonedBtnIcon = clonedDoc.getElementById("toggle-toolbars-icon");
        if (clonedBtnIcon) clonedBtnIcon.textContent = toolbarBtnIcon ? toolbarBtnIcon.textContent : "👁️";
        var clonedContainer = clonedDoc.querySelector(".planner-container");
        if (clonedContainer) {
          clonedContainer.scrollLeft = container.scrollLeft;
          clonedContainer.scrollTop = container.scrollTop;
        }
      }
    }).then(function (canvas) {
      canvas.toBlob(function (blob) {
        if (!blob) { alert("JPEG export failed."); return; }
        var a = document.createElement("a");
        var d = new Date();
        var dateStr = d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0");
        a.download = "planner_" + dateStr + ".jpg";
        a.href = URL.createObjectURL(blob);
        a.click();
        URL.revokeObjectURL(a.href);
      }, "image/jpeg", 0.92);
    }).catch(function (err) {
      console.error("Export error:", err);
      alert("JPEG export failed: " + err.message);
    });
  }).catch(function (err) {
      alert("Could not load the export library. Check your internet connection.");
      console.error(err);
  });

  function getExportBackground(el) {
    var nodes = [el, document.body, document.documentElement];
    for (var i = 0; i < nodes.length; i++) {
      var bg = window.getComputedStyle(nodes[i]).backgroundColor;
      if (bg && bg !== "transparent" && bg !== "rgba(0, 0, 0, 0)") return bg;
    }
    return "#ffffff";
  }
}
