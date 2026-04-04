// ═══════════════════════════════════════════════════════════════
//  CSV EXPORT / IMPORT / MERGE
// ═══════════════════════════════════════════════════════════════

/** Build and return the full CSV string for the current planner state. */
function generateCSVString() {
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
  customTemplates.forEach((t) => {
    const safeComposition = t.composition
      ? '"' + JSON.stringify(t.composition).replace(/"/g, '""') + '"'
      : '""';
    csv.push(
      'Template,"' +
      t.name.replace(/"/g, '""') +
      '",' +
      t.duration +
      ',"' +
      t.color +
      '",' +
      safeComposition,
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

  csv.push(
    "Id,Type,Name,StartWeek,Duration,ParentId,IsExpanded,Comment,Color,IsLocked,MilestoneRow,BlockCommentData,AssigneesData,Completion",
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
      (item.completion || 0),
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
        highlightedWeek = values[6] ? parseInt(values[6]) : null;
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
      if (values[0] === "Template") {
        const tName = values[1];
        const tDur = parseFloat(values[2]);
        const tColor = values[3] || "#4f46e5";
        let tComp = null;
        if (values[4] && values[4] !== "") {
          try {
            tComp = JSON.parse(values[4]);
          } catch (e) { }
        }
        // Only add if it doesn't already exist locally
        if (
          !customTemplates.some(
            (t) => t.name.toLowerCase() === tName.toLowerCase(),
          )
        ) {
          const newT = {
            id: "t_" + Date.now() + Math.random(),
            name: tName,
            duration: tDur,
            color: tColor,
          };
          if (tComp) newT.composition = tComp;
          customTemplates.push(newT);
        }
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
      nextLinkId =
        newLinks.length > 0
          ? Math.max(...newLinks.map((l) => l.id)) + 1
          : 1;
      nextId = tempId;
      nextAlarmId = tempAlarmId;
      nextPersonId = tempPersonId;
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
      } else if (v[0] === "Template") {
        const tName = v[1];
        const tDur = parseFloat(v[2]);
        const tColor = v[3] || "#4f46e5";
        if (
          !customTemplates.some(
            (t) => t.name.toLowerCase() === tName.toLowerCase(),
          )
        ) {
          customTemplates.push({
            id: "t_" + Date.now() + Math.random(),
            name: tName,
            duration: tDur,
            color: tColor,
          });
        }
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

    render();
    if (typeof markChanged === 'function') markChanged();
    let msg = "Merged " + newItems.length + " item rows";
    if (incomingAlarms.length)
      msg += " and " + incomingAlarms.length + " alarms";
    msg += ' from "' + file.name + '" into the planner.';
    alert(msg);
  };
  reader.readAsText(file);
  e.target.value = "";
}

function exportJPEG() {
  // Use browser print — works fully offline, no CDN needed.
  // In the print dialog choose "Save as PDF" or your printer.
  // Print CSS (in style.css @media print) hides all UI chrome automatically.
  window.print();
}
