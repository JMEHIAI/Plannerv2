// ═══════════════════════════════════════════════════════════════
//  ALARM SYSTEM
// ═══════════════════════════════════════════════════════════════

function openAlarmForDate(dateStr, itemId) {
  if (alarms && getAlarmAbsWeek) {
    currentAlarmEditId = null;
    currentAlarmItemId = itemId || null;
    setAlarmDateValue(dateStr);
    updateAlarmDateWeek();
    syncAlarmEditorState();

    let p = document.getElementById("alarm-panel");
    if (!p.classList.contains("open")) p.classList.add("open");
    renderAlarmPanel();
  }
}

function openAlarmForWeek(w, itemId, dayIndex) {
  if (alarms && getAlarmAbsWeek) {
    currentAlarmEditId = null;
    currentAlarmItemId = itemId || null;
    const weekInfo = getYearWeekInfo(w);
    let yrIndex = weekInfo.yearIndex;
    let yrNum = parseInt(years[yrIndex]);
    if (isNaN(yrNum)) yrNum = new Date().getFullYear();
    let wkNum = weekInfo.relWeek;

    let jan4 = new Date(yrNum, 0, 4);
    let startOfYear = new Date(jan4.getTime());
    startOfYear.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1);

    let target = new Date(startOfYear.getTime());
    target.setDate(startOfYear.getDate() + (wkNum - 1) * 7 + (dayIndex || 0));

    let dStr =
      target.getFullYear() +
      "-" +
      pad2(target.getMonth() + 1) +
      "-" +
      pad2(target.getDate());

    setAlarmDateValue(dStr);
    updateAlarmDateWeek();
    syncAlarmEditorState();

    let p = document.getElementById("alarm-panel");
    if (!p.classList.contains("open")) p.classList.add("open");
    renderAlarmPanel();
  }
}

function openAlarmEditor(alarmId) {
  const alarm = alarms.find((a) => a.id === alarmId);
  if (!alarm) return;
  currentAlarmEditId = alarm.id;
  currentAlarmItemId = alarm.itemId || null;
  document.getElementById("alarm-title").value = alarm.title || "";
  setAlarmDateValue(alarm.date || "");
  document.getElementById("alarm-time").value = alarm.time || "09:00";
  document.getElementById("alarm-duration").value = alarm.duration || 60;
  updateAlarmDateWeek();
  syncAlarmEditorState();
  const p = document.getElementById("alarm-panel");
  if (p && !p.classList.contains("open")) p.classList.add("open");
  renderAlarmPanel();
}

function parseAlarmDisplayDate(dateText) {
  if (!dateText) return "";
  const trimmed = dateText.trim();
  let m = trimmed.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return m[3] + "-" + m[2] + "-" + m[1];
  m = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return trimmed;
  return "";
}

function setAlarmDateValue(dateStr) {
  const nativeInput = document.getElementById("alarm-date");
  const textInput = document.getElementById("alarm-date-text");
  if (nativeInput) nativeInput.value = dateStr || "";
  if (textInput) textInput.value = dateStr ? formatAlarmDisplayDate(dateStr) : "";
}

function syncAlarmDateFromNative() {
  const nativeInput = document.getElementById("alarm-date");
  setAlarmDateValue(nativeInput ? nativeInput.value : "");
  updateAlarmDateWeek();
}

function commitAlarmDateText() {
  const textInput = document.getElementById("alarm-date-text");
  if (!textInput) return;
  const parsed = parseAlarmDisplayDate(textInput.value);
  if (parsed) {
    setAlarmDateValue(parsed);
    updateAlarmDateWeek();
  } else if (textInput.value.trim() === "") {
    setAlarmDateValue("");
    updateAlarmDateWeek();
  } else {
    const nativeInput = document.getElementById("alarm-date");
    textInput.value = nativeInput && nativeInput.value ? formatAlarmDisplayDate(nativeInput.value) : "";
  }
}

function openAlarmDatePicker() {
  const nativeInput = document.getElementById("alarm-date");
  if (!nativeInput) return;
  if (typeof nativeInput.showPicker === "function") {
    nativeInput.showPicker();
  } else {
    nativeInput.focus();
    nativeInput.click();
  }
}

function updateAlarmDateWeek() {
  const dt = document.getElementById("alarm-date").value;
  const badge = document.getElementById("alarm-date-week");
  if (badge && dt) {
    const weekLabel = getAlarmDisplayWeekLabel({ date: dt });
    badge.innerText = weekLabel ? "(" + weekLabel + ")" : "";
  } else if (badge) {
    badge.innerText = "";
  }
  const textInput = document.getElementById("alarm-date-text");
  if (textInput && dt) textInput.value = formatAlarmDisplayDate(dt);
}

function getAlarmDisplayWeekLabel(alarm) {
  if (!alarm || !alarm.date) return "";
  const parts = alarm.date.split("-");
  if (parts.length !== 3) return "";
  const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  const isoInfo = getIsoYearWeekFromDate(d);
  return isoInfo && isoInfo.isoWeek ? "W" + isoInfo.isoWeek : "";
}

function formatAlarmDisplayDate(dateStr) {
  if (!dateStr) return "";
  const parts = dateStr.split("-");
  if (parts.length !== 3) return dateStr;
  return parts[2] + "/" + parts[1] + "/" + parts[0];
}

function getAlarmAbsWeek(alarm) {
  if (!alarm || !alarm.date) return -1;
  const exactWeek = getAlarmExactWeek(alarm);
  return exactWeek > 0 ? Math.floor(exactWeek) : -1;
}

function getAlarmExactWeek(alarm) {
  if (!alarm || !alarm.date) return -1;
  // Handle the ISO format
  const parts = alarm.date.split("-");
  if (parts.length !== 3) return -1;
  // Parse dates explicitly as local midnight to avoid timezone shifts
  const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  const isoInfo = getIsoYearWeekFromDate(d);
  const yi = years.findIndex((y) => String(parseInt(y)) === String(isoInfo.isoYear));
  if (yi < 0) return -1;
  const absWeek = getAbsWeekFromYearWeek(yi, isoInfo.isoWeek);
  const jsDay = d.getDay();
  const weekdayIndex = jsDay === 0 ? 6 : jsDay - 1; // Monday = 0
  const clampedWeekday = Math.max(0, Math.min(4, weekdayIndex));
  return parseFloat((absWeek + clampedWeekday * 0.2).toFixed(1));
}

function toggleAlarmPanel() {
  const p = document.getElementById("alarm-panel");
  if (p) {
    p.classList.toggle("open");
    if (p.classList.contains("open")) syncAlarmEditorState();
    renderAlarmPanel();
  }
}

function resetAlarmEditor(keepDate) {
  currentAlarmEditId = null;
  currentAlarmItemId = null;
  document.getElementById("alarm-title").value = "";
  if (!keepDate) setAlarmDateValue("");
  document.getElementById("alarm-time").value = "09:00";
  document.getElementById("alarm-duration").value = 60;
  updateAlarmDateWeek();
  syncAlarmEditorState();
}

function syncAlarmEditorState() {
  const saveBtn = document.getElementById("alarm-save-btn");
  const deleteBtn = document.getElementById("alarm-delete-btn");
  const cancelBtn = document.getElementById("alarm-cancel-btn");
  if (saveBtn) saveBtn.textContent = currentAlarmEditId ? "Save Alarm" : "+ Add Alarm";
  if (deleteBtn) deleteBtn.style.display = currentAlarmEditId ? "inline-flex" : "none";
  if (cancelBtn) cancelBtn.style.display = currentAlarmEditId ? "inline-flex" : "none";
}

function saveAlarm() {
  commitAlarmDateText();
  const title = (
    document.getElementById("alarm-title").value || ""
  ).trim();
  const date = document.getElementById("alarm-date").value;
  const time = document.getElementById("alarm-time").value || "09:00";
  const dur =
    parseInt(document.getElementById("alarm-duration").value) || 60;
  if (!title) {
    alert("Please enter a title.");
    return;
  }
  if (!date) {
    alert("Please select a date.");
    return;
  }
  if (currentAlarmEditId) {
    const alarm = alarms.find((a) => a.id === currentAlarmEditId);
    if (alarm) {
      alarm.title = title;
      alarm.date = date;
      alarm.time = time;
      alarm.duration = dur;
      alarm.itemId = currentAlarmItemId;
    }
  } else {
    alarms.push({
      id: nextAlarmId++,
      title,
      date,
      time,
      duration: dur,
      itemId: currentAlarmItemId,
    });
  }
  resetAlarmEditor(true);
  renderAlarmPanel();
  render();
}

function deleteAlarm(id) {
  alarms = alarms.filter((a) => a.id !== id);
  if (currentAlarmEditId === id) resetAlarmEditor(true);
  renderAlarmPanel();
  render();
}

function deleteCurrentAlarm() {
  if (currentAlarmEditId !== null) deleteAlarm(currentAlarmEditId);
}

function renderAlarmPanel() {
  const list = document.getElementById("alarm-list");
  if (!list) return;
  syncAlarmEditorState();
  if (alarms.length === 0) {
    list.innerHTML =
      '<p style="color:#94a3b8;font-size:13px;text-align:center;padding:12px 0;">No alarms yet.</p>';
    return;
  }
  list.innerHTML = alarms
    .map((a) => {
      const displayWeekLabel = getAlarmDisplayWeekLabel(a);
      const wkLabel =
        displayWeekLabel
          ? '<span class="alarm-week-badge">' + displayWeekLabel + "</span>"
          : "";

      let itemName = "";
      if (a.itemId) {
        const mappedItem = items.find((i) => i.id === a.itemId);
        if (mappedItem)
          itemName =
            ' <span style="color:#10b981; font-size:10px;">(' +
            mappedItem.name.substring(0, 20) +
            ")</span>";
      }
      return (
        '<div class="alarm-item' + (currentAlarmEditId === a.id ? " active" : "") + '" onclick="openAlarmEditor(' + a.id + ')">' +
        '<div style="min-width:0;">' +
        '<div style="font-weight:600;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' +
        a.title +
        itemName +
        "</div>" +
        '<div style="font-size:11px;color:#64748b;margin-top:2px;">' +
        formatAlarmDisplayDate(a.date) +
        "&nbsp;&nbsp;" +
        a.time +
        "&nbsp;&nbsp;" +
        a.duration +
        "&nbsp;min</div>" +
        wkLabel +
        "</div>" +
        '<button class="alarm-del-btn" onclick="event.stopPropagation(); deleteAlarm(' +
        a.id +
        ')" title="Delete alarm">×</button>' +
        "</div>"
      );
    })
    .join("");
}

function exportICS() {
  if (alarms.length === 0) {
    alert("No alarms to export.");
    return;
  }
  const crlf = "\r\n";
  const nowStamp =
    new Date()
      .toISOString()
      .replace(/[-:.Z]/g, "")
      .slice(0, 15) + "Z";
  let ics = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//H-Customer Performance Planner//EN",
    "CALSCALE:GREGORIAN",
    "METHOD:PUBLISH",
  ];
  alarms.forEach((a) => {
    const dtStart =
      a.date.replace(/-/g, "") + "T" + a.time.replace(":", "") + "00";
    const endMs =
      new Date(a.date + "T" + a.time).getTime() + a.duration * 60000;
    const endD = new Date(endMs);
    const dtEnd =
      endD.getFullYear() +
      "" +
      pad2(endD.getMonth() + 1) +
      "" +
      pad2(endD.getDate()) +
      "T" +
      pad2(endD.getHours()) +
      "" +
      pad2(endD.getMinutes()) +
      "00";
    ics.push(
      "BEGIN:VEVENT",
      "UID:alarm-" + a.id + "-" + nowStamp + "@hcustomer-planner",
      "DTSTAMP:" + nowStamp,
      "DTSTART:" + dtStart,
      "DTEND:" + dtEnd,
      "SUMMARY:" + a.title.replace(/,/g, "\\,").replace(/;/g, "\\;"),
      "DESCRIPTION:Alarm from H-Customer Performance Planner",
      "BEGIN:VALARM",
      "TRIGGER:-PT15M",
      "ACTION:DISPLAY",
      "DESCRIPTION:Reminder: " + a.title,
      "END:VALARM",
      "END:VEVENT",
    );
  });
  ics.push("END:VCALENDAR");
  const blob = new Blob([ics.join(crlf)], {
    type: "text/calendar;charset=utf-8;",
  });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "planner_alarms.ics";
  link.click();
}
