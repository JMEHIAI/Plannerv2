// ═══════════════════════════════════════════════════════════════
//  ALARM SYSTEM
// ═══════════════════════════════════════════════════════════════

function openAlarmForDate(dateStr, itemId) {
  if (alarms && getAlarmAbsWeek) {
    currentAlarmItemId = itemId || null;
    document.getElementById("alarm-date").value = dateStr;
    updateAlarmDateWeek();

    let p = document.getElementById("alarm-panel");
    if (!p.classList.contains("open")) p.classList.add("open");
    renderAlarmPanel();
  }
}

function openAlarmForWeek(w, itemId, dayIndex) {
  if (alarms && getAlarmAbsWeek) {
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

    document.getElementById("alarm-date").value = dStr;
    updateAlarmDateWeek();

    let p = document.getElementById("alarm-panel");
    if (!p.classList.contains("open")) p.classList.add("open");
    renderAlarmPanel();
  }
}

function updateAlarmDateWeek() {
  const dt = document.getElementById("alarm-date").value;
  const badge = document.getElementById("alarm-date-week");
  if (badge && dt) {
    let absWeek = getAlarmAbsWeek({ date: dt });
    badge.innerText = absWeek > 0 ? "(W" + absWeek + ")" : "";
  } else if (badge) {
    badge.innerText = "";
  }
}

function getAlarmAbsWeek(alarm) {
  if (!alarm || !alarm.date) return -1;
  // Handle the ISO format
  const parts = alarm.date.split("-");
  if (parts.length !== 3) return -1;
  // Parse dates explicitly as local midnight to avoid timezone shifts
  const d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  const isoInfo = getIsoYearWeekFromDate(d);
  const yi = years.findIndex((y) => String(parseInt(y)) === String(isoInfo.isoYear));
  if (yi < 0) return -1;
  return getAbsWeekFromYearWeek(yi, isoInfo.isoWeek);
}

function toggleAlarmPanel() {
  const p = document.getElementById("alarm-panel");
  if (p) {
    p.classList.toggle("open");
    renderAlarmPanel();
  }
}

function addAlarm() {
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
  alarms.push({
    id: nextAlarmId++,
    title,
    date,
    time,
    duration: dur,
    itemId: currentAlarmItemId,
  });
  currentAlarmItemId = null;
  document.getElementById("alarm-title").value = "";
  renderAlarmPanel();
  render();
}

function deleteAlarm(id) {
  alarms = alarms.filter((a) => a.id !== id);
  renderAlarmPanel();
  render();
}

function renderAlarmPanel() {
  const list = document.getElementById("alarm-list");
  if (!list) return;
  if (alarms.length === 0) {
    list.innerHTML =
      '<p style="color:#94a3b8;font-size:13px;text-align:center;padding:12px 0;">No alarms yet.</p>';
    return;
  }
  list.innerHTML = alarms
    .map((a) => {
      const absWeek = getAlarmAbsWeek(a);
      const wkLabel =
        absWeek > 0
          ? '<span class="alarm-week-badge">W' + absWeek + "</span>"
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
        '<div class="alarm-item">' +
        '<div style="min-width:0;">' +
        '<div style="font-weight:600;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">' +
        a.title +
        itemName +
        "</div>" +
        '<div style="font-size:11px;color:#64748b;margin-top:2px;">' +
        a.date +
        "&nbsp;&nbsp;" +
        a.time +
        "&nbsp;&nbsp;" +
        a.duration +
        "&nbsp;min</div>" +
        wkLabel +
        "</div>" +
        '<button class="alarm-del-btn" onclick="deleteAlarm(' +
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
