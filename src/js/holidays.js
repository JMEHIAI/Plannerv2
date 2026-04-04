// ═══════════════════════════════════════════════════════════════
//  HOLIDAYS MANAGEMENT
// ═══════════════════════════════════════════════════════════════

function openHolidaysModal() {
  document.getElementById("holidaysModal").style.display = "flex";
  renderHolidaysList();
}

function closeHolidaysModal() {
  document.getElementById("holidaysModal").style.display = "none";
}

function addHoliday() {
  const input = document.getElementById("newHolidayDate");
  if (input.value && !holidays.some(h => h.date === input.value)) {
    holidays.push({ date: input.value, people: [] });
    holidays.sort((a, b) => a.date.localeCompare(b.date));
    renderHolidaysList();
    render();
  }
  input.value = "";
}

function addHolidayRange() {
  const fromEl = document.getElementById("holidayRangeFrom");
  const toEl = document.getElementById("holidayRangeTo");
  if (!fromEl.value || !toEl.value) {
    alert("Please select both a start and end date.");
    return;
  }
  const from = new Date(fromEl.value);
  const to = new Date(toEl.value);
  if (from > to) {
    alert("Start date must be before end date.");
    return;
  }
  const cur = new Date(from);
  while (cur <= to) {
    const ds = formatDateStr(cur);
    if (!holidays.some(h => h.date === ds)) {
      holidays.push({ date: ds, people: [] });
    }
    cur.setDate(cur.getDate() + 1);
  }
  holidays.sort((a, b) => a.date.localeCompare(b.date));
  fromEl.value = "";
  toEl.value = "";
  renderHolidaysList();
  render();
}

function removeHoliday(index) {
  holidays.splice(index, 1);
  renderHolidaysList();
  render();
}

function addHolidayPerson(hIdx, personId) {
  if (!personId) return;
  if (!holidays[hIdx].people.includes(personId)) {
    holidays[hIdx].people.push(personId);
  }
  renderHolidaysList();
  render();
}

function removeHolidayPerson(hIdx, personId) {
  holidays[hIdx].people = holidays[hIdx].people.filter(p => p !== personId);
  renderHolidaysList();
  render();
}

function renderHolidaysList() {
  const listEl = document.getElementById("holidaysList");
  const countEl = document.getElementById("holidays-count");
  if (countEl) countEl.textContent = "(" + holidays.length + ")";
  if (holidays.length === 0) {
    listEl.innerHTML = '<div style="color:#64748b; font-size:14px; font-style:italic;">No holidays added yet.</div>';
    return;
  }
  listEl.innerHTML = holidays.map((h, i) => {
    const peopleChips = (h.people || []).map(pid => {
      const p = people.find(x => x.id === pid);
      const name = p ? p.name : pid;
      return `<span style="display:inline-flex;align-items:center;gap:2px;background:#e0e7ff;color:#3730a3;border-radius:10px;padding:1px 7px 1px 8px;font-size:11px;font-weight:600;">
        ${name}<button onclick="removeHolidayPerson(${i},'${pid}')" style="background:none;border:none;cursor:pointer;color:#6366f1;font-size:13px;line-height:1;padding:0 0 0 2px;">&times;</button>
      </span>`;
    }).join("");
    const availablePeople = people.filter(p => !(h.people || []).includes(p.id));
    const addPersonSelect = availablePeople.length > 0
      ? `<select onchange="addHolidayPerson(${i},this.value);this.value=''" style="font-size:11px;border:1px solid #cbd5e1;border-radius:4px;padding:1px 4px;cursor:pointer;color:#64748b;background:white;">
          <option value="">+ Assign person…</option>
          ${availablePeople.map(p => `<option value="${p.id}">${p.name}</option>`).join("")}
        </select>` : "";
    const scopeLabel = (h.people || []).length === 0
      ? `<span style="font-size:11px;color:#64748b;font-style:italic;">👥 All staff</span>`
      : "";
    return `<div style="padding:6px 10px;background:#f8fafc;border-radius:4px;margin-bottom:4px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
        <span style="font-size:13px;font-weight:600;">${h.date}</span>
        <button onclick="removeHoliday(${i})" style="color:#ef4444;background:none;border:none;cursor:pointer;font-size:16px;line-height:1;" title="Remove">&times;</button>
      </div>
      <div style="display:flex;align-items:center;gap:4px;flex-wrap:wrap;">
        ${scopeLabel}${peopleChips}${addPersonSelect}
      </div>
    </div>`;
  }).join("");
}
