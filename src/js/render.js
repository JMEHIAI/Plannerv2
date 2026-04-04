// ═══════════════════════════════════════════════════════════════
//  RENDER ENGINE + DISPLAY ITEMS + EFFECTIVE DURATION
// ═══════════════════════════════════════════════════════════════

const gridEl = document.getElementById("grid");

// Month names constant (never changes, no need to recreate every render)
const _MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

// Cache for dynamicMonthSpans — recomputed only when `years` array changes
let _monthSpansCache = null;
let _monthSpansCacheKey = "";

// Cache for day-column metadata (days mode) — recomputed only when years change
let _dayColMetaCache = null;
let _dayColMetaCacheKey = "";

// Cache for sticky header heights — recomputed only when zoomMode changes
let _cachedHeaderHeights = null;
let _cachedHeaderZoom = null;

// Shared empty Map sentinel — avoids allocating a new Map() on every cache miss
const _emptyMap = new Map();

// Most-recent day-column metadata (exposed for click handlers)
let _currentDayColMeta = null;

function getDisplayItems() {
  const rawList = [];

  // Build children map for O(1) child lookups inside this function
  const _childMap = new Map();
  items.forEach((it) => {
    const pid = it.parentId;
    if (!_childMap.has(pid)) _childMap.set(pid, []);
    _childMap.get(pid).push(it);
  });
  const _getChildren = (pid) => _childMap.get(pid) || [];

  function isFilteredOut(item) {
    if (filters.name.length > 0) {
      if (
        filters.name.includes((item.name || "").trim()) ||
        filters.name.includes(item.name)
      )
        return true;
    }
    if (filters.type.length > 0) {
      if (filters.type.includes(item.type)) return true;
    }
    return false;
  }

  function addChildren(parentId, level) {
    const children = _getChildren(parentId).slice(); // copy before sort to avoid mutating the map
    children.sort((a, b) => {
      if (a.type === "milestone" && b.type !== "milestone") return -1;
      if (b.type === "milestone" && a.type !== "milestone") return 1;
      return 0;
    });
    children.forEach((child) => {
      if (isFilteredOut(child)) return;
      child.level = level;
      rawList.push(child);
      if (child.isExpanded !== false) {
        addChildren(child.id, level + 1);
      }
    });
  }

  const topLevelItems = items.filter((i) => !i.parentId);
  topLevelItems.forEach((p) => {
    if (isFilteredOut(p)) return;
    p.level = 0;
    rawList.push(p);
    if (p.isExpanded !== false) {
      addChildren(p.id, 1);
    }
  });

  return rawList;
}

/**
 * Given an activity's startWeek (1-based, can be fractional for days mode) and
 * its stored duration in weeks, return the effective VISUAL duration in weeks
 * by counting forward calendar working-days and skipping any holidays.
 * Safe: calendarDays always increases, so no infinite loop is possible.
 */
function computeEffectiveDuration(startWeek, duration, itemAssignees) {
  // Only count holidays that apply to this item's assignees
  const applicableHolidays = holidays.filter(h => holidayApplies(h, itemAssignees || []));
  if (!applicableHolidays.length || duration <= 0) return duration;
  const workDaysNeeded = Math.round(duration * 5);
  if (workDaysNeeded === 0) return duration;

  // Resolve fractional startWeek to absolute week + day index
  const startW = Math.floor(startWeek);
  const startDayFract = Math.round((startWeek - startW) * 5); // 0=Mon

  const yearIndex = Math.floor((startW - 1) / 52);
  const yearNum = parseInt(years[yearIndex]) || new Date().getFullYear();
  const relWeek = ((startW - 1) % 52) + 1;
  const jan4 = new Date(yearNum, 0, 4);
  const weekStart = new Date(jan4.getTime());
  weekStart.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1); // Mon of ISO wk1

  // Date of the first day of this block
  const blockStart = new Date(weekStart.getTime());
  blockStart.setDate(
    weekStart.getDate() + (relWeek - 1) * 7 + startDayFract,
  );

  let workDone = 0;
  let holidaysHit = 0;
  let calDays = 0;
  const maxCalDays = workDaysNeeded + applicableHolidays.length + 30; // safety cap
  const applicableHolidayDates = new Set(applicableHolidays.map(h => h.date));

  while (workDone < workDaysNeeded && calDays < maxCalDays) {
    const cur = new Date(blockStart.getTime());
    cur.setDate(blockStart.getDate() + calDays);
    const dow = cur.getDay(); // 0=Sun, 6=Sat
    if (dow !== 0 && dow !== 6) {
      if (!applicableHolidayDates.has(formatDateStr(cur))) {
        workDone++;
      } else {
        holidaysHit++;
      }
    }
    calDays++;
  }

  // Each holiday in the span adds exactly 1 working-day slot (1/5 week)
  return (workDaysNeeded + holidaysHit) / 5;
}

// ── Year layout helper (used by render and click handlers) ──────
function _computeYearLayout() {
  const cw = zoomMode === "days" ? 20 : zoomMode === "months" ? 80 : 28;
  const cols = zoomMode === "days" ? 260 : zoomMode === "months" ? 12 : 52;
  const collW = 2;
  const widths = years.map((_, yi) => hiddenYears.has(yi) ? cols * collW : cols * cw);
  const offsets = [];
  let acc = 0;
  widths.forEach(w => { offsets.push(acc); acc += w; });
  return { cw, cols, collW, widths, offsets, total: acc };
}

// ── Global click handlers for timeline cells ────────────────────
function handleTlDblClick(e, itemId) {
  if (e.target.closest('.bar, .milestone, .range-bar, .resize-handle')) return;
  const x = e.offsetX;
  const { cw, cols, collW, widths } = _computeYearLayout();
  let remaining = x, snapW = 1, dayIdx = 0;
  for (let yi = 0; yi < years.length; yi++) {
    const yw = widths[yi];
    if (remaining <= yw || yi === years.length - 1) {
      const effCw = hiddenYears.has(yi) ? collW : cw;
      if (zoomMode === "days") {
        const col = Math.max(0, Math.min(Math.floor(remaining / effCw), 259));
        snapW = yi * 52 + Math.floor(col / 5) + 1;
        dayIdx = col % 5;
      } else if (zoomMode === "weeks") {
        snapW = yi * 52 + Math.max(0, Math.floor(remaining / effCw)) + 1;
      } else {
        snapW = yi * 12 + Math.max(0, Math.floor(remaining / effCw)) + 1;
      }
      break;
    }
    remaining -= yw;
  }
  openAlarmForWeek(snapW, itemId, dayIdx);
}

function handleTimelineHeaderClick(e) {
  const x = e.offsetX;
  const { cw, cols, collW, widths } = _computeYearLayout();
  let remaining = x;
  for (let yi = 0; yi < years.length; yi++) {
    const yw = widths[yi];
    if (remaining <= yw || yi === years.length - 1) {
      const effCw = hiddenYears.has(yi) ? collW : cw;
      if (zoomMode === "days") {
        const col = Math.max(0, Math.min(Math.floor(remaining / effCw), 259));
        const w = yi * 52 + Math.floor(col / 5) + 1;
        const dayFrac = (col % 5) * 0.2;
        toggleHighlight(dayFrac > 0 ? parseFloat((w + dayFrac).toFixed(1)) : "W" + w);
      } else if (zoomMode === "weeks") {
        const w = yi * 52 + Math.max(0, Math.floor(remaining / effCw)) + 1;
        toggleHighlight("W" + w);
      } else {
        const c = yi * 12 + Math.max(0, Math.floor(remaining / effCw)) + 1;
        toggleHighlight(c);
      }
      return;
    }
    remaining -= yw;
  }
}

function handleTimelineHeaderDblClick(e) {
  const x = e.offsetX;
  const { cw, cols, collW, widths } = _computeYearLayout();
  let remaining = x;
  for (let yi = 0; yi < years.length; yi++) {
    const yw = widths[yi];
    if (remaining <= yw || yi === years.length - 1) {
      const effCw = hiddenYears.has(yi) ? collW : cw;
      if (zoomMode === "days" && _currentDayColMeta) {
        const col = Math.max(0, Math.min(Math.floor(remaining / effCw), 259));
        const metaIdx = yi * 260 + col;
        if (_currentDayColMeta[metaIdx]) openAlarmForDate(_currentDayColMeta[metaIdx].fullDateStr);
      } else {
        const w = zoomMode === "months"
          ? yi * 12 + Math.max(0, Math.floor(remaining / effCw)) + 1
          : yi * 52 + Math.max(0, Math.floor(remaining / effCw)) + 1;
        openAlarmForWeek(w);
      }
      return;
    }
    remaining -= yw;
  }
}

// ═══════════════════════════════════════════════════════════════
//  RENDER — 3-column grid (comment | name | single timeline cell)
// ═══════════════════════════════════════════════════════════════
function render() {
  let html = "";

  // Pre-build a parent→children map once so every child-lookup is O(1)
  const childrenMap = new Map();
  items.forEach((it) => {
    const pid = it.parentId;
    if (!childrenMap.has(pid)) childrenMap.set(pid, []);
    childrenMap.get(pid).push(it);
  });
  const getChildren = (parentId) => childrenMap.get(parentId) || [];

  // Sync project/family/milestones-group startWeek+duration from descendants (bottom-up post-order)
  (function syncParentBounds() {
    function syncItem(it) {
      let minS = Infinity, maxE = -Infinity;
      function traverse(pid) {
        (childrenMap.get(pid) || []).forEach(k => {
          if (k.startWeek < minS) minS = k.startWeek;
          const kEnd = k.startWeek + (k.duration || 0);
          if (kEnd > maxE) maxE = kEnd;
          traverse(k.id);
        });
      }
      traverse(it.id);
      if (minS !== Infinity) {
        it.startWeek = minS;
        it.duration = Math.max(0, maxE - minS);
      }
    }
    function syncTree(pid) {
      (childrenMap.get(pid) || []).forEach(child => {
        syncTree(child.id);
        if (child.type === "project" || child.type === "family" || child.type === "milestones-group") {
          syncItem(child);
        }
      });
    }
    items.filter(i => !i.parentId).forEach(root => {
      syncTree(root.id);
      if (root.type === "project" || root.type === "family" || root.type === "milestones-group") {
        syncItem(root);
      }
    });
  })();

  const getCommentIndicator = (it) =>
    it.blockComment && it.blockComment.text
      ? '<div class="comment-indicator"></div>'
      : "";

  const fmtMsDate = (sw) => {
    const yi = Math.max(0, Math.min(years.length - 1, Math.floor((sw - 1) / 52)));
    const yy = String(parseInt(years[yi]) || new Date().getFullYear()).slice(-2);
    const wiy = ((sw - 1) % 52) + 1;
    const ww = String(Math.floor(wiy)).padStart(2, "0");
    const frac = wiy - Math.floor(wiy);
    return yy + ww + (frac > 0.05 ? "." + Math.round(frac * 10) : "");
  };

  // Clear all dynamic SVG tails from previous render
  const svg = document.getElementById("comment-tail-svg-container");
  if (svg) svg.innerHTML = "";

  // ── CSS variables ──────────────────────────────────────────────
  const actualCommentWidth = showComments ? commentWidth + "px" : "0px";
  const actualNameWidth = nameWidthBase + (showSettings ? 280 : 0) + "px";
  gridEl.style.setProperty("--comment-width", actualCommentWidth);
  gridEl.style.setProperty("--name-width", actualNameWidth);

  if (!showComments) gridEl.classList.add("hide-comments");
  else gridEl.classList.remove("hide-comments");

  // ── Year layout ────────────────────────────────────────────────
  const totalYears = years.length;
  const colsPerYear = zoomMode === "days" ? 260 : zoomMode === "months" ? 12 : 52;
  const cellW = zoomMode === "days" ? 20 : zoomMode === "months" ? 80 : 28;
  const collapsedCellW = 2;

  const yearWidths = years.map((_, yi) =>
    hiddenYears.has(yi) ? colsPerYear * collapsedCellW : colsPerYear * cellW
  );
  const yearOffsets = [];
  let _tlAcc = 0;
  yearWidths.forEach(w => { yearOffsets.push(_tlAcc); _tlAcc += w; });
  const totalTimelineW = _tlAcc;

  // 3-column grid: comment | name | single wide timeline column
  gridEl.style.gridTemplateColumns =
    "var(--comment-width) var(--name-width) " + totalTimelineW + "px";

  // ── Month spans cache ──────────────────────────────────────────
  const _cacheKey = years.join(",");
  if (!_monthSpansCache || _monthSpansCacheKey !== _cacheKey) {
    const computed = [];
    for (let y = 0; y < years.length; y++) {
      const yearNum = parseInt(years[y]) || new Date().getFullYear();
      const jan4 = new Date(yearNum, 0, 4);
      const startOfYear = new Date(jan4.getTime());
      startOfYear.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1);
      const yearSpans = Array(12).fill(0).map((_, i) => ({ name: _MONTH_NAMES[i], weeks: 0 }));
      for (let w = 1; w <= 52; w++) {
        const thursday = new Date(startOfYear.getTime());
        thursday.setDate(startOfYear.getDate() + (w - 1) * 7 + 3);
        yearSpans[thursday.getMonth()].weeks++;
      }
      computed.push(yearSpans);
    }
    _monthSpansCache = computed;
    _monthSpansCacheKey = _cacheKey;
  }
  const dynamicMonthSpans = _monthSpansCache;

  // Month-end week Sets for O(1) lookups
  const _monthEndWeekSets = dynamicMonthSpans.map(yearSpans => {
    const s = new Set();
    let acc = 0;
    yearSpans.forEach(m => { acc += m.weeks; s.add(acc); });
    return s;
  });

  // ── Alarm week map ─────────────────────────────────────────────
  const _alarmWeekMap = new Map();
  alarms.forEach(a => {
    const w = getAlarmAbsWeek(a);
    if (!_alarmWeekMap.has(w)) _alarmWeekMap.set(w, []);
    _alarmWeekMap.get(w).push(a);
  });

  // ── Holiday pre-computation ────────────────────────────────────
  const _holidayAbsWeeks = new Set();
  const _holidayMonthKeys = new Set();
  if (holidays.length > 0) {
    holidays.forEach(h => {
      const hd = new Date(h.date + "T12:00:00");
      const hYear = hd.getFullYear();
      const hMon = hd.getMonth() + 1;
      _holidayMonthKeys.add(hYear + "-" + String(hMon).padStart(2, "0"));
      for (let yi = 0; yi < years.length; yi++) {
        const yn = parseInt(years[yi]);
        if (yn === hYear) {
          const jan4 = new Date(yn, 0, 4);
          const soy = new Date(jan4.getTime());
          soy.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1);
          const diff = Math.floor((hd - soy) / 86400000);
          if (diff >= 0) {
            const relW = Math.floor(diff / 7) + 1;
            if (relW >= 1 && relW <= 52) _holidayAbsWeeks.add(yi * 52 + relW);
          }
          break;
        }
      }
    });
  }

  const _holidayDayCols = new Set();
  const _generalHolidayDayCols = new Set();
  const _dayColToHoliday = new Map();
  if (holidays.length > 0 && zoomMode === "days") {
    holidays.forEach(h => {
      const hd = new Date(h.date + "T12:00:00");
      const hYear = hd.getFullYear();
      for (let yi = 0; yi < years.length; yi++) {
        const yn = parseInt(years[yi]);
        if (yn === hYear) {
          const jan4 = new Date(yn, 0, 4);
          const soy = new Date(jan4.getTime());
          soy.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1);
          const diff = Math.floor((hd - soy) / 86400000);
          if (diff >= 0 && diff < 52 * 7) {
            const dow = diff % 7;
            if (dow < 5) {
              const col = (yi * 52 + Math.floor(diff / 7)) * 5 + dow + 1;
              _holidayDayCols.add(col);
              _dayColToHoliday.set(col, h);
              if (!h.people || h.people.length === 0) _generalHolidayDayCols.add(col);
            }
          }
          break;
        }
      }
    });
  }

  // ── Day-column metadata cache (days mode) ──────────────────────
  let _dayColMeta = null;
  if (zoomMode === "days") {
    if (_dayColMetaCache && _dayColMetaCacheKey === _cacheKey) {
      _dayColMeta = _dayColMetaCache;
    } else {
      _dayColMeta = new Array(totalYears * 260);
      const _yearStarts = years.map((yr) => {
        const yearNum = parseInt(yr) || new Date().getFullYear();
        const jan4 = new Date(yearNum, 0, 4);
        const soy = new Date(jan4.getTime());
        soy.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1);
        return soy;
      });
      const _DAY_LETTERS = ["M", "T", "W", "T", "F"];
      for (let c = 1; c <= totalYears * 260; c++) {
        const w = Math.floor((c - 1) / 5) + 1;
        const dayIndex = (c - 1) % 5;
        const yearIndex = Math.floor((w - 1) / 52);
        const relativeWeek = ((w - 1) % 52) + 1;
        const soy = _yearStarts[yearIndex];
        const d = new Date(soy.getTime());
        d.setDate(soy.getDate() + (relativeWeek - 1) * 7 + dayIndex);
        const dd = pad2(d.getDate()), mm = pad2(d.getMonth() + 1);
        const isWeekEnd = dayIndex === 4;
        const isMonthEnd = isWeekEnd && _monthEndWeekSets[yearIndex].has(relativeWeek);
        const exactWeekValue = w + dayIndex * 0.2;
        _dayColMeta[c - 1] = {
          letter: _DAY_LETTERS[dayIndex],
          dateStr: dd,
          fullDateStr: d.getFullYear() + "-" + mm + "-" + dd,
          exactWeekValue,
          isWeekEnd,
          isMonthEnd,
        };
      }
      _dayColMetaCache = _dayColMeta;
      _dayColMetaCacheKey = _cacheKey;
    }
    _currentDayColMeta = _dayColMeta;
  }

  // ── Per-item alarm map ─────────────────────────────────────────
  const _itemAlarmMap = new Map();
  alarms.forEach(a => {
    if (!_itemAlarmMap.has(a.itemId)) _itemAlarmMap.set(a.itemId, new Map());
    const wMap = _itemAlarmMap.get(a.itemId);
    const w = getAlarmAbsWeek(a);
    if (!wMap.has(w)) wMap.set(w, []);
    wMap.get(w).push(a);
  });

  const displayItems = getDisplayItems();
  const totalRows = displayItems.length;

  // ── Months-mode week→month mapper ─────────────────────────────
  function mapWeekToMonth(wk) {
    let wCount = 1;
    for (let yy = 0; yy < years.length; yy++) {
      for (let mm = 0; mm < 12; mm++) {
        const mSpan = dynamicMonthSpans[yy][mm].weeks;
        if (wk >= wCount && wk < wCount + mSpan) {
          return yy * 12 + mm + 1 + (wk - wCount) / mSpan;
        }
        wCount += mSpan;
      }
    }
    return totalYears * 12 + 0.99;
  }

  // ── Timeline X-coordinate helpers ─────────────────────────────
  // Convert absolute week (1-based, can be fractional) to px offset from timeline start
  function _wkX(sw) {
    const baseW = Math.max(1, Math.floor(sw));
    const yi = Math.max(0, Math.min(Math.floor((baseW - 1) / 52), years.length - 1));
    const relW = (baseW - 1) % 52; // 0-based week within year
    const eff = hiddenYears.has(yi) ? collapsedCellW : cellW;
    if (zoomMode === "days") {
      const dayI = Math.round((sw - Math.floor(sw)) * 5);
      return yearOffsets[yi] + (relW * 5 + dayI) * eff;
    }
    // weeks mode: fractional offset within column
    const fracOffset = (sw - Math.floor(sw)) * eff;
    return yearOffsets[yi] + relW * eff + fracOffset;
  }

  // Convert absolute month number (1-based, can be fractional) to px offset
  function _monthX(mm) {
    const absM0 = Math.max(0, mm - 1); // 0-based
    const yi = Math.max(0, Math.min(Math.floor(absM0 / 12), years.length - 1));
    const relM = absM0 - yi * 12; // 0-based within year, can be fractional
    const eff = hiddenYears.has(yi) ? collapsedCellW : cellW;
    return yearOffsets[yi] + relM * eff;
  }

  // ── HTML GENERATION ───────────────────────────────────────────

  // ── Header Row 1: Year labels ──────────────────────────────────
  html += '<div class="row-bg header-row-1">';
  html += '<div class="cell header-cell comment-cell" style="border-bottom:none; z-index:30;"></div>';
  html +=
    '<div class="cell header-cell name-cell" style="justify-content:flex-start; align-items:flex-end; padding:0 12px 10px 18px; border-bottom:none; z-index:30;">' +
    '<div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">' +
    '<span class="filter-icon' + (filters.name.length > 0 ? " active" : "") + '" onclick="toggleFilterMenu(\'name\', event)" title="Filter names">▼ Filter name</span>' +
    '<span class="filter-icon' + (filters.type.length > 0 ? " active" : "") + '" onclick="toggleFilterMenu(\'type\', event)" title="Filter by type">&#128203; Filter type</span>' +
    "</div>" +
    "</div>";
  html += '<div class="cell header-cell tl-header" style="padding:0; overflow:visible; border-right:none; height:48px;">';
  years.forEach((yr, yi) => {
    const x = yearOffsets[yi], w = yearWidths[yi];
    const isHidden = hiddenYears.has(yi);
    if (isHidden) {
      html += '<div onclick="toggleYearVisibility(' + yi + ')" title="Expand ' + yr + '"' +
        ' style="position:absolute; left:' + x + 'px; width:' + w + 'px; height:100%; background:#e8edf5; border-right:2px solid #94a3b8; display:flex; align-items:center; justify-content:center; cursor:pointer; box-sizing:border-box;">' +
        '<span style="font-size:9px; color:#64748b; writing-mode:vertical-rl; transform:rotate(180deg); font-weight:bold; line-height:1;">▶ ' + yr + '</span></div>';
    } else {
      html += '<div style="position:absolute; left:' + x + 'px; width:' + w + 'px; height:100%; font-weight:bold; font-size:14px; display:flex; align-items:center; padding:0 4px; gap:2px; box-sizing:border-box; border-right:2px solid #94a3b8;">' +
        '<button onclick="toggleYearVisibility(' + yi + ')" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:11px; color:#94a3b8; flex-shrink:0; line-height:1;" title="Collapse ' + yr + '">◀</button>' +
        '<input type="text" value="' + yr.replace(/"/g, '&quot;') + '" onchange="updateYear(' + yi + ', this.value)" style="border:none; background:transparent; font-weight:bold; font-size:14px; text-align:center; flex:1; outline:none; min-width:0;">' +
        '<div style="width:20px; flex-shrink:0;"></div></div>';
    }
  });
  html += '</div></div>';

  // ── Header Row 2: Month labels (not in months mode) ────────────
  if (zoomMode !== "months") {
    html += '<div class="row-bg header-row-2">';
    html += '<div class="cell header-cell comment-cell" style="border-bottom:none; z-index:30;"></div>';
    html += '<div class="cell header-cell name-cell" style="border-bottom:none; z-index:30;"></div>';
    html += '<div class="cell header-cell tl-header" style="padding:0; overflow:visible; height:48px;">';
    let _mX2 = 0;
    years.forEach((yr, yi) => {
      if (hiddenYears.has(yi)) {
        html += '<div style="position:absolute; left:' + _mX2 + 'px; width:' + yearWidths[yi] + 'px; height:100%; background:#e8edf5; border-right:2px solid #94a3b8; box-sizing:border-box;"></div>';
        _mX2 += yearWidths[yi];
      } else {
        dynamicMonthSpans[yi].forEach((m) => {
          const mW = m.weeks * (zoomMode === "days" ? 5 * cellW : cellW);
          html += '<div style="position:absolute; left:' + _mX2 + 'px; width:' + mW + 'px; height:100%; font-size:12px; display:flex; align-items:center; justify-content:center; border-right:1px solid #e2e8f0; box-sizing:border-box;">' + m.name + '</div>';
          _mX2 += mW;
        });
      }
    });
    html += '</div></div>';
  }

  // ── Header Row 3: Week numbers (days mode only) ────────────────
  if (zoomMode === "days") {
    html += '<div class="row-bg header-row-3">';
    html += '<div class="cell header-cell comment-cell" style="border-bottom:none; z-index:30; height:24px;"></div>';
    html += '<div class="cell header-cell name-cell" style="border-bottom:none; z-index:30; height:24px;"></div>';
    html += '<div class="cell header-cell tl-header" style="padding:0; overflow:visible; height:24px;">';
    for (let w = 1; w <= years.length * 52; w++) {
      const yi = Math.floor((w - 1) / 52);
      const relW = ((w - 1) % 52);
      const x3 = yearOffsets[yi] + relW * 5 * (hiddenYears.has(yi) ? collapsedCellW : cellW);
      const w3 = 5 * (hiddenYears.has(yi) ? collapsedCellW : cellW);
      if (hiddenYears.has(yi)) {
        if (relW === 0) {
          html += '<div style="position:absolute; left:' + x3 + 'px; width:' + yearWidths[yi] + 'px; height:100%; background:#e8edf5;"></div>';
        }
        continue;
      }
      const isHighlighted = isHighlightActive(w, highlightedWeek);
      const _row3Alarms = _alarmWeekMap.get(w) || [];
      const hasAlarm = _row3Alarms.length > 0;
      html += '<div style="position:absolute; left:' + x3 + 'px; width:' + w3 + 'px; height:100%; cursor:pointer; font-size:11px; font-weight:bold; color:#2563eb; display:flex; align-items:center; justify-content:center; border-right:2px solid #94a3b8; box-sizing:border-box;' +
        (isHighlighted ? 'background:#fef08a;' : '') +
        '" onclick="toggleHighlight(\'W\' + ' + w + ')" ondblclick="openAlarmForWeek(' + w + ')" title="Click to highlight week">W' +
        (((w - 1) % 52) + 1) +
        (hasAlarm ? ' <span title="' + _row3Alarms.map(a => a.title + " " + a.time).join(", ") + '" style="font-size:10px;">🔔</span>' : '') +
        '</div>';
    }
    html += '</div></div>';
  }

  // ── Header Row 4: Main header (Comments / Activity + column labels) ──
  html += '<div class="row-bg header-row-4" style="height:56px;">';
  html += '<div class="cell header-cell comment-cell" style="justify-content:flex-start; padding-left:16px; align-items:flex-end; padding-bottom:10px; z-index:30; font-size:13px; font-weight:bold; height:56px;">Comments</div>';

  // Pre-compute which absolute weeks and months contain holidays
  html +=
    '<div class="cell header-cell name-cell" style="justify-content: space-between; align-items: flex-end; padding-left: 18px; padding-right: 8px; z-index: 30; height: 56px; border-bottom: 2px solid var(--border); box-sizing: border-box;">' +
    '<div style="display:flex; align-items:flex-end; width: 100%; padding-bottom: 10px; min-width:0;">' +
    '<span style="font-size: 13px; font-weight: bold; line-height: 1.05;">Family / Project / Milestone / Activity</span>' +
    "</div>" +
    '<button class="settings-btn' + (showSettings ? " active" : "") + '" onclick="toggleAllSettings()" title="Toggle Settings" style="align-self: flex-end; margin-bottom: 8px;">&#9881;</button>' +
    "</div>";

  // Column header timeline cell — individual absolutely positioned labels
  html += '<div class="cell header-cell tl-header" style="padding:0; overflow:visible; height:56px;">';
  {
    const totalCols4 = zoomMode === "weeks" ? totalYears * 52 : zoomMode === "days" ? totalYears * 260 : totalYears * 12;
    let _x4 = 0;
    for (let c = 1; c <= totalCols4; c++) {
      const yi = Math.floor((c - 1) / colsPerYear);
      const isHiddenYr = hiddenYears.has(yi);
      const eff4 = isHiddenYr ? collapsedCellW : cellW;

      if (isHiddenYr) {
        // Draw single collapsed-year block once (at start of year)
        const relC = (c - 1) % colsPerYear;
        if (relC === 0) {
          html += '<div style="position:absolute; left:' + _x4 + 'px; width:' + yearWidths[yi] + 'px; height:100%; background:#e8edf5;"></div>';
          _x4 += yearWidths[yi];
          c += colsPerYear - 1; // skip rest of year
        }
        continue;
      }

      let label = "", extraStyle = "", clickH = "", dblClickH = "", titleH = "";
      if (zoomMode === "weeks") {
        const relW = ((c - 1) % 52) + 1;
        const _yi2 = Math.floor((c - 1) / 52);
        const isHighlighted = isHighlightActive(c, highlightedWeek);
        const _wkAlarms = _alarmWeekMap.get(c) || [];
        const hasAlarm = _wkAlarms.length > 0;
        const isMonthEnd = _monthEndWeekSets[_yi2].has(relW);
        const _hldBorder = _holidayAbsWeeks.has(c) ? "border-bottom:3px solid #f59e0b;" : "";
        label = String(relW) + (hasAlarm ? ' <span style="font-size:9px;">🔔</span>' : '');
        extraStyle = (isHighlighted ? "background:#fef08a;" : "") + (isMonthEnd ? "border-right:4px solid #94a3b8;" : "border-right:1px solid #e2e8f0;") + _hldBorder;
        clickH = 'onclick="toggleHighlight(\'W\' + ' + c + ')" ';
        dblClickH = 'ondblclick="openAlarmForWeek(' + c + ')" ';
        titleH = 'title="Click to highlight week"';
      } else if (zoomMode === "days") {
        const _m = _dayColMeta[c - 1];
        const isHighlighted = isHighlightActive(_m.exactWeekValue, highlightedWeek);
        const isHoliday = _holidayDayCols.has(c);
        const borderStyle = _m.isMonthEnd ? "border-right:4px solid #94a3b8;" : _m.isWeekEnd ? "border-right:2px solid #94a3b8;" : "border-right:1px dashed #e2e8f0;";
        label = '<div style="display:flex;flex-direction:column;align-items:center;line-height:1.2;"><span>' + _m.letter + '</span><span style="font-size:9px;color:#64748b;">' + _m.dateStr + '</span></div>';
        extraStyle = (isHighlighted ? "background:#fef08a;" : "") + (isHoliday ? "background-color:#b0bec5;" : "") + borderStyle;
        clickH = 'onclick="toggleHighlight(' + _m.exactWeekValue + ')" ';
        dblClickH = 'ondblclick="openAlarmForDate(\'' + _m.fullDateStr + '\')" ';
        titleH = 'title="Click to highlight day"';
      } else { // months
        const _yi3 = Math.floor((c - 1) / 12);
        const _mi = (c - 1) % 12;
        const monthName = dynamicMonthSpans[_yi3][_mi].name;
        const _mYN = parseInt(years[_yi3]);
        const _mKey = _mYN + "-" + String(_mi + 1).padStart(2, "0");
        const _mHldBorder = _holidayMonthKeys.has(_mKey) ? "border-bottom:3px solid #f59e0b;" : "";
        const isHighlighted = isHighlightActive(c, highlightedWeek);
        label = '<div style="font-size:11px;font-weight:bold;color:#64748b;text-align:center;">' + monthName + '</div>';
        extraStyle = (isHighlighted ? "background:#fef08a;" : "") + "border-right:1px solid #e2e8f0;" + _mHldBorder;
        clickH = 'onclick="toggleHighlight(' + c + ')" ';
        dblClickH = '';
        titleH = 'title="Click to highlight month\n' + monthName + '"';
      }

      html += '<div ' + clickH + dblClickH + titleH + ' style="position:absolute; left:' + _x4 + 'px; width:' + eff4 + 'px; height:100%; display:flex; align-items:center; justify-content:center; cursor:pointer; font-size:12px; font-weight:600; color:#475569; box-sizing:border-box; ' + extraStyle + '">' + label + '</div>';
      _x4 += eff4;
    }
  }
  html += '</div></div>'; // end tl-header + header-row-4

  // ── Data rows ──────────────────────────────────────────────────
  displayItems.forEach((item, index) => {
    const level = item.level || 0;
    const margin = level * 20;
    const children = getChildren(item.id);
    const hasChildren = children.length > 0;

    let computedStartWeek = item.startWeek;
    let computedDuration = item.duration;

    if (item.type === "task" && holidays.length > 0) {
      computedDuration = computeEffectiveDuration(
        computedStartWeek, computedDuration, item.assignees || []
      );
    }

    if (hasChildren) {
      let minStart = 999, maxEnd = 0;
      function traverse(parentId) {
        const kids = getChildren(parentId);
        kids.forEach((k) => {
          if (k.startWeek < minStart) minStart = k.startWeek;
          const kDur = k.type === "task" && holidays.length > 0
            ? computeEffectiveDuration(k.startWeek, k.duration || 0, k.assignees || [])
            : k.duration || 0;
          const end = k.startWeek + kDur;
          if (end > maxEnd) maxEnd = end;
          traverse(k.id);
        });
      }
      traverse(item.id);
      if (minStart !== 999) {
        computedStartWeek = minStart;
        computedDuration = Math.max(maxEnd - minStart, 0);
      }
    }

    html += '<div class="row-bg">';

    let peerLabel = "";
    if (item._groupPeers && item._groupPeers.length > 0) {
      peerLabel =
        '<div style="font-size:10px;color:#b45309;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.2;" title="' +
        item._groupPeers.map((p) => p.name).join(", ").replace(/"/g, "&quot;") +
        '">+' + item._groupPeers.map((p) => p.name).join(", ") + "</div>";
    }
    html +=
      '<div class="cell comment-cell">' + peerLabel +
      '<input class="comment-input" type="text" value="' + (item.comment || "").replace(/"/g, "&quot;") +
      '" onchange="updateComment(' + item.id + ', this.value)" onclick="handleCommentLinkClick(event)" placeholder="Add comment...">' +
      "</div>";

    html +=
      '<div class="cell name-cell" ondragover="rowDragOver(event)" ondragleave="rowDragLeave(event)" ondrop="rowDrop(event, ' + item.id + ')">' +
      '<div style="display: flex; align-items: center; margin-left: ' + margin + 'px; flex-grow: 1; min-width: 0; overflow: hidden; margin-right: 8px;">' +
      '<div class="drag-handle" draggable="true" ondragstart="startRowDrag(event, ' + item.id + ')" title="Drag to reorder ▪ Alt+Drop onto row to nest inside ▪ Ctrl+Drop to make top-level">&#8942;&#8942;</div>';

    if (item.type === "task" || item.type === "project" || item.type === "family" || item.type === "milestones-group" || item.type === "milestone") {
      const toggleIcon = item.isExpanded === false ? "&#9654;" : "&#9660;";
      const visibility = hasChildren ? "visible" : "hidden";
      html += '<button class="toggle-btn" onclick="toggleExpand(' + item.id + ')" style="font-size: 10px; width: 20px; visibility: ' + visibility + ';">' + toggleIcon + "</button>";
    } else {
      html += '<div style="width: 24px;"></div>';
    }

    if (item.type === "project" || item.type === "family") {
      if (item.name === "New Project") {
        html += '<select class="name-input" onchange="updateProjectName(' + item.id + ', this.value)">';
        html += '<option value="New Project" disabled ' + (item.name === "New Project" ? "selected" : "") + ">Select Project Type...</option>";
        ["Child Project", "Brother Project", "Mother Project"].forEach((m) => {
          html += '<option value="' + m + '" ' + (m === item.name ? "selected" : "") + ">" + m + "</option>";
        });
        html += '<option value="Custom Project">Custom Project...</option>';
        html += "</select>";
      } else {
        html += '<input class="name-input" type="text" value="' + item.name.replace(/"/g, "&quot;") + '" onchange="updateName(' + item.id + ', this.value)" placeholder="Project name...">';
      }
    } else if (item.type === "milestones-group") {
      html += '<span class="name-input" style="font-style:italic;color:#b45309;font-weight:600;cursor:default;">Milestones</span>';
    } else if (item.type === "milestone") {
      const isPredefined = predefinedMilestones.includes(item.name);
      if (isPredefined || item.name === "New Milestone") {
        html += '<select class="name-input" onchange="updateMilestoneName(' + item.id + ', this.value)">';
        html += '<option value="New Milestone" disabled ' + (item.name === "New Milestone" ? "selected" : "") + ">Select Milestone...</option>";
        predefinedMilestones.forEach((m) => {
          html += '<option value="' + m + '" ' + (m === item.name ? "selected" : "") + ">" + m + "</option>";
        });
        html += '<option value="Custom Milestone">Custom Milestone...</option>';
        html += "</select>";
      } else {
        html += '<input class="name-input" type="text" value="' + item.name.replace(/"/g, "&quot;") + '" onchange="updateName(' + item.id + ', this.value)" placeholder="Milestone name...">';
      }
      const weekDisplay = formatYYWWD(item.startWeek);
      html += '<input type="text" maxlength="6" value="' + weekDisplay + '" onchange="setMilestoneWeek(' + item.id + ', this.value)" title="Type YYWW or YYWW.D" style="width:52px;border:1px solid #cbd5e1;border-radius:4px;padding:2px 4px;font-size:11px;font-family:monospace;color:#92400e;background:#fffbeb;text-align:center;flex-shrink:0;">';
    } else {
      if ((item.name === "New Activity" || item.name === "New Sub-activity" || /^Activity \d+$/.test(item.name)) && customTemplates.length > 0) {
        html += '<select class="name-input" onchange="applyActivityTemplate(' + item.id + ', this.value)">';
        html += '<option value="" disabled selected>' + item.name + " (Select Template...)</option>";
        customTemplates.forEach((t) => {
          const durDays = (t.duration * 5).toFixed(1);
          html += '<option value="' + t.id + '">' + t.name.replace(/"/g, "&quot;") + " (" + durDays + " days)</option>";
        });
        html += '<option value="Custom Activity">Custom Activity...</option>';
        html += "</select>";
      } else {
        html += '<input class="name-input" type="text" value="' + item.name.replace(/"/g, "&quot;") + '" onchange="updateName(' + item.id + ', this.value)" placeholder="Activity name...">';
      }
    }

    if (!showSettings && (item.type === "task" || item.type === "project" || item.type === "family")) {
      let displayVal = parseFloat(computedDuration);
      let displayUnit = "wks";
      if (zoomMode === "days") { displayVal = (computedDuration * 5).toFixed(1); displayUnit = "days"; }
      else if (zoomMode === "months") { displayVal = (computedDuration / 4).toFixed(2); displayUnit = "mos"; }
      else { displayVal = displayVal.toFixed(2); }
      displayVal = parseFloat(displayVal);
      html += '<span style="font-size: 11px; color: #64748b; margin-left: 8px; flex-shrink: 0; white-space: nowrap;" title="Duration">(' + displayVal + ' ' + displayUnit + ')</span>';
    }

    html += "</div>";

    if (showSettings) {
      html += '<div class="controls-group" style="flex-shrink: 0; display: flex; gap: 4px; align-items: center;">';
      if (item.type !== "milestones-group") {
        html += '<input type="color" class="color-picker" value="' +
          (item.color || (item.type === "task" ? "#4f46e5" : item.type === "project" ? "#10b981" : "#f59e0b")) +
          '" onchange="updateColor(' + item.id + ', this.value)" title="Choose color">';
      }
      if (item.type === "task" || item.type === "project" || item.type === "family") {
        let displayVal = computedDuration, displayUnit = "wks", displayStep = "0.04", displayTitle = "Duration (weeks. 0.04 = 0.2 day, 0.2 = 1 day)";
        if (zoomMode === "days") { displayVal = (computedDuration * 5).toFixed(1); displayUnit = "days"; displayStep = "0.2"; displayTitle = "Duration (days)"; }
        else if (zoomMode === "months") { displayVal = (computedDuration / 4).toFixed(2); displayUnit = "mos"; displayStep = "0.1"; displayTitle = "Duration (months. 1 month = 4 weeks)"; }
        else { displayVal = parseFloat(computedDuration).toFixed(2); }
        html += '<div class="duration-wrapper"><input class="duration-input" type="number" min="0.04" max="500" step="' + displayStep + '" value="' + displayVal + '" ' +
          (hasChildren ? 'disabled title="Duration is computed from sub-items"' : 'onchange="updateDuration(' + item.id + ', this.value)" title="' + displayTitle + '"') + '>' + displayUnit + '</div>';
      } else if (item.type === "milestone") {
        const displayDate = formatYYWWD(item.startWeek);
        html += '<div class="duration-wrapper" title="Milestone Date (YYWW or YYWW.D)"><input class="milestone-date-input" type="text" value="' + displayDate + '" onchange="updateMilestoneDate(' + item.id + ', this.value)" style="text-transform: uppercase;" inputmode="decimal" spellcheck="false"></div>';
      } else {
        html += '<div style="width: 74px"></div>';
      }

      if (item.type !== "milestone" && item.type !== "milestones-group") {
        const comp = item.completion || 0;
        html += '<div class="duration-wrapper" title="Completion %"><input class="duration-input" type="number" min="0" max="100" step="1" value="' + comp + '" onchange="updateCompletion(' + item.id + ', this.value)">%</div>';
      }

      if (item.type !== "family") {
        html += '<div class="assignee-select" style="position:relative; display:inline-block; margin-right:2px;">' +
          '<button class="add-sub-btn" onclick="toggleAssigneeDropdown(' + item.id + ', event)" style="font-size:11px; padding:2px 4px; line-height:1;" title="Assign Team Members">👥</button></div>';
      }

      const itemIsLocked = item.isLocked === true;
      html += '<button class="lock-btn ' + (itemIsLocked ? "locked" : "") + '" onclick="toggleLock(' + item.id + ', event)" title="' +
        (itemIsLocked ? "🔒 Locked – click to enable movement (also unlocks children and linked items)" : "🔓 Unlocked – click to lock again") + '">' +
        (itemIsLocked ? "🔒" : "🔓") + "</button>";

      if (item.type === "family") {
        html += '<button class="add-sub-btn" onclick="addSubProject(' + item.id + ')" title="Add Sub-project" style="margin-right: 2px;">+P</button>';
      }
      if (item.type === "task" || item.type === "project" || item.type === "family" || item.type === "milestone") {
        html += '<button class="add-sub-btn" onclick="addSubTask(' + item.id + ')" title="Add Sub-activity">+A</button>';
      }
      if (item.type === "task" || item.type === "project" || item.type === "family" || item.type === "milestones-group") {
        html += '<button class="add-sub-btn" onclick="addSubMilestone(' + item.id + ')" title="Add Sub-milestone" style="margin-left: 2px;">+M</button>';
      }
      html += '<button class="duplicate-btn" onclick="duplicateItem(' + item.id + ')" title="Duplicate">D</button>';
      html += '<button class="delete-btn" onclick="deleteItem(' + item.id + ')" title="Delete">&times;</button></div>';
    }
    html += "</div>"; // end name-cell

    // ── Single timeline-row cell ──────────────────────────────────
    html += '<div class="cell timeline-row" ondblclick="handleTlDblClick(event,' + item.id + ')" style="position:relative; overflow:visible; height:48px; border-bottom:1px solid var(--border); box-sizing:border-box;">';

    // ── Alarm indicators ──────────────────────────────────────────
    const itemAlarmsByWeek = _itemAlarmMap.get(item.id);
    if (itemAlarmsByWeek) {
      itemAlarmsByWeek.forEach((alarmList, w) => {
        const ax = _wkX(w);
        const alarmTitles = alarmList.map(a => a.title + " " + a.time).join(", ").replace(/"/g, '&quot;');
        html += '<span title="' + alarmTitles + '" style="position:absolute; left:' + (ax + 2) + 'px; top:2px; font-size:10px; z-index:11; pointer-events:none;">🔔</span>';
      });
    }

    // ── Bar / Milestone rendering ─────────────────────────────────
    if (item.type === "task" || item.type === "project" || item.type === "family") {
      let barLeft, barWidth;
      if (zoomMode === "months") {
        const ms = mapWeekToMonth(computedStartWeek);
        const me = mapWeekToMonth(computedStartWeek + computedDuration);
        barLeft = _monthX(ms);
        barWidth = Math.max(cellW / 4, _monthX(me) - barLeft);
      } else {
        barLeft = _wkX(computedStartWeek);
        const endX = _wkX(computedStartWeek + computedDuration);
        barWidth = Math.max(cellW / 2, endX - barLeft);
      }
      barWidth = Math.min(barWidth, totalTimelineW - barLeft);
      barLeft = Math.max(0, barLeft);
      const finalWidth = Math.max(1, barWidth);

      function getAssigneeLabels(assignees, barWidth2) {
        if (!assignees || assignees.length === 0) return "";
        const showFull = barWidth2 / assignees.length >= 80;
        return assignees.map((aId) => {
          const p = people.find((person) => person.id === aId);
          if (!p) return "";
          const initials = p.name.split(" ").map((n) => n[0]).join("").substring(0, 2).toUpperCase();
          const displayLabel = showFull ? p.name : initials;
          return '<span style="display:inline-block; background:rgba(0,0,0,0.1); color:inherit; font-size:9px; padding:1px 3px; border-radius:3px; margin-left:4px; font-weight:bold; line-height:1; white-space:nowrap;" title="' + p.name + '">' + displayLabel + '</span>';
        }).join("");
      }
      const assigneesHtml = item.assignees && item.assignees.length > 0 ? getAssigneeLabels(item.assignees, finalWidth) : "";
      const label = computedDuration >= 1 ? item.name.replace(/</g, "&lt;") : "";
      const bgColor = item.color || (item.type === "project" ? "#10b981" : "#4f46e5");

      if (hasChildren) {
        const parentPct = Math.min(100, Math.max(0, item.completion || 0));
        const parentProgressHtml = parentPct > 0
          ? '<div style="position:absolute;bottom:0;left:0;width:' + parentPct + '%;height:6px;background:' + bgColor + ';opacity:0.85;pointer-events:none;border-radius:0 0 0 1px;"></div>' +
          '<span style="position:absolute;bottom:1px;right:5px;font-size:8px;font-weight:bold;color:' + bgColor + ';line-height:1;pointer-events:none;text-shadow:0 1px 1px white;">' + parentPct + '%</span>'
          : '';
        html +=
          '<div id="block-' + item.id + '" class="range-bar" style="position:absolute; left:' + barLeft + 'px; width:' + finalWidth + 'px; top:0; margin-left:0; border-color:' + bgColor + ';" onmousedown="startDrag(event,' + item.id + ')" ondblclick="event.stopPropagation()">' +
          '<div class="range-label" style="color:' + bgColor + '; display:flex; align-items:center;">' + label + assigneesHtml + "</div>" +
          parentProgressHtml + getCommentIndicator(item) + "</div>";
      } else {
        const completionPct = Math.min(100, Math.max(0, item.completion || 0));
        const progressHtml = completionPct > 0
          ? '<div style="position:absolute;left:0;top:0;height:100%;width:' + completionPct + '%;background:rgba(255,255,255,0.45);border-right:1px solid rgba(255,255,255,0.4);border-radius:6px' + (completionPct >= 100 ? '' : ' 0 0 6px') + ';pointer-events:none;"></div>'
          : '';

        const CHAR_W = 7, PAD = 12;
        const innerW = Math.max(0, finalWidth - PAD);
        const pctStr = completionPct > 0 ? completionPct + "%" : "";
        const pctPx = pctStr ? pctStr.length * CHAR_W + 12 : 0;
        const rawName = item.name;
        const fullNamePx = rawName.length * CHAR_W;
        let displayLabel = "", displayPct = "";

        if (innerW <= 0) {
          displayLabel = ""; displayPct = "";
        } else if (pctStr === "") {
          const charsAvail = Math.floor(innerW / CHAR_W);
          if (charsAvail >= rawName.length) displayLabel = rawName;
          else if (charsAvail >= 2) displayLabel = rawName.substring(0, charsAvail - 1) + "…";
          else if (charsAvail === 1) displayLabel = rawName.charAt(0);
          displayPct = "";
        } else if (fullNamePx + pctPx <= innerW) {
          displayLabel = rawName; displayPct = pctStr;
        } else {
          const nameW = innerW - pctPx;
          if (nameW >= CHAR_W * 2) {
            const charsAvail = Math.floor(nameW / CHAR_W);
            displayLabel = charsAvail >= rawName.length ? rawName : rawName.substring(0, charsAvail - 1) + "…";
            displayPct = pctStr;
          } else if (nameW >= CHAR_W) {
            displayLabel = rawName.charAt(0); displayPct = pctStr;
          } else if (innerW >= CHAR_W) {
            displayLabel = rawName.charAt(0); displayPct = "";
          }
        }

        const safeLabel = displayLabel ? displayLabel.replace(/</g, "&lt;") : "";
        const pctLabelHtml = displayPct
          ? '<span style="flex-shrink:0;font-size:10px;font-weight:600;color:white;padding:0 8px 0 2px;white-space:nowrap;pointer-events:none;text-shadow:0 1px 2px rgba(0,0,0,0.3);">' + displayPct + "</span>"
          : "";
        const _isAutoSized = hasChildren && (item.type === "project" || item.type === "family" || item.type === "milestones-group");

        html +=
          '<div id="block-' + item.id + '" class="bar" style="position:absolute; left:' + barLeft + 'px; width:' + finalWidth + 'px; margin-left:0; background-color:' + bgColor + ';" onmousedown="startDrag(event,' + item.id + ')" ondblclick="event.stopPropagation()">' +
          progressHtml +
          (_isAutoSized ? '' : '<div class="resize-handle left" onmousedown="startResize(event,' + item.id + ",'left')\"></div>") +
          '<div class="bar-label" style="display:flex; align-items:center; overflow:hidden; white-space:nowrap; min-width:0;">' + safeLabel + assigneesHtml + "</div>" +
          pctLabelHtml +
          (_isAutoSized ? '' : '<div class="resize-handle right" onmousedown="startResize(event,' + item.id + ",'right')\"></div>") +
          getCommentIndicator(item) + "</div>";
      }
    } else if (item.type === "milestone") {
      const _parentItem = items.find(i => i.id === item.parentId);
      const _parentIsMilestonesGroup = _parentItem && _parentItem.type === "milestones-group";
      if (!_parentIsMilestonesGroup) {
        const msX = zoomMode === "months" ? _monthX(mapWeekToMonth(item.startWeek)) : _wkX(item.startWeek);
        const lineHeight = (totalRows - index + 1) * 48;
        const bgColor = item.color || "#f59e0b";
        let r = 245, g = 158, b = 11;
        if (bgColor.match(/^#[0-9a-fA-F]{6}$/)) {
          r = parseInt(bgColor.slice(1, 3), 16);
          g = parseInt(bgColor.slice(3, 5), 16);
          b = parseInt(bgColor.slice(5, 7), 16);
        }
        const lineBg = "rgba(" + r + "," + g + "," + b + ",0.4)";
        const shortLabel = item.name.replace(/\s*\(.*/, "").trim().toUpperCase();
        html +=
          '<div class="milestone-wrapper" style="position:absolute; left:' + msX + 'px; top:0; width:0; height:48px; overflow:visible;">' +
          '<div class="milestone-line" style="height:' + lineHeight + 'px; background-color:' + lineBg + ';"></div>' +
          '<div style="display:flex;flex-direction:column;align-items:center;pointer-events:none;position:relative;z-index:10;">' +
          '<span style="font-size:9px;font-weight:700;color:' + bgColor + ';white-space:nowrap;line-height:1.1;letter-spacing:0.03em;background:rgba(255,255,255,0.88);border-radius:2px;padding:0 2px;">' + shortLabel.replace(/</g, "&lt;") + "</span>" +
          '<span style="font-size:8px;font-weight:600;color:' + bgColor + ';white-space:nowrap;line-height:1;background:rgba(255,255,255,0.88);border-radius:2px;padding:0 2px;margin-top:1px;">' + fmtMsDate(item.startWeek) + "</span>" +
          '<div id="block-' + item.id + '" class="milestone" style="background-color:' + bgColor + '; pointer-events:auto;" onmousedown="startDrag(event,' + item.id + ')" ondblclick="event.stopPropagation()" title="' + item.name.replace(/"/g, "&quot;") + '">' +
          getCommentIndicator(item) + "</div></div></div>";
      }
    }

    // ── Descendant milestones (collapsed or milestones-group) ─────
    if (item.type === "milestones-group" || ((item.type === "task" || item.type === "project" || item.type === "family") && item.isExpanded === false)) {
      let _tailHeight = 0;
      if (item.type === "milestones-group") {
        const _mgLevel = item.level || 0;
        let _lastDescIdx = index;
        for (let k = index + 1; k < displayItems.length; k++) {
          const kLevel = displayItems[k].level || 0;
          if (_mgLevel === 0 ? kLevel === 0 : kLevel < _mgLevel) break;
          _lastDescIdx = k;
        }
        _tailHeight = (_lastDescIdx - index + 1) * 48;
      }

      const descMilestones = getDescendantMilestones(item.id);
      descMilestones.forEach((ms) => {
        const msX = zoomMode === "months" ? _monthX(mapWeekToMonth(ms.startWeek)) : _wkX(ms.startWeek);
        const msColor = ms.color || "#f59e0b";
        const msShortLabel = ms.name.replace(/\s*\(.*/, "").trim().toUpperCase();
        let _lineHtml = "";
        if (item.type === "milestones-group" && _tailHeight > 0) {
          let _lr = 245, _lg = 158, _lb = 11;
          if (msColor.match(/^#[0-9a-fA-F]{6}$/)) {
            _lr = parseInt(msColor.slice(1, 3), 16);
            _lg = parseInt(msColor.slice(3, 5), 16);
            _lb = parseInt(msColor.slice(5, 7), 16);
          }
          _lineHtml = '<div class="milestone-line" style="height:' + _tailHeight + 'px;background-color:rgba(' + _lr + ',' + _lg + ',' + _lb + ',0.4);"></div>';
        }
        html +=
          '<div class="milestone-wrapper" style="position:absolute; left:' + msX + 'px; top:0; width:0; height:48px; overflow:visible;">' +
          _lineHtml +
          '<div style="display:flex;flex-direction:column;align-items:center;pointer-events:none;position:relative;z-index:12;">' +
          '<div id="block-' + ms.id + '" class="milestone" style="background-color:' + msColor + ';pointer-events:auto;" onmousedown="startDrag(event,' + ms.id + ')" ondblclick="event.stopPropagation()" title="' + ms.name.replace(/"/g, "&quot;") + '">' + getCommentIndicator(ms) + "</div>" +
          '<span style="font-size:9px;font-weight:700;color:' + msColor + ';white-space:nowrap;line-height:1.2;letter-spacing:0.03em;background:rgba(255,255,255,0.88);border-radius:2px;padding:0 2px;margin-top:1px;">' + msShortLabel.replace(/</g, "&lt;") + "</span>" +
          '<span style="font-size:8px;font-weight:600;color:' + msColor + ';white-space:nowrap;line-height:1;background:rgba(255,255,255,0.88);border-radius:2px;padding:0 2px;margin-top:1px;">' + fmtMsDate(ms.startWeek) + "</span></div></div>";
      });
    }

    html += "</div>"; // end timeline-row
    html += "</div>"; // end row-bg
  });

  // ── Bottom drop row ────────────────────────────────────────────
  html += '<div class="row-bg">';
  html += '<div class="cell comment-cell" style="border-bottom:none"></div>';
  html += '<div class="cell name-cell" style="border-bottom:none; color:#cbd5e1; font-style:italic; justify-content:center;" ondragover="rowDragOverBottom(event)" ondragleave="rowDragLeaveBottom(event)" ondrop="rowDrop(event, \'bottom\')"></div>';
  html += '<div class="cell timeline-row" ondblclick="handleTlDblClick(event,null)" style="border-bottom:none; height:48px; position:relative; overflow:visible;"></div>';
  html += '</div>';

  // ── Filter menus ───────────────────────────────────────────────
  if (activeFilterMenu) {
    let menuHtml =
      '<div class="filter-menu open" style="top: 150px; left: ' +
      (activeFilterMenu === "name" ? 250 : 20) + 'px;" onclick="event.stopPropagation()">';
    menuHtml += "<strong>Filter " + (activeFilterMenu === "name" ? "by Name" : "by Type") + "</strong><br><br>";

    if (activeFilterMenu === "type") {
      ["family", "project", "milestone", "task"].forEach((t) => {
        const checked = !filters.type.includes(t) ? "checked" : "";
        const label = t.charAt(0).toUpperCase() + t.slice(1);
        menuHtml += '<label><input type="checkbox" ' + checked + " onclick=\"handleFilterChange('type', '" + t + "', event)\"> " + label + "</label>";
      });
    } else if (activeFilterMenu === "name") {
      menuHtml += '<input class="filter-menu-search" type="text" placeholder="Type project or item name" oninput="filterNameFilterOptions(this.value)" onclick="event.stopPropagation()">';
      const eligibleItems = items.filter((i) => {
        if (!i.parentId) return true;
        const parent = items.find((p) => p.id === i.parentId);
        if (parent && !parent.parentId) return true;
        return false;
      });
      const uniqueNames = [...new Set(eligibleItems.map((i) => (i.name || "").trim()).filter(Boolean))].sort();
      uniqueNames.slice(0, 50).forEach((n) => {
        const checked = !filters.name.includes(n) ? "checked" : "";
        const safeName = n.replace(/'/g, "\\'").replace(/"/g, "&quot;");
        const searchText = n.toLowerCase().replace(/"/g, "&quot;");
        menuHtml += '<label class="filter-name-option" data-filter-search="' + searchText + '"><input type="checkbox" ' + checked + " onclick=\"handleFilterChange('name', '" + safeName + "', event)\"> " + n + "</label>";
      });
      menuHtml += '<div class="filter-menu-empty" style="display:none;">No matching names</div>';
      if (uniqueNames.length > 50) menuHtml += '<div style="font-size:10px; color:#64748b; margin-top:4px;">(Showing first 50 results)</div>';
    }

    menuHtml += '<div class="filter-menu-actions">';
    menuHtml += '<button class="success" style="padding: 2px 8px; font-size:11px;" onclick="clearFilter(\'' + activeFilterMenu + "')\">Clear All</button>";
    menuHtml += '<button class="outline" style="padding: 2px 8px; font-size:11px;" onclick="toggleFilterMenu(null, event)">Close</button>';
    menuHtml += "</div></div>";
    html += menuHtml;
  }

  // ── Write DOM ─────────────────────────────────────────────────
  gridEl.innerHTML = html;
  if (activeFilterMenu === "name") {
    const searchInput = gridEl.querySelector(".filter-menu-search");
    if (searchInput) searchInput.focus();
  }

  // ── Draw background grid canvas ───────────────────────────────
  (function drawBgCanvas() {
    // Compute header heights (cached by zoomMode)
    if (_cachedHeaderZoom !== zoomMode || _cachedHeaderHeights === null) {
      const r1 = gridEl.querySelector(".header-row-1 .cell");
      const r2 = gridEl.querySelector(".header-row-2 .cell");
      const r3 = gridEl.querySelector(".header-row-3 .cell");
      const r4 = gridEl.querySelector(".header-row-4 .cell");
      _cachedHeaderHeights = {
        h1: r1 ? r1.offsetHeight : 48,
        h2: r2 ? r2.offsetHeight : 0,
        h3: r3 ? r3.offsetHeight : 0,
        h4: r4 ? r4.offsetHeight : 56,
      };
      _cachedHeaderZoom = zoomMode;
    }
    const { h1, h2, h3, h4 } = _cachedHeaderHeights;
    gridEl.style.setProperty("--row2-top", h1 + "px");
    gridEl.style.setProperty("--row3-top", (h1 + h2) + "px");
    gridEl.style.setProperty("--main-header-top", (h1 + h2 + h3) + "px");

    const headerH = h1 + h2 + h3 + h4;
    const bgH = headerH + (totalRows + 1) * 48;

    // Measure actual timeline column offset from DOM to ensure canvas aligns with headers
    const firstTlHeader = gridEl.querySelector(".tl-header");
    const timelineLeft = firstTlHeader ? firstTlHeader.offsetLeft : (showComments ? commentWidth : 0) + nameWidthBase + (showSettings ? 280 : 0);

    // Get or create background canvas
    let bgCanvas = document.getElementById("tl-bg");
    if (!bgCanvas) {
      bgCanvas = document.createElement("canvas");
      bgCanvas.id = "tl-bg";
      bgCanvas.style.cssText = "position:absolute; pointer-events:none; z-index:0;";
      gridEl.appendChild(bgCanvas);
    }
    bgCanvas.style.left = timelineLeft + "px";
    bgCanvas.style.top = "0px";
    bgCanvas.width = totalTimelineW;
    bgCanvas.height = bgH;

    const ctx = bgCanvas.getContext("2d");
    ctx.clearRect(0, 0, totalTimelineW, bgH);

    // Draw vertical column lines
    const totalCols = zoomMode === "weeks" ? totalYears * 52 : zoomMode === "days" ? totalYears * 260 : totalYears * 12;
    let _bgX = 0;
    for (let c = 1; c <= totalCols; c++) {
      const yi = Math.floor((c - 1) / colsPerYear);
      const isHiddenYr = hiddenYears.has(yi);
      const eff = isHiddenYr ? collapsedCellW : cellW;

      if (isHiddenYr) {
        const relC = (c - 1) % colsPerYear;
        if (relC === 0) {
          // Draw collapsed year zone
          ctx.fillStyle = "#f1f4f9";
          ctx.fillRect(_bgX, 0, yearWidths[yi], bgH);
          // Right border — offset by lw/2 to match CSS border-box positioning
          ctx.strokeStyle = "#94a3b8";
          ctx.lineWidth = 2;
          ctx.setLineDash([]);
          ctx.beginPath();
          ctx.moveTo(_bgX + yearWidths[yi] - 1, 0);
          ctx.lineTo(_bgX + yearWidths[yi] - 1, bgH);
          ctx.stroke();
          _bgX += yearWidths[yi];
          c += colsPerYear - 1;
        }
        continue;
      }

      const relC = (c - 1) % colsPerYear;
      const x = _bgX + eff; // right edge of this column

      // Holiday column tint (data rows only)
      if (zoomMode === "weeks" && _holidayAbsWeeks.has(c)) {
        ctx.fillStyle = "rgba(245,158,11,0.12)";
        ctx.fillRect(_bgX, headerH, eff, bgH - headerH);
      } else if (zoomMode === "days" && _generalHolidayDayCols.has(c)) {
        ctx.fillStyle = "rgba(176,190,197,0.35)";
        ctx.fillRect(_bgX, headerH, eff, bgH - headerH);
      }

      // Highlighted column
      let isHL = false;
      if (zoomMode === "weeks") {
        isHL = isHighlightActive(c, highlightedWeek);
      } else if (zoomMode === "days") {
        const _m = _dayColMeta ? _dayColMeta[c - 1] : null;
        isHL = _m ? isHighlightActive(_m.exactWeekValue, highlightedWeek) : false;
      } else {
        isHL = isHighlightActive(c, highlightedWeek);
      }
      if (isHL) {
        ctx.fillStyle = "rgba(254,240,138,0.4)";
        ctx.fillRect(_bgX, 0, eff, bgH);
      }

      // Determine line style
      let isMonthEnd = false, isWeekEnd = false;
      if (zoomMode === "weeks") {
        const relW = relC + 1;
        isMonthEnd = _monthEndWeekSets[yi].has(relW);
      } else if (zoomMode === "days") {
        const intWeek = Math.floor(relC / 5) + 1;
        isWeekEnd = (relC + 1) % 5 === 0;
        isMonthEnd = isWeekEnd && _monthEndWeekSets[yi].has(intWeek);
      }
      // months: every column is a month, use regular line

      // Determine line style first so we can offset the stroke to match CSS border-box positioning
      // CSS border-right is drawn INSIDE the box: [x-borderW, x]
      // Canvas stroke is centered on the path: [x-lw/2, x+lw/2]
      // To align: draw at x - lw/2 so canvas stroke spans [x-lw, x] matching CSS border
      let lw, lColor, lDash;
      if (isMonthEnd) {
        lColor = "#94a3b8"; lw = 4; lDash = null;
      } else if (isWeekEnd) {
        lColor = "#94a3b8"; lw = 2; lDash = null;
      } else if (zoomMode === "days") {
        lColor = "#e2e8f0"; lw = 1; lDash = [3, 3];
      } else {
        lColor = "#e2e8f0"; lw = 1; lDash = null;
      }
      const lx = x - lw / 2;
      ctx.beginPath();
      ctx.moveTo(lx, 0);
      ctx.lineTo(lx, bgH);
      ctx.strokeStyle = lColor;
      ctx.lineWidth = lw;
      if (lDash) ctx.setLineDash(lDash); else ctx.setLineDash([]);
      ctx.stroke();
      ctx.setLineDash([]);
      _bgX += eff;
    }

    // Draw horizontal row separator lines (data area only)
    ctx.strokeStyle = "#e2e8f0"; ctx.lineWidth = 1; ctx.setLineDash([]);
    for (let r = 0; r <= totalRows + 1; r++) {
      const y = headerH + r * 48;
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(totalTimelineW, y);
      ctx.stroke();
    }
  })();

  // ── Inject col-resizers ────────────────────────────────────────
  if (showComments) {
    const commentResizer = document.createElement("div");
    commentResizer.className = "col-resizer";
    commentResizer.style.left = "calc(var(--comment-width) - 2px)";
    commentResizer.addEventListener("mousedown", (e) => startColResize(e, "comment"));
    gridEl.appendChild(commentResizer);
  }
  const nameResizer = document.createElement("div");
  nameResizer.className = "col-resizer";
  nameResizer.style.left = "calc(var(--comment-width) + var(--name-width) - 2px)";
  nameResizer.addEventListener("mousedown", (e) => startColResize(e, "name"));
  gridEl.appendChild(nameResizer);

  // ── Inject Link Anchor Dots ────────────────────────────────────
  if (linkMode) {
    displayItems.forEach((item) => {
      const block = document.getElementById("block-" + item.id);
      if (!block) return;
      if (block.style.position !== "absolute") block.style.position = "relative";
      const isActiveFrom = linkSource && linkSource.itemId === item.id && linkSource.anchor === "start";
      const isActiveTo = linkSource && linkSource.itemId === item.id && linkSource.anchor === "end";
      const startDot = document.createElement("div");
      startDot.className = "link-anchor start-anchor" + (isActiveFrom ? " active" : "");
      startDot.title = 'Link: start of "' + item.name + '"';
      startDot.onclick = (e) => { e.stopPropagation(); clickAnchor(item.id, "start"); };
      const endDot = document.createElement("div");
      endDot.className = "link-anchor end-anchor" + (isActiveTo ? " active" : "");
      endDot.title = 'Link: end of "' + item.name + '"';
      endDot.onclick = (e) => { e.stopPropagation(); clickAnchor(item.id, "end"); };
      block.style.overflow = "visible";
      block.appendChild(startDot);
      block.appendChild(endDot);
    });
  }

  // ── Draw dependency link SVG overlay ──────────────────────────
  drawLinks();

  // ── Block Comment Popups ──────────────────────────────────────
  items.forEach((item) => {
    const existing = document.getElementById(`comment-popup-${item.id}`);
    if (existing) existing.remove();

    if (item.blockComment && item.blockComment.isOpen) {
      const popup = document.createElement("div");
      popup.id = `comment-popup-${item.id}`;
      popup.className = "comment-popup";
      popup.style.left = (item.blockComment.x || 100) + "px";
      popup.style.top = (item.blockComment.y || 100) + "px";
      popup.style.width = (item.blockComment.width || 200) + "px";
      popup.style.height = (item.blockComment.height || 150) + "px";
      popup.innerHTML = `
        <div class="comment-popup-header" onmousedown="startCommentDrag(event, ${item.id})">
          <div style="display:flex; gap: 8px;">
            <span class="comment-popup-delete" title="Delete comment" onmousedown="event.stopPropagation()" onclick="deleteComment(${item.id}); event.stopPropagation()">🗑️</span>
            <span class="comment-popup-close" title="Close" onmousedown="event.stopPropagation()" onclick="closeComment(${item.id}); event.stopPropagation()">&times;</span>
          </div>
        </div>
        <textarea class="comment-popup-textarea" oninput="updateBlockCommentText(${item.id}, this.value)" onclick="handleCommentLinkClick(event)" placeholder="Write a comment...">${item.blockComment.text || ""}</textarea>
        <div class="comment-resizer" onmousedown="startCommentResize(event, ${item.id})"></div>
      `;
      gridEl.appendChild(popup);
      updatePopupPointer(item, popup);
    }
  });
}
