// ═══════════════════════════════════════════════════════════════
//  RENDER ENGINE + DISPLAY ITEMS + EFFECTIVE DURATION
// ═══════════════════════════════════════════════════════════════

const gridEl = document.getElementById("grid");

function syncStickyColumnResizers() {
  const container = document.querySelector(".planner-container");
  if (!gridEl || !container) return;
  gridEl.style.setProperty("--sticky-scroll-left", container.scrollLeft + "px");
}

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

// Closure reference to the most recent drawBgCanvas function (captures render() scope)
let _drawCanvas = null;

// Redraw the background canvas (grid lines + today line) without touching the DOM.
// Safe to call during today-bar drag — uses the state from the last full render().
function redrawBgCanvas() {
  if (_drawCanvas) _drawCanvas();
}

// Shared empty Map sentinel — avoids allocating a new Map() on every cache miss
const _emptyMap = new Map();

// Most-recent day-column metadata (exposed for click handlers)
let _currentDayColMeta = null;

// Today line — absolute week position (with day fraction), null = auto from system date
let _todayLineWeek = null;
let _todayLineVisible = true;

function _getSystemTodayWeek() {
  const now = new Date();
  const yr = String(now.getFullYear());
  const yi = years.indexOf(yr);
  if (yi === -1) return null;
  // ISO week: find Monday of ISO week 1
  const jan4 = new Date(now.getFullYear(), 0, 4);
  const soy = new Date(jan4.getTime());
  soy.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1);
  const diffMs = now.getTime() - soy.getTime();
  const diffDays = Math.floor(diffMs / 86400000);
  const isoWeek = Math.floor(diffDays / 7) + 1;
  const dayOfWeek = now.getDay(); // 0=Sun,1=Mon..5=Fri,6=Sat
  // Map to fraction: Mon=0, Tue=0.2, Wed=0.4, Thu=0.6, Fri=0.8
  // Weekend days clamp to Friday
  const dayFrac = dayOfWeek === 0 ? 0.8 : dayOfWeek === 6 ? 0.8 : (dayOfWeek - 1) * 0.2;
  const absWeek = getAbsWeekFromYearWeek(yi, Math.min(isoWeek, getIsoWeeksInYear(yr)));
  return absWeek + dayFrac;
}

function getTodayLineWeek() {
  if (_todayLineWeek !== null) return _todayLineWeek;
  return _getSystemTodayWeek();
}

function setTodayLineWeek(wk) {
  _todayLineWeek = wk;
}

function resetTodayLine() {
  _todayLineWeek = null;
  render();
}

function toggleTodayLine() {
  _todayLineVisible = !_todayLineVisible;
  render();
}

function getTodayLineTimelineX() {
  const todayWk = getTodayLineWeek();
  if (todayWk === null) return null;

  const { cw, collW, offsets } = _computeYearLayout();
  const wholeWeek = Math.max(1, Math.floor(todayWk));
  const info = getYearWeekInfo(wholeWeek);
  const yi = info.yearIndex;
  if (yi < 0) return null;

  if (hiddenYears.has(yi)) {
    return offsets[yi] + collW / 2;
  }

  if (zoomMode === "months") {
    let weekCursor = 1;
    for (let y = 0; y < years.length; y++) {
      const spans = getMonthWeekSpans(years[y]);
      for (let m = 0; m < 12; m++) {
        const spanWeeks = spans[m].weeks;
        if (todayWk >= weekCursor && todayWk < weekCursor + spanWeeks) {
          return (y * 12 + m + (todayWk - weekCursor) / spanWeeks) * cw;
        }
        weekCursor += spanWeeks;
      }
    }
    return years.length * 12 * cw;
  }

  if (zoomMode === "days") {
    const dayOffset = Math.max(0, Math.min(4.999, (todayWk - wholeWeek) * 5));
    return offsets[yi] + ((info.relWeek - 1) * 5 + dayOffset) * cw + cw / 2;
  }

  return offsets[yi] + (info.relWeek - 1 + (todayWk - wholeWeek)) * cw;
}

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
    // Custom column filters
    var colFilterKeys = Object.keys(columnFilters);
    for (var cf = 0; cf < colFilterKeys.length; cf++) {
      var cfId = colFilterKeys[cf];
      var cfVals = columnFilters[cfId];
      if (cfVals && cfVals.length > 0) {
        var cellVal = (item.customData && item.customData[cfId]) || "";
        if (cfVals.includes(cellVal)) return true;
      }
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
  let remainingWorkDays = duration * 5;
  if (remainingWorkDays <= 0) return duration;

  // Resolve fractional startWeek to absolute week + day index
  const startW = Math.floor(startWeek);
  const startDayFract = Math.round((startWeek - startW) * 5); // 0=Mon

  const weekInfo = getYearWeekInfo(startW);
  const yearNum = parseInt(weekInfo.yearLabel) || new Date().getFullYear();
  const jan4 = new Date(yearNum, 0, 4);
  const weekStart = new Date(jan4.getTime());
  weekStart.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1); // Mon of ISO wk1

  // Date of the first day of this block
  const blockStart = new Date(weekStart.getTime());
  blockStart.setDate(
    weekStart.getDate() + (weekInfo.relWeek - 1) * 7 + startDayFract,
  );

  let occupiedWorkDays = 0;
  let calDays = 0;
  const maxCalDays = Math.ceil(remainingWorkDays) + applicableHolidays.length + 30; // safety cap
  const applicableHolidayDates = new Set(applicableHolidays.map(h => h.date));

  while (remainingWorkDays > 0.0001 && calDays < maxCalDays) {
    const cur = new Date(blockStart.getTime());
    cur.setDate(blockStart.getDate() + calDays);
    const dow = cur.getDay(); // 0=Sun, 6=Sat
    if (dow !== 0 && dow !== 6) {
      if (applicableHolidayDates.has(formatDateStr(cur))) {
        occupiedWorkDays += 1;
      } else {
        const workedToday = Math.min(1, remainingWorkDays);
        occupiedWorkDays += workedToday;
        remainingWorkDays -= workedToday;
      }
    }
    calDays++;
  }

  return parseFloat((occupiedWorkDays / 5).toFixed(4));
}

// ── Year layout helper (used by render and click handlers) ──────
function _computeYearLayout() {
  const cw = zoomMode === "days" ? 20 : zoomMode === "months" ? 80 : 28;
  const collW = 24; // fixed narrow width for any collapsed year
  const yearWeekCounts = years.map((yearLabel) => getIsoWeeksInYear(yearLabel));
  const yearDayCounts = yearWeekCounts.map((count) => count * 5);
  const widths = years.map((_, yi) => {
    if (hiddenYears.has(yi)) return collW;
    const cols = zoomMode === "days" ? yearDayCounts[yi] : zoomMode === "months" ? 12 : yearWeekCounts[yi];
    return cols * cw;
  });
  const offsets = [];
  let acc = 0;
  widths.forEach(w => { offsets.push(acc); acc += w; });
  return { cw, collW, widths, offsets, total: acc, yearWeekCounts, yearDayCounts };
}

function _monthColumnToAbsWeek(yearIndex, monthIndex) {
  let absWeek = 1;
  for (let yi = 0; yi < years.length; yi++) {
    const spans = getMonthWeekSpans(years[yi]).map((span) => span.weeks);
    for (let mi = 0; mi < 12; mi++) {
      if (yi === yearIndex && mi === monthIndex) return absWeek;
      absWeek += spans[mi];
    }
  }
  return getAbsWeekFromYearWeek(yearIndex, 1);
}

function _getRelativeTimelineX(e, targetEl) {
  const anchorEl = targetEl || e.currentTarget || e.target;
  if (
    anchorEl &&
    typeof anchorEl.getBoundingClientRect === "function" &&
    typeof e.clientX === "number"
  ) {
    const rect = anchorEl.getBoundingClientRect();
    let zoomFactor = 1;
    if (typeof getComputedStyle === "function" && document.body) {
      zoomFactor = parseFloat(getComputedStyle(document.body).zoom) || 1;
    } else if (document.body && document.body.style) {
      zoomFactor = parseFloat(document.body.style.zoom) || 1;
    }
    return Math.max(0, (e.clientX - rect.left) / zoomFactor);
  }
  return typeof e.offsetX === "number" ? e.offsetX : 0;
}

// ── Global click handlers for timeline cells ────────────────────
function handleTlDblClick(e, targetEl, itemId) {
  if (e.target.closest('.bar, .milestone, .range-bar, .resize-handle')) return;
  const x = _getRelativeTimelineX(e, targetEl);
  const { cw, collW, widths, yearWeekCounts, yearDayCounts } = _computeYearLayout();
  let remaining = x, snapW = 1, dayIdx = 0;
  for (let yi = 0; yi < years.length; yi++) {
    const yw = widths[yi];
    if (remaining <= yw || yi === years.length - 1) {
      if (hiddenYears.has(yi)) return; // collapsed year — ignore click
      if (zoomMode === "days") {
        const col = Math.max(0, Math.min(Math.floor(remaining / cw), yearDayCounts[yi] - 1));
        snapW = getAbsWeekFromYearWeek(yi, Math.floor(col / 5) + 1);
        dayIdx = col % 5;
      } else if (zoomMode === "weeks") {
        snapW = getAbsWeekFromYearWeek(yi, Math.max(0, Math.min(Math.floor(remaining / cw), yearWeekCounts[yi] - 1)) + 1);
      } else {
        const monthIndex = Math.max(0, Math.min(Math.floor(remaining / cw), 11));
        snapW = _monthColumnToAbsWeek(yi, monthIndex);
      }
      break;
    }
    remaining -= yw;
  }
  openAlarmForWeek(snapW, itemId, dayIdx);
}

function handleTimelineRowClick(e, itemId) {
  if (typeof _suppressRowHighlightUntil !== "undefined" && Date.now() < _suppressRowHighlightUntil) return;
  if (e.target.closest('.bar, .milestone, .range-bar, .resize-handle, button, input, select, textarea, a')) return;
  toggleRowHighlight(itemId);
}

function handleTimelineHeaderClick(e) {
  const x = _getRelativeTimelineX(e);
  const { cw, collW, widths, yearWeekCounts, yearDayCounts } = _computeYearLayout();
  const { offsets: yearDayOffsets } = getYearDayOffsets();
  let remaining = x;
  for (let yi = 0; yi < years.length; yi++) {
    const yw = widths[yi];
    if (remaining <= yw || yi === years.length - 1) {
      if (hiddenYears.has(yi)) return; // collapsed year — ignore click
      if (zoomMode === "days") {
        const col = Math.max(0, Math.min(Math.floor(remaining / cw), yearDayCounts[yi] - 1));
        const w = getAbsWeekFromYearWeek(yi, Math.floor(col / 5) + 1);
        const dayFrac = (col % 5) * 0.2;
        toggleHighlight(dayFrac > 0 ? parseFloat((w + dayFrac).toFixed(1)) : "W" + w);
      } else if (zoomMode === "weeks") {
        const w = getAbsWeekFromYearWeek(yi, Math.max(0, Math.min(Math.floor(remaining / cw), yearWeekCounts[yi] - 1)) + 1);
        toggleHighlight("W" + w);
      } else {
        const c = yi * 12 + Math.max(0, Math.floor(remaining / cw)) + 1;
        toggleHighlight(c);
      }
      return;
    }
    remaining -= yw;
  }
}

function handleTimelineHeaderDblClick(e) {
  const x = _getRelativeTimelineX(e);
  const { cw, collW, widths, yearWeekCounts, yearDayCounts } = _computeYearLayout();
  const { offsets: yearDayOffsets } = getYearDayOffsets();
  let remaining = x;
  for (let yi = 0; yi < years.length; yi++) {
    const yw = widths[yi];
    if (remaining <= yw || yi === years.length - 1) {
      if (hiddenYears.has(yi)) return; // collapsed year — ignore dblclick
      if (zoomMode === "days" && _currentDayColMeta) {
        const col = Math.max(0, Math.min(Math.floor(remaining / cw), yearDayCounts[yi] - 1));
        const metaIdx = yearDayOffsets[yi] + col;
        if (_currentDayColMeta[metaIdx]) openAlarmForDate(_currentDayColMeta[metaIdx].fullDateStr);
      } else {
        const w = zoomMode === "months"
          ? _monthColumnToAbsWeek(yi, Math.max(0, Math.min(Math.floor(remaining / cw), 11)))
          : getAbsWeekFromYearWeek(yi, Math.max(0, Math.min(Math.floor(remaining / cw), yearWeekCounts[yi] - 1)) + 1);
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
          // Milestones have duration 0 — extend by half a week so the parent block ends at the diamond center
          const kDur = (k.type === "milestone" && (k.duration || 0) === 0) ? 0.5 : (k.duration || 0);
          const kEnd = k.startWeek + kDur;
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

  const fmtMsDate = (sw) => formatYYWWD(sw);

  // Clear all dynamic SVG tails from previous render
  const svg = document.getElementById("comment-tail-svg-container");
  if (svg) svg.innerHTML = "";

  // ── CSS variables ──────────────────────────────────────────────
  const actualCommentWidth = showComments ? commentWidth + "px" : "0px";
  const actualNameWidth = nameWidthBase + (showSettings ? 280 : 0) + "px";
  gridEl.style.setProperty("--comment-width", actualCommentWidth);
  gridEl.style.setProperty("--name-width", actualNameWidth);

  // Custom columns layout
  const visibleCols = customColumns.filter(function (c) { return c.visible; });
  var ccTotalW = 0;
  visibleCols.forEach(function (c) { ccTotalW += c.width; });
  gridEl.style.setProperty("--custom-cols-width", ccTotalW + "px");

  if (!showComments) gridEl.classList.add("hide-comments");
  else gridEl.classList.remove("hide-comments");

  // ── Year layout ────────────────────────────────────────────────
  const totalYears = years.length;
  const cellW = zoomMode === "days" ? 20 : zoomMode === "months" ? 80 : 28;
  const { offsets: yearWeekOffsets, totalWeeks } = getYearWeekOffsets();
  const { offsets: yearDayOffsets, totalDays } = getYearDayOffsets();
  const yearWeekCounts = years.map((yearLabel) => getIsoWeeksInYear(yearLabel));
  const yearDayCounts = yearWeekCounts.map((count) => count * 5);

  const collapsedYearW = 24; // fixed narrow width for any collapsed year
  const yearWidths = years.map((_, yi) =>
    hiddenYears.has(yi)
      ? collapsedYearW
      : (zoomMode === "days" ? yearDayCounts[yi] : zoomMode === "months" ? 12 : yearWeekCounts[yi]) * cellW
  );
  const yearOffsets = [];
  let _tlAcc = 0;
  yearWidths.forEach(w => { yearOffsets.push(_tlAcc); _tlAcc += w; });
  const totalTimelineW = _tlAcc;

  // Dynamic grid: comment | [custom cols...] | name | timeline
  var gridColDef = "var(--comment-width) ";
  visibleCols.forEach(function (c) { gridColDef += c.width + "px "; });
  gridColDef += "var(--name-width) " + totalTimelineW + "px";
  gridEl.style.gridTemplateColumns = gridColDef;

  // ── Month spans cache ──────────────────────────────────────────
  const _cacheKey = years.join(",");
  if (!_monthSpansCache || _monthSpansCacheKey !== _cacheKey) {
    _monthSpansCache = years.map((yearLabel) => getMonthWeekSpans(yearLabel));
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
  const _alarmExactMap = new Map();
  alarms.forEach(a => {
    const w = getAlarmAbsWeek(a);
    if (!_alarmWeekMap.has(w)) _alarmWeekMap.set(w, []);
    _alarmWeekMap.get(w).push(a);
    const exactW = getAlarmExactWeek(a);
    if (!_alarmExactMap.has(exactW)) _alarmExactMap.set(exactW, []);
    _alarmExactMap.get(exactW).push(a);
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
            if (relW >= 1 && relW <= yearWeekCounts[yi]) _holidayAbsWeeks.add(getAbsWeekFromYearWeek(yi, relW));
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
          if (diff >= 0 && diff < yearWeekCounts[yi] * 7) {
            const dow = diff % 7;
            if (dow < 5) {
              const col = yearDayOffsets[yi] + Math.floor(diff / 7) * 5 + dow + 1;
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
      _dayColMeta = new Array(totalDays);
      const _yearStarts = years.map((yr) => {
        const yearNum = parseInt(yr) || new Date().getFullYear();
        const jan4 = new Date(yearNum, 0, 4);
        const soy = new Date(jan4.getTime());
        soy.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1);
        return soy;
      });
      const _DAY_LETTERS = ["M", "T", "W", "T", "F"];
      let c = 0;
      for (let yearIndex = 0; yearIndex < totalYears; yearIndex++) {
        const soy = _yearStarts[yearIndex];
        for (let relativeWeek = 1; relativeWeek <= yearWeekCounts[yearIndex]; relativeWeek++) {
          const absWeek = getAbsWeekFromYearWeek(yearIndex, relativeWeek);
          for (let dayIndex = 0; dayIndex < 5; dayIndex++) {
            const d = new Date(soy.getTime());
            d.setDate(soy.getDate() + (relativeWeek - 1) * 7 + dayIndex);
            const dd = pad2(d.getDate()), mm = pad2(d.getMonth() + 1);
            const isWeekEnd = dayIndex === 4;
            const isMonthEnd = isWeekEnd && _monthEndWeekSets[yearIndex].has(relativeWeek);
            const exactWeekValue = absWeek + dayIndex * 0.2;
            _dayColMeta[c++] = {
              letter: _DAY_LETTERS[dayIndex],
              dateStr: dd,
              fullDateStr: d.getFullYear() + "-" + mm + "-" + dd,
              exactWeekValue,
              isWeekEnd,
              isMonthEnd,
            };
          }
        }
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
    const alarmMap = _itemAlarmMap.get(a.itemId);
    const exactW = getAlarmExactWeek(a);
    if (!alarmMap.has(exactW)) alarmMap.set(exactW, []);
    alarmMap.get(exactW).push(a);
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
    const info = getYearWeekInfo(baseW);
    const yi = info.yearIndex;
    const relW = info.relWeek - 1; // 0-based week within year
    if (hiddenYears.has(yi)) {
      // Map proportionally into the fixed collapsed year width
      const totalCols = zoomMode === "days" ? yearDayCounts[yi] : yearWeekCounts[yi];
      const dayI = zoomMode === "days" ? Math.max(0, Math.min(4.999, (sw - Math.floor(sw)) * 5)) : 0;
      const col = zoomMode === "days" ? relW * 5 + dayI : relW + (sw - Math.floor(sw));
      return yearOffsets[yi] + (col / totalCols) * collapsedYearW;
    }
    if (zoomMode === "days") {
      const dayI = Math.max(0, Math.min(4.999, (sw - Math.floor(sw)) * 5));
      return yearOffsets[yi] + (relW * 5 + dayI) * cellW;
    }
    // weeks mode: fractional offset within column
    const fracOffset = (sw - Math.floor(sw)) * cellW;
    return yearOffsets[yi] + relW * cellW + fracOffset;
  }

  // Convert absolute month number (1-based, can be fractional) to px offset
  function _monthX(mm) {
    const absM0 = Math.max(0, mm - 1); // 0-based
    const yi = Math.max(0, Math.min(Math.floor(absM0 / 12), years.length - 1));
    const relM = absM0 - yi * 12; // 0-based within year, can be fractional
    if (hiddenYears.has(yi)) {
      return yearOffsets[yi] + (relM / 12) * collapsedYearW;
    }
    return yearOffsets[yi] + relM * cellW;
  }

  // ── HTML GENERATION ───────────────────────────────────────────

  // Helper: compute sticky left for each visible custom column
  var _ccBaseLeft = showComments ? commentWidth : 0;
  function _ccStickyLeft(colIndex) {
    var left = _ccBaseLeft;
    for (var i = 0; i < colIndex; i++) left += visibleCols[i].width;
    return left;
  }

  // Helper: generate empty custom column cells for header rows
  function _ccHeaderCells(extraStyle, extraClass) {
    var h = "";
    for (var ci = 0; ci < visibleCols.length; ci++) {
      h += '<div class="cell header-cell custom-col-cell' + (extraClass || "") +
        '" style="left:' + _ccStickyLeft(ci) + 'px; z-index:28; position:sticky;' + (extraStyle || "") + '"></div>';
    }
    return h;
  }

  // ── Header Row 1: Year labels ──────────────────────────────────
  html += '<div class="row-bg header-row-1">';
  html += '<div class="cell header-cell comment-cell" style="border-bottom:none; z-index:30;"></div>';
  html += _ccHeaderCells("border-bottom:none;");
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
  // Today handle inline in year header row
  if (_todayLineVisible) {
    const _todayWkHdr = getTodayLineWeek();
    if (_todayWkHdr !== null) {
      const _todayRawX = zoomMode === "months"
        ? _monthX(mapWeekToMonth(_todayWkHdr))
        : _wkX(_todayWkHdr);
      const _todayHdrX = _todayRawX + (zoomMode === "days" ? cellW / 2 : 0);
      html += '<div id="today-line-handle" onmousedown="_startTodayLineDrag(event)" ondblclick="resetTodayLine()"' +
        ' title="Drag to move · Double-click to reset to today"' +
        ' style="position:absolute; left:' + _todayHdrX + 'px; top:0; transform:translateX(-50%);' +
        ' cursor:ew-resize; z-index:5; display:flex; flex-direction:column; align-items:center;">' +
        '<div style="background:#ef4444; color:#fff; font-size:12px; font-weight:800;' +
        ' padding:3px 10px; border-radius:4px; white-space:nowrap; pointer-events:none;' +
        ' line-height:1.3; letter-spacing:0.5px;">TODAY</div>' +
        '<div style="width:0; height:0; border-left:6px solid transparent;' +
        ' border-right:6px solid transparent; border-top:10px solid #ef4444;' +
        ' pointer-events:none;"></div></div>';
    }
  }
  html += '</div></div>';

  // ── Header Row 2: Month labels (not in months mode) ────────────
  if (zoomMode !== "months") {
    html += '<div class="row-bg header-row-2">';
    html += '<div class="cell header-cell comment-cell" style="border-bottom:none; z-index:30;"></div>';
    html += _ccHeaderCells("border-bottom:none;");
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
    html += _ccHeaderCells("border-bottom:none; height:24px;");
    html += '<div class="cell header-cell name-cell" style="border-bottom:none; z-index:30; height:24px;"></div>';
    html += '<div class="cell header-cell tl-header" style="padding:0; overflow:visible; height:24px;">';
    for (let yi = 0; yi < years.length; yi++) {
      const eff = cellW;
      if (hiddenYears.has(yi)) {
        html += '<div style="position:absolute; left:' + yearOffsets[yi] + 'px; width:' + yearWidths[yi] + 'px; height:100%; background:#e8edf5;"></div>';
        continue;
      }
      for (let relW = 1; relW <= yearWeekCounts[yi]; relW++) {
        const absWeek = getAbsWeekFromYearWeek(yi, relW);
        const x3 = yearOffsets[yi] + (relW - 1) * 5 * eff;
        const w3 = 5 * eff;
        const isHighlighted = isHighlightActive(absWeek, highlightedWeek);
        const _row3Alarms = _alarmWeekMap.get(absWeek) || [];
        const hasAlarm = _row3Alarms.length > 0;
        const row3AlarmTitle = _row3Alarms.map(a => a.title + " " + a.time).join(", ").replace(/"/g, '&quot;');
        const row3AlarmId = hasAlarm ? _row3Alarms[0].id : null;
        html += '<div style="position:absolute; left:' + x3 + 'px; width:' + w3 + 'px; height:100%; cursor:pointer; font-size:11px; font-weight:bold; color:#2563eb; display:flex; align-items:center; justify-content:center; border-right:2px solid #94a3b8; box-sizing:border-box;' +
          (isHighlighted ? 'background:#fef08a;' : '') +
          '" onclick="toggleHighlight(\'W\' + ' + absWeek + ')" ondblclick="openAlarmForWeek(' + absWeek + ')" title="Click to highlight week">W' +
          relW +
          (hasAlarm ? ' <button onclick="event.stopPropagation(); openAlarmEditor(' + row3AlarmId + ')" title="' + row3AlarmTitle + '" style="font-size:10px; background:none; border:none; cursor:pointer; padding:0; margin-left:2px;">🔔</button>' : '') +
          '</div>';
      }
    }
    html += '</div></div>';
  }

  // ── Header Row 4: Main header (Comments / Activity + column labels) ──
  html += '<div class="row-bg header-row-4" style="height:56px;">';
  html += '<div class="cell header-cell comment-cell" style="justify-content:flex-start; padding-left:16px; align-items:flex-end; padding-bottom:10px; z-index:30; font-size:13px; font-weight:bold; height:56px;">Comments</div>';

  // Custom column headers with editable name + filter icon
  for (var _cci = 0; _cci < visibleCols.length; _cci++) {
    var _cc = visibleCols[_cci];
    var _ccHasFilter = columnFilters[_cc.id] && columnFilters[_cc.id].length > 0;
    html += '<div class="cell header-cell custom-col-cell custom-col-header" style="left:' +
      _ccStickyLeft(_cci) + 'px; z-index:28; position:sticky; height:56px;">' +
      '<input type="text" value="' + (_cc.name || "").replace(/"/g, "&quot;") +
      '" onchange="renameCustomColumn(\'' + _cc.id + '\', this.value)" title="Edit column name">' +
      '<span class="cc-filter-icon' + (_ccHasFilter ? " active" : "") +
      '" onclick="toggleColFilterMenu(\'' + _cc.id + '\', event)" title="Filter this column">&#9660;</span>' +
      '</div>';
  }

  // Pre-compute which absolute weeks and months contain holidays
  html +=
    '<div class="cell header-cell name-cell" style="justify-content: space-between; align-items: flex-end; padding-left: 18px; padding-right: 8px; z-index: 30; height: 56px; border-bottom: 2px solid var(--border); box-sizing: border-box;">' +
    '<div style="display:flex; align-items:flex-end; width: 100%; padding-bottom: 10px; min-width:0;">' +
    '<span style="font-size: 13px; font-weight: bold; line-height: 1.05;">Family / Project / Milestone / Activity</span>' +
    "</div>" +
    '<button class="settings-btn' + (showSettings ? " active" : "") + '" onclick="toggleAllSettings()" title="Toggle Settings" style="align-self: flex-end; margin-bottom: 8px; font-size: 13px; gap: 4px;"><span style="font-size:18px;">&#9881;</span> Settings</button>' +
    "</div>";

  // Column header timeline cell — individual absolutely positioned labels
  html += '<div class="cell header-cell tl-header" style="padding:0; overflow:visible; height:56px;">';
  {
    let _x4 = 0;
    for (let yi = 0; yi < years.length; yi++) {
      const isHiddenYr = hiddenYears.has(yi);
      const eff4 = cellW;

      if (isHiddenYr) {
        html += '<div style="position:absolute; left:' + _x4 + 'px; width:' + yearWidths[yi] + 'px; height:100%; background:#e8edf5;"></div>';
        _x4 += yearWidths[yi];
        continue;
      }

      if (zoomMode === "weeks") {
        for (let relW = 1; relW <= yearWeekCounts[yi]; relW++) {
          const absWeek = getAbsWeekFromYearWeek(yi, relW);
          const isHighlighted = isHighlightActive(absWeek, highlightedWeek);
          const _wkAlarms = _alarmWeekMap.get(absWeek) || [];
          const hasAlarm = _wkAlarms.length > 0;
          const isMonthEnd = _monthEndWeekSets[yi].has(relW);
          const _hldBorder = _holidayAbsWeeks.has(absWeek) ? "border-bottom:3px solid #f59e0b;" : "";
          const weekAlarmTitle = _wkAlarms.map(a => a.title + " " + a.time).join(", ").replace(/"/g, '&quot;');
          const weekAlarmId = hasAlarm ? _wkAlarms[0].id : null;
          const label = String(relW) + (hasAlarm ? ' <button onclick="event.stopPropagation(); openAlarmEditor(' + weekAlarmId + ')" title="' + weekAlarmTitle + '" style="font-size:9px; background:none; border:none; cursor:pointer; padding:0; margin-left:2px;">🔔</button>' : '');
          const extraStyle = (isHighlighted ? "background:#fef08a;" : "") + (isMonthEnd ? "border-right:4px solid #94a3b8;" : "border-right:1px solid #e2e8f0;") + _hldBorder;
          html += '<div onclick="toggleHighlight(\'W\' + ' + absWeek + ')" ondblclick="openAlarmForWeek(' + absWeek + ')" title="Click to highlight week" style="position:absolute; left:' + _x4 + 'px; width:' + eff4 + 'px; height:100%; display:flex; align-items:center; justify-content:center; cursor:pointer; font-size:12px; font-weight:600; color:#475569; box-sizing:border-box; ' + extraStyle + '">' + label + '</div>';
          _x4 += eff4;
        }
      } else if (zoomMode === "days") {
        for (let relC = 0; relC < yearDayCounts[yi]; relC++) {
          const absCol = yearDayOffsets[yi] + relC + 1;
          const _m = _dayColMeta[absCol - 1];
          const isHighlighted = isHighlightActive(_m.exactWeekValue, highlightedWeek);
          const isHoliday = _holidayDayCols.has(absCol);
          const _dayAlarms = _alarmExactMap.get(_m.exactWeekValue) || [];
          const hasAlarm = _dayAlarms.length > 0;
          const dayAlarmTitle = _dayAlarms.map(a => a.title + " " + a.time).join(", ").replace(/"/g, '&quot;');
          const dayAlarmId = hasAlarm ? _dayAlarms[0].id : null;
          const borderStyle = _m.isMonthEnd ? "border-right:4px solid #94a3b8;" : _m.isWeekEnd ? "border-right:2px solid #94a3b8;" : "border-right:1px dashed #e2e8f0;";
          const label = '<div style="display:flex;flex-direction:column;align-items:center;line-height:1.2;"><span>' + _m.letter + (hasAlarm ? ' <button onclick="event.stopPropagation(); openAlarmEditor(' + dayAlarmId + ')" title="' + dayAlarmTitle + '" style="font-size:9px; background:none; border:none; cursor:pointer; padding:0; margin-left:2px;">🔔</button>' : '') + '</span><span style="font-size:9px;color:#64748b;">' + _m.dateStr + '</span></div>';
          const extraStyle = (isHighlighted ? "background:#fef08a;" : "") + (isHoliday ? "background-color:#b0bec5;" : "") + borderStyle;
          html += '<div onclick="toggleHighlight(' + _m.exactWeekValue + ')" ondblclick="openAlarmForDate(\'' + _m.fullDateStr + '\')" title="Click to highlight day" style="position:absolute; left:' + _x4 + 'px; width:' + eff4 + 'px; height:100%; display:flex; align-items:center; justify-content:center; cursor:pointer; font-size:12px; font-weight:600; color:#475569; box-sizing:border-box; ' + extraStyle + '">' + label + '</div>';
          _x4 += eff4;
        }
      } else { // months
        for (let mi = 0; mi < 12; mi++) {
          const monthCol = yi * 12 + mi + 1;
          const monthName = dynamicMonthSpans[yi][mi].name;
          const _mYN = parseInt(years[yi]);
          const _mKey = _mYN + "-" + String(mi + 1).padStart(2, "0");
          const _mHldBorder = _holidayMonthKeys.has(_mKey) ? "border-bottom:3px solid #f59e0b;" : "";
          const isHighlighted = isHighlightActive(monthCol, highlightedWeek);
          const label = '<div style="font-size:11px;font-weight:bold;color:#64748b;text-align:center;">' + monthName + '</div>';
          const extraStyle = (isHighlighted ? "background:#fef08a;" : "") + "border-right:1px solid #e2e8f0;" + _mHldBorder;
          html += '<div onclick="toggleHighlight(' + monthCol + ')" title="Click to highlight month\n' + monthName + '" style="position:absolute; left:' + _x4 + 'px; width:' + eff4 + 'px; height:100%; display:flex; align-items:center; justify-content:center; cursor:pointer; font-size:12px; font-weight:600; color:#475569; box-sizing:border-box; ' + extraStyle + '">' + label + '</div>';
          _x4 += eff4;
        }
      }
    }
  }
  html += '</div></div>'; // end tl-header + header-row-4

  // In day view, milestones with no explicit day (integer startWeek) should
  // appear on Wednesday (mid-week) rather than Monday. 0.4 fractional = day 2
  // (Wed) in the 5-day-per-week layout.
  function _msWeek(sw) {
    return (zoomMode === "days" && sw % 1 === 0) ? sw + 0.4 : sw;
  }

  // ── Data rows ──────────────────────────────────────────────────
  displayItems.forEach((item, index) => {
    const level = item.level || 0;
    const margin = level * 20;
    const children = getChildren(item.id);
    const hasChildren = children.length > 0;
    const rowHighlightClass = highlightedRowId === item.id ? " row-highlighted" : "";

    let computedStartWeek = item.startWeek;
    let computedDuration = item.duration;

    if (item.type === "task" && holidays.length > 0) {
      computedDuration = computeEffectiveDuration(
        computedStartWeek, computedDuration, item.assignees || []
      );
    }

    // For parent items, compute bar bounds in pixel-space so the end
    // aligns exactly with the milestone diamond tip in every zoom mode.
    // Also track week-space bounds for label visibility and duration display.
    let _parentBarLeftPx = null, _parentBarRightPx = null;
    if (hasChildren) {
      let minStartPx = Infinity, maxEndPx = -Infinity;
      let minStartWk = Infinity, maxEndWk = -Infinity;
      function traverse(parentId) {
        const kids = getChildren(parentId);
        kids.forEach((k) => {
          const kSw = k.type === "milestone" && (k.duration || 0) === 0 ? _msWeek(k.startWeek) : k.startWeek;
          const kStartPx = zoomMode === "months"
            ? _monthX(mapWeekToMonth(kSw))
            : _wkX(kSw);
          if (kStartPx < minStartPx) minStartPx = kStartPx;
          if (k.startWeek < minStartWk) minStartWk = k.startWeek;

          let kDur = k.type === "task" && holidays.length > 0
            ? computeEffectiveDuration(k.startWeek, k.duration || 0, k.assignees || [])
            : k.duration || 0;

          let kEndPx;
          if (k.type === "milestone" && kDur === 0) {
            // End at the diamond center (half a cell past start)
            const _kYi = getYearWeekInfo(Math.max(1, Math.floor(k.startWeek))).yearIndex;
            const kEff = hiddenYears.has(_kYi) ? collapsedYearW / (zoomMode === "days" ? yearDayCounts[_kYi] : zoomMode === "months" ? 12 : yearWeekCounts[_kYi]) : cellW;
            kEndPx = kStartPx + kEff / 2;
            kDur = 0.5; // week-space approximation for label/duration display
          } else {
            kEndPx = zoomMode === "months"
              ? _monthX(mapWeekToMonth(k.startWeek + kDur))
              : _wkX(k.startWeek + kDur);
          }
          const kEndWk = k.startWeek + kDur;
          if (kEndPx > maxEndPx) maxEndPx = kEndPx;
          if (kEndWk > maxEndWk) maxEndWk = kEndWk;
          traverse(k.id);
        });
      }
      traverse(item.id);
      if (minStartPx !== Infinity) {
        _parentBarLeftPx = minStartPx;
        _parentBarRightPx = maxEndPx;
        computedStartWeek = minStartWk;
        computedDuration = Math.max(0, maxEndWk - minStartWk);
      }
    }

    html += '<div class="row-bg' + (selectedIds.has(item.id) ? ' row-selected' : '') + '" data-id="' + item.id + '">';

    let peerLabel = "";
    if (item._groupPeers && item._groupPeers.length > 0) {
      peerLabel =
        '<div style="font-size:10px;color:#b45309;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.2;" title="' +
        item._groupPeers.map((p) => p.name).join(", ").replace(/"/g, "&quot;") +
        '">+' + item._groupPeers.map((p) => p.name).join(", ") + "</div>";
    }
    html +=
      '<div class="cell comment-cell' + rowHighlightClass + '">' + peerLabel +
      '<input class="comment-input' + (getCommentLinks(item.comment || "").length ? " has-link" : "") +
      '" type="text" value="' + (item.comment || "").replace(/"/g, "&quot;") +
      '" onchange="updateComment(' + item.id + ', this.value)" onclick="handleCommentLinkClick(event)" ' +
      'onfocus="this.classList.remove(\'has-link\')" onblur="updateComment(' + item.id + ', this.value); if(getCommentLinks(this.value).length)this.classList.add(\'has-link\')" ' +
      'placeholder="Add comment...">' +
      '<div class="comment-link-wrap" title="Double-click to edit" ondblclick="editCommentLink(this)">' +
      getCommentLinksHtml(item.comment || "", true) +
      '</div></div>';

    // Custom column data cells
    for (var _dci = 0; _dci < visibleCols.length; _dci++) {
      var _dc = visibleCols[_dci];
      var _dcVal = (item.customData && item.customData[_dc.id]) || "";
      html += '<div class="cell custom-col-cell' + rowHighlightClass +
        '" style="left:' + _ccStickyLeft(_dci) + 'px;">' +
        '<input type="text" value="' + _dcVal.replace(/"/g, "&quot;") +
        '" onchange="updateCustomColumnCell(' + item.id + ', \'' + _dc.id + '\', this.value)"' +
        ' placeholder="...">' +
        '</div>';
    }

    html +=
      '<div class="cell name-cell' + rowHighlightClass + '" ondragover="rowDragOver(event)" ondragleave="rowDragLeave(event)" ondrop="rowDrop(event, ' + item.id + ')">' +
      '<input type="checkbox" class="row-select-cb" data-id="' + item.id + '"' + (selectedIds.has(item.id) ? ' checked' : '') + ' onclick="toggleItemSelection(' + item.id + ', event)" title="Select row">' +
      '<div style="display: flex; align-items: center; margin-left: ' + margin + 'px; flex-grow: 1; min-width: 0; overflow: hidden; margin-right: 8px;">' +
      '<div class="drag-handle" draggable="true" ondragstart="startRowDrag(event, ' + item.id + ')" title="Drag to reorder ▪ Alt+Drop onto row to nest inside ▪ Ctrl+Drop to make top-level">&#8942;&#8942;</div>';

    if (item.type === "task" || item.type === "project" || item.type === "family" || item.type === "milestones-group" || item.type === "milestone") {
      const toggleIcon = item.isExpanded === false ? "&#9654;" : "&#9660;";
      const visibility = hasChildren ? "visible" : "hidden";
      html += '<button class="toggle-btn" onclick="toggleExpand(' + item.id + ')" style="visibility: ' + visibility + ';">' + toggleIcon + "</button>";
    } else {
      html += '<div style="width: 32px;"></div>';
    }

    const nameClass =
      item.type === "family" ? "name-input name-input-family" :
      item.type === "project" ? "name-input name-input-project" :
      "name-input";

    if (item.type === "project" || item.type === "family") {
      if (item.name === "New Project") {
        html += '<select class="' + nameClass + '" onchange="updateProjectName(' + item.id + ', this.value)">';
        html += '<option value="New Project" disabled ' + (item.name === "New Project" ? "selected" : "") + ">Select Project Type...</option>";
        ["Child Project", "Brother Project", "Mother Project"].forEach((m) => {
          html += '<option value="' + m + '" ' + (m === item.name ? "selected" : "") + ">" + m + "</option>";
        });
        html += '<option value="Custom Project">Custom Project...</option>';
        html += "</select>";
      } else {
        html += '<input class="' + nameClass + '" type="text" value="' + item.name.replace(/"/g, "&quot;") + '" onchange="updateName(' + item.id + ', this.value)" onblur="updateName(' + item.id + ', this.value)" placeholder="Project name...">';
      }
    } else if (item.type === "milestones-group") {
      html += '<span class="name-input" style="font-style:italic;color:#b45309;font-weight:600;cursor:default;">Milestones</span>';
    } else if (item.type === "milestone") {
      const isPredefined = predefinedMilestones.includes(item.name);
      if (isPredefined || item.name === "New Milestone") {
        html += '<select class="' + nameClass + '" onchange="updateMilestoneName(' + item.id + ', this.value)">';
        html += '<option value="New Milestone" disabled ' + (item.name === "New Milestone" ? "selected" : "") + ">Select Milestone...</option>";
        predefinedMilestones.forEach((m) => {
          html += '<option value="' + m + '" ' + (m === item.name ? "selected" : "") + ">" + m + "</option>";
        });
        html += '<option value="Custom Milestone">Custom Milestone...</option>';
        html += "</select>";
      } else {
        html += '<input class="' + nameClass + '" type="text" value="' + item.name.replace(/"/g, "&quot;") + '" onchange="updateName(' + item.id + ', this.value)" onblur="updateName(' + item.id + ', this.value)" placeholder="Milestone name...">';
      }
      const weekDisplay = formatYYWWD(item.startWeek);
      html += '<input type="text" maxlength="6" value="' + weekDisplay + '" onchange="updateMilestoneDate(' + item.id + ', this.value)" title="Type YYWW or YYWW.D" style="width:52px;border:1px solid #cbd5e1;border-radius:4px;padding:2px 4px;font-size:13px;font-family:monospace;color:#92400e;background:#fffbeb;text-align:center;flex-shrink:0;">';
    } else {
      if ((item.name === "New Activity" || item.name === "New Sub-activity" || /^Activity \d+$/.test(item.name)) && customTemplates.length > 0) {
        html += '<select class="' + nameClass + '" onchange="applyActivityTemplate(' + item.id + ', this.value); if(!(customTemplates.find(t=>t.id===this.value)||{}).askName && !(customTemplates.find(t=>t.id===this.value)||{}).askDuration && !(customTemplates.find(t=>t.id===this.value)||{}).askColor) this.value=\'\';">';
        html += '<option value="" disabled selected>' + item.name + " (Select Template...)</option>";
        // Group by category using <optgroup>
        (function() {
          const allCats = (typeof templateCategories !== "undefined" ? templateCategories : []);
          const groups = [];
          allCats.forEach(function(cat) {
            const catTemplates = customTemplates.filter(function(t) { return t.category === cat.id && !(t.composition && t.composition.length > 0); });
            if (catTemplates.length) groups.push({ label: cat.name, templates: catTemplates });
          });
          const uncategorized = customTemplates.filter(function(t) {
            return (!(t.composition && t.composition.length > 0)) && (!t.category || !allCats.find(function(c) { return c.id === t.category; }));
          });
          const masters = customTemplates.filter(function(t) { return t.composition && t.composition.length > 0; });
          if (uncategorized.length) groups.push({ label: "General", templates: uncategorized });
          if (masters.length) groups.push({ label: "Master Templates", templates: masters });
          groups.forEach(function(g) {
            html += '<optgroup label="' + g.label.replace(/"/g, "&quot;") + '">';
            g.templates.forEach(function(t) {
              const durDays = (t.duration * 5).toFixed(1);
              const paramHint = (t.askName || t.askDuration || t.askColor) ? " ✦" : "";
              html += '<option value="' + t.id + '">' + t.name.replace(/"/g, "&quot;") + " (" + durDays + " days)" + paramHint + "</option>";
            });
            html += '</optgroup>';
          });
        })();
        html += '<option value="Custom Activity">Custom Activity...</option>';
        html += "</select>";
      } else {
        html += '<input class="' + nameClass + '" type="text" value="' + item.name.replace(/"/g, "&quot;") + '" onchange="updateName(' + item.id + ', this.value)" onblur="updateName(' + item.id + ', this.value)" placeholder="Activity name...">';
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
        // date input already shown inline next to the name
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
    html += '<div class="cell timeline-row' + rowHighlightClass + '" onclick="handleTimelineRowClick(event,' + item.id + ')" ondblclick="handleTlDblClick(event,this,' + item.id + ')" style="position:relative; overflow:visible; height:48px; border-bottom:1px solid var(--border); box-sizing:border-box;">';

    // ── Alarm indicators ──────────────────────────────────────────
    const itemAlarmsByWeek = _itemAlarmMap.get(item.id);
    if (itemAlarmsByWeek) {
      itemAlarmsByWeek.forEach((alarmList, exactW) => {
        const ax = _wkX(exactW);
        const alarmTitles = alarmList.map(a => a.title + " " + a.time).join(", ").replace(/"/g, '&quot;');
        const alarmId = alarmList[0].id;
        html += '<button onclick="event.stopPropagation(); openAlarmEditor(' + alarmId + ')" title="' + alarmTitles + '" style="position:absolute; left:' + (ax + 2) + 'px; top:2px; font-size:10px; z-index:11; background:none; border:none; cursor:pointer; padding:0;">🔔</button>';
      });
    }

    // ── Bar / Milestone rendering ─────────────────────────────────
    if (item.type === "task" || item.type === "project" || item.type === "family") {
      let barLeft, barWidth;
      if (_parentBarLeftPx !== null) {
        // Parent items: use pixel-space bounds computed from children
        barLeft = _parentBarLeftPx;
        barWidth = Math.max(cellW / (zoomMode === "months" ? 4 : 2), _parentBarRightPx - _parentBarLeftPx);
      } else if (zoomMode === "months") {
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
          '<div id="block-' + item.id + '" class="range-bar' + (selectedIds.has(item.id) ? ' block-selected' : '') + '" draggable="false" ondragstart="return false" style="position:absolute; left:' + barLeft + 'px; width:' + finalWidth + 'px; margin-left:0; border-color:' + bgColor + '; background:' + bgColor + '18;" onmousedown="startDrag(event,' + item.id + ')" ondblclick="event.stopPropagation()">' +
          '<div class="range-label" style="color:' + bgColor + '; display:flex; align-items:center; justify-content:center; font-weight:700; letter-spacing:0.01em;">' + label + assigneesHtml + "</div>" +
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
          '<div id="block-' + item.id + '" class="bar' + (selectedIds.has(item.id) ? ' block-selected' : '') + '" draggable="false" ondragstart="return false" style="position:absolute; left:' + barLeft + 'px; width:' + finalWidth + 'px; margin-left:0; background-color:' + bgColor + ';" onmousedown="startDrag(event,' + item.id + ')" ondblclick="event.stopPropagation()">' +
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
        const _msSw = _msWeek(item.startWeek);
        const msXRaw = zoomMode === "months" ? _monthX(mapWeekToMonth(_msSw)) : _wkX(_msSw);
        const _msYi = getYearWeekInfo(item.startWeek).yearIndex;
        const msEff = hiddenYears.has(_msYi) ? collapsedYearW / (zoomMode === "months" ? 12 : yearWeekCounts[_msYi]) : cellW;
        const msX = msXRaw + msEff / 2;
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
          '<div id="block-' + item.id + '" class="milestone' + (selectedIds.has(item.id) ? ' block-selected' : '') + '" draggable="false" ondragstart="return false" style="background-color:' + bgColor + '; pointer-events:auto;" onmousedown="startDrag(event,' + item.id + ')" ondblclick="event.stopPropagation()" title="' + item.name.replace(/"/g, "&quot;") + '">' +
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
        _tailHeight = (_lastDescIdx - index) * 48 + 24;
      }

      const descMilestones = getDescendantMilestones(item.id);
      descMilestones.forEach((ms) => {
        const _msSw2 = _msWeek(ms.startWeek);
        const msXRaw = zoomMode === "months" ? _monthX(mapWeekToMonth(_msSw2)) : _wkX(_msSw2);
        const _msYi2 = getYearWeekInfo(ms.startWeek).yearIndex;
        const msEff = hiddenYears.has(_msYi2) ? collapsedYearW / (zoomMode === "months" ? 12 : yearWeekCounts[_msYi2]) : cellW;
        const msX = msXRaw + msEff / 2;
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
  for (var _bci = 0; _bci < visibleCols.length; _bci++) {
    html += '<div class="cell custom-col-cell" style="border-bottom:none; left:' + _ccStickyLeft(_bci) + 'px;"></div>';
  }
  html += '<div class="cell name-cell" style="border-bottom:none; color:#cbd5e1; font-style:italic; justify-content:center;" ondragover="rowDragOverBottom(event)" ondragleave="rowDragLeaveBottom(event)" ondrop="rowDrop(event, \'bottom\')"></div>';
  html += '<div class="cell timeline-row" ondblclick="handleTlDblClick(event,this,null)" style="border-bottom:none; height:48px; position:relative; overflow:visible;"></div>';
  html += '</div>';

  // ── Filter menus ───────────────────────────────────────────────
  if (activeFilterMenu) {
    var _isColFilter = activeFilterMenu.indexOf("col_") === 0;
    let menuHtml =
      '<div class="filter-menu open" style="top: ' + filterMenuPosition.top + 'px; left: ' +
      filterMenuPosition.left + 'px;" onclick="event.stopPropagation()">';

    if (_isColFilter) {
      var _fCol = customColumns.find(function (c) { return c.id === activeFilterMenu; });
      var _fColName = _fCol ? _fCol.name : "Column";
      menuHtml += "<strong>Filter: " + _fColName.replace(/</g, "&lt;") + "</strong><br><br>";
      var _cfVals = columnFilters[activeFilterMenu] || [];
      var _allColVals = [];
      items.forEach(function (it) {
        var v = (it.customData && it.customData[activeFilterMenu]) || "";
        if (v && _allColVals.indexOf(v) === -1) _allColVals.push(v);
      });
      _allColVals.sort();
      if (_allColVals.length === 0) {
        menuHtml += '<div style="color:#94a3b8; font-size:12px;">No values in this column yet</div>';
      } else {
        _allColVals.forEach(function (v) {
          var checked = _cfVals.indexOf(v) === -1 ? "checked" : "";
          var safeVal = v.replace(/'/g, "\\'").replace(/"/g, "&quot;");
          menuHtml += '<label><input type="checkbox" ' + checked +
            " onclick=\"handleColFilterChange('" + activeFilterMenu.replace(/'/g, "\\'") + "', '" + safeVal + "', event)\"> " +
            v.replace(/</g, "&lt;") + "</label>";
        });
      }
      menuHtml += '<div class="filter-menu-actions">';
      menuHtml += '<button class="success" style="padding: 2px 8px; font-size:11px;" onclick="clearColFilter(\'' + activeFilterMenu.replace(/'/g, "\\'") + '\')">Clear All</button>';
      menuHtml += '<button class="outline" style="padding: 2px 8px; font-size:11px;" onclick="toggleFilterMenu(null, event)">Close</button>';
      menuHtml += "</div></div>";
    } else {
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
    }
    html += menuHtml;
  }

  // ── Write DOM ─────────────────────────────────────────────────
  gridEl.innerHTML = html;
  if (activeFilterMenu === "name") {
    const searchInput = gridEl.querySelector(".filter-menu-search");
    if (searchInput) searchInput.focus();
  }

  // ── Draw background grid canvas ───────────────────────────────
  function drawBgCanvas() {
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
    let _bgX = 0;
    for (let yi = 0; yi < years.length; yi++) {
      const isHiddenYr = hiddenYears.has(yi);
      const eff = cellW;
      const yearColCount = zoomMode === "weeks" ? yearWeekCounts[yi] : zoomMode === "days" ? yearDayCounts[yi] : 12;

      if (isHiddenYr) {
        ctx.fillStyle = "#f1f4f9";
        ctx.fillRect(_bgX, 0, yearWidths[yi], bgH);
        ctx.strokeStyle = "#94a3b8";
        ctx.lineWidth = 2;
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(_bgX + yearWidths[yi] - 1, 0);
        ctx.lineTo(_bgX + yearWidths[yi] - 1, bgH);
        ctx.stroke();
        _bgX += yearWidths[yi];
        continue;
      }

      for (let relC = 0; relC < yearColCount; relC++) {
        const x = _bgX + eff;

        let isHL = false;
        let isMonthEnd = false;
        let isWeekEnd = false;

        if (zoomMode === "weeks") {
          const absWeek = getAbsWeekFromYearWeek(yi, relC + 1);
          if (_holidayAbsWeeks.has(absWeek)) {
            ctx.fillStyle = "rgba(245,158,11,0.12)";
            ctx.fillRect(_bgX, headerH, eff, bgH - headerH);
          }
          isHL = isHighlightActive(absWeek, highlightedWeek);
          isMonthEnd = _monthEndWeekSets[yi].has(relC + 1);
        } else if (zoomMode === "days") {
          const absCol = yearDayOffsets[yi] + relC + 1;
          const _m = _dayColMeta ? _dayColMeta[absCol - 1] : null;
          if (_generalHolidayDayCols.has(absCol)) {
            ctx.fillStyle = "rgba(176,190,197,0.35)";
            ctx.fillRect(_bgX, headerH, eff, bgH - headerH);
          }
          isHL = _m ? isHighlightActive(_m.exactWeekValue, highlightedWeek) : false;
          isWeekEnd = (relC + 1) % 5 === 0;
          isMonthEnd = isWeekEnd && _m && _m.isMonthEnd;
        } else {
          const monthCol = yi * 12 + relC + 1;
          isHL = isHighlightActive(monthCol, highlightedWeek);
        }

        if (isHL) {
          ctx.fillStyle = "rgba(254,240,138,0.4)";
          ctx.fillRect(_bgX, 0, eff, bgH);
        }

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

    // ── Today line on canvas ──────────────────────────────────────
    if (_todayLineVisible) {
      const todayWk = getTodayLineWeek();
      if (todayWk !== null) {
        // Center the line in the middle of the current cell
        const rawX = zoomMode === "months"
          ? _monthX(mapWeekToMonth(todayWk))
          : _wkX(todayWk);
        const todayX = rawX + (zoomMode === "days" ? cellW / 2 : 0);
        // Red vertical line
        ctx.save();
        ctx.strokeStyle = "#ef4444";
        ctx.lineWidth = 2.5;
        ctx.setLineDash([]);
        ctx.beginPath();
        ctx.moveTo(todayX, 0);
        ctx.lineTo(todayX, bgH);
        ctx.stroke();
        ctx.restore();
      }
    }
  }
  drawBgCanvas();
  _drawCanvas = drawBgCanvas;
  syncStickyColumnResizers();

  // ── Inject col-resizers ────────────────────────────────────────
  if (showComments) {
    const commentResizer = document.createElement("div");
    commentResizer.className = "col-resizer";
    commentResizer.style.left = "calc(var(--comment-width) - 2px)";
    commentResizer.addEventListener("mousedown", (e) => startColResize(e, "comment"));
    gridEl.appendChild(commentResizer);
  }
  // Custom column resizers
  for (var _cri = 0; _cri < visibleCols.length; _cri++) {
    (function(ci, col) {
      var resizerLeft = _ccStickyLeft(ci) + col.width - 2;
      var resizer = document.createElement("div");
      resizer.className = "col-resizer";
      resizer.style.left = resizerLeft + "px";
      resizer.addEventListener("mousedown", function(e) { startColResize(e, "customcol_" + col.id); });
      gridEl.appendChild(resizer);
    })(_cri, visibleCols[_cri]);
  }
  const nameResizer = document.createElement("div");
  nameResizer.className = "col-resizer";
  nameResizer.style.left = "calc(var(--comment-width) + var(--custom-cols-width, 0px) + var(--name-width) - 2px)";
  nameResizer.addEventListener("mousedown", (e) => startColResize(e, "name"));
  gridEl.appendChild(nameResizer);

  // ── Inject Link Anchor Dots ────────────────────────────────────
  if (linkMode) {
    displayItems.forEach((item) => {
      const block = document.getElementById("block-" + item.id);
      if (!block) return;
      if (block.style.position !== "absolute") block.style.position = "relative";
      const milestoneAnchorClass = item.type === "milestone" ? " milestone-link-anchor" : "";
      const isActiveFrom = linkSource && linkSource.itemId === item.id && linkSource.anchor === "start";
      const isActiveTo = linkSource && linkSource.itemId === item.id && linkSource.anchor === "end";
      const startDot = document.createElement("div");
      startDot.className = "link-anchor start-anchor" + milestoneAnchorClass + (isActiveFrom ? " active" : "");
      startDot.title = 'Link: start of "' + item.name + '"';
      startDot.onclick = (e) => { e.stopPropagation(); clickAnchor(item.id, "start"); };
      const endDot = document.createElement("div");
      endDot.className = "link-anchor end-anchor" + milestoneAnchorClass + (isActiveTo ? " active" : "");
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
        ${getCommentLinksHtml(item.blockComment.text || "", false)}
        <div class="comment-resizer" onmousedown="startCommentResize(event, ${item.id})"></div>
      `;
      gridEl.appendChild(popup);
      updatePopupPointer(item, popup);
    }
  });
}
