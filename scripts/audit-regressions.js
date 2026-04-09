const fs = require("fs");
const path = require("path");
const vm = require("vm");

function makeElement(id = "") {
  return {
    id,
    style: {},
    innerHTML: "",
    textContent: "",
    value: "",
    disabled: false,
    className: "",
    classList: {
      add() {},
      remove() {},
      contains() { return false; },
      toggle() {},
    },
    appendChild() {},
    remove() {},
    addEventListener() {},
    removeEventListener() {},
    querySelector() { return null; },
    querySelectorAll() { return []; },
    closest() { return null; },
    getBoundingClientRect() {
      return { left: 0, right: 0, top: 0, bottom: 0, width: 0, height: 0 };
    },
    setAttribute() {},
    focus() {},
    select() {},
    click() {},
    offsetWidth: 0,
    offsetHeight: 0,
    offsetLeft: 0,
    offsetTop: 0,
    parentElement: null,
    offsetParent: null,
  };
}

function createHarness() {
  const elementCache = new Map();
  const alerts = [];
  const localStore = new Map();

  const document = {
    body: makeElement("body"),
    getElementById(id) {
      if (!elementCache.has(id)) elementCache.set(id, makeElement(id));
      return elementCache.get(id);
    },
    querySelector() { return null; },
    querySelectorAll() { return []; },
    createElement() { return makeElement(); },
    createElementNS() { return makeElement(); },
    addEventListener() {},
    removeEventListener() {},
  };

  class FakeFileReader {
    readAsText(file) {
      if (this.onload) this.onload({ target: { result: file.__text } });
    }
  }

  const context = {
    console,
    document,
    window: {
      addEventListener() {},
      removeEventListener() {},
      open() {},
      print() {},
      prompt() { return "QA"; },
      showSaveFilePicker: null,
    },
    localStorage: {
      getItem(key) { return localStore.has(key) ? localStore.get(key) : null; },
      setItem(key, value) { localStore.set(key, String(value)); },
      removeItem(key) { localStore.delete(key); },
    },
    indexedDB: {
      open() {
        throw new Error("indexedDB is not available in the audit harness");
      },
    },
    FileReader: FakeFileReader,
    alert(message) { alerts.push(message); },
    confirm() { return true; },
    requestAnimationFrame(fn) { fn(); return 1; },
    cancelAnimationFrame() {},
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    Blob,
    URL: {
      createObjectURL() { return "blob:test"; },
      revokeObjectURL() {},
    },
    fetch() {
      return Promise.resolve({ ok: false, text: async () => "" });
    },
    Date,
    Math,
    JSON,
    Map,
    Set,
    parseInt,
    parseFloat,
    isNaN,
    Array,
    String,
    Number,
    RegExp,
  };
  context.getComputedStyle = (el) => el && el.style ? el.style : {};

  context.global = context;
  vm.createContext(context);

  const root = path.resolve(__dirname, "..");
  [
    "src/js/state.js",
    "src/js/links.js",
    "src/js/items.js",
    "src/js/interactions.js",
    "src/js/alarms.js",
    "src/js/csv.js",
    "src/js/render.js",
    "src/js/ui.js",
    "src/js/autosave.js",
  ].forEach((file) => {
    const source = fs.readFileSync(path.join(root, file), "utf8");
    vm.runInContext(source, context, { filename: file });
  });

  vm.runInContext(
    "render = function(){}; drawLinks = function(){}; _updateYearCountLabel = function(){};",
    context,
  );

  return {
    alerts,
    document,
    run(code) {
      return vm.runInContext(code, context);
    },
    set(name, value) {
      context[name] = value;
    },
  };
}

const results = [];

function pass(name) {
  results.push({ name, ok: true });
}

function fail(name, details) {
  results.push({ name, ok: false, details });
}

function assertEqual(actual, expected, name) {
  if (JSON.stringify(actual) === JSON.stringify(expected)) pass(name);
  else fail(name, `expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
}

function assertOk(condition, name, details) {
  if (condition) pass(name);
  else fail(name, details);
}

const h = createHarness();

h.run(`
years = ["2025"];
holidays = [{ date: "2025-01-06", people: [] }];
`);
assertEqual(h.run("computeEffectiveDuration(1, 0.1, [])"), 0.1, "Half-day durations stay half-day");
assertEqual(h.run("computeEffectiveDuration(1, 0.5, [])"), 0.5, "Fractional durations keep their precision");
assertEqual(h.run("computeEffectiveDuration(1.8, 0.4, [])"), 0.6, "Holiday overlap extends by one working day");

h.run(`
years = ["2025", "2026", "2027", "2028"];
items = [
  { id: 1, type: "milestone", name: "M1", startWeek: 5, duration: 0 },
  { id: 2, type: "task", name: "T2", startWeek: 7, duration: 2 }
];
links = [{ id: 1, fromId: 1, fromAnchor: "end", toId: 2, toAnchor: "start" }];
`);
h.run(`updateMilestoneDate(1, "2508");`);
assertEqual(
  h.run("items.map(i => ({ id: i.id, startWeek: i.startWeek }))"),
  [{ id: 1, startWeek: 8 }, { id: 2, startWeek: 10 }],
  "Typed milestone moves propagate through dependency links",
);

h.run(`
highlightedWeek = "W12";
items = [{ id: 1, type: "task", name: "A", startWeek: 1, duration: 1 }];
alarms = [];
holidays = [];
people = [];
links = [];
customTemplates = [];
filters = { name: [], type: [] };
zoomMode = "weeks";
showSettings = false;
showComments = false;
commentWidth = 200;
nameWidthBase = 350;
showLinks = true;
`);
const weekCsv = h.run("generateCSVString()");
h.set("__csv", weekCsv);
h.run(`
highlightedWeek = null;
items = [];
alarms = [];
holidays = [];
people = [];
links = [];
_importCSVText(__csv);
`);
assertEqual(h.run("highlightedWeek"), "W12", "Week highlights survive CSV round-trip");

h.run(`
highlightedWeek = 12.2;
items = [{ id: 1, type: "task", name: "A", startWeek: 1, duration: 1 }];
`);
const dayCsv = h.run("generateCSVString()");
h.set("__csvDay", dayCsv);
h.run(`
highlightedWeek = null;
items = [];
alarms = [];
holidays = [];
people = [];
links = [];
_importCSVText(__csvDay);
`);
assertEqual(h.run("highlightedWeek"), 12.2, "Day highlights survive CSV round-trip");

h.run(`
years = ["2025", "2026", "2027"];
items = [
  { id: 1, type: "task", name: "A", startWeek: 1, duration: 1 },
  { id: 2, type: "task", name: "B", startWeek: 60, duration: 1 }
];
`);
h.run(`updateYear(1, "2028");`);
assertEqual(
  h.run("({ years: years.slice(), starts: items.map(i => i.startWeek) })"),
  { years: ["2027", "2028", "2029"], starts: [1, 60] },
  "Editing a year re-anchors labels without moving work",
);

h.run(`
years = ["2025", "2026", "2027", "2028"];
holidays = [];
filters = { name: [], type: [] };
items = [
  { id: 1, type: "task", name: "Late", startWeek: 80, duration: 3 },
  { id: 2, type: "task", name: "Early", startWeek: 1, duration: 1 }
];
`);
assertEqual(h.run("getLastItemEndWeek()"), 83, "New top-level items start after the furthest work, not the last row");

h.run(`
_getSystemTodayWeek = function() { return 67.4; };
years = ["2025", "2026"];
items = [
  { id: 10, type: "project", name: "Parent", startWeek: 1, duration: 2 },
  { id: 11, type: "task", name: "Child Late", parentId: 10, startWeek: 20, duration: 1 },
  { id: 12, type: "task", name: "Child Early", parentId: 10, startWeek: 5, duration: 1 }
];
nextId = 100;
`);
h.run("addSubTask(10);");
assertEqual(h.run("items.find(i => i.id === 100).startWeek"), 68, "New child items prefer the current Monday when it is inside the visible timeline");

h.run(`
_getSystemTodayWeek = function() { return null; };
items = [
  { id: 10, type: "project", name: "Parent", startWeek: 1, duration: 2 },
  { id: 11, type: "task", name: "Child Late", parentId: 10, startWeek: 20, duration: 1 },
  { id: 12, type: "task", name: "Child Early", parentId: 10, startWeek: 5, duration: 1 }
];
nextId = 101;
`);
h.run("addSubTask(10);");
assertEqual(h.run("items.find(i => i.id === 101).startWeek"), 21, "New child items still append after the latest sibling end when no current-day preference is available");

h.run(`
years = ["2025"];
items = [
  { id: 1, type: "milestone", name: "M1", startWeek: 5, duration: 0, color: "#f59e0b" }
];
nextId = 2;
`);
h.run("addTask();");
assertEqual(h.run("items.find(i => i.id === 2).startWeek"), 6, "New activities snap to Monday after milestone-only timelines");

h.run(`
years = ["2025"];
items = [{ id: 1, type: "task", name: "New Activity", startWeek: 50, duration: 2, color: "#4f46e5" }];
customTemplates = [{ id: "tpl_1", name: "Long Template", duration: 8, color: "#123456", composition: [] }];
`);
h.run(`applyActivityTemplate(1, "tpl_1");`);
assertEqual(h.run("years.slice()"), ["2025", "2026"], "Applying a template that exceeds the current year auto-adds the next year");

h.run(`
years = ["2025", "2026"];
hiddenYears = new Set();
`);
h.run("__focusWeek = _timelineXToAbsWeek(_absWeekToTimelineX(60.4, 'weeks'), 'weeks');");
h.run("__focusMonth = _timelineXToAbsWeek(_absWeekToTimelineX(60.4, 'months'), 'months');");
assertEqual(h.run("Math.round(__focusWeek * 10) / 10"), 60.4, "Week-view focus helpers preserve the same timeline position");
assertEqual(h.run("Math.floor(__focusMonth)"), 60, "Month-view focus helpers preserve the same calendar area");

h.run(`
__origGetElementById = document.getElementById;
const planner = document.getElementById("planner-test");
planner.scrollLeft = 120;
planner.clientWidth = 600;
planner.scrollWidth = 3000;
planner.getBoundingClientRect = () => ({ left: 0, right: 600 });
const grid = document.getElementById("grid");
grid.getBoundingClientRect = () => ({ left: 0, right: 3000 });
grid.querySelector = (sel) => {
  if (sel === ".header-row-4 .name-cell" || sel === ".name-cell") {
    return { getBoundingClientRect: () => ({ right: 430 }) };
  }
  return { getBoundingClientRect: () => ({ left: 430 }), offsetLeft: 550 };
};
document.querySelector = (sel) => sel === ".planner-container" ? planner : null;
draggingItems = [99];
document.getElementById = function(id) {
  if (id === "grid") return grid;
  if (id === "planner-test") return planner;
  if (id === "block-99") {
    return {
      getBoundingClientRect() { return { left: 9999, right: 10039 }; },
    };
  }
  return {
    id,
    style: {},
    classList: { add() {}, remove() {}, contains() { return false; }, toggle() {} },
    querySelector() { return null; },
    querySelectorAll() { return []; },
    appendChild() {},
  };
};
`);
h.run("_dragStartViewportBounds = new Map([[99, { left: 460, right: 500 }]]); _dragPreviewSnapPx = -35;");
assertOk(h.run("_getDragAutoScrollDelta(520) < 0"), "Dragging near the activity column triggers left auto-scroll", "Expected negative auto-scroll delta when the dragged block touches the frozen columns");
h.run("_dragPreviewSnapPx = 60;");
assertOk(h.run("_getDragAutoScrollDelta(590) > 0"), "Dragging near the right edge triggers right auto-scroll", "Expected positive auto-scroll delta near the viewport edge");
h.run("document.getElementById = __origGetElementById;");
assertOk(
  fs.readFileSync(path.join(__dirname, "..", "src/css/style.css"), "utf8").includes(".milestone-link-anchor"),
  "Milestone-specific link anchor styling is present",
  "Expected milestone link anchor CSS rules",
);
assertEqual(
  h.run("getCommentLinks('see https://example.com and file://server/doc').length"),
  2,
  "Comment link detection finds URLs and file links",
);
assertOk(
  h.run("getCommentLinksHtml('see https://example.com', true).includes('comment-detected-link')"),
  "Comment link previews render as clickable link markup",
  "Expected clickable comment link preview markup",
);

h.run(`
years = ["2025"];
zoomMode = "days";
hiddenYears = new Set();
alarms = [];
document.getElementById("alarm-date").value = "";
document.body.style.zoom = "0.75";
`);
h.set("__dayEvent", {
  clientX: 115,
  target: {
    closest() { return false; },
  },
});
h.set("__dayTarget", {
  getBoundingClientRect() {
    return { left: 10 };
  },
});
h.run("handleTlDblClick(__dayEvent, __dayTarget, null);");
assertEqual(h.run('document.getElementById("alarm-date").value'), "2025-01-08", "Day-view double-click opens the alarm on the clicked calendar day");
h.run(`document.body.style.zoom = "";`);

h.run(`
years = ["2025"];
people = [{ id: "p_1", name: "Alice" }, { id: "p_2", name: "Bob" }];
items = [
  { id: 1, type: "task", name: "Task A", startWeek: 1, duration: 1, assignees: ["p_1"] },
  { id: 2, type: "task", name: "Task B", startWeek: 1, duration: 1, assignees: ["p_1"] }
];
zoomMode = "weeks";
renderWorkloadView();
`);
const workloadHtml = h.document.getElementById("workload-view").innerHTML;
assertOk(
  workloadHtml.includes("overlapping assignments") && workloadHtml.includes("overloaded"),
  "Workload view flags overlapping assignments as overloads",
  "Expected overload markers in workload HTML",
);

h.run(`
years = ["2025"];
people = [{ id: "p_1", name: "Alice" }];
items = [
  { id: 3, type: "task", name: "Ghost", startWeek: 1, duration: 1, assignees: ["deleted_person"] }
];
zoomMode = "weeks";
renderWorkloadView();
`);
assertOk(
  h.document.getElementById("workload-view").innerHTML.includes("1 of 1 free"),
  "Deleted assignees do not consume live team capacity",
  "Expected stale assignees to be ignored in workload totals",
);

h.run(`
years = ["2025"];
zoomMode = "months";
hiddenYears = new Set();
__openedWeek = null;
openAlarmForWeek = function(week) { __openedWeek = week; };
`);
h.set("__monthEvent", {
  offsetX: 120,
  target: {
    closest() { return false; },
  },
});
h.run("handleTlDblClick(__monthEvent, null);");
assertEqual(h.run("__openedWeek"), 6, "Month-view alarms open on the mapped month start week");

h.run(`
years = ["2026", "2027"];
`);
assertEqual(h.run("getIsoWeeksInYear('2026')"), 53, "ISO calendar detects 53-week years");
assertEqual(h.run("getTotalWeekCount()"), 105, "Total week count follows real ISO year lengths");
assertEqual(h.run("parseYYWWD('2653')"), 53, "YYWW parsing accepts ISO week 53 when valid");
assertEqual(h.run("formatYYWWD(53)"), "2653", "YYWW formatting preserves ISO week 53");
assertEqual(h.run("getAlarmAbsWeek({ date: '2027-01-01' })"), 53, "Year-boundary dates map to the correct ISO week");
assertEqual(h.run("getAlarmExactWeek({ date: '2027-01-01' })"), 53.8, "Alarm dates keep their weekday offset within the ISO week");
h.run(`document.getElementById("alarm-date").value = "2026-03-09"; updateAlarmDateWeek();`);
assertEqual(h.run('document.getElementById("alarm-date-week").innerText'), "(W11)", "Alarm popup shows the ISO week number within the selected year");
assertEqual(h.run('document.getElementById("alarm-date-text").value'), "09/03/2026", "Alarm popup displays dates in day/month/year format");
h.run('document.getElementById("alarm-date-text").value = "03/08/2026"; commitAlarmDateText();');
assertEqual(h.run('document.getElementById("alarm-date").value'), "2026-08-03", "Typed alarm dates are parsed from day/month/year");
assertEqual(h.run('document.getElementById("alarm-date-text").value'), "03/08/2026", "Typed alarm dates stay visible in day/month/year format");
h.run(`years = ["2025"];`);
assertEqual(h.run("getAlarmExactWeek({ date: '2025-01-07' })"), 2.2, "Tuesday alarms map to the Tuesday column within their ISO week");
h.run(`
years = ["2026"];
items = [
  { id: 1, type: "milestone", name: "M", startWeek: 1, duration: 0 },
  { id: 2, type: "task", name: "T", startWeek: 2, duration: 1 }
];
links = [];
_undoStack = [];
`);
h.run(`updateMilestoneDate(1, "2554");`);
assertEqual(
  h.run("({ years: years.slice(), starts: items.map(i => i.startWeek), undo: _undoStack.length })"),
  { years: ["2026"], starts: [1, 2], undo: 0 },
  "Rejected milestone dates do not mutate the timeline or create undo entries",
);
h.run(`
alarms = [{ id: 11, title: "Review", date: "2025-01-07", time: "10:30", duration: 45, itemId: 7 }];
items = [{ id: 7, type: "task", name: "Task A", startWeek: 1, duration: 1 }];
openAlarmEditor(11);
`);
assertEqual(
  h.run(`({
    editId: currentAlarmEditId,
    itemId: currentAlarmItemId,
    title: document.getElementById("alarm-title").value,
    date: document.getElementById("alarm-date").value,
    time: document.getElementById("alarm-time").value,
    duration: document.getElementById("alarm-duration").value
  })`),
  { editId: 11, itemId: 7, title: "Review", date: "2025-01-07", time: "10:30", duration: 45 },
  "Opening an alarm loads its details into the editor",
);
h.run("deleteCurrentAlarm();");
assertEqual(h.run("alarms.length"), 0, "Deleting the selected alarm removes it from the planner");

h.run(`
years = ["2026"];
items = [{ id: 1, type: "task", name: "A", startWeek: 1, duration: 1 }];
alarms = [];
holidays = [];
people = [];
links = [];
customTemplates = [];
templateCategories = [];
customColumns = [];
columnFilters = {};
filters = { name: [], type: [] };
showSettings = false;
showComments = false;
commentWidth = 200;
nameWidthBase = 350;
showLinks = true;
localStorage.setItem("diaglo_planner_editor_name", "QA");
_masterFileHandle = { name: "planner_master.csv" };
_canWrite = true;
_changesSinceLastSave = false;
_lastSavedAt = Date.parse("2026-01-01T12:34:00.000Z");
`);
const lockCsv = h.run("generateCSVString({ includeEditorLock: true })");
h.set("__lockCsv", lockCsv);
assertOk(
  lockCsv.includes("EditorLock"),
  "Shared-master CSV writes include a saved-by metadata row",
  "Expected EditorLock-formatted metadata row in autosave CSV output",
);
assertEqual(
  h.run("(_parseSavedByFromText(__lockCsv) || {}).editorName"),
  "QA",
  "Saved-by rows round-trip through CSV parsing",
);
h.run("_updateAutoSaveUI();");
assertEqual(
  h.run('document.getElementById("autosave-label").textContent'),
  "Auto-save",
  "Auto-save UI labels the shared-file workflow correctly when a master file is active",
);
assertOk(
  h.run('document.getElementById("autosave-status").textContent.includes("Saved")'),
  "Auto-save status shows a saved indicator for the active master file",
  "Expected saved status text after updating the auto-save UI",
);

h.run(`
years = ["2025"];
items = [{ id: 1, type: "task", name: "Existing", startWeek: 1, duration: 1 }];
people = [];
alarms = [];
holidays = [];
links = [];
customTemplates = [];
templateCategories = [];
customColumns = [];
columnFilters = {};
nextId = 10;
nextPersonId = 1;
nextAlarmId = 1;
nextLinkId = 1;
`);
h.set("__mergeCsvText", [
  'Years,2025',
  'TemplateCategory,"cat_old","Ops"',
  'Template,"Merge Base",1,"#335577","","cat_old",true,false,false,"tpl_old"',
  'Template,"Merge Master",2,"[{""templateId"":""tpl_old"",""quantity"":2}]","cat_old",false,true,false,"tpl_master_old"',
  'CustomCol,"col_old","Risk",160,true',
  'Id,Type,Name,StartWeek,Duration,ParentId,IsExpanded,Comment,Color,IsLocked,MilestoneRow,BlockCommentData,AssigneesData,Completion,CustomData',
  '5,task,"Merged",3,2,,true,"","#123456",false,0,"","",75,"{""col_old"":""High""}"',
].join("\n"));
h.set("__mergeEvent", { target: { files: [{ name: "merge.csv", __text: h.run("__mergeCsvText") }], value: "x" } });
h.run("mergeCSV(__mergeEvent);");
assertEqual(
  h.run('({ completion: items.find(i => i.name === "Merged").completion, customData: items.find(i => i.name === "Merged").customData, customColumns: customColumns.map(c => c.name), templateCount: customTemplates.length, categoryCount: templateCategories.length })'),
  { completion: 75, customData: { col_old: "High" }, customColumns: ["Risk"], templateCount: 2, categoryCount: 1 },
  "CSV merge preserves completion, custom column data, templates, and template categories",
);

h.run('_masterFileHandle = null; clearTimeout(markChanged._debounce); _stopTimers();');

const failed = results.filter((result) => !result.ok);
results.forEach((result) => {
  const prefix = result.ok ? "PASS" : "FAIL";
  console.log(`${prefix} ${result.name}${result.details ? ` - ${result.details}` : ""}`);
});

if (failed.length > 0) {
  console.error(`\\n${failed.length} audit checks failed.`);
  process.exit(1);
}

console.log(`\\nAll ${results.length} audit checks passed.`);
