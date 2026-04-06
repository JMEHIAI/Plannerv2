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

  context.global = context;
  vm.createContext(context);

  const root = path.resolve(__dirname, "..");
  [
    "src/js/state.js",
    "src/js/links.js",
    "src/js/items.js",
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
items = [
  { id: 10, type: "project", name: "Parent", startWeek: 1, duration: 2 },
  { id: 11, type: "task", name: "Child Late", parentId: 10, startWeek: 20, duration: 1 },
  { id: 12, type: "task", name: "Child Early", parentId: 10, startWeek: 5, duration: 1 }
];
nextId = 100;
`);
h.run("addSubTask(10);");
assertEqual(h.run("items.find(i => i.id === 100).startWeek"), 21, "New child items append after the latest sibling end");

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

h.run(`
years = ["2026"];
items = [{ id: 1, type: "task", name: "A", startWeek: 1, duration: 1 }];
alarms = [];
holidays = [];
people = [];
links = [];
customTemplates = [];
filters = { name: [], type: [] };
showSettings = false;
showComments = false;
commentWidth = 200;
nameWidthBase = 350;
showLinks = true;
localStorage.setItem("diaglo_planner_editor_name", "QA");
_masterFileHandle = { name: "planner_master.csv" };
_hasEditLock = true;
_editorLockInfo = { sessionId: _editorSessionId, editorName: "QA", acquiredAt: "2026-01-01T00:00:00.000Z" };
`);
const lockCsv = h.run("generateCSVString({ includeEditorLock: true })");
h.set("__lockCsv", lockCsv);
assertOk(
  lockCsv.includes("EditorLock"),
  "Shared-master CSV writes include an editor-lock row",
  "Expected EditorLock row in autosave CSV output",
);
assertEqual(
  h.run("(_parseEditorLockFromText(__lockCsv) || {}).editorName"),
  "QA",
  "Editor-lock rows round-trip through CSV parsing",
);
h.run("_hasEditLock = false; _masterFileHandle = null; clearTimeout(markChanged._debounce); _stopTimer();");

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
