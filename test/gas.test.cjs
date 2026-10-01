/* gas.js against a fake planner tab, in node: `node test/gas.test.cjs` (or pass another copy of gas.js). 1 Oct 2026: Suggested only; Actual is never written, even when a caller sends actualWeighting. */
const fs = require("fs"), vm = require("vm");
const src = fs.readFileSync(process.argv[2] || require("path").join(__dirname, "..", "gas.js"), "utf8");
// A tab: header on row 1, columns like Ottimate's (campaign 12, group 4, Actual 20, Suggested 21, engine 18).
function makeSheet() {
  const W = 24;
  const hdr = Array(W).fill("");
  Object.assign(hdr, { 0: "Channel", 4: "Budget Group", 12: "Campaign", 13: "Cost", 14: "Clicks", 15: "CPC", 16: "Leads", 17: "CPL", 18: "Engine", 20: "Percentage Weighting Actual", 21: "Percentage Weighting Suggested" });
  const rows = [hdr];
  const mk = (g, c, a) => { const r = Array(W).fill(""); r[4] = g; r[12] = c; r[18] = "Leads"; r[20] = a; return r; };
  rows.push(mk("G1", "A", 0.5), mk("", "B", 0.5), mk("G2", "C", 1));
  const formulas = rows.map((r) => r.map(() => ""));
  formulas[3][20] = "=1-SUM(T2:T3)"; // C's Actual is a formula
  const writes = [];
  const sheet = {
    getName: () => "Pacing & Optimization", getLastRow: () => rows.length, getLastColumn: () => W,
    getRange: (r, c, nr, nc) => ({
      getValues: () => rows.slice(r - 1, r - 1 + nr).map((x) => x.slice(c - 1, c - 1 + nc)),
      getDisplayValues: () => rows.slice(r - 1, r - 1 + nr).map((x) => x.slice(c - 1, c - 1 + nc).map(String)),
      getFormulas: () => formulas.slice(r - 1, r - 1 + nr).map((x) => x.slice(c - 1, c - 1 + nc)),
      setValues: (v) => { writes.push({ r, c, v }); v.forEach((row, i) => row.forEach((x, j) => { rows[r - 1 + i][c - 1 + j] = x; })); },
    }),
  };
  return { sheet, rows, writes };
}
function run(payload) {
  const t = makeSheet();
  const ctx = {
    SpreadsheetApp: { openById: () => ({ getSheetByName: (n) => (n === "Pacing & Optimization" ? t.sheet : null), getSheets: () => [t.sheet] }) },
    ContentService: { createTextOutput: (s) => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 1 } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    console,
  };
  vm.createContext(ctx); vm.runInContext(src, ctx);
  const out = ctx.doPost({ postData: { contents: JSON.stringify(payload) } });
  return { res: JSON.parse(out.s), rows: t.rows };
}
const res = (o) => Object.assign({ budgetGroup: "G1", campaign: "A", rowIndex: 2, cost: 1 }, o);
let pass = 0, fail = 0; const ok = (n, c, x) => { c ? pass++ : fail++; console.log((c ? "  ok   " : "  FAIL ") + n + (c ? "" : " -> " + JSON.stringify(x))); };
const base = { action: "write", sheetId: "1234567890abcdefghijklmnop", tabName: "Pacing & Optimization" };
// 1. A caller that sends actualWeighting anyway: Actual stays as it was, Suggested gets the number.
const a = run(Object.assign({}, base, { results: [res({ suggestedWeighting: 70, actualWeighting: 70 }), res({ campaign: "B", rowIndex: 3, suggestedWeighting: 30, actualWeighting: 30 }), res({ budgetGroup: "G2", campaign: "C", rowIndex: 4, suggestedWeighting: 100, actualWeighting: 100 })] }));
ok("success, three rows", a.res.success && a.res.rowsWritten === 3, a.res);
ok("Actual is untouched on every row, formula or not", a.rows[1][20] === 0.5 && a.rows[2][20] === 0.5 && a.rows[3][20] === 1, [a.rows[1][20], a.rows[2][20], a.rows[3][20]]);
ok("Suggested holds the new weighting as a fraction", a.rows[1][21] === 0.7 && a.rows[2][21] === 0.3 && a.rows[3][21] === 1, [a.rows[1][21], a.rows[2][21], a.rows[3][21]]);
ok("columns lists what landed, and Actual is not among them", a.res.columns.indexOf("suggestedWeighting") >= 0 && a.res.columns.indexOf("cost") >= 0 && a.res.columns.indexOf("actualWeighting") < 0, a.res.columns);
// 2. The performance columns still land beside it.
ok("cost is written", a.rows[1][13] === 1, a.rows[1][13]);
// 3. Reading still maps Actual as current and Suggested separately.
const c = run({ action: "read", sheetId: base.sheetId, tabName: base.tabName });
ok("read still maps Actual as current", c.res.success && c.res.columnIndices.currentWeighting === 20 && c.res.columnIndices.suggestedWeighting === 21, c.res.columnIndices || c.res);
console.log(pass + " passed, " + fail + " failed"); process.exit(fail ? 1 : 0);
