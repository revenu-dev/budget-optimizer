/* Contract v4 (1 Nov 2026, tools E47 WP-P7): `cells` writes named fields into checked rows, Actual only for a planner
   Tools owns; `read` gives each group's Daily Target row. `node test/gas-v4.test.cjs` */
const fs = require("fs"), vm = require("vm");
const src = fs.readFileSync(process.argv[2] || require("path").join(__dirname, "..", "gas.js"), "utf8");

function tab(name, rows) {
  const W = Math.max(...rows.map((r) => r.length));
  const vals = rows.map((r) => r.concat(Array(W - r.length).fill("")));
  const writes = [];
  return {
    vals, writes, getName: () => name, getLastRow: () => vals.length, getLastColumn: () => W,
    getRange: (r, c, nr, nc) => ({
      getValues: () => vals.slice(r - 1, r - 1 + (nr || 1)).map((x) => x.slice(c - 1, c - 1 + (nc || 1))),
      getDisplayValues: () => vals.slice(r - 1, r - 1 + (nr || 1)).map((x) => x.slice(c - 1, c - 1 + (nc || 1)).map(String)),
      setValues: (v) => { writes.push({ r, c, n: v.length }); v.forEach((row, i) => row.forEach((x, j) => { vals[r - 1 + i][c - 1 + j] = x; })); },
      setValue: (x) => { writes.push({ r, c, n: 1 }); vals[r - 1][c - 1] = x; },
    }),
  };
}
const H = ["Channel", "Google/LinkedIn Account ID, Meta Campaign ID", "Google Budget Group ID, LinkedIn Campaign ID, Meta Ad Group ID", "Daily Budget Make", "Budget Group", "Campaign", "Engine", "Daily Budget", "% Weighting Actual", "% Weighting Suggested", "Status", "Campaign Status"];
const row = (o) => { const r = Array(H.length).fill(""); for (const [k, v] of Object.entries(o)) r[H.indexOf(k)] = v; return r; };
function sheets() {
  const pacing = tab("Pacing & Optimization", [H,
    row({ Channel: "Google", "Budget Group": "Non Brand", Campaign: "NB US", Engine: "Leads", "Daily Budget": 70, "% Weighting Actual": 0.7, "Daily Budget Make": 70, Status: "Updated", "Campaign Status": "ENABLED" }),
    row({ Channel: "Meta", Campaign: "Meta RT", Engine: "Leads", "Daily Budget": 30, "% Weighting Actual": 0.3, "Daily Budget Make": 3000, Status: "Updated" }),
    row({ "Budget Group": "Non Brand", Campaign: "Budget Group Daily Target", "Daily Budget": 100 }),
  ]);
  const budgets = tab("Budgets", [["Budget Group", "Monthly Budget"], ["Non Brand", 3000]]);
  return { pacing, budgets };
}
function run(payload, sh) {
  const map = { "Pacing & Optimization": sh.pacing, Budgets: sh.budgets };
  const ctx = {
    SpreadsheetApp: { openById: () => ({ getSheetByName: (n) => map[n] || null, getSheets: () => Object.values(map) }) },
    ContentService: { createTextOutput: (s) => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 1 } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) }, console,
  };
  vm.createContext(ctx); vm.runInContext(src, ctx);
  return JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(payload) } }).s);
}
let pass = 0, fail = 0; const ok = (n, c, x) => { c ? pass++ : fail++; console.log((c ? "  ok   " : "  FAIL ") + n + (c ? "" : " -> " + JSON.stringify(x))); };
const base = { action: "cells", sheetId: "1234567890abcdefghijklmnop", tabName: "Pacing & Optimization" };
const C = (h) => H.indexOf(h);

{
  const sh = sheets();
  const r = run({ action: "read", sheetId: base.sheetId, tabName: base.tabName }, sh);
  ok("read: contract 4 and each group's Daily Target row", r.contract === 4 && r.targetRows["Non Brand"] === 4, r.targetRows);
}
{
  const sh = sheets();
  const r = run(Object.assign({}, base, { mirror: true,
    pacing: [{ rowIndex: 2, campaign: "NB US", set: { currentWeighting: 80, dailyBudget: 80, makeAmount: 80, status: "" } },
      { rowIndex: 3, campaign: "Meta RT", set: { currentWeighting: 20, dailyBudget: 20, makeAmount: 2000, status: "" } },
      { rowIndex: 4, campaign: "Budget Group Daily Target", set: { dailyBudget: 100 } }],
    budgets: [{ rowIndex: 2, group: "Non Brand", monthlyBudget: 3000 }] }), sh);
  const v = sh.pacing.vals;
  ok("mirror: Actual as a fraction, daily budget, Make amount, status cleared, the target row, the monthly budget", r.success && r.written === 4 &&
    v[1][C("% Weighting Actual")] === 0.8 && v[1][C("Daily Budget")] === 80 && v[1][C("Daily Budget Make")] === 80 && v[1][C("Status")] === "" &&
    v[2][C("% Weighting Actual")] === 0.2 && v[2][C("Daily Budget Make")] === 2000 && v[3][C("Daily Budget")] === 100 && sh.budgets.vals[1][1] === 3000, { r, v });
  ok("...and nothing else in the sheet moved: Suggested, Campaign Status, names untouched", v[1][C("Campaign Status")] === "ENABLED" && v[1][C("% Weighting Suggested")] === "" && v[1][C("Campaign")] === "NB US");
  ok("...each column in one run where the rows are contiguous", sh.pacing.writes.filter((w) => w.c === C("% Weighting Actual") + 1).length === 1);
}
{
  const sh = sheets();
  const r = run(Object.assign({}, base, { pacing: [{ rowIndex: 2, campaign: "NB US", set: { currentWeighting: 80, dailyBudget: 80 } }] }), sh);
  ok("without mirror, Actual is never written (1 Oct 2026) and says so; the daily budget still is", sh.pacing.vals[1][C("% Weighting Actual")] === 0.7 && sh.pacing.vals[1][C("Daily Budget")] === 80 && r.warnings.some((w) => /Actual is only written for a planner Tools owns/.test(w)), r);
}
{
  const sh = sheets();
  const r = run(Object.assign({}, base, { mirror: true, pacing: [{ rowIndex: 2, campaign: "Somebody else", set: { dailyBudget: 1 } }, { rowIndex: 99, campaign: "x", set: { dailyBudget: 1 } }],
    budgets: [{ rowIndex: 2, group: "Brand", monthlyBudget: 1 }] }), sh);
  ok("a row that is not the campaign named, a row off the tab and a Budgets row of another group are not written, each named", r.written === 0 &&
    sh.pacing.vals[1][C("Daily Budget")] === 70 && sh.budgets.vals[1][1] === 3000 && r.warnings.length === 3 && /holds "NB US", not "Somebody else"/.test(r.warnings[0]), r);
}
{
  const sh = sheets();
  const r = run(Object.assign({}, base, { mirror: true, pacing: [{ rowIndex: 3, campaign: "Meta RT", set: { status: "Held by Tools" } }] }), sh);
  ok("a status can be set (Tools holds a row back from Make)", r.written === 1 && sh.pacing.vals[2][C("Status")] === "Held by Tools", r);
}
console.log(pass + " passed, " + fail + " failed"); process.exit(fail ? 1 : 0);
