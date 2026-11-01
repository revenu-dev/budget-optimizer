/* Contract v3 read (1 Nov 2026, tools E47): IDs, daily budgets, Make amount, status columns, group targets and the
   Budgets tab, on a tab shaped like OneRail's planner. `node test/gas-v3.test.cjs` */
const fs = require("fs"), vm = require("vm");
const src = fs.readFileSync(process.argv[2] || require("path").join(__dirname, "..", "gas.js"), "utf8");

function tab(name, rows, shownRows) {
  const W = Math.max(...rows.map((r) => r.length));
  const pad = (r) => r.concat(Array(W - r.length).fill(""));
  const vals = rows.map(pad), shown = (shownRows || rows.map((r) => r.map(String))).map(pad);
  return {
    getName: () => name, getLastRow: () => vals.length, getLastColumn: () => W,
    getRange: (r, c, nr, nc) => ({
      getValues: () => vals.slice(r - 1, r - 1 + nr).map((x) => x.slice(c - 1, c - 1 + nc)),
      getDisplayValues: () => shown.slice(r - 1, r - 1 + nr).map((x) => x.slice(c - 1, c - 1 + nc).map(String)),
    }),
  };
}
// OneRail's columns: A Channel, B account, C budget/campaign id, D Daily Budget Make, E Budget Group, L Campaign, R Engine,
// S Daily Budget, T Actual, U Suggested, X Status, Y Campaign Status.
const H = ["Channel", "Google/LinkedIn Account ID, Meta Campaign ID", "Google Budget Group ID, LinkedIn Campaign ID, Meta Ad Group ID", "Daily Budget Make",
  "Budget Group", "Strategy", "Targeting Type", "Targeting Audience", "Advert Type", "Targeting Country", "Variant", "Campaign", "Cost", "Clicks", "CPC", "Leads", "CPL",
  "Engine", "Daily Budget", "% Weighting Actual", "% Weighting Suggested", "Bidding Strategy", "Bid Value", "Status", "Campaign Status"];
const row = (o) => { const r = Array(H.length).fill(""); for (const [k, v] of Object.entries(o)) r[H.indexOf(k)] = v; return r; };
const pacingRows = [H,
  row({ Channel: "Google", "Google/LinkedIn Account ID, Meta Campaign ID": 1112223333, "Google Budget Group ID, LinkedIn Campaign ID, Meta Ad Group ID": 14567890123, "Daily Budget Make": 116.06, "Budget Group": "Non Brand", Campaign: "NB US", Engine: "Leads", "Daily Budget": 116.06, "% Weighting Actual": 0.7, Status: "Updated", "Campaign Status": "ENABLED" }),
  row({ Channel: "Meta", "Google/LinkedIn Account ID, Meta Campaign ID": 120212345678901234, "Google Budget Group ID, LinkedIn Campaign ID, Meta Ad Group ID": 120298765432109876, "Daily Budget Make": 5017, Campaign: "Meta RT", Engine: "Leads", "Daily Budget": 50.17, "% Weighting Actual": 0.3 }),
  row({ "Budget Group": "Non Brand", Campaign: "Budget Group Daily Target", "Daily Budget": 167.23, "% Weighting Actual": 1 }),
  row({ Channel: "Google", "Budget Group": "Brand", Campaign: "Brand US", Engine: "Leads", "Daily Budget": 0, "% Weighting Actual": 1 }),
  row({ "Budget Group": "Brand", Campaign: "Budget Group Daily Target", "Daily Budget": "-" }),
];
// Display values: big ids as Sheets shows plain-text cells (the SOP says plain text), money formatted.
const pacingShown = pacingRows.map((r) => r.map(String));
pacingShown[1][1] = "111-222-3333"; pacingShown[2][1] = "120212345678901234"; pacingShown[2][2] = "120298765432109876"; pacingShown[1][19] = "70%"; pacingShown[2][19] = "30%"; pacingShown[4][19] = "100%";
const budgetsRows = [["Region", "Budget Group", "Strategy", "Channels", "Monthly Budget", "Weekly Pacing"], ["US", "Non Brand", "Search", "Google", 5900, ""], ["US", "Brand", "Search", "Google", "-", ""], ["", "", "", "", "", ""]];

function run(payload, opts) {
  const o = opts || {};
  const sheets = { "Pacing & Optimization": tab("Pacing & Optimization", pacingRows, pacingShown) };
  if (!o.noBudgets) sheets.Budgets = tab("Budgets", budgetsRows);
  const ctx = {
    SpreadsheetApp: { openById: () => ({ getSheetByName: (n) => sheets[n] || null, getSheets: () => Object.values(sheets) }) },
    ContentService: { createTextOutput: (s) => ({ s, setMimeType() { return this; } }), MimeType: { JSON: 1 } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    console,
  };
  vm.createContext(ctx); vm.runInContext(src, ctx);
  return JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(payload) } }).s);
}
let pass = 0, fail = 0; const ok = (n, c, x) => { c ? pass++ : fail++; console.log((c ? "  ok   " : "  FAIL ") + n + (c ? "" : " -> " + JSON.stringify(x))); };
const read = { action: "read", sheetId: "1234567890abcdefghijklmnop", tabName: "Pacing & Optimization" };

const r = run(read);
const ci = r.columnIndices;
ok("v3+: success and contract 4", r.success && r.contract === 4, r);
ok("the SOP's id columns map to account (B) and platform id (C), never to Campaign or Budget Group", ci.accountId === 1 && ci.platformId === 2 && ci.campaign === 11 && ci.budgetGroup === 4, ci);
ok("the Make amount (D) and the plain Daily Budget (S) are told apart", ci.makeAmount === 3 && ci.dailyBudget === 18, ci);
ok("Status (X) and Campaign Status (Y) are told apart", ci.status === 23 && ci.campaignStatus === 24, ci);
ok("v2 fields unchanged: Actual is current, Suggested separate", ci.currentWeighting === 19 && ci.suggestedWeighting === 20, ci);
const nb = r.rows.find((x) => x.campaign === "NB US"), meta = r.rows.find((x) => x.campaign === "Meta RT");
ok("ids come back as the sheet shows them; an 18-digit Meta id keeps every digit", nb.accountId === "111-222-3333" && meta.accountId === "120212345678901234" && meta.platformId === "120298765432109876", [nb, meta]);
ok("daily budget and Make amount as numbers; status texts as shown", nb.dailyBudget === 116.06 && meta.makeAmount === 5017 && nb.status === "Updated" && nb.campaignStatus === "ENABLED" && meta.status === "", [nb, meta]);
ok("the Daily Target rows are not campaigns", r.rows.every((x) => !/daily target/i.test(x.campaign)) && r.rows.length === 3, r.rows.map((x) => x.campaign));
ok("each group's typed daily target is read; a '-' target is null, never 0", r.groupTargets["Non Brand"] === 167.23 && r.groupTargets.Brand === null && "Brand" in r.groupTargets, r.groupTargets);
ok("the Budgets tab: one row per group, monthly budget a number, blank rows skipped", r.budgets.length === 2 && r.budgets[0].group === "Non Brand" && r.budgets[0].monthlyBudget === 5900 && r.budgets[0].rowIndex === 2, r.budgets);
ok("a '-' monthly budget is null and warned about by name (Hatch, 28 Oct)", r.budgets[1].monthlyBudget === null && r.warnings.some((w) => /"Brand" has a monthly budget that is not a number/.test(w)), [r.budgets[1], r.warnings]);
const nob = run(read, { noBudgets: true });
ok("no Budgets tab: an empty list and a warning, not an error", nob.success && nob.budgets.length === 0 && nob.warnings.some((w) => /No "Budgets" tab/.test(w)), nob);
console.log(pass + " passed, " + fail + " failed"); process.exit(fail ? 1 : 0);
