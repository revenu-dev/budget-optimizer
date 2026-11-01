/**
 * Budget Optimization: the planner's reader and writer, as a Google Apps Script web app.
 *
 * CONTRACT v2, 17 September 2026 (E9 WP-B12b / WP-B2). Same URL, new deployment.
 * CONTRACT v3, 1 November 2026 (tools E47): `read` adds, per row, accountId, platformId, dailyBudget, makeAmount,
 *   status, campaignStatus; per group, `groupTargets` (the typed Daily Target rows); and `budgets`, the Budgets tab
 *   ([{group, monthlyBudget, rowIndex}]). Nothing a v2 caller reads changed.
 * CONTRACT v4, 1 November 2026 (tools E47 WP-P7): `read` adds `targetRows` (each group's Daily Target row number);
 *   a new `cells` action writes named fields into checked rows (below). Nothing a v3 caller reads changed.
 *
 *   POST {action:"read",  token?, sheetId|sheetUrl, tabName?}
 *     -> {success, sheetId, tabName, headerRow, columnIndices,
 *         budgetGroups:{<group>:{campaigns:{<campaign>:{currentWeighting, engine, rowIndex}}}},   (v1 shape, kept)
 *         rows:[{rowIndex, group, campaign, channel, engine, currentWeighting}],                  (v2: every row, in sheet order)
 *         warnings:[...]}
 *   POST {action:"write", token?, sheetId|sheetUrl, tabName?, results:[{rowIndex?, budgetGroup, campaign, cost, clicks, cpc, leads, cpl,
 *                                                                        grossPipeline, qualifiedPipeline, suggestedWeighting, lostIsRank, lostIsBudget,
 *                                                                        biddingStrategy, bidValue ("" clears a bid that is no longer set)}]}
 *     -> {success, rowsWritten, total, warnings:[...], columns:[the fields this write put in the sheet]}
 *   GET   -> {success:false, error:"POST only"}
 *
 * WHAT CHANGED AND WHY.
 *  - `rows` beside `budgetGroups`: v1 keyed campaigns by name inside a group, so two rows with the same
 *    campaign name in one group (a Reddit and a LinkedIn campaign, say) collapsed to the last one. The
 *    caller now gets every row and decides.
 *  - `write` goes by ROW NUMBER when the caller sends one, and only if the campaign cell on that row is
 *    the campaign it says it is. Names are what the sheet is keyed by for a person; rows are what it is
 *    keyed by for a script. Results without `rowIndex` still match by group|campaign, with the group
 *    key normalised the same way on both sides (v1 lowercased the sheet's and collapsed the caller's,
 *    so a double space or a non-breaking space in the sheet never matched).
 *  - `token`: Script Property BUDGET_TOKEN. When the property is set, a request without the same token
 *    is refused. Apps Script web apps cannot read request headers, so it travels in the body.
 *  - PERCENTAGE WEIGHTING ACTUAL IS NEVER WRITTEN, whatever the caller sends. On 1 October 2026 this
 *    script briefly wrote the approved weighting into Actual as well; Kavinda's review the same day
 *    reversed it (Pavlo and Joe agreed): Actual is the PMM's baseline, judged against client-specific
 *    lead data from Looker, and overwriting it hid the very number the suggestion is weighed against.
 *    Suggested is the AI's column; moving Actual stays a person's decision in the sheet.
 *  - `columns` says which fields actually landed, so the caller can check Actual was not among them.
 *  - `tabName`: when given, the tab must exist by that name; nothing falls back to the first sheet.
 *    Without it the old heuristic runs (a tab whose name contains "pacing" or "budget" and whose
 *    first ten rows contain "campaign"), and when nothing qualifies it is an error, not sheets[0].
 *  - `doGet` no longer reads or writes anything.
 *  - Weightings are read from DISPLAY values, so a percent-formatted cell showing 100% arrives as
 *    "100%" and parses to 100, instead of the raw 1 that made the tool treat a full campaign as 1%.
 */

// ─── Column Matching Configuration ──────────────────────────────────

var READ_COLUMNS = {
  budgetGroup:       { patterns: [['budget group'], ['strategic campaign']], required: true,  exact: true },
  campaign:          { patterns: [['campaign']],                            required: true,  exact: true },
  channel:           { patterns: [['channel']],                             required: false, exact: true },
  currentWeighting:  { patterns: [['weighting', 'actual'], ['weighting', 'before']], required: true,  exact: false },
  engine:            { patterns: [['engine']],                              required: true,  exact: false },
  cost:              { patterns: [['cost']],                                required: false, exact: false },
  clicks:            { patterns: [['clicks']],                              required: false, exact: false },
  cpc:               { patterns: [['cpc']],                                 required: false, exact: false },
  leads:             { patterns: [['leads']],                               required: false, exact: false },
  cpl:               { patterns: [['cpl']],                                 required: false, exact: false },
  grossPipeline:     { patterns: [['gross', 'pipeline'], ['gp']],           required: false, exact: false },
  qualifiedPipeline: { patterns: [['qualified', 'pipeline'], ['qp']],       required: false, exact: false },
  suggestedWeighting:{ patterns: [['weighting', 'suggested']],              required: false, exact: false },
  /* CONTRACT v3 (1 Nov 2026, tools E47 Budget Planner import): the Pacing Automation SOP's columns. IDs are read as
     DISPLAY values (a Meta id is 17-18 digits and loses precision as a number). The Make amount is matched before the
     plain daily budget, and Campaign Status before Status, so neither is taken for the other. */
  accountId:         { patterns: [['account', 'id']],                       required: false, exact: false },
  platformId:        { patterns: [['budget', 'group', 'id'], ['ad', 'group', 'id'], ['campaign', 'id']], required: false, exact: false },
  makeAmount:        { patterns: [['daily', 'budget', 'make']],             required: false, exact: false },
  dailyBudget:       { patterns: [['daily', 'budget']],                     required: false, exact: true },
  campaignStatus:    { patterns: [['campaign', 'status']],                  required: false, exact: true },
  status:            { patterns: [['status']],                              required: false, exact: true },
};

var WRITE_COLUMNS = {
  campaign:           { patterns: [['campaign']],               exact: true },
  budgetGroup:        { patterns: [['budget group'], ['strategic campaign']], exact: true },
  cost:               { patterns: [['cost']],                   exact: false },
  clicks:             { patterns: [['clicks']],                 exact: false },
  cpc:                { patterns: [['cpc']],                    exact: false },
  leads:              { patterns: [['leads']],                  exact: false },
  cpl:                { patterns: [['cpl']],                    exact: false },
  grossPipeline:      { patterns: [['gross', 'pipeline'], ['gp']], exact: false },
  qualifiedPipeline:  { patterns: [['qualified', 'pipeline'], ['qp']], exact: false },
  suggestedWeighting: { patterns: [['weighting', 'suggested']], exact: false },
  lostIsRank:         { patterns: [['lost', 'rank']],           exact: false },
  lostIsBudget:       { patterns: [['lost', 'budget']],         exact: false },
  /* The live bid (20 Oct 2026, Joe's "budget planner with actual bids"): tools reads it from the ad
     account at Write and overwrites these two. Strategy is matched first, so "Bidding Strategy" can never
     be taken for the bid; "Bid Value" (Abnormal) and "Manual Bid" (the template) are the bid. */
  biddingStrategy:    { patterns: [['bidding', 'strategy']],   exact: false },
  bidValue:           { patterns: [['bid', 'value'], ['manual', 'bid']], exact: false },
};

var READ_MATCH_ORDER = ['budgetGroup', 'currentWeighting', 'suggestedWeighting', 'engine', 'channel', 'accountId', 'platformId', 'makeAmount', 'dailyBudget', 'campaignStatus', 'status', 'cost', 'clicks', 'cpc', 'leads', 'cpl', 'grossPipeline', 'qualifiedPipeline', 'campaign'];
var WRITE_MATCH_ORDER = ['budgetGroup', 'suggestedWeighting', 'cost', 'clicks', 'cpc', 'leads', 'cpl', 'grossPipeline', 'qualifiedPipeline', 'lostIsRank', 'lostIsBudget', 'biddingStrategy', 'bidValue', 'campaign'];
var WRITE_FIELDS = ['cost', 'clicks', 'cpc', 'leads', 'cpl', 'grossPipeline', 'qualifiedPipeline', 'suggestedWeighting', 'lostIsRank', 'lostIsBudget', 'biddingStrategy', 'bidValue'];

// ─── Shared Helpers ─────────────────────────────────────────────────

function normalizeHeader(raw) {
  return String(raw).replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
}

/** Case, whitespace and non-breaking spaces do not matter. Used for campaigns AND groups. */
function normalizeCampaignName(name) {
  return String(name || '').replace(/[ \s]+/g, ' ').trim().toLowerCase();
}

function parseWeighting(raw) {
  if (raw === null || raw === undefined || raw === '') return 0;
  var s = String(raw).replace(/[%$,\s]/g, '');
  var num = parseFloat(s);
  if (isNaN(num)) return 0;
  /* An unformatted fraction (0.6 for 60%). A display value "60%" has already lost its sign above and is 60. */
  if (Math.abs(num) > 0 && Math.abs(num) < 1) num = num * 100;
  return Math.round(num * 100) / 100;
}

function extractSheetId(input) {
  if (!input) return null;
  var match = String(input).match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if (match) return match[1];
  if (/^[a-zA-Z0-9_-]+$/.test(String(input))) return String(input);
  return null;
}

/**
 * The tab to work on. A named tab must exist; an unnamed one is found by the old heuristic and
 * never by falling back to the first sheet.
 */
function findBudgetTab(spreadsheet, tabName) {
  if (tabName) {
    var named = spreadsheet.getSheetByName(String(tabName));
    if (!named) throw new Error('Tab "' + tabName + '" not found in this spreadsheet.');
    return { sheet: named, warnings: [] };
  }
  var sheets = spreadsheet.getSheets();
  var searchTerms = ['pacing', 'budget'];
  var potentialSheets = [];
  for (var s = 0; s < searchTerms.length; s++) {
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase().indexOf(searchTerms[s]) !== -1) potentialSheets.push(sheets[i]);
    }
    if (potentialSheets.length > 0) break;
  }
  if (potentialSheets.length === 0) throw new Error('No tab named like "pacing" or "budget". Send tabName.');
  for (var j = 0; j < potentialSheets.length; j++) {
    var testSheet = potentialSheets[j];
    var testData = testSheet.getRange(1, 1, Math.min(testSheet.getLastRow(), 10), testSheet.getLastColumn()).getValues();
    for (var row = 0; row < testData.length; row++) {
      if (testData[row].join(' ').toLowerCase().indexOf('campaign') !== -1) return { sheet: testSheet, warnings: [] };
    }
  }
  throw new Error('A tab matched "pacing" or "budget" but none has a "campaign" header in its first 10 rows. Send tabName.');
}

function buildColumnMap(sheet, columns, matchOrder) {
  var data = sheet.getRange(1, 1, Math.min(sheet.getLastRow(), 10), sheet.getLastColumn()).getValues();
  for (var rowIdx = 0; rowIdx < data.length; rowIdx++) {
    var normalized = data[rowIdx].map(normalizeHeader);
    var columnMap = {};
    var matched = {};
    for (var m = 0; m < matchOrder.length; m++) {
      var target = matchOrder[m];
      var config = columns[target];
      if (!config) continue;
      for (var p = 0; p < config.patterns.length; p++) {
        var pattern = config.patterns[p];
        for (var c = 0; c < normalized.length; c++) {
          if (matched[c]) continue;
          var header = normalized[c];
          var allWordsMatch = pattern.every(function (word) { return header.indexOf(word) !== -1; });
          if (!allWordsMatch) continue;
          if (config.exact) {
            var headerWordCount = header.split(/\s+/).filter(Boolean).length;
            if (headerWordCount > pattern.length + 1) continue;
          }
          columnMap[target] = c;
          matched[c] = target;
          break;
        }
        if (columnMap[target] !== undefined) break;
      }
    }
    if (columnMap.campaign !== undefined) return { columnMap: columnMap, headerRow: rowIdx + 1, missing: [] };
  }
  return { columnMap: {}, headerRow: null, missing: ['Campaign column not found'] };
}

function buildResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

/** The token check. A missing Script Property means the check is off, which is how a fresh deployment behaves until BUDGET_TOKEN is set. */
function authorized_(payload) {
  var want = PropertiesService.getScriptProperties().getProperty('BUDGET_TOKEN');
  if (!want) return true;
  return String(payload.token || '') === String(want);
}

// ─── Router ─────────────────────────────────────────────────────────

function doGet() {
  return buildResponse_({ success: false, error: 'POST only' });
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (!authorized_(payload)) return buildResponse_({ success: false, error: 'unauthorized' });
    var action = (payload.action || '').toLowerCase();
    if (action === 'read') return processRead_(payload);
    if (action === 'write') return processWrite_(payload);
    if (action === 'list') return processList_(payload);
    if (action === 'cells') return processCells_(payload);
    return buildResponse_({ success: false, error: 'Unknown action.' });
  } catch (err) { return buildResponse_({ success: false, error: err.message }); }
}

// ─── Action: List Planners (19 October 2026) ────────────────────────

/** Every planner sheet this script's owner can open, for tools' Settings > Switchboard dropdown: the newest first,
 *  at most 500. "Landing Page Planner"s carry no budget, so they are left out here rather than in every caller. */
function processList_(payload) {
  try {
    var q = 'mimeType = "application/vnd.google-apps.spreadsheet" and title contains "Planner" and trashed = false';
    var it = DriveApp.searchFiles(q);
    var files = [];
    while (it.hasNext() && files.length < 500) {
      var f = it.next();
      var name = f.getName();
      if (/landing page planner/i.test(name)) continue;
      files.push({ id: f.getId(), name: name, updatedAt: f.getLastUpdated().toISOString() });
    }
    files.sort(function (a, b) { return a.updatedAt < b.updatedAt ? 1 : -1; });
    return buildResponse_({ success: true, files: files });
  } catch (err) { return buildResponse_({ success: false, error: err.message }); }
}

// ─── Action: Read Planner ───────────────────────────────────────────

function processRead_(payload) {
  try {
    var sheetId = extractSheetId(payload.sheetUrl || payload.sheetId);
    if (!sheetId) throw new Error('No sheet id.');
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var tab = findBudgetTab(spreadsheet, payload.tabName);
    var sheet = tab.sheet;
    var warnings = tab.warnings.slice();
    var colResult = buildColumnMap(sheet, READ_COLUMNS, READ_MATCH_ORDER);
    if (colResult.headerRow === null) throw new Error(colResult.missing.join('; '));
    var columnMap = colResult.columnMap;
    var headerRow = colResult.headerRow;
    var numRows = sheet.getLastRow() - headerRow;
    var budgetGroups = {};
    var rows = [];
    var groupTargets = {};
    var targetRows = {};
    if (numRows > 0) {
      var range = sheet.getRange(headerRow + 1, 1, numRows, sheet.getLastColumn());
      var allData = range.getValues();
      /* Display values for the weighting only: "100%" parses to 100, the raw value 1 would not. */
      var shown = columnMap.currentWeighting !== undefined ? range.getDisplayValues() : null;
      if (!shown) shown = range.getDisplayValues();
      var currentBudgetGroup = null;
      for (var r = 0; r < allData.length; r++) {
        var row = allData[r];
        var campaignName = String(row[columnMap.campaign] || '').trim();
        var budgetGroupValue = String(row[columnMap.budgetGroup] || '').trim();
        var engineValue = String(row[columnMap.engine] || '').trim();
        if (budgetGroupValue) currentBudgetGroup = budgetGroupValue;
        /* v3: the group's typed daily target, from its "Budget Group Daily Target" row. */
        if (campaignName.toLowerCase().indexOf('daily target') !== -1 && columnMap.dailyBudget !== undefined) {
          groupTargets[budgetGroupValue || currentBudgetGroup || 'Unknown'] = parseMoney_(row[columnMap.dailyBudget]);
          targetRows[budgetGroupValue || currentBudgetGroup || 'Unknown'] = headerRow + r + 1;
          continue;
        }
        if (!campaignName || !engineValue || campaignName.toLowerCase().indexOf('daily target') !== -1) continue;
        var group = currentBudgetGroup || 'Unknown';
        var weightRaw = shown ? shown[r][columnMap.currentWeighting] : row[columnMap.currentWeighting];
        var rec = {
          rowIndex: headerRow + r + 1,
          group: group,
          campaign: campaignName,
          channel: columnMap.channel !== undefined ? String(row[columnMap.channel] || '').trim() : '',
          engine: engineValue,
          currentWeighting: parseWeighting(weightRaw),
          accountId: cellText_(shown, r, columnMap.accountId),
          platformId: cellText_(shown, r, columnMap.platformId),
          dailyBudget: columnMap.dailyBudget !== undefined ? parseMoney_(row[columnMap.dailyBudget]) : null,
          makeAmount: columnMap.makeAmount !== undefined ? parseMoney_(row[columnMap.makeAmount]) : null,
          status: cellText_(shown, r, columnMap.status),
          campaignStatus: cellText_(shown, r, columnMap.campaignStatus),
        };
        rows.push(rec);
        if (!budgetGroups[group]) budgetGroups[group] = { campaigns: {} };
        budgetGroups[group].campaigns[campaignName] = { currentWeighting: rec.currentWeighting, engine: rec.engine, rowIndex: rec.rowIndex };
      }
    }
    return buildResponse_({
      success: true, sheetId: sheetId, tabName: sheet.getName(), headerRow: headerRow, columnIndices: columnMap,
      budgetGroups: budgetGroups, rows: rows, warnings: warnings,
      contract: 4, groupTargets: groupTargets, targetRows: targetRows, budgets: readBudgetsTab_(spreadsheet, warnings),
    });
  } catch (err) { return buildResponse_({ success: false, error: err.message }); }
}

/** A cell's display text, trimmed; '' when the column is absent. */
function cellText_(shown, r, col) {
  if (col === undefined || !shown) return '';
  return String(shown[r][col] === undefined || shown[r][col] === null ? '' : shown[r][col]).trim();
}

/** Money from a cell: a number, or a string like "$1,500.00"; anything else ("-", "", "#VALUE!") is null, never 0. */
function parseMoney_(raw) {
  if (typeof raw === 'number') return isFinite(raw) ? Math.round(raw * 100) / 100 : null;
  var s = String(raw === null || raw === undefined ? '' : raw).replace(/[$£€,\s]/g, '');
  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  return Math.round(parseFloat(s) * 100) / 100;
}

/**
 * v3: the Budgets tab, one row per group: [{group, monthlyBudget, rowIndex}]. Absent tab = [] and a warning.
 * A monthly budget that is not a number is null, with a warning naming the group (Hatch, 28 Oct: a "-").
 */
function readBudgetsTab_(spreadsheet, warnings) {
  var sheet = spreadsheet.getSheetByName('Budgets');
  if (!sheet) { warnings.push('No "Budgets" tab: monthly budgets were not read.'); return []; }
  var last = sheet.getLastRow(), width = sheet.getLastColumn();
  if (last < 1 || width < 1) return [];
  var data = sheet.getRange(1, 1, last, width).getValues();
  for (var h = 0; h < Math.min(10, data.length); h++) {
    var hdr = data[h].map(normalizeHeader);
    var gi = -1, mi = -1;
    for (var c = 0; c < hdr.length; c++) {
      if (gi < 0 && (hdr[c] === 'budget group' || hdr[c] === 'strategic campaign')) gi = c;
      if (mi < 0 && hdr[c].indexOf('monthly') !== -1 && hdr[c].indexOf('budget') !== -1) mi = c;
    }
    if (gi < 0 || mi < 0) continue;
    var out = [];
    for (var r = h + 1; r < data.length; r++) {
      var group = String(data[r][gi] || '').trim();
      if (!group) continue;
      var money = parseMoney_(data[r][mi]);
      if (money === null) warnings.push('Budgets tab: "' + group + '" has a monthly budget that is not a number.');
      out.push({ group: group, monthlyBudget: money, rowIndex: r + 1 });
    }
    return out;
  }
  warnings.push('The "Budgets" tab has no "Budget Group" and "Monthly Budget" headers in its first 10 rows.');
  return [];
}

// ─── Action: Cells (contract v4, 1 Nov 2026, tools E47 WP-P7) ───────

/**
 * Named fields into named rows, each row checked first: on the Pacing tab its Campaign cell must be the campaign the
 * caller names (or "Budget Group Daily Target" for a target row), on the Budgets tab its Budget Group cell the group.
 * A row that does not match is not written and is named in `warnings`; nothing else in the sheet is touched, no row
 * or column is added or moved (each client's Make scenario reads columns by position).
 *
 *   {action:"cells", token, sheetId, tabName,
 *    pacing:[{rowIndex, campaign, set:{currentWeighting?, dailyBudget?, makeAmount?, status?}}],
 *    budgets:[{rowIndex, group, monthlyBudget}],
 *    mirror:true}
 *
 * PERCENTAGE WEIGHTING ACTUAL is written only when `mirror` is true: the client's planner is then a copy of the
 * groups Tools holds (its per-client "Tools owns this planner" switch), and Actual has to say what Tools says or
 * Funnel and Looker pace against the wrong number. The optimizer's `write` still never writes it (1 October 2026).
 * Weightings go in as fractions (60 -> 0.6, the planner's percent format); `status: ""` clears the cell.
 */
var CELL_FIELDS = ['currentWeighting', 'dailyBudget', 'makeAmount', 'status'];
function processCells_(payload) {
  try {
    var sheetId = extractSheetId(payload.sheetId || payload.sheetUrl);
    if (!sheetId) throw new Error('No sheet id.');
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var warnings = [], written = 0;
    var pacing = payload.pacing || [], budgets = payload.budgets || [];
    if (pacing.length) {
      var sheet = findBudgetTab(spreadsheet, payload.tabName).sheet;
      var colResult = buildColumnMap(sheet, READ_COLUMNS, READ_MATCH_ORDER);
      if (colResult.headerRow === null) throw new Error(colResult.missing.join('; '));
      var map = colResult.columnMap, headerRow = colResult.headerRow;
      var numRows = sheet.getLastRow() - headerRow;
      var values = numRows > 0 ? sheet.getRange(headerRow + 1, 1, numRows, sheet.getLastColumn()).getValues() : [];
      var fields = CELL_FIELDS.filter(function (f) { return map[f] !== undefined && (f !== 'currentWeighting' || payload.mirror === true); });
      var cols = {};
      fields.forEach(function (f) { cols[f] = new Array(numRows).fill(null); });
      for (var i = 0; i < pacing.length; i++) {
        var u = pacing[i], off = Number(u.rowIndex) - headerRow - 1;
        if (!(off >= 0 && off < numRows)) { warnings.push('Row ' + u.rowIndex + ' is outside the tab. Not written.'); continue; }
        if (normalizeCampaignName(values[off][map.campaign]) !== normalizeCampaignName(u.campaign)) {
          warnings.push('Row ' + u.rowIndex + ' holds "' + String(values[off][map.campaign] || '') + '", not "' + u.campaign + '". Not written; the sheet has moved since it was read.');
          continue;
        }
        var set = u.set || {}, any = false;
        fields.forEach(function (f) {
          if (set[f] === undefined || set[f] === null) return;
          cols[f][off] = f === 'currentWeighting' ? Number(set[f]) / 100 : set[f];
          any = true;
        });
        if (set.currentWeighting !== undefined && payload.mirror !== true) warnings.push('Row ' + u.rowIndex + ': Actual is only written for a planner Tools owns. Not written.');
        if (any) written++;
      }
      fields.forEach(function (f) { writeRuns_(sheet, headerRow, map[f] + 1, cols[f]); });
    }
    if (budgets.length) {
      var bs = spreadsheet.getSheetByName('Budgets');
      if (!bs) warnings.push('No "Budgets" tab: monthly budgets were not written.');
      else {
        var data = bs.getRange(1, 1, bs.getLastRow(), bs.getLastColumn()).getValues();
        var gi = -1, mi = -1;
        for (var h = 0; h < Math.min(10, data.length) && (gi < 0 || mi < 0); h++) {
          gi = -1; mi = -1;
          var hdr = data[h].map(normalizeHeader);
          for (var c = 0; c < hdr.length; c++) {
            if (gi < 0 && (hdr[c] === 'budget group' || hdr[c] === 'strategic campaign')) gi = c;
            if (mi < 0 && hdr[c].indexOf('monthly') !== -1 && hdr[c].indexOf('budget') !== -1) mi = c;
          }
        }
        if (gi < 0 || mi < 0) warnings.push('The "Budgets" tab has no "Budget Group" and "Monthly Budget" headers. Not written.');
        else for (var j = 0; j < budgets.length; j++) {
          var b = budgets[j], r0 = Number(b.rowIndex) - 1;
          if (!(r0 >= 0 && r0 < data.length) || normalizeCampaignName(data[r0][gi]) !== normalizeCampaignName(b.group)) { warnings.push('Budgets row ' + b.rowIndex + ' is not "' + b.group + '". Not written.'); continue; }
          bs.getRange(r0 + 1, mi + 1).setValue(Number(b.monthlyBudget));
          written++;
        }
      }
    }
    return buildResponse_({ success: true, written: written, warnings: warnings });
  } catch (err) { return buildResponse_({ success: false, error: err.message }); }
}

/** One column's values (null = leave the cell alone) in contiguous runs. */
function writeRuns_(sheet, headerRow, colIdx, colValues) {
  var runStart = null;
  for (var row = 0; row <= colValues.length; row++) {
    var val = row < colValues.length ? colValues[row] : null;
    if (val !== null && runStart === null) runStart = row;
    else if (val === null && runStart !== null) {
      sheet.getRange(headerRow + 1 + runStart, colIdx, row - runStart, 1).setValues(colValues.slice(runStart, row).map(function (v) { return [v]; }));
      runStart = null;
    }
  }
}

// ─── Action: Write Planner ──────────────────────────────────────────

function processWrite_(payload) {
  try {
    var sheetId = extractSheetId(payload.sheetId || payload.sheetUrl);
    if (!sheetId) throw new Error('No sheet id.');
    var results = payload.results || [];
    if (results.length === 0) return buildResponse_({ success: false, error: 'No results.' });

    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var sheet = findBudgetTab(spreadsheet, payload.tabName).sheet;
    var colResult = buildColumnMap(sheet, WRITE_COLUMNS, WRITE_MATCH_ORDER);
    if (colResult.headerRow === null) throw new Error(colResult.missing.join('; '));
    var columnMap = colResult.columnMap;
    var headerRow = colResult.headerRow;
    var numRows = sheet.getLastRow() - headerRow;
    var lookupValues = numRows > 0 ? sheet.getRange(headerRow + 1, 1, numRows, sheet.getLastColumn()).getValues() : [];

    // 1. Composite lookup (group|campaign -> 0-based row offset), for results without a row number.
    var compositeMap = {};
    var campaignAt = [];
    var currentGroup = '';
    for (var r = 0; r < lookupValues.length; r++) {
      var groupInSheet = String(lookupValues[r][columnMap.budgetGroup] || '').trim();
      if (groupInSheet) currentGroup = normalizeCampaignName(groupInSheet);
      var campName = normalizeCampaignName(lookupValues[r][columnMap.campaign]);
      campaignAt[r] = campName;
      if (campName) {
        var key = currentGroup + '|' + campName;
        if (compositeMap[key] === undefined) compositeMap[key] = r;
      }
    }

    // 2. Per-column value arrays (null = leave this row alone).
    var available = WRITE_FIELDS.filter(function (f) { return columnMap[f] !== undefined; });
    var columnData = {};
    available.forEach(function (f) { columnData[f] = new Array(numRows).fill(null); });
    var rowsWritten = 0;
    var warnings = [];

    for (var i = 0; i < results.length; i++) {
      var res = results[i];
      var rowOffset;
      if (res.rowIndex !== undefined && res.rowIndex !== null && res.rowIndex !== '') {
        /* By row number, and only if the row is the campaign it claims to be. */
        var off = Number(res.rowIndex) - headerRow - 1;
        if (!(off >= 0 && off < numRows)) { warnings.push('Row ' + res.rowIndex + ' is outside the tab for: ' + res.budgetGroup + ' | ' + res.campaign + '. Not written.'); continue; }
        if (campaignAt[off] !== normalizeCampaignName(res.campaign)) {
          warnings.push('Row ' + res.rowIndex + ' holds "' + String(lookupValues[off][columnMap.campaign] || '') + '", not "' + res.campaign + '". Not written; the sheet has moved since it was read.');
          continue;
        }
        rowOffset = off;
      } else {
        rowOffset = compositeMap[normalizeCampaignName(res.budgetGroup) + '|' + normalizeCampaignName(res.campaign)];
        if (rowOffset === undefined) { warnings.push('No sheet row found for: ' + res.budgetGroup + ' | ' + res.campaign); continue; }
      }
      available.forEach(function (field) {
        var val = res[field];
        if (val !== undefined && val !== null) columnData[field][rowOffset] = (field === 'suggestedWeighting') ? val / 100 : val;
      });
      rowsWritten++;
    }

    // 3. Each target column in contiguous runs; other columns are never touched.
    available.forEach(function (field) {
      var colIdx = columnMap[field] + 1;
      var colValues = columnData[field];
      var runStart = null;
      for (var row = 0; row <= colValues.length; row++) {
        var val = row < colValues.length ? colValues[row] : null;
        if (val !== null && runStart === null) {
          runStart = row;
        } else if (val === null && runStart !== null) {
          var runData = colValues.slice(runStart, row).map(function (v) { return [v]; });
          sheet.getRange(headerRow + 1 + runStart, colIdx, row - runStart, 1).setValues(runData);
          runStart = null;
        }
      }
    });

    var columns = available.filter(function (f) { return columnData[f].some(function (v) { return v !== null; }); });
    return buildResponse_({ success: true, rowsWritten: rowsWritten, total: results.length, warnings: warnings, columns: columns });
  } catch (err) { return buildResponse_({ success: false, error: err.message }); }
}
