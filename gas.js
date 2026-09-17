/**
 * Budget Optimization: the planner's reader and writer, as a Google Apps Script web app.
 *
 * CONTRACT v2, 17 September 2026 (E9 WP-B12b / WP-B2). Same URL, new deployment.
 *
 *   POST {action:"read",  token?, sheetId|sheetUrl, tabName?}
 *     -> {success, sheetId, tabName, headerRow, columnIndices,
 *         budgetGroups:{<group>:{campaigns:{<campaign>:{currentWeighting, engine, rowIndex}}}},   (v1 shape, kept)
 *         rows:[{rowIndex, group, campaign, channel, engine, currentWeighting}],                  (v2: every row, in sheet order)
 *         warnings:[...]}
 *   POST {action:"write", token?, sheetId|sheetUrl, tabName?, results:[{rowIndex?, budgetGroup, campaign, cost, clicks, cpc, leads, cpl,
 *                                                                        grossPipeline, qualifiedPipeline, suggestedWeighting, lostIsRank, lostIsBudget}]}
 *     -> {success, rowsWritten, total, warnings:[...]}
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
};

var READ_MATCH_ORDER = ['budgetGroup', 'currentWeighting', 'suggestedWeighting', 'engine', 'channel', 'cost', 'clicks', 'cpc', 'leads', 'cpl', 'grossPipeline', 'qualifiedPipeline', 'campaign'];
var WRITE_MATCH_ORDER = ['budgetGroup', 'suggestedWeighting', 'cost', 'clicks', 'cpc', 'leads', 'cpl', 'grossPipeline', 'qualifiedPipeline', 'lostIsRank', 'lostIsBudget', 'campaign'];
var WRITE_FIELDS = ['cost', 'clicks', 'cpc', 'leads', 'cpl', 'grossPipeline', 'qualifiedPipeline', 'suggestedWeighting', 'lostIsRank', 'lostIsBudget'];

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
    return buildResponse_({ success: false, error: 'Unknown action.' });
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
    if (numRows > 0) {
      var range = sheet.getRange(headerRow + 1, 1, numRows, sheet.getLastColumn());
      var allData = range.getValues();
      /* Display values for the weighting only: "100%" parses to 100, the raw value 1 would not. */
      var shown = columnMap.currentWeighting !== undefined ? range.getDisplayValues() : null;
      var currentBudgetGroup = null;
      for (var r = 0; r < allData.length; r++) {
        var row = allData[r];
        var campaignName = String(row[columnMap.campaign] || '').trim();
        var budgetGroupValue = String(row[columnMap.budgetGroup] || '').trim();
        var engineValue = String(row[columnMap.engine] || '').trim();
        if (budgetGroupValue) currentBudgetGroup = budgetGroupValue;
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
        };
        rows.push(rec);
        if (!budgetGroups[group]) budgetGroups[group] = { campaigns: {} };
        budgetGroups[group].campaigns[campaignName] = { currentWeighting: rec.currentWeighting, engine: rec.engine, rowIndex: rec.rowIndex };
      }
    }
    return buildResponse_({
      success: true, sheetId: sheetId, tabName: sheet.getName(), headerRow: headerRow, columnIndices: columnMap,
      budgetGroups: budgetGroups, rows: rows, warnings: warnings,
    });
  } catch (err) { return buildResponse_({ success: false, error: err.message }); }
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

    return buildResponse_({ success: true, rowsWritten: rowsWritten, total: results.length, warnings: warnings });
  } catch (err) { return buildResponse_({ success: false, error: err.message }); }
}
