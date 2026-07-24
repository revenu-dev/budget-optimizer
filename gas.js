/**
 * Budget Optimization — Google Apps Script (Composite Key + Batch Optimized)
 */

// ─── Column Matching Configuration ──────────────────────────────────

var READ_COLUMNS = {
  budgetGroup:       { patterns: [['budget group'], ['strategic campaign']], required: true,  exact: true },
  campaign:          { patterns: [['campaign']],                            required: true,  exact: true },
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

var READ_MATCH_ORDER = ['budgetGroup', 'currentWeighting', 'suggestedWeighting', 'engine', 'cost', 'clicks', 'cpc', 'leads', 'cpl', 'grossPipeline', 'qualifiedPipeline', 'campaign'];
var WRITE_MATCH_ORDER = ['budgetGroup', 'suggestedWeighting', 'cost', 'clicks', 'cpc', 'leads', 'cpl', 'grossPipeline', 'qualifiedPipeline', 'lostIsRank', 'lostIsBudget', 'campaign'];

// ─── Shared Helpers ─────────────────────────────────────────────────

function normalizeHeader(raw) {
  return String(raw).replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
}

function normalizeCampaignName(name) {
  return String(name || '').replace(/[\u00A0\s]+/g, ' ').trim().toLowerCase();
}

function parseWeighting(raw) {
  if (raw === null || raw === undefined || raw === '') return 0;
  var s = String(raw).replace(/[%$,\s]/g, '');
  var num = parseFloat(s);
  if (isNaN(num)) return 0;
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

function findBudgetTab(spreadsheet) {
  var sheets = spreadsheet.getSheets();
  var searchTerms = ['pacing', 'budget'];
  var potentialSheets = [];

  for (var s = 0; s < searchTerms.length; s++) {
    var term = searchTerms[s];
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase().indexOf(term) !== -1) {
        potentialSheets.push(sheets[i]);
      }
    }
    if (potentialSheets.length > 0) break;
  }

  if (potentialSheets.length === 0) return { sheet: sheets[0], warnings: ['No tab matched "pacing" or "budget" — using first tab'] };

  for (var j = 0; j < potentialSheets.length; j++) {
    var testSheet = potentialSheets[j];
    var testData = testSheet.getRange(1, 1, Math.min(testSheet.getLastRow(), 10), testSheet.getLastColumn()).getValues();
    for (var row = 0; row < testData.length; row++) {
      var joinedRow = testData[row].join(' ').toLowerCase();
      if (joinedRow.indexOf('campaign') !== -1) return { sheet: testSheet, warnings: [] };
    }
  }
  return { sheet: potentialSheets[0], warnings: ['Tab matched keyword but no "campaign" header found in first 10 rows — defaulting to first match'] };
}

function buildColumnMap(sheet, columns, matchOrder) {
  var data = sheet.getRange(1, 1, Math.min(sheet.getLastRow(), 10), sheet.getLastColumn()).getValues();

  for (var rowIdx = 0; rowIdx < data.length; rowIdx++) {
    var headers = data[rowIdx];
    var normalized = headers.map(normalizeHeader);
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

          // All pattern words must be present in the header
          var allWordsMatch = pattern.every(function(word) { return header.indexOf(word) !== -1; });
          if (!allWordsMatch) continue;

          // For exact fields: header word count must be within 1 of the pattern word count.
          // This allows "Budget Group" to match ['budget group'] but rejects
          // "Strategic Campaign Budget Group Overview" (too many extra words).
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

// ─── Router ─────────────────────────────────────────────────────────

function doGet(e) {
  var action = (e.parameter.action || '').toLowerCase();
  if (action === 'read') return processRead_(e.parameter);
  if (action === 'write') return processWrite_(JSON.parse(e.parameter.payload));
  return buildResponse_({ success: false, error: 'Unknown action.' });
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
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
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var sheet = findBudgetTab(spreadsheet).sheet;
    var colResult = buildColumnMap(sheet, READ_COLUMNS, READ_MATCH_ORDER);

    var columnMap = colResult.columnMap;
    var headerRow = colResult.headerRow;
    var allData = sheet.getRange(headerRow + 1, 1, sheet.getLastRow() - headerRow, sheet.getLastColumn()).getValues();

    var budgetGroups = {};
    var currentBudgetGroup = null;

    for (var r = 0; r < allData.length; r++) {
      var row = allData[r];
      var campaignName = String(row[columnMap.campaign] || '').trim();
      var budgetGroupValue = String(row[columnMap.budgetGroup] || '').trim();
      var engineValue = String(row[columnMap.engine] || '').trim();

      if (budgetGroupValue) currentBudgetGroup = budgetGroupValue;
      if (!campaignName || !engineValue || campaignName.toLowerCase().indexOf('daily target') !== -1) continue;

      var group = currentBudgetGroup || 'Unknown';
      if (!budgetGroups[group]) budgetGroups[group] = { campaigns: {} };

      budgetGroups[group].campaigns[campaignName] = {
        currentWeighting: parseWeighting(row[columnMap.currentWeighting]),
        engine: engineValue,
        rowIndex: headerRow + r + 1
      };
    }

    return buildResponse_({ success: true, sheetId: sheetId, budgetGroups: budgetGroups });
  } catch (err) { return buildResponse_({ success: false, error: err.message }); }
}

// ─── Action: Write Planner (COMPOSITE KEY BATCH) ────────────────────

function processWrite_(payload) {
  try {
    var sheetId = extractSheetId(payload.sheetId || payload.sheetUrl);
    var results = payload.results || [];
    if (results.length === 0) return buildResponse_({ success: false, error: 'No results.' });

    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var sheet = findBudgetTab(spreadsheet).sheet;
    var colResult = buildColumnMap(sheet, WRITE_COLUMNS, WRITE_MATCH_ORDER);

    var columnMap = colResult.columnMap;
    var headerRow = colResult.headerRow;
    var numRows = sheet.getLastRow() - headerRow;

    // Read full rows only to build the composite lookup map
    var lookupValues = sheet.getRange(headerRow + 1, 1, numRows, sheet.getLastColumn()).getValues();

    // 1. Build composite lookup map (budgetGroup|campaign → 0-based row offset)
    var compositeMap = {};
    var currentGroup = '';

    for (var r = 0; r < lookupValues.length; r++) {
      var groupInSheet = String(lookupValues[r][columnMap.budgetGroup] || '').trim();
      if (groupInSheet) currentGroup = groupInSheet.toLowerCase();

      var campName = normalizeCampaignName(lookupValues[r][columnMap.campaign]);
      if (campName) {
        var key = currentGroup + '|' + campName;
        if (!compositeMap[key]) compositeMap[key] = r;
      }
    }

    // 2. Build per-column value arrays (null = don't write this row)
    var writeFields = ['cost', 'clicks', 'cpc', 'leads', 'cpl', 'grossPipeline', 'qualifiedPipeline', 'suggestedWeighting', 'lostIsRank', 'lostIsBudget'];
    var available = writeFields.filter(function(f) { return columnMap[f] !== undefined; });
    var rowsWritten = 0;
    var warnings = [];

    var columnData = {};
    available.forEach(function(f) {
      columnData[f] = new Array(numRows).fill(null);
    });

    for (var i = 0; i < results.length; i++) {
      var res = results[i];
      var resKey = normalizeCampaignName(res.budgetGroup) + '|' + normalizeCampaignName(res.campaign);
      var rowOffset = compositeMap[resKey];

      if (rowOffset !== undefined) {
        available.forEach(function(field) {
          var val = res[field];
          if (val !== undefined && val !== null) {
            columnData[field][rowOffset] = (field === 'suggestedWeighting') ? val / 100 : val;
          }
        });
        rowsWritten++;
      } else {
        warnings.push('No sheet row found for: ' + res.budgetGroup + ' | ' + res.campaign);
      }
    }

    // 3. Write each target column in contiguous runs — never touches other columns
    available.forEach(function(field) {
      var colIdx = columnMap[field] + 1; // 1-based for getRange
      var colValues = columnData[field];
      var runStart = null;

      for (var row = 0; row <= colValues.length; row++) {
        var val = colValues[row];
        if (val !== null && runStart === null) {
          runStart = row;
        } else if ((val === null || row === colValues.length) && runStart !== null) {
          var runData = colValues.slice(runStart, row).map(function(v) { return [v]; });
          sheet.getRange(headerRow + 1 + runStart, colIdx, row - runStart, 1).setValues(runData);
          runStart = null;
        }
      }
    });

    return buildResponse_({ success: true, rowsWritten: rowsWritten, total: results.length, warnings: warnings });
  } catch (err) { return buildResponse_({ success: false, error: err.message }); }
}