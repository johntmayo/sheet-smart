// ============================================================
// MergeEngine.gs — Core Merge Infrastructure
// ============================================================
// Lookup-and-merge engine for integrating data across Google
// Spreadsheets. Reads a config spreadsheet to know which source
// to pull from, which column to match on, and which columns to
// map. Fills blank target cells from the source and logs
// conflicts (non-blank cells with differing values) for manual
// review. Designed to be pasted into the same Apps Script
// project as Corrections.gs; all functions share one namespace.
// ============================================================

/**
 * Reads the Settings and Column Mapping tabs from the config
 * spreadsheet and returns a structured config object.
 *
 * @param {SpreadsheetApp.Spreadsheet} configSpreadsheet
 * @return {{ sourceId: string, matchColumn: string, folderId: string, masterId: string, columnMap: Array<{source: string, target: string}> }}
 */
function readMergeConfig_(configSpreadsheet) {
  var settingsSheet = configSpreadsheet.getSheetByName('Settings');
  var settingsData = settingsSheet.getDataRange().getValues();

  var settings = {};
  for (var i = 0; i < settingsData.length; i++) {
    var key = String(settingsData[i][0]).trim();
    var val = String(settingsData[i][1]).trim();
    if (key !== '') settings[key] = val;
  }

  var mappingSheet = configSpreadsheet.getSheetByName('Column Mapping');
  var mappingData = mappingSheet.getDataRange().getValues();
  var columnMap = [];

  for (var j = 1; j < mappingData.length; j++) {
    var src = String(mappingData[j][0]).trim();
    var tgt = String(mappingData[j][1]).trim();
    // Skip blank rows and placeholder text like "(header name in master)"
    if (src === '' || tgt === '') continue;
    if (src.charAt(0) === '(' && src.charAt(src.length - 1) === ')') continue;
    if (tgt.charAt(0) === '(' && tgt.charAt(tgt.length - 1) === ')') continue;
    columnMap.push({ source: src, target: tgt });
  }

  return {
    masterId:     settings['Master Spreadsheet'] || '',
    sourceId:     settings['External Source'] || settings['Source Spreadsheet'] || '',
    userSheetId:  settings['User Sheet'] || '',
    folderId:     settings['User Sheets Folder'] || settings['Target Folder'] || '',
    matchColumn:  settings['Match Column'] || '',
    columnMap:    columnMap
  };
}

/**
 * Builds a lookup object keyed by the match-column value.
 * Each value is an object mapping column headers to cell values
 * for that row.
 *
 * @param {Array<Array>} sourceData  2D array including header row
 * @param {Array<string>} headers    Header row (already trimmed)
 * @param {string} matchColumn       Column name to key on
 * @return {Object}  e.g. { '123-456': { 'APN': '123-456', 'Sale Price': 500000 } }
 */
function buildSourceLookup_(sourceData, headers, matchColumn) {
  var matchIdx = headers.indexOf(matchColumn);
  if (matchIdx === -1) return {};

  var lookup = {};
  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    var key = String(row[matchIdx]).trim();
    if (key === '' || key === 'undefined' || key === 'null') continue;

    var record = {};
    for (var c = 0; c < headers.length; c++) {
      record[headers[c]] = row[c];
    }
    lookup[key] = record;
  }

  return lookup;
}

/**
 * Merges source data into a single target sheet. Fills blank
 * cells and records conflicts without overwriting.
 *
 * @param {SpreadsheetApp.Sheet} targetSheet   First tab of a target spreadsheet
 * @param {Object} sourceLookup                From buildSourceLookup_
 * @param {string} matchColumn                 Column name to match on
 * @param {Array<{source: string, target: string}>} columnMap
 * @param {boolean} dryRun                     If true, log only — never write
 * @return {{ filled: Array, conflicts: Array, errors: Array }}
 */
function mergeIntoTarget_(targetSheet, sourceLookup, matchColumn, columnMap, dryRun) {
  var filled = [];
  var conflicts = [];
  var errors = [];

  var data = targetSheet.getDataRange().getValues();
  if (data.length === 0) {
    errors.push({ row: 0, column: '', existingValue: '', newValue: 'Target sheet is empty' });
    return { filled: filled, conflicts: conflicts, errors: errors };
  }

  var targetHeaders = data[0].map(function (h) { return String(h).trim(); });

  var matchIdx = targetHeaders.indexOf(matchColumn);
  if (matchIdx === -1) {
    errors.push({
      row: 0,
      column: matchColumn,
      existingValue: '',
      newValue: 'Match column "' + matchColumn + '" not found in target headers'
    });
    return { filled: filled, conflicts: conflicts, errors: errors };
  }

  var targetColIndices = [];
  for (var m = 0; m < columnMap.length; m++) {
    var tgtIdx = targetHeaders.indexOf(columnMap[m].target);
    if (tgtIdx === -1) {
      errors.push({
        row: 0,
        column: columnMap[m].target,
        existingValue: '',
        newValue: 'Mapped target column "' + columnMap[m].target + '" not found in target sheet'
      });
      targetColIndices.push(-1);
    } else {
      targetColIndices.push(tgtIdx);
    }
  }

  var writeQueue = [];

  for (var r = 1; r < data.length; r++) {
    var matchKey = String(data[r][matchIdx]).trim();
    if (matchKey === '' || matchKey === 'undefined' || matchKey === 'null') continue;

    var sourceRow = sourceLookup[matchKey];
    if (!sourceRow) continue;

    for (var c = 0; c < columnMap.length; c++) {
      if (targetColIndices[c] === -1) continue;

      var sourceVal = sourceRow[columnMap[c].source];
      if (typeof sourceVal === 'string' && sourceVal.toLowerCase() === 'true') {
        sourceVal = true;
      } else if (typeof sourceVal === 'string' && sourceVal.toLowerCase() === 'false') {
        sourceVal = false;
      }
      var targetVal = data[r][targetColIndices[c]];

      // Treat boolean false as blank: it is the default state of an unchecked
      // checkbox, not a user-entered value. This prevents checkbox-formatted
      // columns from generating spurious conflicts.
      var targetBlank = (targetVal === '' || targetVal === null || targetVal === undefined || targetVal === false);
      var sourceBlank = (sourceVal === '' || sourceVal === null || sourceVal === undefined);

      if (targetBlank && !sourceBlank) {
        filled.push({
          row: r + 1,
          column: columnMap[c].target,
          existingValue: '',
          newValue: sourceVal
        });
        if (!dryRun) {
          // wasCheckbox: cell held the boolean false default, so its column
          // likely has checkbox data validation — clear that before writing
          // so the new value displays as plain text/number, not a checkbox.
          writeQueue.push({
            row: r + 1,
            col: targetColIndices[c] + 1,
            val: sourceVal,
            wasCheckbox: (targetVal === false)
          });
        }
      } else if (!targetBlank && !sourceBlank && targetVal !== sourceVal) {
        conflicts.push({
          row: r + 1,
          column: columnMap[c].target,
          existingValue: targetVal,
          newValue: sourceVal
        });
      }
    }
  }

  if (!dryRun) {
    for (var w = 0; w < writeQueue.length; w++) {
      var cell = targetSheet.getRange(writeQueue[w].row, writeQueue[w].col);
      if (writeQueue[w].wasCheckbox) {
        cell.clearDataValidations();
        cell.setNumberFormat('General');
      }
      cell.setValue(writeQueue[w].val);
    }
  }

  return { filled: filled, conflicts: conflicts, errors: errors };
}

/**
 * Creates (or clears) a results log tab and writes all merge
 * results as a batch.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet  Where to write the log
 * @param {string} tabName                          Tab name for the log
 * @param {string} sheetName                        Name of the sheet being merged (for context)
 * @param {Array} filled                            Fill entries from mergeIntoTarget_
 * @param {Array} conflicts                         Conflict entries
 * @param {Array} errors                            Error entries
 */
function writeResultsLog_(spreadsheet, tabName, sheetName, filled, conflicts, errors) {
  var tab = spreadsheet.getSheetByName(tabName);
  if (tab) {
    tab.clear();
  } else {
    tab = spreadsheet.insertSheet(tabName);
  }

  var header = ['Type', 'Spreadsheet', 'Row', 'Column', 'Existing Value', 'New Value'];
  var rows = [header];

  for (var i = 0; i < filled.length; i++) {
    rows.push(['Filled', sheetName, filled[i].row, filled[i].column, filled[i].existingValue, filled[i].newValue]);
  }
  for (var j = 0; j < conflicts.length; j++) {
    rows.push(['Conflict', sheetName, conflicts[j].row, conflicts[j].column, conflicts[j].existingValue, conflicts[j].newValue]);
  }
  for (var k = 0; k < errors.length; k++) {
    rows.push(['Error', sheetName, errors[k].row, errors[k].column, errors[k].existingValue, errors[k].newValue]);
  }

  if (rows.length > 0) {
    tab.getRange(1, 1, rows.length, header.length).setValues(rows);
  }

  formatHeaderRow_(tab, header.length);
  tab.setFrozenRows(1);
}

/**
 * Bolds and freezes the header row of a log tab.
 *
 * @param {SpreadsheetApp.Sheet} tab
 * @param {number} numCols
 */
function formatHeaderRow_(tab, numCols) {
  var headerRange = tab.getRange(1, 1, 1, numCols);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f3f3');
}

/**
 * Adds missing column headers to the first row of a target sheet.
 * New headers are appended after the last existing column in a
 * single batch write. Already-present columns are skipped silently.
 *
 * @param {SpreadsheetApp.Sheet} targetSheet
 * @param {Array<string>} columnNames  Ordered list of column names to ensure exist
 * @param {boolean} dryRun             If true, log only — never write
 * @return {{ added: Array<{column: string}>, skipped: Array<{column: string}>, errors: Array<{column: string, message: string}> }}
 */
function addColumnsToTarget_(targetSheet, columnNames, dryRun) {
  var added = [];
  var skipped = [];
  var errors = [];

  var data = targetSheet.getDataRange().getValues();
  if (data.length === 0) {
    errors.push({ column: '', message: 'Target sheet is empty' });
    return { added: added, skipped: skipped, errors: errors };
  }

  var existingHeaders = data[0].map(function (h) { return String(h).trim(); });
  var toAdd = [];

  for (var i = 0; i < columnNames.length; i++) {
    var colName = String(columnNames[i]).trim();
    if (colName === '') continue;
    if (existingHeaders.indexOf(colName) !== -1) {
      skipped.push({ column: colName });
    } else {
      added.push({ column: colName });
      toAdd.push(colName);
    }
  }

  if (!dryRun && toAdd.length > 0) {
    var startCol = existingHeaders.length + 1;
    targetSheet.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]);
    // Flush so the header write is committed before subsequent reads.
    SpreadsheetApp.flush();
    // Data rows: clear inherited checkbox validation, any false values that
    // Google Sheets propagates from an adjacent checkbox column, and reset
    // format to plain General so cells start genuinely empty.
    var lastRow = targetSheet.getMaxRows();
    if (lastRow > 1) {
      var dataRange = targetSheet.getRange(2, startCol, lastRow - 1, toAdd.length);
      dataRange.clearDataValidations();
      dataRange.clearContent();
      dataRange.setNumberFormat('General');
    }
  }

  return { added: added, skipped: skipped, errors: errors };
}

/**
 * Appends combined sync results (column additions + fills + conflicts + errors)
 * to a single log tab, creating it with a header row if needed.
 * Used by the unified Sync operations so all results land in one place.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} tabName
 * @param {string} sheetName
 * @param {{ added: Array, skipped: Array, errors: Array }} addResult   From addColumnsToTarget_
 * @param {{ filled: Array, conflicts: Array, errors: Array }} mergeResult  From mergeIntoTarget_
 */
function appendToSyncLog_(spreadsheet, tabName, sheetName, addResult, mergeResult) {
  var tab = spreadsheet.getSheetByName(tabName);
  var header = ['Type', 'Spreadsheet', 'Row', 'Column', 'Existing Value', 'New / Source Value'];

  if (!tab) {
    tab = spreadsheet.insertSheet(tabName);
    tab.getRange(1, 1, 1, header.length).setValues([header]);
    formatHeaderRow_(tab, header.length);
    tab.setFrozenRows(1);
  }

  var newRows = [];

  for (var i = 0; i < addResult.added.length; i++) {
    newRows.push(['Column Added', sheetName, 1, addResult.added[i].column, '(new column)', '']);
  }
  for (var j = 0; j < addResult.errors.length; j++) {
    newRows.push(['Error', sheetName, 0, addResult.errors[j].column, '', addResult.errors[j].message]);
  }
  for (var k = 0; k < mergeResult.filled.length; k++) {
    var f = mergeResult.filled[k];
    newRows.push(['Filled', sheetName, f.row, f.column, '(blank)', f.newValue]);
  }
  for (var l = 0; l < mergeResult.conflicts.length; l++) {
    var c = mergeResult.conflicts[l];
    newRows.push(['Conflict', sheetName, c.row, c.column, c.existingValue, c.newValue]);
  }
  for (var m = 0; m < mergeResult.errors.length; m++) {
    var e = mergeResult.errors[m];
    newRows.push(['Error', sheetName, e.row, e.column, e.existingValue, e.newValue]);
  }

  if (newRows.length > 0) {
    var startRow = tab.getLastRow() + 1;
    tab.getRange(startRow, 1, newRows.length, header.length).setValues(newRows);
  }
}

/**
 * Appends add-columns results to a log tab, creating the tab with
 * a header row if it does not yet exist. Used when accumulating
 * results across many target sheets.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} tabName
 * @param {string} sheetName
 * @param {{ added: Array, skipped: Array, errors: Array }} results
 */
function appendToAddColumnsLog_(spreadsheet, tabName, sheetName, results) {
  var tab = spreadsheet.getSheetByName(tabName);
  var header = ['Type', 'Spreadsheet', 'Column', 'Notes'];

  if (!tab) {
    tab = spreadsheet.insertSheet(tabName);
    tab.getRange(1, 1, 1, header.length).setValues([header]);
    formatHeaderRow_(tab, header.length);
    tab.setFrozenRows(1);
  }

  var newRows = [];

  for (var i = 0; i < results.added.length; i++) {
    newRows.push(['Added', sheetName, results.added[i].column, '']);
  }
  for (var j = 0; j < results.skipped.length; j++) {
    newRows.push(['Skipped', sheetName, results.skipped[j].column, 'Already exists']);
  }
  for (var k = 0; k < results.errors.length; k++) {
    newRows.push(['Error', sheetName, results.errors[k].column, results.errors[k].message]);
  }

  if (newRows.length > 0) {
    var startRow = tab.getLastRow() + 1;
    tab.getRange(startRow, 1, newRows.length, header.length).setValues(newRows);
  }
}

/**
 * Appends merge results to an existing log tab, creating it
 * with a header row if it doesn't exist yet. Used when
 * accumulating results across many target sheets.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} tabName
 * @param {string} sheetName
 * @param {{ filled: Array, conflicts: Array, errors: Array }} results
 */
function appendToResultsLog_(spreadsheet, tabName, sheetName, results) {
  var tab = spreadsheet.getSheetByName(tabName);
  var header = ['Type', 'Spreadsheet', 'Row', 'Column', 'Existing Value', 'New Value'];

  if (!tab) {
    tab = spreadsheet.insertSheet(tabName);
    tab.getRange(1, 1, 1, header.length).setValues([header]);
    formatHeaderRow_(tab, header.length);
    tab.setFrozenRows(1);
  }

  var newRows = [];

  for (var i = 0; i < results.filled.length; i++) {
    var f = results.filled[i];
    newRows.push(['Filled', sheetName, f.row, f.column, f.existingValue, f.newValue]);
  }
  for (var j = 0; j < results.conflicts.length; j++) {
    var c = results.conflicts[j];
    newRows.push(['Conflict', sheetName, c.row, c.column, c.existingValue, c.newValue]);
  }
  for (var k = 0; k < results.errors.length; k++) {
    var e = results.errors[k];
    newRows.push(['Error', sheetName, e.row, e.column, e.existingValue, e.newValue]);
  }

  if (newRows.length > 0) {
    var startRow = tab.getLastRow() + 1;
    tab.getRange(startRow, 1, newRows.length, header.length).setValues(newRows);
  }
}
