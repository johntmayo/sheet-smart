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
 * @return {{ sourceId: string, matchColumn: string, folderId: string, masterId: string, renameFrom: string, renameTo: string, columnMap: Array<{source: string, target: string}> }}
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
    masterId:         settings['Master Spreadsheet'] || '',
    sourceId:         settings['External Source'] || settings['Source Spreadsheet'] || '',
    userSheetId:      settings['User Sheet'] || '',
    folderId:         settings['User Sheets Folder'] || settings['Target Folder'] || '',
    matchColumn:      settings['Match Column'] || '',
    renameFrom:       settings['Rename Column - From'] || '',
    renameTo:         settings['Rename Column - To'] || '',
    columnMap:        columnMap,
    sensitiveColumns: parseSensitiveColumns_(settings['Sensitive Columns'] || '')
  };
}

/**
 * Parses the "Sensitive Columns" setting value (a comma-separated list
 * of column headers) into a trimmed array of header names. Blank
 * entries are removed.
 *
 * @param {string} raw
 * @return {Array<string>}
 */
function parseSensitiveColumns_(raw) {
  if (!raw) return [];
  return String(raw).split(',')
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s !== ''; });
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
 * @param {Array<string>=} virtualAddedColumns Column headers that the caller
 *        just "added" in dry-run mode (i.e. not actually written to the sheet
 *        yet). Treated as if present in the target header row so fills can be
 *        logged instead of reported as missing-column errors.
 * @return {{ filled: Array, conflicts: Array, errors: Array }}
 */
function mergeIntoTarget_(targetSheet, sourceLookup, matchColumn, columnMap, dryRun, virtualAddedColumns) {
  var filled = [];
  var conflicts = [];
  var errors = [];

  var data = targetSheet.getDataRange().getValues();
  if (data.length === 0) {
    errors.push({ row: 0, column: '', existingValue: '', newValue: 'Target sheet is empty' });
    return { filled: filled, conflicts: conflicts, errors: errors };
  }

  var targetHeaders = data[0].map(function (h) { return String(h).trim(); });

  // Dry-run bridge: addColumnsToTarget_ doesn't actually write headers in dry
  // mode, so those columns are absent from the sheet we just read. Treat them
  // as present here; their per-row cells will index out of the real row array
  // and resolve to undefined, which the blank check below handles correctly.
  if (virtualAddedColumns && virtualAddedColumns.length) {
    for (var v = 0; v < virtualAddedColumns.length; v++) {
      var vcol = String(virtualAddedColumns[v]).trim();
      if (vcol !== '' && targetHeaders.indexOf(vcol) === -1) {
        targetHeaders.push(vcol);
      }
    }
  }

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

/**
 * Renames a header in the first row of a target sheet.
 * Only the header row is changed; data rows are untouched.
 *
 * @param {SpreadsheetApp.Sheet} targetSheet
 * @param {string} oldHeader
 * @param {string} newHeader
 * @param {boolean} dryRun
 * @return {{ renamed: boolean, skipped: boolean, error: string, note: string }}
 */
function renameColumnInTarget_(targetSheet, oldHeader, newHeader, dryRun) {
  var data = targetSheet.getDataRange().getValues();
  if (data.length === 0) {
    return { renamed: false, skipped: false, error: 'Target sheet is empty', note: '' };
  }

  var oldName = String(oldHeader).trim();
  var newName = String(newHeader).trim();
  if (oldName === '' || newName === '') {
    return { renamed: false, skipped: false, error: 'Old and new header names are required', note: '' };
  }
  if (oldName === newName) {
    return { renamed: false, skipped: true, error: '', note: 'Old and new header are identical' };
  }

  var headers = data[0].map(function (h) { return String(h).trim(); });
  var oldIdx = headers.indexOf(oldName);
  if (oldIdx === -1) {
    return { renamed: false, skipped: true, error: '', note: 'Old header not found' };
  }

  var newIdx = headers.indexOf(newName);
  if (newIdx !== -1 && newIdx !== oldIdx) {
    return { renamed: false, skipped: true, error: '', note: 'New header already exists' };
  }

  if (!dryRun) {
    data[0][oldIdx] = newName;
    targetSheet.getRange(1, 1, 1, data[0].length).setValues([data[0]]);
  }

  return { renamed: true, skipped: false, error: '', note: '' };
}

/**
 * Appends folder-wide rename results to a log tab.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} tabName
 * @param {string} sheetName
 * @param {string} oldHeader
 * @param {string} newHeader
 * @param {{ renamed: boolean, skipped: boolean, error: string, note: string }} result
 */
function appendToRenameLog_(spreadsheet, tabName, sheetName, oldHeader, newHeader, result) {
  var tab = spreadsheet.getSheetByName(tabName);
  var header = ['Type', 'Spreadsheet', 'Old Header', 'New Header', 'Notes'];
  if (!tab) {
    tab = spreadsheet.insertSheet(tabName);
    tab.getRange(1, 1, 1, header.length).setValues([header]);
    formatHeaderRow_(tab, header.length);
    tab.setFrozenRows(1);
  }

  var rowType = result.renamed ? 'Renamed' : (result.skipped ? 'Skipped' : 'Error');
  var note = result.error || result.note || '';
  tab.getRange(tab.getLastRow() + 1, 1, 1, header.length).setValues([
    [rowType, sheetName, oldHeader, newHeader, note]
  ]);
}

/**
 * Appends rows from master into a target user sheet for residents whose
 * ZoneName matches the target's detected zone but whose resident_id
 * isn't already present in the target. The target's assigned zone is
 * inferred by taking the most common non-blank value in its ZoneName
 * column (same approach as the Phase 1 audit).
 *
 * Column values on the new rows are populated by header-name join:
 * every column where the target header exactly matches a master header
 * is filled from master. Columns present in the target but absent in
 * master are left blank on the new rows.
 *
 * Pure addition. Existing rows are never modified, and rows are never
 * removed. Duplicate resident_ids in master for the same zone are
 * appended once (the first occurrence wins); on subsequent runs, the
 * resident_id will already exist in the target and be skipped.
 *
 * @param {SpreadsheetApp.Sheet} targetSheet
 * @param {Array<Array>} masterData            Full 2D array of master, including header row
 * @param {Array<string>} masterHeaders        Trimmed master header row
 * @param {Array<string>} sensitiveColumns     Column headers considered privacy-sensitive
 * @param {boolean} dryRun                     If true, compute only — never write
 * @return {{ appended: Array, flagged: Array, errors: Array, detectedZone: string }}
 */
function appendMissingRowsToSheet_(targetSheet, masterData, masterHeaders, sensitiveColumns, dryRun) {
  var result = {
    appended: [],
    flagged: [],
    errors: [],
    detectedZone: ''
  };

  var targetData = targetSheet.getDataRange().getValues();
  if (targetData.length === 0) {
    result.errors.push({ message: 'Target sheet is empty (no header row)' });
    return result;
  }

  var targetHeaders = targetData[0].map(function (h) { return String(h).trim(); });
  var targetZoneCol = targetHeaders.indexOf('ZoneName');
  var targetIdCol = targetHeaders.indexOf('resident_id');

  if (targetZoneCol === -1) {
    result.errors.push({ message: 'Target sheet has no ZoneName column' });
    return result;
  }
  if (targetIdCol === -1) {
    result.errors.push({ message: 'Target sheet has no resident_id column' });
    return result;
  }

  var existing = {};
  var zoneCounts = {};
  for (var r = 1; r < targetData.length; r++) {
    var row = targetData[r];

    var id = String(row[targetIdCol] || '').trim();
    if (id !== '' && id !== 'undefined' && id !== 'null') {
      existing[id] = true;
    }

    var zv = String(row[targetZoneCol] || '').trim();
    if (zv !== '') zoneCounts[zv] = (zoneCounts[zv] || 0) + 1;
  }

  var detectedZone = '';
  var topCount = 0;
  Object.keys(zoneCounts).forEach(function (z) {
    if (zoneCounts[z] > topCount) {
      detectedZone = z;
      topCount = zoneCounts[z];
    }
  });
  if (detectedZone === '') {
    result.errors.push({ message: 'No zone detected (ZoneName column has no non-blank values)' });
    return result;
  }
  result.detectedZone = detectedZone;

  var masterIdCol = masterHeaders.indexOf('resident_id');
  var masterZoneCol = masterHeaders.indexOf('ZoneName');
  var masterNameCol = masterHeaders.indexOf('Resident Name');
  if (masterIdCol === -1) {
    result.errors.push({ message: 'Master has no resident_id column' });
    return result;
  }
  if (masterZoneCol === -1) {
    result.errors.push({ message: 'Master has no ZoneName column' });
    return result;
  }

  // Header-name join: target column index -> master column index (or -1 if
  // the target column has no matching master header).
  var colMap = [];
  for (var tc = 0; tc < targetHeaders.length; tc++) {
    var h = targetHeaders[tc];
    colMap.push(h === '' ? -1 : masterHeaders.indexOf(h));
  }

  var sensitiveMasterCols = [];
  var sensitiveColumnNames = [];
  for (var s = 0; s < sensitiveColumns.length; s++) {
    var sName = String(sensitiveColumns[s]).trim();
    if (sName === '') continue;
    var sIdx = masterHeaders.indexOf(sName);
    if (sIdx !== -1) {
      sensitiveMasterCols.push(sIdx);
      sensitiveColumnNames.push(sName);
    }
  }

  var newRows = [];
  for (var mr = 1; mr < masterData.length; mr++) {
    var masterRow = masterData[mr];
    var masterId = String(masterRow[masterIdCol] || '').trim();
    if (masterId === '' || masterId === 'undefined' || masterId === 'null') continue;

    var masterZone = String(masterRow[masterZoneCol] || '').trim();
    if (masterZone !== detectedZone) continue;
    if (existing[masterId]) continue;

    var newRow = [];
    for (var c = 0; c < targetHeaders.length; c++) {
      var src = colMap[c];
      newRow.push(src === -1 ? '' : masterRow[src]);
    }

    var residentName = (masterNameCol !== -1) ? String(masterRow[masterNameCol] || '').trim() : '';

    var flaggedCols = [];
    for (var sc = 0; sc < sensitiveMasterCols.length; sc++) {
      var sv = masterRow[sensitiveMasterCols[sc]];
      if (sv !== '' && sv !== null && sv !== undefined) {
        flaggedCols.push(sensitiveColumnNames[sc]);
      }
    }

    result.appended.push({
      residentId: masterId,
      residentName: residentName,
      masterRow: mr + 1
    });

    if (flaggedCols.length > 0) {
      result.flagged.push({
        residentId: masterId,
        residentName: residentName,
        flaggedColumns: flaggedCols.join(', ')
      });
    }

    newRows.push(newRow);

    // Guard against duplicate master rows with the same resident_id in the
    // same zone — first occurrence wins for this run.
    existing[masterId] = true;
  }

  if (!dryRun && newRows.length > 0) {
    var startRow = targetSheet.getLastRow() + 1;
    targetSheet.getRange(startRow, 1, newRows.length, targetHeaders.length).setValues(newRows);
  }

  return result;
}

/**
 * Appends rows-appended results to a log tab, creating the tab with
 * a header row on first write. Errors are recorded inline.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} tabName
 * @param {string} sheetName
 * @param {string} detectedZone
 * @param {{ appended: Array, flagged: Array, errors: Array, detectedZone: string }} result
 */
function appendToMissingRowsLog_(spreadsheet, tabName, sheetName, detectedZone, result) {
  var tab = spreadsheet.getSheetByName(tabName);
  var header = ['Spreadsheet', 'Detected Zone', 'resident_id', 'Resident Name', 'Master Row #', 'Status'];
  if (!tab) {
    tab = spreadsheet.insertSheet(tabName);
    tab.getRange(1, 1, 1, header.length).setValues([header]);
    formatHeaderRow_(tab, header.length);
    tab.setFrozenRows(1);
  }

  var newRows = [];
  for (var i = 0; i < result.appended.length; i++) {
    var a = result.appended[i];
    newRows.push([sheetName, detectedZone, a.residentId, a.residentName, a.masterRow, 'Appended']);
  }
  for (var j = 0; j < result.errors.length; j++) {
    newRows.push([sheetName, detectedZone || '(none)', '', '', '', 'Error: ' + result.errors[j].message]);
  }

  if (newRows.length > 0) {
    var startRow = tab.getLastRow() + 1;
    tab.getRange(startRow, 1, newRows.length, header.length).setValues(newRows);
  }
}

/**
 * Appends flagged-sensitive-data entries to a log tab, creating the
 * tab with a header row on first write. Only rows that had a non-blank
 * value in at least one sensitive column are written here.
 *
 * @param {SpreadsheetApp.Spreadsheet} spreadsheet
 * @param {string} tabName
 * @param {string} sheetName
 * @param {string} detectedZone
 * @param {{ flagged: Array }} result
 */
function appendToFlaggedSensitiveLog_(spreadsheet, tabName, sheetName, detectedZone, result) {
  var tab = spreadsheet.getSheetByName(tabName);
  var header = ['Spreadsheet', 'Detected Zone', 'resident_id', 'Resident Name', 'Flagged Columns'];
  if (!tab) {
    tab = spreadsheet.insertSheet(tabName);
    tab.getRange(1, 1, 1, header.length).setValues([header]);
    formatHeaderRow_(tab, header.length);
    tab.setFrozenRows(1);
  }

  var newRows = [];
  for (var i = 0; i < result.flagged.length; i++) {
    var f = result.flagged[i];
    newRows.push([sheetName, detectedZone, f.residentId, f.residentName, f.flaggedColumns]);
  }

  if (newRows.length > 0) {
    var startRow = tab.getLastRow() + 1;
    tab.getRange(startRow, 1, newRows.length, header.length).setValues(newRows);
  }
}
