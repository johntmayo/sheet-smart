// ============================================================
// Corrections.gs — UI & Entry Points for Sheet Smart
// ============================================================
// Container-bound script inside the "Sheet Smart Config"
// spreadsheet. Provides a custom menu with these actions:
//
//   Import → Master              bring data INTO the master from
//                                an external source spreadsheet
//
//   Push → User Sheet            push data FROM the master into
//                                a single user spreadsheet
//                                (fills blank cells in existing rows)
//
//   Push → User Sheets Folder    push data FROM the master into
//                                every sheet in a Drive folder
//                                (fills blank cells in existing rows)
//
//   Push Missing Rows → User Sheet
//                                append rows from master whose
//                                ZoneName matches the sheet's zone
//                                and whose resident_id isn't present
//
//   Push Missing Rows → User Sheets Folder
//                                same, across every sheet in a folder
//
//   Rename Columns → Folder      rename one header across every
//                                sheet in a Drive folder
//
//   Set Up Config Tabs           format Settings & Column Mapping
//                                with labels and instructions
//
// Cell-fill operations automatically add any missing target columns
// before filling. Row-append operations never modify or remove
// existing rows. All heavy lifting is in MergeEngine.gs.
// ============================================================

/**
 * Adds the Sheet Smart menu to the spreadsheet menu bar.
 * Runs automatically when the spreadsheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sheet Smart')
    .addItem('Import → Master', 'importToMaster')
    .addItem('Import → Master (Dry Run)', 'importToMasterDryRun')
    .addSeparator()
    .addItem('Push → User Sheet', 'pushToUserSheet')
    .addItem('Push → User Sheet (Dry Run)', 'pushToUserSheetDryRun')
    .addSeparator()
    .addItem('Push → User Sheets Folder', 'pushToFolder')
    .addItem('Push → User Sheets Folder (Dry Run)', 'pushToFolderDryRun')
    .addSeparator()
    .addItem('Push Missing Rows → User Sheet', 'pushMissingRowsToUserSheet')
    .addItem('Push Missing Rows → User Sheet (Dry Run)', 'pushMissingRowsToUserSheetDryRun')
    .addSeparator()
    .addItem('Push Missing Rows → User Sheets Folder', 'pushMissingRowsToFolder')
    .addItem('Push Missing Rows → User Sheets Folder (Dry Run)', 'pushMissingRowsToFolderDryRun')
    .addSeparator()
    .addItem('Rename Columns → User Sheets Folder', 'renameColumnsInFolder')
    .addItem('Rename Columns → User Sheets Folder (Dry Run)', 'renameColumnsInFolderDryRun')
    .addSeparator()
    .addItem('Set Up Config Tabs', 'setupConfigTabs')
    .addToUi();
}

// -------  Import → Master  -------
//
// Reads data FROM the External Source spreadsheet and writes it
// INTO the Master. Use this when new data has arrived from an
// outside source (e.g. a sales feed, a contractor's sheet) and
// you want to bring it into the master.

function importToMaster()        { runImport_(false); }
function importToMasterDryRun()  { runImport_(true);  }

function runImport_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config   = readMergeConfig_(configSs);

    if (!config.sourceId)    throw new Error('External Source is not set in the Settings tab.');
    if (!config.masterId)    throw new Error('Master Spreadsheet is not set in the Settings tab.');
    if (!config.matchColumn) throw new Error('Match Column is not set in the Settings tab.');
    if (config.columnMap.length === 0) throw new Error(
      'Column Mapping tab has no valid rows. Add at least one Source Column → Target Column pair.'
    );

    var sourceSs     = SpreadsheetApp.openById(config.sourceId);
    var sourceSheet  = sourceSs.getSheets()[0];
    var sourceData   = sourceSheet.getDataRange().getValues();
    var sourceHdrs   = sourceData[0].map(function (h) { return String(h).trim(); });
    var sourceLookup = buildSourceLookup_(sourceData, sourceHdrs, config.matchColumn);

    var masterSs    = SpreadsheetApp.openById(config.masterId);
    var masterSheet = masterSs.getSheets()[0];

    var targetCols  = config.columnMap.map(function (m) { return m.target; });
    var addResult   = addColumnsToTarget_(masterSheet, targetCols, dryRun);
    var virtualCols = addResult.added.map(function (a) { return a.column; });
    var mergeResult = mergeIntoTarget_(masterSheet, sourceLookup, config.matchColumn, config.columnMap, dryRun, virtualCols);

    var logTab = dryRun ? 'Dry Run - Import' : 'Last Import';
    var old = configSs.getSheetByName(logTab);
    if (old) configSs.deleteSheet(old);
    appendToSyncLog_(configSs, logTab, masterSs.getName(), addResult, mergeResult);

    var prefix = dryRun ? 'DRY RUN — Import → Master\n\n' : 'Import → Master complete.\n\n';
    ui.alert(
      prefix +
      'Columns added: '               + addResult.added.length + '\n' +
      'Cells filled: '                + mergeResult.filled.length + '\n' +
      'Conflicts (not overwritten): ' + mergeResult.conflicts.length + '\n' +
      'Errors: '                      + (addResult.errors.length + mergeResult.errors.length) + '\n\n' +
      'See the "' + logTab + '" tab for details.'
    );
  } catch (e) {
    ui.alert('Import → Master failed:\n\n' + e.message);
  }
}

// -------  Push → User Sheet  -------
//
// Reads data FROM the Master and writes it INTO a single user
// spreadsheet. Use this to push master data to one specific sheet
// without touching the rest of the folder.

function pushToUserSheet()       { runPushToSheet_(false); }
function pushToUserSheetDryRun() { runPushToSheet_(true);  }

function runPushToSheet_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config   = readMergeConfig_(configSs);

    if (!config.masterId)     throw new Error('Master Spreadsheet is not set in the Settings tab.');
    if (!config.userSheetId)  throw new Error('User Sheet is not set in the Settings tab.');
    if (!config.matchColumn)  throw new Error('Match Column is not set in the Settings tab.');
    if (config.columnMap.length === 0) throw new Error(
      'Column Mapping tab has no valid rows. Add at least one Source Column → Target Column pair.'
    );

    var masterSs     = SpreadsheetApp.openById(config.masterId);
    var masterSheet  = masterSs.getSheets()[0];
    var masterData   = masterSheet.getDataRange().getValues();
    var masterHdrs   = masterData[0].map(function (h) { return String(h).trim(); });
    var sourceLookup = buildSourceLookup_(masterData, masterHdrs, config.matchColumn);

    var userSs    = SpreadsheetApp.openById(config.userSheetId);
    var userSheet = userSs.getSheets()[0];

    var targetCols  = config.columnMap.map(function (m) { return m.target; });
    var addResult   = addColumnsToTarget_(userSheet, targetCols, dryRun);
    var virtualCols = addResult.added.map(function (a) { return a.column; });
    var mergeResult = mergeIntoTarget_(userSheet, sourceLookup, config.matchColumn, config.columnMap, dryRun, virtualCols);

    var logTab = dryRun ? 'Dry Run - Push User Sheet' : 'Last Push - User Sheet';
    var old = configSs.getSheetByName(logTab);
    if (old) configSs.deleteSheet(old);
    appendToSyncLog_(configSs, logTab, userSs.getName(), addResult, mergeResult);

    var prefix = dryRun ? 'DRY RUN — Push → User Sheet\n\n' : 'Push → User Sheet complete.\n\n';
    ui.alert(
      prefix +
      'Columns added: '               + addResult.added.length + '\n' +
      'Cells filled: '                + mergeResult.filled.length + '\n' +
      'Conflicts (not overwritten): ' + mergeResult.conflicts.length + '\n' +
      'Errors: '                      + (addResult.errors.length + mergeResult.errors.length) + '\n\n' +
      'See the "' + logTab + '" tab for details.'
    );
  } catch (e) {
    ui.alert('Push → User Sheet failed:\n\n' + e.message);
  }
}

// -------  Push → User Sheets Folder  -------
//
// Reads data FROM the Master and writes it INTO every Google
// Sheet in the configured Drive folder. Use this to propagate
// master data out to all ~55 user sheets at once.

function pushToFolder()       { runPushToFolder_(false); }
function pushToFolderDryRun() { runPushToFolder_(true);  }

function runPushToFolder_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config   = readMergeConfig_(configSs);

    if (!config.masterId)    throw new Error('Master Spreadsheet is not set in the Settings tab.');
    if (!config.folderId)    throw new Error('User Sheets Folder is not set in the Settings tab.');
    if (!config.matchColumn) throw new Error('Match Column is not set in the Settings tab.');
    if (config.columnMap.length === 0) throw new Error(
      'Column Mapping tab has no valid rows. Add at least one Source Column → Target Column pair.'
    );

    var masterSs     = SpreadsheetApp.openById(config.masterId);
    var masterSheet  = masterSs.getSheets()[0];
    var masterData   = masterSheet.getDataRange().getValues();
    var masterHdrs   = masterData[0].map(function (h) { return String(h).trim(); });
    var sourceLookup = buildSourceLookup_(masterData, masterHdrs, config.matchColumn);

    var targetCols = config.columnMap.map(function (m) { return m.target; });

    var folder = DriveApp.getFolderById(config.folderId);
    var iter   = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    var files  = [];
    while (iter.hasNext()) files.push(iter.next());
    files.sort(function (a, b) { return a.getName().localeCompare(b.getName()); });

    var logTab = dryRun ? 'Dry Run - Push Folder' : 'Last Push - Folder';
    var old = configSs.getSheetByName(logTab);
    if (old) configSs.deleteSheet(old);

    var totalCols = 0, totalFills = 0, totalConflicts = 0, totalErrors = 0;

    for (var i = 0; i < files.length; i++) {
      var file     = files[i];
      var fileName = file.getName();
      try {
        var ss    = SpreadsheetApp.openById(file.getId());
        var sheet = ss.getSheets()[0];

        var addResult   = addColumnsToTarget_(sheet, targetCols, dryRun);
        var virtualCols = addResult.added.map(function (a) { return a.column; });
        var mergeResult = mergeIntoTarget_(sheet, sourceLookup, config.matchColumn, config.columnMap, dryRun, virtualCols);

        appendToSyncLog_(configSs, logTab, fileName, addResult, mergeResult);
        totalCols      += addResult.added.length;
        totalFills     += mergeResult.filled.length;
        totalConflicts += mergeResult.conflicts.length;
        totalErrors    += addResult.errors.length + mergeResult.errors.length;
      } catch (e) {
        appendToSyncLog_(configSs, logTab, fileName,
          { added: [], skipped: [], errors: [] },
          { filled: [], conflicts: [], errors: [{ row: 0, column: '', existingValue: '', newValue: e.message }] }
        );
        totalErrors++;
      }
    }

    var prefix = dryRun ? 'DRY RUN — Push → Folder\n\n' : 'Push → Folder complete.\n\n';
    ui.alert(
      prefix +
      'Sheets processed: '            + files.length + '\n' +
      'Total columns added: '         + totalCols + '\n' +
      'Total cells filled: '          + totalFills + '\n' +
      'Conflicts (not overwritten): ' + totalConflicts + '\n' +
      'Total errors: '                + totalErrors + '\n\n' +
      'See the "' + logTab + '" tab for details.'
    );
  } catch (e) {
    ui.alert('Push → Folder failed:\n\n' + e.message);
  }
}

// -------  Push Missing Rows → User Sheet  -------
//
// For a single user sheet: detects its zone (mode of the ZoneName
// column), finds master rows whose ZoneName matches the detected
// zone, and appends the ones whose resident_id isn't already in the
// sheet. Columns are filled by header-name join from master.
//
// Rows whose master record has a non-blank value in any column listed
// under "Sensitive Columns" in Settings are still appended, but also
// recorded in the "Flagged - Sensitive Data" log tab for admin review.

function pushMissingRowsToUserSheet()       { runPushMissingRowsToSheet_(false); }
function pushMissingRowsToUserSheetDryRun() { runPushMissingRowsToSheet_(true);  }

function runPushMissingRowsToSheet_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config   = readMergeConfig_(configSs);

    if (!config.masterId)    throw new Error('Master Spreadsheet is not set in the Settings tab.');
    if (!config.userSheetId) throw new Error('User Sheet is not set in the Settings tab.');

    var masterSs    = SpreadsheetApp.openById(config.masterId);
    var masterSheet = masterSs.getSheets()[0];
    var masterData  = masterSheet.getDataRange().getValues();
    var masterHdrs  = masterData[0].map(function (h) { return String(h).trim(); });

    var userSs    = SpreadsheetApp.openById(config.userSheetId);
    var userSheet = userSs.getSheets()[0];

    var appendLogTab = dryRun ? 'Dry Run - Push Missing Rows' : 'Last Push - Missing Rows';
    var flagLogTab   = dryRun ? 'Dry Run - Flagged Sensitive Data' : 'Flagged - Sensitive Data';
    var oldA = configSs.getSheetByName(appendLogTab); if (oldA) configSs.deleteSheet(oldA);
    var oldF = configSs.getSheetByName(flagLogTab);   if (oldF) configSs.deleteSheet(oldF);

    var result = appendMissingRowsToSheet_(userSheet, masterData, masterHdrs, config.sensitiveColumns, dryRun);
    appendToMissingRowsLog_(configSs, appendLogTab, userSs.getName(), result.detectedZone, result);
    if (result.flagged.length > 0) {
      appendToFlaggedSensitiveLog_(configSs, flagLogTab, userSs.getName(), result.detectedZone, result);
    }

    var prefix = dryRun ? 'DRY RUN — Push Missing Rows → User Sheet\n\n' : 'Push Missing Rows → User Sheet complete.\n\n';
    ui.alert(
      prefix +
      'Detected zone: '                      + (result.detectedZone || '(none)') + '\n' +
      'Rows appended: '                      + result.appended.length + '\n' +
      'Flagged (sensitive data present): '   + result.flagged.length + '\n' +
      'Errors: '                             + result.errors.length + '\n\n' +
      'See the "' + appendLogTab + '" tab' +
      (result.flagged.length > 0 ? ' and "' + flagLogTab + '" tab' : '') + '.'
    );
  } catch (e) {
    ui.alert('Push Missing Rows → User Sheet failed:\n\n' + e.message);
  }
}

// -------  Push Missing Rows → User Sheets Folder  -------
//
// Same as above but iterates every sheet in the configured folder.
// Each sheet's zone is detected independently. Results for all
// sheets accumulate into one "Last Push - Missing Rows" log tab
// and one "Flagged - Sensitive Data" log tab.

function pushMissingRowsToFolder()       { runPushMissingRowsToFolder_(false); }
function pushMissingRowsToFolderDryRun() { runPushMissingRowsToFolder_(true);  }

function runPushMissingRowsToFolder_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config   = readMergeConfig_(configSs);

    if (!config.masterId) throw new Error('Master Spreadsheet is not set in the Settings tab.');
    if (!config.folderId) throw new Error('User Sheets Folder is not set in the Settings tab.');

    var masterSs    = SpreadsheetApp.openById(config.masterId);
    var masterSheet = masterSs.getSheets()[0];
    var masterData  = masterSheet.getDataRange().getValues();
    var masterHdrs  = masterData[0].map(function (h) { return String(h).trim(); });

    var folder = DriveApp.getFolderById(config.folderId);
    var iter   = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    var files  = [];
    while (iter.hasNext()) files.push(iter.next());
    files.sort(function (a, b) { return a.getName().localeCompare(b.getName()); });

    var appendLogTab = dryRun ? 'Dry Run - Push Missing Rows' : 'Last Push - Missing Rows';
    var flagLogTab   = dryRun ? 'Dry Run - Flagged Sensitive Data' : 'Flagged - Sensitive Data';
    var oldA = configSs.getSheetByName(appendLogTab); if (oldA) configSs.deleteSheet(oldA);
    var oldF = configSs.getSheetByName(flagLogTab);   if (oldF) configSs.deleteSheet(oldF);

    var totalAppended = 0, totalFlagged = 0, totalErrors = 0;

    for (var i = 0; i < files.length; i++) {
      var file     = files[i];
      var fileName = file.getName();
      try {
        var ss     = SpreadsheetApp.openById(file.getId());
        var sheet  = ss.getSheets()[0];
        var result = appendMissingRowsToSheet_(sheet, masterData, masterHdrs, config.sensitiveColumns, dryRun);

        appendToMissingRowsLog_(configSs, appendLogTab, fileName, result.detectedZone, result);
        if (result.flagged.length > 0) {
          appendToFlaggedSensitiveLog_(configSs, flagLogTab, fileName, result.detectedZone, result);
        }

        totalAppended += result.appended.length;
        totalFlagged  += result.flagged.length;
        totalErrors   += result.errors.length;
      } catch (e) {
        appendToMissingRowsLog_(configSs, appendLogTab, fileName, '', {
          appended: [], flagged: [], errors: [{ message: e.message }], detectedZone: ''
        });
        totalErrors++;
      }
    }

    var prefix = dryRun ? 'DRY RUN — Push Missing Rows → Folder\n\n' : 'Push Missing Rows → Folder complete.\n\n';
    ui.alert(
      prefix +
      'Sheets processed: '                   + files.length + '\n' +
      'Total rows appended: '                + totalAppended + '\n' +
      'Flagged (sensitive data present): '   + totalFlagged + '\n' +
      'Errors: '                             + totalErrors + '\n\n' +
      'See the "' + appendLogTab + '" tab' +
      (totalFlagged > 0 ? ' and "' + flagLogTab + '" tab' : '') + '.'
    );
  } catch (e) {
    ui.alert('Push Missing Rows → Folder failed:\n\n' + e.message);
  }
}

// -------  Rename Columns → User Sheets Folder  -------
//
// Renames a single header across every Google Sheet in the
// configured user folder. Only row-1 header cells are changed.
// Data rows are never edited by this operation.

function renameColumnsInFolder()       { runRenameColumnsInFolder_(false); }
function renameColumnsInFolderDryRun() { runRenameColumnsInFolder_(true);  }

function runRenameColumnsInFolder_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config = readMergeConfig_(configSs);

    if (!config.folderId)   throw new Error('User Sheets Folder is not set in the Settings tab.');
    if (!config.renameFrom) throw new Error('Rename Column - From is not set in the Settings tab.');
    if (!config.renameTo)   throw new Error('Rename Column - To is not set in the Settings tab.');

    var folder = DriveApp.getFolderById(config.folderId);
    var iter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    var files = [];
    while (iter.hasNext()) files.push(iter.next());
    files.sort(function (a, b) { return a.getName().localeCompare(b.getName()); });

    var logTab = dryRun ? 'Dry Run - Rename Folder' : 'Last Rename - Folder';
    var old = configSs.getSheetByName(logTab);
    if (old) configSs.deleteSheet(old);

    var totalRenamed = 0, totalSkipped = 0, totalErrors = 0;

    for (var i = 0; i < files.length; i++) {
      var file = files[i];
      var fileName = file.getName();
      try {
        var ss = SpreadsheetApp.openById(file.getId());
        var sheet = ss.getSheets()[0];
        var result = renameColumnInTarget_(sheet, config.renameFrom, config.renameTo, dryRun);

        appendToRenameLog_(configSs, logTab, fileName, config.renameFrom, config.renameTo, result);

        if (result.renamed) totalRenamed++;
        else if (result.skipped) totalSkipped++;
        else totalErrors++;
      } catch (e) {
        appendToRenameLog_(configSs, logTab, fileName, config.renameFrom, config.renameTo, {
          renamed: false,
          skipped: false,
          error: e.message,
          note: ''
        });
        totalErrors++;
      }
    }

    var prefix = dryRun ? 'DRY RUN — Rename Columns → Folder\n\n' : 'Rename Columns → Folder complete.\n\n';
    ui.alert(
      prefix +
      'Sheets processed: ' + files.length + '\n' +
      'Headers renamed: ' + totalRenamed + '\n' +
      'Skipped: ' + totalSkipped + '\n' +
      'Errors: ' + totalErrors + '\n\n' +
      'See the "' + logTab + '" tab for details.'
    );
  } catch (e) {
    ui.alert('Rename Columns → User Sheets Folder failed:\n\n' + e.message);
  }
}

// -------  Set Up Config Tabs  -------
//
// Formats the Settings and Column Mapping tabs with labels,
// column headers, and plain-English instructions. Run this once
// when setting up a new Sheet Smart Config spreadsheet, or any
// time you want to restore the correct structure.

function setupConfigTabs() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Set Up Config Tabs',
    'This will format the Settings and Column Mapping tabs with labels and instructions.\n\n' +
    '• Existing values in the Settings tab (column B) will be preserved.\n' +
    '• The Column Mapping tab header row will be updated; existing mapping rows will not change.\n\n' +
    'Continue?',
    ui.ButtonSet.OK_CANCEL
  );
  if (response !== ui.Button.OK) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettingsTab_(ss);
  setupColumnMappingTab_(ss);

  ui.alert(
    'Config tabs are ready.\n\n' +
    'Next steps:\n' +
    '1. Fill in the Value column in the Settings tab.\n' +
    '2. Add your column pairs to the Column Mapping tab.\n' +
    '3. Run a Dry Run first to preview results before a live operation.'
  );
}

/**
 * Writes (or rewrites) the Settings tab with labeled rows,
 * descriptions, and proper formatting. Preserves any values
 * already in column B.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss
 */
function setupSettingsTab_(ss) {
  var tab = ss.getSheetByName('Settings');
  if (!tab) tab = ss.insertSheet('Settings');

  // Preserve existing values before clearing
  var existingValues = {};
  var existingData = tab.getDataRange().getValues();
  for (var i = 0; i < existingData.length; i++) {
    var key = String(existingData[i][0]).trim();
    var val = existingData[i].length > 1 ? existingData[i][1] : '';
    if (key !== '') existingValues[key] = val;
  }

  tab.clear();

  var rows = [
    ['Setting', 'Value', 'What to put here'],
    [
      'Master Spreadsheet',
      existingValues['Master Spreadsheet'] || '',
      'Your master spreadsheet ID. This is the DESTINATION for "Import → Master" ' +
      'and the SOURCE for both "Push" operations. ' +
      'Find it in the URL: docs.google.com/spreadsheets/d/[COPY THIS]/edit'
    ],
    [
      'External Source',
      existingValues['External Source'] || existingValues['Source Spreadsheet'] || '',
      'ID of a spreadsheet to import data FROM into the master. ' +
      'Only used by "Import → Master". ' +
      'Find it in the URL: docs.google.com/spreadsheets/d/[COPY THIS]/edit'
    ],
    [
      'User Sheet',
      existingValues['User Sheet'] || '',
      'ID of a single user spreadsheet to push master data TO. ' +
      'Only used by "Push → User Sheet". ' +
      'Find it in the URL: docs.google.com/spreadsheets/d/[COPY THIS]/edit'
    ],
    [
      'User Sheets Folder',
      existingValues['User Sheets Folder'] || existingValues['Target Folder'] || '',
      'ID of the Drive folder containing all user sheets. ' +
      'Only used by "Push → User Sheets Folder". ' +
      'Find it at the end of the folder URL: drive.google.com/drive/folders/[COPY THIS]'
    ],
    [
      'Match Column',
      existingValues['Match Column'] || '',
      'The column header that exists in both master and user sheets and uniquely ' +
      'identifies each row (e.g. APN). Must match exactly — same spelling and capitalization.'
    ],
    [
      'Rename Column - From',
      existingValues['Rename Column - From'] || '',
      'For "Rename Columns → User Sheets Folder": the existing header name to rename ' +
      '(exact match, including spelling and capitalization).'
    ],
    [
      'Rename Column - To',
      existingValues['Rename Column - To'] || '',
      'For "Rename Columns → User Sheets Folder": the new header name to write. ' +
      'If this header already exists in a sheet, that sheet is skipped and logged.'
    ],
    [
      'Sensitive Columns',
      existingValues['Sensitive Columns'] || '',
      'Comma-separated list of column headers whose values are considered privacy-sensitive ' +
      '(e.g. "Person Notes, Contact Notes, Address Notes"). Only used by the "Push Missing Rows" ' +
      'operations. When a missing row is appended to a user sheet and any of these columns has ' +
      'a value in master, the row is listed in the "Flagged - Sensitive Data" log tab for you ' +
      'to review. The row is still appended — this flag is informational.'
    ]
  ];

  tab.getRange(1, 1, rows.length, 3).setValues(rows);

  // Header row
  var headerRange = tab.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e8eaf6');
  tab.setFrozenRows(1);

  // Setting name column: bold
  tab.getRange(2, 1, rows.length - 1, 1).setFontWeight('bold');

  // Description column: grey italic wrapped
  tab.getRange(2, 3, rows.length - 1, 1)
    .setFontStyle('italic')
    .setFontColor('#757575')
    .setWrap(true);

  tab.setColumnWidth(1, 210);
  tab.setColumnWidth(2, 340);
  tab.setColumnWidth(3, 520);
}

/**
 * Writes (or updates) the Column Mapping tab header row with
 * formatting and an instructional cell note. Existing mapping
 * rows (row 2 onward) are never modified.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss
 */
function setupColumnMappingTab_(ss) {
  var tab = ss.getSheetByName('Column Mapping');
  if (!tab) tab = ss.insertSheet('Column Mapping');

  tab.getRange(1, 1, 1, 2).setValues([['Source Column', 'Target Column']]);

  var headerRange = tab.getRange(1, 1, 1, 2);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e8eaf6');
  tab.setFrozenRows(1);

  tab.getRange(1, 1).setNote(
    'One row per column to sync.\n\n' +
    '• Source Column: the header name in the source spreadsheet (master for Push, ' +
    'external sheet for Import).\n' +
    '• Target Column: the header name in the destination spreadsheet.\n\n' +
    'If both use the same column name, put that name in both columns.\n\n' +
    'Rows with a blank cell in either column are automatically skipped.'
  );

  tab.setColumnWidth(1, 240);
  tab.setColumnWidth(2, 240);
}
