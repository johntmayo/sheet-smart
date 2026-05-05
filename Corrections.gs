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
//   Pull Missing Rows ← User Sheet
//                                append captain-created rows from one
//                                user sheet into the master
//
//   Pull Missing Rows ← User Sheets Folder
//                                same, across every sheet in a folder
//
//   Pull Data ← User Sheet        pull captain-entered data into the
//                                master using Pull Column Policy
//
//   Pull Data ← User Sheets Folder
//                                same, across every sheet in a folder
//
//   Rename Columns → Folder      rename one header across every
//                                sheet in a Drive folder
//
//   Set Up Config Tabs           format Settings, Column Mapping,
//                                and Pull Column Policy
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
    .addItem('Open Dashboard', 'openSheetSmartDashboard')
    .addSeparator()
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
    .addItem('Pull Missing Rows ← User Sheet', 'pullMissingRowsFromUserSheet')
    .addItem('Pull Missing Rows ← User Sheet (Dry Run)', 'pullMissingRowsFromUserSheetDryRun')
    .addSeparator()
    .addItem('Pull Missing Rows ← User Sheets Folder', 'pullMissingRowsFromFolder')
    .addItem('Pull Missing Rows ← User Sheets Folder (Dry Run)', 'pullMissingRowsFromFolderDryRun')
    .addSeparator()
    .addItem('Pull Data ← User Sheet', 'pullDataFromUserSheet')
    .addItem('Pull Data ← User Sheet (Dry Run)', 'pullDataFromUserSheetDryRun')
    .addSeparator()
    .addItem('Pull Data ← User Sheets Folder', 'pullDataFromFolder')
    .addItem('Pull Data ← User Sheets Folder (Dry Run)', 'pullDataFromFolderDryRun')
    .addSeparator()
    .addItem('Rename Columns → User Sheets Folder', 'renameColumnsInFolder')
    .addItem('Rename Columns → User Sheets Folder (Dry Run)', 'renameColumnsInFolderDryRun')
    .addSeparator()
    .addItem('Set Up Config Tabs', 'setupConfigTabs')
    .addToUi();
}

/**
 * Opens the guided Sheet Smart sidebar. The existing menu operations remain
 * available, but the sidebar presents the common tasks as named workflows.
 */
function openSheetSmartDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Sheet Smart');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Returns sidebar-ready workflow summaries.
 *
 * @return {{ workflows: Array }}
 */
function getSheetSmartDashboardModel() {
  var configSs = SpreadsheetApp.getActiveSpreadsheet();
  var presets = readWorkflowPresets_(configSs);
  var workflows = [];

  for (var i = 0; i < presets.length; i++) {
    if (!presets[i].enabled) continue;
    workflows.push(buildWorkflowSidebarModel_(configSs, presets[i], false));
  }

  return { workflows: workflows };
}

/**
 * Returns details and readiness checks for one workflow.
 *
 * @param {string} workflowId
 * @return {Object}
 */
function getWorkflowDetails(workflowId) {
  var configSs = SpreadsheetApp.getActiveSpreadsheet();
  var preset = getWorkflowPreset_(configSs, workflowId);
  if (!preset) throw new Error('Workflow "' + workflowId + '" was not found.');
  return buildWorkflowSidebarModel_(configSs, preset, true);
}

/**
 * Runs a workflow in dry-run mode from the sidebar.
 *
 * @param {string} workflowId
 * @return {Object}
 */
function runWorkflowDryRun(workflowId) {
  return runWorkflowFromSidebar_(workflowId, true);
}

/**
 * Runs a workflow live from the sidebar.
 *
 * @param {string} workflowId
 * @return {Object}
 */
function runWorkflowLive(workflowId) {
  return runWorkflowFromSidebar_(workflowId, false);
}

/**
 * Dispatches a named workflow to the matching backend operation.
 *
 * @param {string} workflowId
 * @param {boolean} dryRun
 * @return {Object}
 */
function runWorkflowFromSidebar_(workflowId, dryRun) {
  var configSs = SpreadsheetApp.getActiveSpreadsheet();
  var preset = getWorkflowPreset_(configSs, workflowId);
  if (!preset) throw new Error('Workflow "' + workflowId + '" was not found.');
  if (!preset.enabled) throw new Error('Workflow "' + preset.name + '" is disabled.');
  if (preset.operation !== 'import_to_master') {
    throw new Error('Workflow operation "' + preset.operation + '" is not supported by the sidebar yet.');
  }

  var logTab = dryRun ? 'Dry Run - ' + preset.name : 'Last Run - ' + preset.name;
  var result = executeImportToMaster_(configSs, preset, dryRun, logTab);
  result.workflowId = preset.id;
  result.workflowName = preset.name;
  result.mode = dryRun ? 'Dry Run' : 'Live Run';
  return result;
}

/**
 * Builds a compact workflow model for the sidebar. When includeChecks is true,
 * spreadsheet/tab/header checks are performed so the UI can show readiness.
 *
 * @param {SpreadsheetApp.Spreadsheet} configSs
 * @param {Object} preset
 * @param {boolean} includeChecks
 * @return {Object}
 */
function buildWorkflowSidebarModel_(configSs, preset, includeChecks) {
  var checks = includeChecks ? validateImportWorkflow_(preset) : [];
  return {
    id: preset.id,
    name: preset.name,
    operation: preset.operation,
    sourceTabName: preset.sourceTabName,
    matchColumn: preset.matchColumn,
    mappingCount: preset.columnMap.length,
    mappings: preset.columnMap,
    notes: preset.notes,
    checks: checks,
    ready: checks.length === 0 || checks.every(function (check) { return check.status !== 'error'; })
  };
}

/**
 * Validates the sidebar import workflow before it runs.
 *
 * @param {Object} preset
 * @return {Array<{status: string, message: string}>}
 */
function validateImportWorkflow_(preset) {
  var checks = [];

  if (!preset.sourceId) checks.push({ status: 'error', message: 'Source spreadsheet is not set.' });
  if (!preset.masterId) checks.push({ status: 'error', message: 'Master spreadsheet is not set.' });
  if (!preset.matchColumn) checks.push({ status: 'error', message: 'Match column is not set.' });
  if (preset.columnMap.length === 0) checks.push({ status: 'error', message: 'No column mappings are configured.' });
  if (checks.length > 0) return checks;

  try {
    var sourceSs = SpreadsheetApp.openById(preset.sourceId);
    var sourceSheet = getConfiguredSheet_(sourceSs, preset.sourceTabName);
    var sourceData = sourceSheet.getDataRange().getValues();
    var sourceHeaders = sourceData.length > 0 ? sourceData[0].map(function (h) { return String(h).trim(); }) : [];
    var sourceColumns = preset.columnMap.map(function (m) { return m.source; });
    sourceColumns.push(preset.matchColumn);
    var missingSource = findMissingHeaders_(sourceHeaders, sourceColumns);

    checks.push({ status: 'ok', message: 'Source tab found: ' + sourceSheet.getName() + ' (' + Math.max(sourceData.length - 1, 0) + ' data rows).' });
    if (missingSource.length > 0) {
      checks.push({ status: 'error', message: 'Source is missing: ' + missingSource.join(', ') + '.' });
    } else {
      checks.push({ status: 'ok', message: 'Source has the match column and mapped columns.' });
    }
  } catch (e) {
    checks.push({ status: 'error', message: e.message });
  }

  try {
    var masterSs = SpreadsheetApp.openById(preset.masterId);
    var masterSheet = masterSs.getSheets()[0];
    var masterData = masterSheet.getDataRange().getValues();
    var masterHeaders = masterData.length > 0 ? masterData[0].map(function (h) { return String(h).trim(); }) : [];
    var missingMasterMatch = findMissingHeaders_(masterHeaders, [preset.matchColumn]);
    var targetColumns = preset.columnMap.map(function (m) { return m.target; });
    var missingTargets = findMissingHeaders_(masterHeaders, targetColumns);

    checks.push({ status: 'ok', message: 'Master found: ' + masterSs.getName() + ' (' + Math.max(masterData.length - 1, 0) + ' data rows).' });
    if (missingMasterMatch.length > 0) {
      checks.push({ status: 'error', message: 'Master is missing match column: ' + preset.matchColumn + '.' });
    }
    if (missingTargets.length > 0) {
      checks.push({ status: 'warning', message: 'These target columns will be added if you run live: ' + missingTargets.join(', ') + '.' });
    } else {
      checks.push({ status: 'ok', message: 'Master already has all mapped target columns.' });
    }
  } catch (e2) {
    checks.push({ status: 'error', message: e2.message });
  }

  return checks;
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

    var logTab = dryRun ? 'Dry Run - Import' : 'Last Import';
    var summary = executeImportToMaster_(configSs, config, dryRun, logTab);

    var prefix = dryRun ? 'DRY RUN — Import → Master\n\n' : 'Import → Master complete.\n\n';
    ui.alert(
      prefix +
      'Source tab: '                  + summary.sourceTab + '\n' +
      'Columns added: '               + summary.columnsAdded + '\n' +
      'Cells filled: '                + summary.cellsFilled + '\n' +
      'Conflicts (not overwritten): ' + summary.conflicts + '\n' +
      'Errors: '                      + summary.errors + '\n\n' +
      'See the "' + logTab + '" tab for details.'
    );
  } catch (e) {
    ui.alert('Import → Master failed:\n\n' + e.message);
  }
}

/**
 * Shared Import -> Master implementation used by both the legacy menu and
 * the sidebar workflows.
 *
 * @param {SpreadsheetApp.Spreadsheet} configSs
 * @param {Object} config
 * @param {boolean} dryRun
 * @param {string} logTab
 * @return {Object}
 */
function executeImportToMaster_(configSs, config, dryRun, logTab) {
  var sourceSs     = SpreadsheetApp.openById(config.sourceId);
  var sourceSheet  = getConfiguredSheet_(sourceSs, config.sourceTabName);
  var sourceData   = sourceSheet.getDataRange().getValues();
  if (sourceData.length === 0) throw new Error('Source tab "' + sourceSheet.getName() + '" is empty.');
  var sourceHdrs   = sourceData[0].map(function (h) { return String(h).trim(); });
  var sourceLookup = buildSourceLookup_(sourceData, sourceHdrs, config.matchColumn);

  var masterSs    = SpreadsheetApp.openById(config.masterId);
  var masterSheet = masterSs.getSheets()[0];

  var targetCols  = config.columnMap.map(function (m) { return m.target; });
  var addResult   = addColumnsToTarget_(masterSheet, targetCols, dryRun);
  var virtualCols = addResult.added.map(function (a) { return a.column; });
  var mergeResult = mergeIntoTarget_(masterSheet, sourceLookup, config.matchColumn, config.columnMap, dryRun, virtualCols);

  var old = configSs.getSheetByName(logTab);
  if (old) configSs.deleteSheet(old);
  appendToSyncLog_(configSs, logTab, masterSs.getName(), addResult, mergeResult);

  return {
    sourceSpreadsheet: sourceSs.getName(),
    sourceTab: sourceSheet.getName(),
    targetSpreadsheet: masterSs.getName(),
    logTab: logTab,
    columnsAdded: addResult.added.length,
    cellsFilled: mergeResult.filled.length,
    conflicts: mergeResult.conflicts.length,
    errors: addResult.errors.length + mergeResult.errors.length
  };
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

// -------  Pull Missing Rows ← User Sheet  -------
//
// For a single user sheet: finds rows whose resident_id does not already
// exist in the master and appends them to the master. Any user-sheet
// columns missing from the master are added first so captain-created
// fields are preserved.

function pullMissingRowsFromUserSheet()       { runPullMissingRowsFromSheet_(false); }
function pullMissingRowsFromUserSheetDryRun() { runPullMissingRowsFromSheet_(true);  }

function runPullMissingRowsFromSheet_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config   = readMergeConfig_(configSs);

    if (!config.masterId)    throw new Error('Master Spreadsheet is not set in the Settings tab.');
    if (!config.userSheetId) throw new Error('User Sheet is not set in the Settings tab.');

    var masterSs    = SpreadsheetApp.openById(config.masterId);
    var masterSheet = masterSs.getSheets()[0];
    var masterState = buildMasterPullState_(masterSheet);

    var userSs    = SpreadsheetApp.openById(config.userSheetId);
    var userSheet = userSs.getSheets()[0];

    var logTab = dryRun ? 'Dry Run - Pull Missing Rows' : 'Last Pull - Missing Rows';
    var old = configSs.getSheetByName(logTab);
    if (old) configSs.deleteSheet(old);

    var result = appendMissingRowsToMaster_(masterSheet, userSheet, userSs.getName(), dryRun, masterState);
    appendToPullMissingRowsLog_(configSs, logTab, userSs.getName(), result);

    var prefix = dryRun ? 'DRY RUN — Pull Missing Rows ← User Sheet\n\n' : 'Pull Missing Rows ← User Sheet complete.\n\n';
    ui.alert(
      prefix +
      'Columns added to master: ' + result.columnsAdded.length + '\n' +
      'Rows appended to master: ' + result.appended.length + '\n' +
      'Skipped: '                + result.skipped.length + '\n' +
      'Errors: '                 + result.errors.length + '\n\n' +
      'See the "' + logTab + '" tab for details.'
    );
  } catch (e) {
    ui.alert('Pull Missing Rows ← User Sheet failed:\n\n' + e.message);
  }
}

// -------  Pull Missing Rows ← User Sheets Folder  -------
//
// Same as above but iterates every sheet in the configured folder.
// Results for all sheets accumulate into one "Last Pull - Missing Rows"
// log tab.

function pullMissingRowsFromFolder()       { runPullMissingRowsFromFolder_(false); }
function pullMissingRowsFromFolderDryRun() { runPullMissingRowsFromFolder_(true);  }

function runPullMissingRowsFromFolder_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config   = readMergeConfig_(configSs);

    if (!config.masterId) throw new Error('Master Spreadsheet is not set in the Settings tab.');
    if (!config.folderId) throw new Error('User Sheets Folder is not set in the Settings tab.');

    var masterSs    = SpreadsheetApp.openById(config.masterId);
    var masterSheet = masterSs.getSheets()[0];
    var masterState = buildMasterPullState_(masterSheet);

    var folder = DriveApp.getFolderById(config.folderId);
    var iter   = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    var files  = [];
    while (iter.hasNext()) files.push(iter.next());
    files.sort(function (a, b) { return a.getName().localeCompare(b.getName()); });

    var logTab = dryRun ? 'Dry Run - Pull Missing Rows' : 'Last Pull - Missing Rows';
    var old = configSs.getSheetByName(logTab);
    if (old) configSs.deleteSheet(old);

    var totalColumnsAdded = 0, totalAppended = 0, totalSkipped = 0, totalErrors = 0;

    for (var i = 0; i < files.length; i++) {
      var file     = files[i];
      var fileName = file.getName();
      try {
        var ss     = SpreadsheetApp.openById(file.getId());
        var sheet  = ss.getSheets()[0];
        var result = appendMissingRowsToMaster_(masterSheet, sheet, fileName, dryRun, masterState);

        appendToPullMissingRowsLog_(configSs, logTab, fileName, result);

        totalColumnsAdded += result.columnsAdded.length;
        totalAppended     += result.appended.length;
        totalSkipped      += result.skipped.length;
        totalErrors       += result.errors.length;
      } catch (e) {
        appendToPullMissingRowsLog_(configSs, logTab, fileName, {
          columnsAdded: [],
          appended: [],
          skipped: [],
          errors: [{ message: e.message }]
        });
        totalErrors++;
      }
    }

    var prefix = dryRun ? 'DRY RUN — Pull Missing Rows ← Folder\n\n' : 'Pull Missing Rows ← Folder complete.\n\n';
    ui.alert(
      prefix +
      'Sheets processed: '          + files.length + '\n' +
      'Total columns added: '       + totalColumnsAdded + '\n' +
      'Total rows appended: '       + totalAppended + '\n' +
      'Skipped: '                   + totalSkipped + '\n' +
      'Errors: '                    + totalErrors + '\n\n' +
      'See the "' + logTab + '" tab for details.'
    );
  } catch (e) {
    ui.alert('Pull Missing Rows ← User Sheets Folder failed:\n\n' + e.message);
  }
}

// -------  Pull Data ← User Sheet  -------
//
// Pulls captain-entered values from one user sheet into the master.
// Existing master rows are updated according to Pull Column Policy;
// rows missing from master are appended.

function pullDataFromUserSheet()       { runPullDataFromSheet_(false); }
function pullDataFromUserSheetDryRun() { runPullDataFromSheet_(true);  }

function runPullDataFromSheet_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config   = readMergeConfig_(configSs);

    if (!config.masterId)    throw new Error('Master Spreadsheet is not set in the Settings tab.');
    if (!config.userSheetId) throw new Error('User Sheet is not set in the Settings tab.');

    var masterSs    = SpreadsheetApp.openById(config.masterId);
    var masterSheet = masterSs.getSheets()[0];
    var pullState   = buildMasterPullDataState_(masterSheet);

    var userSs    = SpreadsheetApp.openById(config.userSheetId);
    var userSheet = userSs.getSheets()[0];

    var logTab = dryRun ? 'Dry Run - Pull Data' : 'Last Pull Data';
    var old = configSs.getSheetByName(logTab);
    if (old) configSs.deleteSheet(old);

    var result = pullDataIntoMaster_(masterSheet, userSheet, userSs.getName(), config.pullColumnPolicies, dryRun, pullState);
    appendToPullDataLog_(configSs, logTab, userSs.getName(), result);

    var prefix = dryRun ? 'DRY RUN — Pull Data ← User Sheet\n\n' : 'Pull Data ← User Sheet complete.\n\n';
    ui.alert(
      prefix +
      'Columns added to master: ' + result.columnsAdded.length + '\n' +
      'Rows appended to master: ' + result.appended.length + '\n' +
      'Cells filled: '           + result.filled.length + '\n' +
      'Cells overwritten: '      + result.overwritten.length + '\n' +
      'Conflicts logged: '       + result.conflicts.length + '\n' +
      'Skipped: '                + result.skipped.length + '\n' +
      'Errors: '                 + result.errors.length + '\n\n' +
      'See the "' + logTab + '" tab for details.'
    );
  } catch (e) {
    ui.alert('Pull Data ← User Sheet failed:\n\n' + e.message);
  }
}

// -------  Pull Data ← User Sheets Folder  -------
//
// Same as above but iterates every sheet in the configured folder.

function pullDataFromFolder()       { runPullDataFromFolder_(false); }
function pullDataFromFolderDryRun() { runPullDataFromFolder_(true);  }

function runPullDataFromFolder_(dryRun) {
  var ui = SpreadsheetApp.getUi();
  try {
    var configSs = SpreadsheetApp.getActiveSpreadsheet();
    var config   = readMergeConfig_(configSs);

    if (!config.masterId) throw new Error('Master Spreadsheet is not set in the Settings tab.');
    if (!config.folderId) throw new Error('User Sheets Folder is not set in the Settings tab.');

    var masterSs    = SpreadsheetApp.openById(config.masterId);
    var masterSheet = masterSs.getSheets()[0];
    var pullState   = buildMasterPullDataState_(masterSheet);

    var folder = DriveApp.getFolderById(config.folderId);
    var iter   = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    var files  = [];
    while (iter.hasNext()) files.push(iter.next());
    files.sort(function (a, b) { return a.getName().localeCompare(b.getName()); });

    var logTab = dryRun ? 'Dry Run - Pull Data' : 'Last Pull Data';
    var old = configSs.getSheetByName(logTab);
    if (old) configSs.deleteSheet(old);

    var totalColumnsAdded = 0, totalAppended = 0, totalFilled = 0, totalOverwritten = 0;
    var totalConflicts = 0, totalSkipped = 0, totalErrors = 0;

    for (var i = 0; i < files.length; i++) {
      var file     = files[i];
      var fileName = file.getName();
      try {
        var ss     = SpreadsheetApp.openById(file.getId());
        var sheet  = ss.getSheets()[0];
        var result = pullDataIntoMaster_(masterSheet, sheet, fileName, config.pullColumnPolicies, dryRun, pullState);

        appendToPullDataLog_(configSs, logTab, fileName, result);

        totalColumnsAdded += result.columnsAdded.length;
        totalAppended     += result.appended.length;
        totalFilled       += result.filled.length;
        totalOverwritten  += result.overwritten.length;
        totalConflicts    += result.conflicts.length;
        totalSkipped      += result.skipped.length;
        totalErrors       += result.errors.length;
      } catch (e) {
        appendToPullDataLog_(configSs, logTab, fileName, {
          columnsAdded: [],
          appended: [],
          filled: [],
          overwritten: [],
          conflicts: [],
          skipped: [],
          errors: [{ message: e.message }]
        });
        totalErrors++;
      }
    }

    var prefix = dryRun ? 'DRY RUN — Pull Data ← Folder\n\n' : 'Pull Data ← Folder complete.\n\n';
    ui.alert(
      prefix +
      'Sheets processed: '          + files.length + '\n' +
      'Total columns added: '       + totalColumnsAdded + '\n' +
      'Total rows appended: '       + totalAppended + '\n' +
      'Total cells filled: '        + totalFilled + '\n' +
      'Total cells overwritten: '   + totalOverwritten + '\n' +
      'Conflicts logged: '          + totalConflicts + '\n' +
      'Skipped: '                   + totalSkipped + '\n' +
      'Errors: '                    + totalErrors + '\n\n' +
      'See the "' + logTab + '" tab for details.'
    );
  } catch (e) {
    ui.alert('Pull Data ← User Sheets Folder failed:\n\n' + e.message);
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
// Formats the Settings, Column Mapping, and Pull Column Policy tabs
// with labels, column headers, and plain-English instructions. Run
// this once when setting up a new Sheet Smart Config spreadsheet, or
// any time you want to restore the correct structure.

function setupConfigTabs() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Set Up Config Tabs',
    'This will format the Settings, Column Mapping, and Pull Column Policy tabs with labels and instructions.\n\n' +
    '• Existing values in the Settings tab (column B) will be preserved.\n' +
    '• The Column Mapping tab header row will be updated; existing mapping rows will not change.\n\n' +
    'Continue?',
    ui.ButtonSet.OK_CANCEL
  );
  if (response !== ui.Button.OK) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettingsTab_(ss);
  setupColumnMappingTab_(ss);
  setupPullColumnPolicyTab_(ss);
  setupWorkflowPresetsTab_(ss);

  ui.alert(
    'Config tabs are ready.\n\n' +
    'Next steps:\n' +
    '1. Fill in the Value column in the Settings tab.\n' +
    '2. Add your column pairs to the Column Mapping tab.\n' +
    '3. Add pull policies for captain-entered columns in the Pull Column Policy tab.\n' +
    '4. Review or edit task presets in the Workflow Presets tab.\n' +
    '5. Run a Dry Run first to preview results before a live operation.'
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
      'and "Pull Missing Rows", and the SOURCE for "Push" operations. ' +
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
      'Source Tab Name',
      existingValues['Source Tab Name'] || '',
      'Optional tab name inside the External Source spreadsheet. If blank, ' +
      '"Import → Master" uses the first tab. For sales imports, use "Sales Rollup by APN".'
    ],
    [
      'User Sheet',
      existingValues['User Sheet'] || '',
      'ID of a single user spreadsheet to push master data TO or pull missing rows FROM. ' +
      'Only used by single-sheet Push and Pull operations. ' +
      'Find it in the URL: docs.google.com/spreadsheets/d/[COPY THIS]/edit'
    ],
    [
      'User Sheets Folder',
      existingValues['User Sheets Folder'] || existingValues['Target Folder'] || '',
      'ID of the Drive folder containing all user sheets. ' +
      'Used by folder-wide Push, Pull, and Rename operations. ' +
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

/**
 * Writes (or updates) the Pull Column Policy tab header row with
 * formatting and examples. Existing policy rows are preserved.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss
 */
function setupPullColumnPolicyTab_(ss) {
  var tab = ss.getSheetByName('Pull Column Policy');
  if (!tab) tab = ss.insertSheet('Pull Column Policy');

  var existingData = tab.getDataRange().getValues();
  var hasExistingRows = existingData.length > 1;

  tab.getRange(1, 1, 1, 3).setValues([['Column Name', 'Policy', 'Notes']]);

  var headerRange = tab.getRange(1, 1, 1, 3);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e8eaf6');
  tab.setFrozenRows(1);

  tab.getRange(1, 1).setNote(
    'Controls Pull Data operations into the master.\n\n' +
    'Policy values:\n' +
    '• fill_blank: write captain value only when master is blank\n' +
    '• overwrite: replace master value when captain has a non-blank value\n' +
    '• conflict: log differences without writing\n' +
    '• never: never pull this column\n\n' +
    'Unlisted columns default to conflict. resident_id is always protected.'
  );

  if (!hasExistingRows) {
    var examples = [
      ['resident_id', 'never', 'Protected identity column; always forced to never.'],
      ['ZoneName', 'conflict', 'Review zone differences instead of changing master automatically.'],
      ['Resident Name', 'fill_blank', 'Fill missing master names but do not replace existing values.'],
      ['Damage', 'overwrite', 'Example captain-owned field; change to match your real policy.'],
      ['Person Notes', 'overwrite', 'Example notes field entered by captains.']
    ];
    tab.getRange(2, 1, examples.length, 3).setValues(examples);
  }

  tab.setColumnWidth(1, 240);
  tab.setColumnWidth(2, 140);
  tab.setColumnWidth(3, 520);
  tab.getRange(2, 3, Math.max(tab.getMaxRows() - 1, 1), 1).setWrap(true);
}

/**
 * Creates the Workflow Presets tab used by the guided sidebar. Existing
 * workflow rows are preserved; a sales import preset is added only when the
 * tab has no presets yet.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss
 */
function setupWorkflowPresetsTab_(ss) {
  var tab = ss.getSheetByName('Workflow Presets');
  if (!tab) tab = ss.insertSheet('Workflow Presets');

  var headers = [
    'Workflow ID',
    'Workflow Name',
    'Operation',
    'Enabled',
    'Source Spreadsheet',
    'Source Tab',
    'Master Spreadsheet',
    'Match Column',
    'Column Mappings',
    'Notes'
  ];

  var existingData = tab.getDataRange().getValues();
  var hasExistingRows = existingData.length > 1 && String(existingData[1][0] || existingData[1][1] || '').trim() !== '';

  tab.getRange(1, 1, 1, headers.length).setValues([headers]);
  var headerRange = tab.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#e8eaf6');
  tab.setFrozenRows(1);

  tab.getRange(1, 9).setNote(
    'One mapping per line. Use either:\n\n' +
    'Source Header -> Target Header\n' +
    'or\n' +
    'Source Header | Target Header\n\n' +
    'If source and target use the same header, write it on both sides.'
  );

  if (!hasExistingRows) {
    var settings = readMergeConfig_(ss);
    var defaultMappings = [
      'Address - Sold Since Fire -> Address - Sold Since Fire',
      'Sales History -> Sales History',
      'Latest Sale Date -> Latest Sale Date',
      'Latest Sale Price -> Latest Sale Price',
      'Latest New Owner -> Latest New Owner'
    ].join('\n');

    tab.getRange(2, 1, 1, headers.length).setValues([[
      'sales-import',
      'Update Master From Sales Tracker',
      'import_to_master',
      'Yes',
      settings.sourceId || '',
      settings.sourceTabName || 'Sales Rollup by APN',
      settings.masterId || '',
      settings.matchColumn || 'APN',
      defaultMappings,
      'Imports the one-row-per-APN sales rollup into the master. Start with Dry Run.'
    ]]);
  }

  tab.setColumnWidth(1, 150);
  tab.setColumnWidth(2, 260);
  tab.setColumnWidth(3, 160);
  tab.setColumnWidth(4, 90);
  tab.setColumnWidth(5, 280);
  tab.setColumnWidth(6, 190);
  tab.setColumnWidth(7, 280);
  tab.setColumnWidth(8, 130);
  tab.setColumnWidth(9, 360);
  tab.setColumnWidth(10, 360);
  tab.getRange(2, 9, Math.max(tab.getMaxRows() - 1, 1), 2).setWrap(true);
}
