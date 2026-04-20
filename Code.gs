// ============================================================
// Sheet Smart — Spreadsheet Audit Script
// ============================================================
// Paste this into a standalone Google Apps Script project at
// https://script.google.com and run `runAudit()`.
// ============================================================

const FOLDER_ID = 'your-folder-id-here';
const MASTER_SPREADSHEET_ID = 'your-master-spreadsheet-id-here';

/**
 * Main entry point. Scans every spreadsheet in the target folder,
 * compares columns against the master, counts non-blank cells,
 * and writes a three-tab audit report to a new spreadsheet.
 */
function runAudit() {
  const masterHeaders = readMasterHeaders_();
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const sheets = collectSpreadsheets_(folder);

  const overviewRows = [];
  const detailRows = [];
  const residentIdEntries = [];

  sheets.forEach(function (file) {
    var id = file.getId();
    var name = file.getName();
    var url = file.getUrl();
    var lastEdited = file.getLastUpdated();
    var lastEditor = getLastEditor_(id);

    try {
      var ss = SpreadsheetApp.openById(id);
    } catch (e) {
      overviewRows.push([name, url, lastEdited, lastEditor, 'ERROR', '', '', '', '', '', '', '', '', e.message]);
      return;
    }

    var sheet = ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();

    if (data.length === 0) {
      overviewRows.push([name, url, lastEdited, lastEditor, 0, 0, '', '', '', '', '', '', '', 'Empty']);
      return;
    }

    var headers = data[0].map(function (h) { return String(h).trim(); });
    var dataRows = data.slice(1);
    var totalDataRows = dataRows.length;

    collectResidentIds_(headers, dataRows, name, url, residentIdEntries);

    var masterSet = {};
    masterHeaders.forEach(function (h) { masterSet[h] = true; });

    var sheetSet = {};
    headers.forEach(function (h) { if (h !== '') sheetSet[h] = true; });

    var missing = masterHeaders.filter(function (h) { return !sheetSet[h]; });
    var extra = headers.filter(function (h) { return h !== '' && !masterSet[h]; });

    var status = 'Match';
    if (missing.length > 0 && extra.length > 0) status = 'Missing + Extra';
    else if (missing.length > 0) status = 'Missing Columns';
    else if (extra.length > 0) status = 'Extra Columns';

    var apnCount = countNonBlankInColumn_(headers, dataRows, 'APN');
    var damageCount = countNonBlankInColumn_(headers, dataRows, 'Damage');
    var forSaleCount = countTrueInColumn_(headers, dataRows, 'Address - For Sale');
    var soldCount = countTrueInColumn_(headers, dataRows, 'Address - Sold Since Fire');
    var missingApn = countUniqueAddressesMissingApn_(headers, dataRows);

    overviewRows.push([
      name,
      url,
      lastEdited,
      lastEditor,
      headers.filter(function (h) { return h !== ''; }).length,
      totalDataRows,
      apnCount,
      damageCount,
      forSaleCount,
      soldCount,
      missingApn,
      missing.join(', '),
      extra.join(', '),
      status
    ]);

    headers.forEach(function (header, colIdx) {
      if (header === '') return;
      var nonBlank = 0;
      dataRows.forEach(function (row) {
        var val = row[colIdx];
        if (val !== '' && val !== null && val !== undefined) nonBlank++;
      });
      detailRows.push([
        name,
        header,
        columnLetter_(colIdx),
        masterSet[header] ? 'Yes' : 'No',
        nonBlank,
        totalDataRows
      ]);
    });
  });

  Logger.log('Total resident_id entries collected: ' + residentIdEntries.length);
  var duplicateRows = findDuplicateResidentIds_(residentIdEntries);
  Logger.log('Duplicate rows found: ' + duplicateRows.length);
  writeReport_(masterHeaders, overviewRows, detailRows, duplicateRows);
}

// -------  Internal helpers  -------

function readMasterHeaders_() {
  var ss = SpreadsheetApp.openById(MASTER_SPREADSHEET_ID);
  var sheet = ss.getSheets()[0];
  var row = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return row.map(function (h) { return String(h).trim(); }).filter(function (h) { return h !== ''; });
}

function collectSpreadsheets_(folder) {
  var files = [];
  var iter = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (iter.hasNext()) {
    files.push(iter.next());
  }
  files.sort(function (a, b) {
    return a.getName().localeCompare(b.getName());
  });
  return files;
}

function getLastEditor_(fileId) {
  try {
    var meta = Drive.Files.get(fileId, { fields: 'lastModifyingUser' });
    var user = meta.lastModifyingUser;
    if (user) return user.displayName || user.emailAddress || 'Unknown';
  } catch (e) {
    // Falls through to return below
  }
  return 'Unknown';
}

function countUniqueAddressesMissingApn_(headers, dataRows) {
  var addrCol = headers.indexOf('Address');
  var apnCol = headers.indexOf('APN');
  if (addrCol === -1) return 'N/A';
  if (apnCol === -1) return 'N/A';
  var seen = {};
  var count = 0;
  dataRows.forEach(function (row) {
    var addr = String(row[addrCol]).trim();
    var apn = row[apnCol];
    var apnBlank = (apn === '' || apn === null || apn === undefined);
    if (addr !== '' && apnBlank && !seen[addr]) {
      seen[addr] = true;
      count++;
    }
  });
  return count;
}

function countTrueInColumn_(headers, dataRows, columnName) {
  var colIdx = headers.indexOf(columnName);
  if (colIdx === -1) return 'N/A';
  var count = 0;
  dataRows.forEach(function (row) {
    if (row[colIdx] === true) count++;
  });
  return count;
}

function countNonBlankInColumn_(headers, dataRows, columnName) {
  var colIdx = headers.indexOf(columnName);
  if (colIdx === -1) return 'N/A';
  var count = 0;
  dataRows.forEach(function (row) {
    var val = row[colIdx];
    if (val !== '' && val !== null && val !== undefined) count++;
  });
  return count;
}

function collectResidentIds_(headers, dataRows, sheetName, url, entries) {
  var col = headers.indexOf('resident_id');
  if (col === -1) {
    Logger.log(sheetName + ': resident_id column NOT found. Headers: ' + headers.join(' | '));
    return;
  }
  Logger.log(sheetName + ': resident_id found at column ' + col);
  dataRows.forEach(function (row, i) {
    var val = row[col];
    if (val !== '' && val !== null && val !== undefined) {
      entries.push({
        id: String(val).trim(),
        sheet: sheetName,
        url: url,
        row: i + 2 // +2: 1-indexed header offset
      });
    }
  });
}

/**
 * Groups entries by resident_ID and returns rows only for IDs that
 * appear more than once (within the same sheet or across sheets).
 */
function findDuplicateResidentIds_(entries) {
  var map = {};
  entries.forEach(function (e) {
    if (!map[e.id]) map[e.id] = [];
    map[e.id].push(e);
  });
  var rows = [];
  Object.keys(map).forEach(function (id) {
    var group = map[id];
    if (group.length < 2) return;
    group.forEach(function (e) {
      var sheetNames = group.map(function (g) { return g.sheet; });
      var unique = sheetNames.filter(function (n, i) { return sheetNames.indexOf(n) === i; });
      var scope = (unique.length === 1) ? 'Within Sheet' : 'Across Sheets';
      rows.push([e.id, e.sheet, e.url, e.row, group.length, scope]);
    });
  });
  return rows;
}

function columnLetter_(index) {
  var letter = '';
  var temp = index;
  while (true) {
    letter = String.fromCharCode(65 + (temp % 26)) + letter;
    temp = Math.floor(temp / 26) - 1;
    if (temp < 0) break;
  }
  return letter;
}

function writeReport_(masterHeaders, overviewRows, detailRows, duplicateRows) {
  var report = SpreadsheetApp.create('Sheet Smart — Audit Report');

  // --- Tab 1: Overview ---
  var overviewSheet = report.getSheets()[0].setName('Overview');
  var ovHeader = [
    'Spreadsheet', 'URL', 'Last Edited', 'Last Editor',
    'Total Columns', 'Data Rows',
    'APN Values', 'Damage Values',
    'For Sale (TRUE)', 'Sold Since Fire (TRUE)',
    'Addresses Missing APN',
    'Missing vs Master', 'Extra vs Master', 'Status'
  ];
  var ovData = [ovHeader].concat(overviewRows);
  overviewSheet
    .getRange(1, 1, ovData.length, ovHeader.length)
    .setValues(ovData);
  formatHeaderRow_(overviewSheet, ovHeader.length);
  autoResize_(overviewSheet, ovHeader.length);

  var lastDataRow = ovData.length;
  if (lastDataRow > 1) {
    applyConditionalFormatting_(overviewSheet, lastDataRow);
  }

  // --- Tab 2: Column Detail ---
  var detailSheet = report.insertSheet('Column Detail');
  var dtHeader = [
    'Spreadsheet', 'Column Name', 'Position', 'In Master',
    'Non-Blank Count', 'Total Data Rows'
  ];
  var dtData = [dtHeader].concat(detailRows);
  detailSheet
    .getRange(1, 1, dtData.length, dtHeader.length)
    .setValues(dtData);
  formatHeaderRow_(detailSheet, dtHeader.length);
  autoResize_(detailSheet, dtHeader.length);

  // --- Tab 3: Master Columns ---
  var masterSheet = report.insertSheet('Master Columns');
  var mcHeader = ['Position', 'Column Name'];
  var mcData = [mcHeader];
  masterHeaders.forEach(function (h, i) {
    mcData.push([columnLetter_(i), h]);
  });
  masterSheet
    .getRange(1, 1, mcData.length, mcHeader.length)
    .setValues(mcData);
  formatHeaderRow_(masterSheet, mcHeader.length);
  autoResize_(masterSheet, mcHeader.length);

  // --- Tab 4: Duplicate Resident IDs ---
  var dupSheet = report.insertSheet('Duplicate Resident IDs');
  var dupHeader = [
    'resident_ID', 'Spreadsheet', 'URL', 'Row #',
    'Total Occurrences', 'Scope'
  ];
  if (duplicateRows.length > 0) {
    var dupData = [dupHeader].concat(duplicateRows);
    dupSheet
      .getRange(1, 1, dupData.length, dupHeader.length)
      .setValues(dupData);
  } else {
    dupSheet
      .getRange(1, 1, 1, dupHeader.length)
      .setValues([dupHeader]);
    dupSheet.getRange(2, 1).setValue('No duplicate resident IDs found.');
  }
  formatHeaderRow_(dupSheet, dupHeader.length);
  autoResize_(dupSheet, dupHeader.length);

  // Freeze header rows on all tabs
  [overviewSheet, detailSheet, masterSheet, dupSheet].forEach(function (s) {
    s.setFrozenRows(1);
  });

  Logger.log('Audit report created: ' + report.getUrl());
}

function formatHeaderRow_(sheet, numCols) {
  var range = sheet.getRange(1, 1, 1, numCols);
  range.setFontWeight('bold');
  range.setBackground('#4a86c8');
  range.setFontColor('#ffffff');
}

function applyConditionalFormatting_(sheet, lastRow) {
  var rules = [];

  // APN Values (G), Damage Values (H), For Sale (I), Sold Since Fire (J): light red when < 20
  var apnRange = sheet.getRange('G2:G' + lastRow);
  var dmgRange = sheet.getRange('H2:H' + lastRow);
  var forSaleRange = sheet.getRange('I2:I' + lastRow);
  var soldRange = sheet.getRange('J2:J' + lastRow);

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(20)
      .setBackground('#f4cccc')
      .setRanges([apnRange, dmgRange, forSaleRange, soldRange])
      .build()
  );

  // Data Rows (F): white-to-green gradient
  var dataRowsRange = sheet.getRange('F2:F' + lastRow);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpoint('#ffffff')
      .setGradientMaxpoint('#57bb8a')
      .setRanges([dataRowsRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

function autoResize_(sheet, numCols) {
  for (var i = 1; i <= numCols; i++) {
    sheet.autoResizeColumn(i);
  }
}
