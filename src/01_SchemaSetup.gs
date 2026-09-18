/**
 * 01_SchemaSetup.gs
 *
 * Builds the canonical "system of record" tabs in the bound spreadsheet.
 * Run once via the menu ("SRM AP Strategy System" -> "1. Setup Schema") to
 * initialize a blank spreadsheet, or safely re-run any time: existing sheets
 * are left in place and only missing headers/sheets are added.
 */

function setupSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheetWithHeaders_(ss, SHEET.CONFIG, COLS.CONFIG);
  seedConfigDefaults_(ss);

  ensureSheetWithHeaders_(ss, SHEET.PILLARS, COLS.PILLARS);
  ensureSheetWithHeaders_(ss, SHEET.GOALS, COLS.GOALS);
  ensureSheetWithHeaders_(ss, SHEET.CLUSTERS, COLS.CLUSTERS);
  ensureSheetWithHeaders_(ss, SHEET.OFFICES, COLS.OFFICES);
  ensureSheetWithHeaders_(ss, SHEET.VERTICALS, COLS.VERTICALS);
  ensureSheetWithHeaders_(ss, SHEET.KRAS, COLS.KRAS);
  ensureSheetWithHeaders_(ss, SHEET.KPIS, COLS.KPIS);
  ensureSheetWithHeaders_(ss, SHEET.KPI_OFFICE_LINKS, COLS.KPI_OFFICE_LINKS);
  ensureSheetWithHeaders_(ss, SHEET.ACTIONS, COLS.ACTIONS);
  ensureSheetWithHeaders_(ss, SHEET.AUDIT_LOG, COLS.AUDIT_LOG);
  ensureSheetWithHeaders_(ss, SHEET.VALIDATION_LOG, COLS.VALIDATION_LOG);
  ensureSheetWithHeaders_(ss, SHEET.IMPORT_LOG, COLS.IMPORT_LOG);

  applyKpiSheetValidation_(ss);

  // Tidy up: remove the default empty "Sheet1" if it's still blank.
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && def.getLastColumn() === 0) {
    ss.deleteSheet(def);
  }

  SpreadsheetApp.getUi().alert('Schema setup complete. ' + Object.keys(SHEET).length + ' system sheets are ready.');
}

/** Creates `name` with `headers` in row 1 if it doesn't exist; adds any missing headers if it does. */
function ensureSheetWithHeaders_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  const existing = getHeaderRow_(sheet);
  if (existing.length === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
    return sheet;
  }

  // Append any headers present in the spec but missing from an existing sheet.
  const missing = headers.filter(function (h) { return existing.indexOf(h) === -1; });
  if (missing.length > 0) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    sheet.getRange(1, existing.length + 1, 1, missing.length).setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');
  }
  return sheet;
}

/** Returns the header row (row 1) as an array of strings, or [] if the sheet is empty. */
function getHeaderRow_(sheet) {
  if (sheet.getLastColumn() === 0 || sheet.getLastRow() === 0) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (v) { return String(v).trim(); })
    .filter(function (v) { return v.length > 0; });
}

/** Returns a { headerName: 1-based column index } map for a sheet. */
function headerIndexMap_(sheet) {
  const headers = getHeaderRow_(sheet);
  const map = {};
  headers.forEach(function (h, i) { map[h] = i + 1; });
  return map;
}

function seedConfigDefaults_(ss) {
  const sheet = ss.getSheetByName(SHEET.CONFIG);
  const existingKeys = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; })
    : [];
  const toAdd = DEFAULT_CONFIG.filter(function (row) { return existingKeys.indexOf(row[0]) === -1; });
  if (toAdd.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, 3).setValues(toAdd);
  }
}

/** Adds dropdown data validation on the KPIs sheet for the controlled-vocabulary columns. */
function applyKpiSheetValidation_(ss) {
  const sheet = ss.getSheetByName(SHEET.KPIS);
  const colMap = headerIndexMap_(sheet);
  const maxRows = 2000; // headroom for growth; re-run setupSchema() after expanding

  const dropdowns = [
    ['KRA Type', KRA_TYPES],
    ['Metric Type', METRIC_TYPES],
    ['Indicator Nature', INDICATOR_NATURES],
    ['Measurement Nature', MEASUREMENT_NATURES],
    ['Frequency', FREQUENCIES],
    ['Goal Role', GOAL_ROLES],
    ['Q1 Status', [RAG.GREEN, RAG.AMBER, RAG.RED, RAG.GREY]],
    ['Q2 Status', [RAG.GREEN, RAG.AMBER, RAG.RED, RAG.GREY]],
    ['Q3 Status', [RAG.GREEN, RAG.AMBER, RAG.RED, RAG.GREY]],
    ['Q4 Status', [RAG.GREEN, RAG.AMBER, RAG.RED, RAG.GREY]]
  ];

  dropdowns.forEach(function (pair) {
    const col = colMap[pair[0]];
    if (!col) return;
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(pair[1], true).setAllowInvalid(false).build();
    sheet.getRange(2, col, maxRows, 1).setDataValidation(rule);
  });
}
