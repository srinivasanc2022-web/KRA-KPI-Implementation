/**
 * 03_ImportPipeline.gs
 *
 * Reads the 38-sheet SRMAP_KRA-KPI_AY26-27 workbook (01_Index,
 * 03_Strategic_Pillars, 04_Cascade_Map, the five Tier 0/1 scorecards, and the
 * Tier 2 office scorecards it lists) and populates the canonical model sheets
 * defined in 00_Constants.gs.
 *
 * Sheets are read BY HEADER NAME, never by column position, because at least
 * one known source tab (C1 Academic Scorecard) has a blank/decorative row
 * above its real header row -- findHeaderRow_() locates the real header row
 * by scanning the first few rows for the row with the most header matches.
 *
 * Entry points (wired to the menu in 07_Menu.gs):
 *   importSampleSourceData()   - imports the SAMPLE_SRC_* fixture sheets
 *                                 built by 02_SampleData.gs, for smoke-testing
 *                                 this pipeline without the real workbook.
 *   importRealWorkbook(id)     - imports from an external Spreadsheet ID
 *                                 (or the active spreadsheet if id is blank),
 *                                 using the real 01_Index-listed sheet names.
 */

function importSampleSourceData() {
  runImport_({ spreadsheetId: '', sheetPrefix: SAMPLE_PREFIX });
}

function importRealWorkbook(spreadsheetId) {
  runImport_({ spreadsheetId: spreadsheetId || '', sheetPrefix: '' });
}

function runImport_(options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSs = options.spreadsheetId
    ? SpreadsheetApp.openById(options.spreadsheetId)
    : ss;
  const prefix = options.sheetPrefix || '';

  setupSchema(); // idempotent -- guarantees canonical sheets exist before we write to them

  const indexRows = importIndex_(ss, sourceSs, prefix);
  importPillarsAndGoals_(ss, sourceSs, prefix);
  importCascadeMap_(ss, sourceSs, prefix);
  TIER0_1_SOURCE_SHEETS.forEach(function (sheetName) {
    importScorecardSheet_(ss, sourceSs, prefix + sheetName, null);
  });
  indexRows
    .filter(function (row) { return String(row.tier) === '2'; })
    .forEach(function (row) {
      importScorecardSheet_(ss, sourceSs, prefix + row.sheetName, row);
    });

  deriveVerticalsAndKras_(ss);
  recomputeOfficeRollupCounts_(ss);
  validateAllWeights(); // see 04_Validation.gs

  SpreadsheetApp.getUi().alert('Import complete. See the "' + SHEET.IMPORT_LOG + '" and "' +
    SHEET.VALIDATION_LOG + '" sheets for details.');
}

// ---------------------------------------------------------------------------
// Header-row discovery (handles the blank-header-row quirk generically)
// ---------------------------------------------------------------------------

/**
 * Scans the first `maxScanRows` rows of `sheet` for the row that best matches
 * `expectedHeaders` (by count of exact string matches, case-insensitive).
 * Returns { rowIndex (1-based), colMap: { headerName: 1-based col index } }
 * or null if no row matches at least 2 expected headers.
 */
function findHeaderRow_(sheet, expectedHeaders, maxScanRows) {
  const scanRows = Math.min(maxScanRows || 5, sheet.getLastRow());
  if (scanRows < 1) return null;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const values = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  const expectedLower = expectedHeaders.map(function (h) { return h.toLowerCase(); });

  let best = { rowIndex: -1, matches: 0, colMap: {} };
  values.forEach(function (row, i) {
    const colMap = {};
    let matches = 0;
    row.forEach(function (cell, j) {
      const cellStr = String(cell).trim();
      const idx = expectedLower.indexOf(cellStr.toLowerCase());
      if (idx !== -1 && cellStr.length > 0) {
        colMap[expectedHeaders[idx]] = j + 1;
        matches++;
      }
    });
    if (matches > best.matches) {
      best = { rowIndex: i + 1, matches: matches, colMap: colMap };
    }
  });

  if (best.matches < 2) return null;
  return best;
}

/** Reads all data rows below `headerRowIndex` into an array of {headerName: value} objects. */
function readRowsAsObjects_(sheet, headerRowIndex, colMap) {
  const lastRow = sheet.getLastRow();
  const firstDataRow = headerRowIndex + 1;
  if (lastRow < firstDataRow) return [];
  const lastCol = sheet.getLastColumn();
  const values = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol).getValues();

  return values
    .filter(function (row) { return row.some(function (c) { return String(c).trim().length > 0; }); })
    .map(function (row) {
      const obj = {};
      Object.keys(colMap).forEach(function (header) {
        obj[header] = row[colMap[header] - 1];
      });
      return obj;
    });
}

// ---------------------------------------------------------------------------
// 01_Index
// ---------------------------------------------------------------------------

function importIndex_(ss, sourceSs, prefix) {
  const sheet = sourceSs.getSheetByName(prefix + SOURCE_SHEET.INDEX);
  if (!sheet) {
    logImport_(ss, SOURCE_SHEET.INDEX, '', '', 0, 'SKIPPED', 'Sheet not found');
    return [];
  }
  const expected = ['Sheet Name', 'Full Office Name', 'Tier', 'Cluster', 'KPI Count', 'Total Wt'];
  const header = findHeaderRow_(sheet, expected, 5);
  if (!header) {
    logImport_(ss, SOURCE_SHEET.INDEX, '', '', 0, 'FAILED', 'Could not locate header row');
    return [];
  }
  const rows = readRowsAsObjects_(sheet, header.rowIndex, header.colMap);
  const result = rows.map(function (r) {
    return {
      sheetName: String(r['Sheet Name'] || '').trim(),
      fullName: String(r['Full Office Name'] || '').trim(),
      tier: r['Tier'],
      cluster: String(r['Cluster'] || '').trim(),
      kpiCount: r['KPI Count'],
      totalWt: r['Total Wt']
    };
  }).filter(function (r) { return r.sheetName.length > 0; });

  logImport_(ss, SOURCE_SHEET.INDEX, '', '', result.length, 'OK', result.length + ' scorecards registered');
  return result;
}

// ---------------------------------------------------------------------------
// 03_Strategic_Pillars -> Pillars + Goals
// ---------------------------------------------------------------------------

function importPillarsAndGoals_(ss, sourceSs, prefix) {
  const sheet = sourceSs.getSheetByName(prefix + SOURCE_SHEET.PILLARS);
  if (!sheet) {
    logImport_(ss, SOURCE_SHEET.PILLARS, '', '', 0, 'SKIPPED', 'Sheet not found');
    return;
  }
  const expected = ['Pillar ID', 'Pillar Name', 'Goal ID', 'Five-Year Goal / Target', 'Owner Office', 'Contributing Offices'];
  const header = findHeaderRow_(sheet, expected, 5);
  if (!header) {
    logImport_(ss, SOURCE_SHEET.PILLARS, '', '', 0, 'FAILED', 'Could not locate header row');
    return;
  }
  const rows = readRowsAsObjects_(sheet, header.rowIndex, header.colMap);

  const pillarSheet = ss.getSheetByName(SHEET.PILLARS);
  const goalSheet = ss.getSheetByName(SHEET.GOALS);
  const pillarMap = headerIndexMap_(pillarSheet);
  const goalMap = headerIndexMap_(goalSheet);

  const knownPillars = existingKeys_(pillarSheet, pillarMap['Pillar ID']);
  const knownGoals = existingKeys_(goalSheet, goalMap['Goal ID']);

  const newPillarRows = [];
  const newGoalRows = [];

  rows.forEach(function (r) {
    const pillarId = String(r['Pillar ID'] || '').trim();
    const goalId = String(r['Goal ID'] || '').trim();
    if (pillarId && knownPillars.indexOf(pillarId) === -1) {
      newPillarRows.push(orderRow_(pillarMap, { 'Pillar ID': pillarId, 'Pillar Name': r['Pillar Name'] }));
      knownPillars.push(pillarId);
    }
    if (goalId && knownGoals.indexOf(goalId) === -1) {
      newGoalRows.push(orderRow_(goalMap, {
        'Goal ID': goalId,
        'Parent Pillar': pillarId,
        'Five-Year Goal / Target': r['Five-Year Goal / Target'],
        'Owner Office': r['Owner Office'],
        'Contributing Offices': r['Contributing Offices'],
        'Owner Weight %': '', // admin-set per section 4; left blank for manual entry
        'Overall Progress': '' // computed by the roll-up engine (phase 2, see NEXT_STEPS.md)
      }));
      knownGoals.push(goalId);
    }
  });

  appendRows_(pillarSheet, newPillarRows);
  appendRows_(goalSheet, newGoalRows);
  logImport_(ss, SOURCE_SHEET.PILLARS, '', '', rows.length,
    'OK', newPillarRows.length + ' pillars, ' + newGoalRows.length + ' goals added');
}

// ---------------------------------------------------------------------------
// 04_Cascade_Map -> Offices (goals owned/contributed, KPI split)
// ---------------------------------------------------------------------------

function importCascadeMap_(ss, sourceSs, prefix) {
  const sheet = sourceSs.getSheetByName(prefix + SOURCE_SHEET.CASCADE_MAP);
  if (!sheet) {
    logImport_(ss, SOURCE_SHEET.CASCADE_MAP, '', '', 0, 'SKIPPED', 'Sheet not found');
    return;
  }
  const expected = ['Office', 'Goals Owned', 'Goals Contributed To', 'Strategic KPI Count', 'Operational KPI Count'];
  const header = findHeaderRow_(sheet, expected, 5);
  if (!header) {
    logImport_(ss, SOURCE_SHEET.CASCADE_MAP, '', '', 0, 'FAILED', 'Could not locate header row');
    return;
  }
  const rows = readRowsAsObjects_(sheet, header.rowIndex, header.colMap);

  const officeSheet = ss.getSheetByName(SHEET.OFFICES);
  const officeMap = headerIndexMap_(officeSheet);
  const officeRowIndex = {};
  existingKeys_(officeSheet, officeMap['Office Code']).forEach(function (code, i) { officeRowIndex[code] = i + 2; });

  rows.forEach(function (r) {
    const code = String(r['Office'] || '').trim();
    if (!code) return;
    let rowNum = officeRowIndex[code];
    if (!rowNum) {
      rowNum = officeSheet.getLastRow() + 1;
      officeSheet.getRange(rowNum, officeMap['Office Code']).setValue(code);
      officeRowIndex[code] = rowNum;
    }
    if (officeMap['Goals Owned']) officeSheet.getRange(rowNum, officeMap['Goals Owned']).setValue(r['Goals Owned']);
    if (officeMap['Goals Contributed To']) officeSheet.getRange(rowNum, officeMap['Goals Contributed To']).setValue(r['Goals Contributed To']);
    if (officeMap['KPI Count (Strategic)']) officeSheet.getRange(rowNum, officeMap['KPI Count (Strategic)']).setValue(r['Strategic KPI Count']);
    if (officeMap['KPI Count (Operational)']) officeSheet.getRange(rowNum, officeMap['KPI Count (Operational)']).setValue(r['Operational KPI Count']);
  });

  logImport_(ss, SOURCE_SHEET.CASCADE_MAP, '', '', rows.length, 'OK', rows.length + ' office RACI-at-goal-level rows applied');
}

// ---------------------------------------------------------------------------
// Tier 0/1/2 scorecards -> KPIs (+ KPI_Office_Links for co-ownership)
// ---------------------------------------------------------------------------

function importScorecardSheet_(ss, sourceSs, sourceSheetName, indexRow) {
  const sheet = sourceSs.getSheetByName(sourceSheetName);
  if (!sheet) {
    logImport_(ss, sourceSheetName, indexRow ? indexRow.tier : '', indexRow ? indexRow.cluster : '', 0, 'SKIPPED', 'Sheet not found');
    return;
  }
  // Ask for a subset of well-known headers; anything else present is copied through if it matches COLS.KPIS.
  const expected = COLS.KPIS.filter(function (h) { return h !== 'KPI_ID' && h !== 'Co-Owner Offices (Name:%)'; });
  const header = findHeaderRow_(sheet, expected, 5);
  if (!header) {
    logImport_(ss, sourceSheetName, indexRow ? indexRow.tier : '', indexRow ? indexRow.cluster : '', 0, 'FAILED', 'Could not locate header row');
    return;
  }
  const rows = readRowsAsObjects_(sheet, header.rowIndex, header.colMap);

  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  const linkSheet = ss.getSheetByName(SHEET.KPI_OFFICE_LINKS);
  const linkMap = headerIndexMap_(linkSheet);

  const officeCode = sourceSheetName.replace(SAMPLE_PREFIX, '');
  const newKpiRows = [];
  const newLinkRows = [];

  rows.forEach(function (r, i) {
    const kpiId = officeCode + '-' + String(i + 1).padStart(3, '0');
    const rowObj = { 'KPI_ID': kpiId };
    COLS.KPIS.forEach(function (h) {
      if (h === 'KPI_ID') return;
      rowObj[h] = r.hasOwnProperty(h) ? r[h] : '';
    });
    // If tier wasn't in the sheet (common on Tier 2 sheets that omit a per-row Tier column), infer it.
    if (!rowObj['Tier'] && indexRow) rowObj['Tier'] = indexRow.tier;
    if (!rowObj['Office'] && indexRow) rowObj['Office'] = indexRow.fullName;
    newKpiRows.push(orderRow_(kpiMap, rowObj));

    // Co-ownership: "AiTI:30, QuTI:20" style values create explicit many-to-many
    // links instead of duplicating the KPI row (section 4).
    const coOwnerText = String(r['Co-Owner Offices (Name:%)'] || '').trim();
    if (coOwnerText) {
      coOwnerText.split(',').forEach(function (pair) {
        const parts = pair.split(':');
        const office = (parts[0] || '').trim();
        const pct = parts[1] ? parseFloat(parts[1]) : '';
        if (office) newLinkRows.push(orderRow_(linkMap, { 'KPI_ID': kpiId, 'Office Code': office, 'Role': 'Contributor', 'Co-Ownership %': pct }));
      });
    }
  });

  appendRows_(kpiSheet, newKpiRows);
  appendRows_(linkSheet, newLinkRows);
  logImport_(ss, sourceSheetName, indexRow ? indexRow.tier : '', indexRow ? indexRow.cluster : '',
    newKpiRows.length, 'OK', newKpiRows.length + ' KPI rows imported' + (newLinkRows.length ? (', ' + newLinkRows.length + ' co-ownership links') : ''));
}

// ---------------------------------------------------------------------------
// Derived tables: Verticals / KRAs (unique combinations found in KPIs)
// ---------------------------------------------------------------------------

function deriveVerticalsAndKras_(ss) {
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  if (kpiSheet.getLastRow() < 2) return;
  const data = kpiSheet.getRange(2, 1, kpiSheet.getLastRow() - 1, kpiSheet.getLastColumn()).getValues();

  const verticalSheet = ss.getSheetByName(SHEET.VERTICALS);
  const kraSheet = ss.getSheetByName(SHEET.KRAS);
  const verticalMap = headerIndexMap_(verticalSheet);
  const kraMap = headerIndexMap_(kraSheet);

  const knownVerticals = {}; // key: office|vertical -> Vertical ID
  existingRows_(verticalSheet).forEach(function (r) {
    knownVerticals[r['Office Code'] + '|' + r['Vertical Name']] = r['Vertical ID'];
  });
  const knownKras = {};
  existingRows_(kraSheet).forEach(function (r) { knownKras[r['Parent Vertical ID'] + '|' + r['KRA Name']] = r['KRA ID']; });

  const newVerticalRows = [];
  const newKraRows = [];
  let vSeq = Object.keys(knownVerticals).length;
  let kSeq = Object.keys(knownKras).length;

  data.forEach(function (row) {
    const office = row[kpiMap['Office'] - 1];
    const verticalName = row[kpiMap['Vertical'] - 1];
    const kraName = row[kpiMap['KRA'] - 1];
    if (!verticalName) return;

    const vKey = office + '|' + verticalName;
    let verticalId = knownVerticals[vKey];
    if (!verticalId) {
      vSeq++;
      verticalId = 'V' + String(vSeq).padStart(3, '0');
      knownVerticals[vKey] = verticalId;
      newVerticalRows.push(orderRow_(verticalMap, {
        'Vertical ID': verticalId, 'Vertical Name': verticalName, 'Office Code': office, 'Linked KRAs': ''
      }));
    }

    if (kraName) {
      const kKey = verticalId + '|' + kraName;
      if (!knownKras[kKey]) {
        kSeq++;
        const kraId = 'K' + String(kSeq).padStart(3, '0');
        knownKras[kKey] = kraId;
        newKraRows.push(orderRow_(kraMap, {
          'KRA ID': kraId, 'KRA Name': kraName, 'Parent Vertical ID': verticalId, 'Office Code': office, 'Linked KPIs': ''
        }));
      }
    }
  });

  appendRows_(verticalSheet, newVerticalRows);
  appendRows_(kraSheet, newKraRows);
}

function recomputeOfficeRollupCounts_(ss) {
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  const officeSheet = ss.getSheetByName(SHEET.OFFICES);
  const officeMap = headerIndexMap_(officeSheet);
  if (kpiSheet.getLastRow() < 2 || officeSheet.getLastRow() < 2) return;

  const kpiRows = kpiSheet.getRange(2, 1, kpiSheet.getLastRow() - 1, kpiSheet.getLastColumn()).getValues();
  const officeRows = officeSheet.getRange(2, 1, officeSheet.getLastRow() - 1, officeSheet.getLastColumn()).getValues();

  officeRows.forEach(function (orow, i) {
    const code = orow[officeMap['Office Code'] - 1];
    const officeKpis = kpiRows.filter(function (k) { return k[kpiMap['Office'] - 1] === code || k[kpiMap['Office'] - 1] === orow[officeMap['Full Name'] - 1]; });
    const totalWt = officeKpis.reduce(function (sum, k) { return sum + (parseFloat(k[kpiMap['Wt (%)'] - 1]) || 0); }, 0);
    if (officeMap['Total Weight (Check)']) officeSheet.getRange(2 + i, officeMap['Total Weight (Check)']).setValue(totalWt);
  });
}

// ---------------------------------------------------------------------------
// Generic sheet helpers
// ---------------------------------------------------------------------------

function orderRow_(colMap, obj) {
  const width = Math.max.apply(null, Object.values(colMap));
  const row = new Array(width).fill('');
  Object.keys(obj).forEach(function (key) {
    if (colMap[key]) row[colMap[key] - 1] = obj[key] === undefined ? '' : obj[key];
  });
  return row;
}

function appendRows_(sheet, rows) {
  if (!rows || rows.length === 0) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function existingKeys_(sheet, col) {
  if (!col || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); });
}

function existingRows_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  const map = headerIndexMap_(sheet);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return values.map(function (row) {
    const obj = {};
    Object.keys(map).forEach(function (h) { obj[h] = row[map[h] - 1]; });
    return obj;
  });
}

function logImport_(ss, sourceSheet, tier, cluster, rowsImported, status, details) {
  const sheet = ss.getSheetByName(SHEET.IMPORT_LOG);
  sheet.appendRow([new Date(), sourceSheet, tier, cluster, rowsImported, status, details]);
}
