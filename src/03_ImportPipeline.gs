/**
 * 03_ImportPipeline.gs
 *
 * Reads the 39-sheet SRMAP_KRA-KPI_AY26-27 workbook (01_Index,
 * 03_Strategic_Pillars, 04_Cascade_Map, the five Tier 0/1 scorecards, and
 * the 30 Tier 2 office scorecards 01_Index lists) and populates the
 * canonical model sheets defined in 00_Constants.gs.
 *
 * Sheets are read BY HEADER NAME, never by column position. Header text is
 * normalized (lowercased, all whitespace/newlines/punctuation stripped)
 * before comparison, because the real workbook is inconsistent about it --
 * e.g. the S.No column is literally "S.\nNo" (an embedded newline) on some
 * office tabs and "S.No" on others, and 03_Strategic_Pillars / 04_Cascade_Map
 * / the Tier 0/1 scorecards all have 2-3 blank/title rows above their real
 * header row (on at least one sheet, C2 Growth Scorecard, the header sits on
 * row 4 while others use row 3) -- findHeaderRow_() locates the real header
 * row generically by scanning for the row with the most header matches,
 * rather than assuming a fixed row number.
 *
 * A few headers are worded differently between Tier 0/1 and Tier 2 sheets
 * (e.g. "Pillar" vs "Strategic Pillar") -- see HEADER_ALIASES in
 * 00_Constants.gs.
 *
 * Entry points (wired to the menu in 07_Menu.gs):
 *   importSampleSourceData()   - imports the SAMPLE_SRC_* fixture sheets
 *                                 built by 02_SampleData.gs, for smoke-testing
 *                                 this pipeline without the real workbook.
 *   importRealWorkbook(id)     - imports from an external Spreadsheet ID
 *                                 (or the active spreadsheet if id is blank),
 *                                 using the real 01_Index-listed sheet names.
 */

// Tier 2 office-sheet fields, in the canonical names they're read into.
// SOURCE_GOAL_COMBINED_HEADER is read but never stored directly -- Goal ID
// is extracted from it (extractGoalId_()) into the KPIs sheet's 'Goal ID'
// column instead, matching the spec's entity model (Goal ID links to Goal).
const TIER2_KPI_FIELDS = [
  'S.No', 'Office', 'Tier', 'Cluster', 'KRA Type', 'Strategic Pillar',
  SOURCE_GOAL_COMBINED_HEADER, 'Goal Role', 'Vertical', 'KRA', 'KPI',
  'Strategic Rationale (5-Yr Goals)', 'Definition of KPI',
  'How to Measure (KPI-Specific Formula)', 'Illustration', 'Unit', 'Frequency',
  'Measurement Nature', 'Metric Type', 'Indicator Nature', 'Wt (%)',
  'Annual Target (AY 26-27)',
  'Q1 Milestone Target', 'Q2 Milestone Target', 'Q3 Milestone Target', 'Q4 Milestone Target',
  'Q1 Actual', 'Q2 Actual', 'Q3 Actual', 'Q4 Actual',
  'Q1 Status', 'Q2 Status', 'Q3 Status', 'Q4 Status',
  'Q1 ATR', 'Q2 ATR', 'Q3 ATR', 'Q4 ATR',
  'Score (Actual vs Target)', 'Data Source', 'KPI Owner',
  'R (Responsible)', 'A (Accountable)', 'C (Consulted)', 'I (Informed)',
  'Responsible Officer', 'Escalation / Remarks', 'Overlap / Remarks'
];

// Tier 0/1 scorecard fields -- a much sparser record than Tier 2 (no Office,
// no Strategic Rationale/Definition/Illustration/Data Source/RACI/Remarks;
// adds "Rolls Up From" and "Owner Sign-off" which Tier 2 sheets don't have).
const TIER01_KPI_FIELDS = [
  'S.No', 'Tier', 'Cluster', 'KRA Type', 'Strategic Pillar',
  SOURCE_GOAL_COMBINED_HEADER, 'Goal Role', 'Vertical', 'KRA', 'KPI',
  'How to Measure (KPI-Specific Formula)', 'Unit', 'Frequency',
  'Annual Target (AY 26-27)', 'Wt (%)', 'Rolls Up From',
  'Q1 Milestone Target', 'Q2 Milestone Target', 'Q3 Milestone Target', 'Q4 Milestone Target',
  'Owner Sign-off'
];

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
    importScorecardSheet_(ss, sourceSs, prefix + sheetName, null, sheetName);
  });
  indexRows
    .filter(function (row) { return String(row.tier) === '2'; })
    .forEach(function (row) {
      importScorecardSheet_(ss, sourceSs, prefix + row.sheetName, row, row.sheetName);
    });

  deriveVerticalsAndKras_(ss);
  deriveClusters_(ss);
  recomputeOfficeRollupCounts_(ss);
  validateAllWeights(); // see 04_Validation.gs

  SpreadsheetApp.getUi().alert('Import complete. See the "' + SHEET.IMPORT_LOG + '" and "' +
    SHEET.VALIDATION_LOG + '" sheets for details.');
}

// ---------------------------------------------------------------------------
// Header normalization + discovery
// ---------------------------------------------------------------------------

/** Lowercases and strips every whitespace/newline/punctuation character. */
function normalizeHeaderText_(s) {
  return String(s === null || s === undefined ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

/** canonicalFields[] -> { normalizedVariant: canonicalName } using HEADER_ALIASES where defined. */
function buildHeaderLookup_(canonicalFields) {
  const lookup = {};
  canonicalFields.forEach(function (canonical) {
    const variants = HEADER_ALIASES[canonical] || [canonical];
    variants.forEach(function (v) { lookup[normalizeHeaderText_(v)] = canonical; });
  });
  return lookup;
}

/**
 * Scans the first `maxScanRows` rows of `sheet` for the row that best
 * matches `canonicalFields` (by count of normalized matches). Returns
 * { rowIndex (1-based), colMap: { canonicalName: 1-based col index } }
 * or null if no row matches at least 2 fields.
 */
function findHeaderRow_(sheet, canonicalFields, maxScanRows) {
  const scanRows = Math.min(maxScanRows || 6, sheet.getLastRow());
  if (scanRows < 1) return null;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const values = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  const lookup = buildHeaderLookup_(canonicalFields);

  let best = { rowIndex: -1, matches: 0, colMap: {} };
  values.forEach(function (row, i) {
    const colMap = {};
    let matches = 0;
    row.forEach(function (cell, j) {
      const norm = normalizeHeaderText_(cell);
      if (norm && lookup[norm]) {
        colMap[lookup[norm]] = j + 1;
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

/** Reads all data rows below `headerRowIndex` into an array of {canonicalName: value} objects. */
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

/** Extracts the leading "P<n>.<n>" goal ID out of a compound "P1.4 — <target text>" cell. */
function extractGoalId_(compoundText) {
  const m = String(compoundText || '').match(/P\d+\.\d+/);
  return m ? m[0] : '';
}

/** "Tier 2 - Office" / "Tier 0 - Vice-Chancellor" -> "2" / "0". Falls back to the raw trimmed text. */
function parseTierNumber_(tierText) {
  const m = String(tierText || '').match(/\d+/);
  return m ? m[0] : String(tierText || '').trim();
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
  const expected = ['Sheet', 'Full name', 'Tier', 'Cluster', 'KPIs', 'Strategic', 'Operational', 'Verticals', 'Total Wt'];
  const header = findHeaderRow_(sheet, expected, 6);
  if (!header) {
    logImport_(ss, SOURCE_SHEET.INDEX, '', '', 0, 'FAILED', 'Could not locate header row');
    return [];
  }
  const rows = readRowsAsObjects_(sheet, header.rowIndex, header.colMap);
  const result = rows.map(function (r) {
    return {
      sheetName: String(r['Sheet'] || '').trim(),
      fullName: String(r['Full name'] || '').trim(),
      tier: parseTierNumber_(r['Tier']),
      cluster: String(r['Cluster'] || '').trim(),
      kpiCount: r['KPIs'],
      strategicCount: r['Strategic'],
      operationalCount: r['Operational'],
      verticalCount: r['Verticals'],
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
  const expected = ['Pillar', 'Pillar name', 'Goal ID', 'Five-year goal / target', 'Owner office', 'Contributing offices'];
  const header = findHeaderRow_(sheet, expected, 6);
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

  // 'Pillar' / 'Pillar name' are only filled on the FIRST goal row of each
  // pillar group in the source sheet (a visual merge); forward-fill them.
  let currentPillarId = '';
  let currentPillarName = '';

  rows.forEach(function (r) {
    const pillarIdCell = String(r['Pillar'] || '').trim();
    const pillarNameCell = String(r['Pillar name'] || '').trim();
    if (pillarIdCell) currentPillarId = pillarIdCell;
    if (pillarNameCell) currentPillarName = pillarNameCell;

    const goalId = String(r['Goal ID'] || '').trim();
    if (!goalId) return;

    if (currentPillarId && knownPillars.indexOf(currentPillarId) === -1) {
      newPillarRows.push(orderRow_(pillarMap, { 'Pillar ID': currentPillarId, 'Pillar Name': currentPillarName }));
      knownPillars.push(currentPillarId);
    }
    if (knownGoals.indexOf(goalId) === -1) {
      newGoalRows.push(orderRow_(goalMap, {
        'Goal ID': goalId,
        'Parent Pillar': currentPillarId,
        'Five-Year Goal / Target': r['Five-year goal / target'],
        'Owner Office': r['Owner office'],
        'Contributing Offices': r['Contributing offices'],
        'Owner Weight %': '', // admin-set per section 4; left blank for manual entry
        'Overall Progress': '', // legacy free-text placeholder; see 'Achievement %'
        'Achievement %': '' // computed by the roll-up engine, 08_RollupEngine.gs
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
// 04_Cascade_Map -> Offices (cluster, goals owned/contributed, KPI split)
// ---------------------------------------------------------------------------

function importCascadeMap_(ss, sourceSs, prefix) {
  const sheet = sourceSs.getSheetByName(prefix + SOURCE_SHEET.CASCADE_MAP);
  if (!sheet) {
    logImport_(ss, SOURCE_SHEET.CASCADE_MAP, '', '', 0, 'SKIPPED', 'Sheet not found');
    return;
  }
  const expected = ['Cluster', 'Office', 'Goals owned', 'Goals contributed to', 'KPIs', 'Strategic', 'Operational'];
  const header = findHeaderRow_(sheet, expected, 6);
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
    const goalsOwned = String(r['Goals owned'] || '').trim();
    const goalsContrib = String(r['Goals contributed to'] || '').trim();
    if (officeMap['Cluster']) officeSheet.getRange(rowNum, officeMap['Cluster']).setValue(r['Cluster']);
    if (officeMap['Goals Owned']) officeSheet.getRange(rowNum, officeMap['Goals Owned']).setValue(goalsOwned === '—' ? '' : goalsOwned);
    if (officeMap['Goals Contributed To']) officeSheet.getRange(rowNum, officeMap['Goals Contributed To']).setValue(goalsContrib === '—' ? '' : goalsContrib);
    if (officeMap['KPI Count (Strategic)']) officeSheet.getRange(rowNum, officeMap['KPI Count (Strategic)']).setValue(r['Strategic']);
    if (officeMap['KPI Count (Operational)']) officeSheet.getRange(rowNum, officeMap['KPI Count (Operational)']).setValue(r['Operational']);
  });

  logImport_(ss, SOURCE_SHEET.CASCADE_MAP, '', '', rows.length, 'OK', rows.length + ' office RACI-at-goal-level rows applied');
}

// ---------------------------------------------------------------------------
// Tier 0/1/2 scorecards -> KPIs (+ KPI_Office_Links for Goal Role = Co-owner)
// ---------------------------------------------------------------------------

function importScorecardSheet_(ss, sourceSs, sourceSheetName, indexRow, plainSheetName) {
  const sheet = sourceSs.getSheetByName(sourceSheetName);
  if (!sheet) {
    logImport_(ss, sourceSheetName, indexRow ? indexRow.tier : '', indexRow ? indexRow.cluster : '', 0, 'SKIPPED', 'Sheet not found');
    return;
  }
  const isTier01 = !indexRow; // Tier 0/1 sheets are passed with indexRow === null
  const expected = isTier01 ? TIER01_KPI_FIELDS : TIER2_KPI_FIELDS;
  const header = findHeaderRow_(sheet, expected, 6);
  if (!header) {
    logImport_(ss, sourceSheetName, indexRow ? indexRow.tier : '', indexRow ? indexRow.cluster : '', 0, 'FAILED', 'Could not locate header row');
    return;
  }
  // The real sheets carry a trailing "TOTAL" row (Wt(%) sums the column,
  // KPI/KRA text is blank or "TOTAL") and, on at least one tab, stray
  // leftover cells below the real data (only "Office" filled in). A row is
  // only a real KPI row if it has KPI text that isn't a total marker.
  const rows = readRowsAsObjects_(sheet, header.rowIndex, header.colMap).filter(function (r) {
    const kpiText = String(r['KPI'] || '').trim();
    return kpiText.length > 0 && kpiText.toUpperCase() !== 'TOTAL';
  });

  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  const linkSheet = ss.getSheetByName(SHEET.KPI_OFFICE_LINKS);
  const linkMap = headerIndexMap_(linkSheet);

  const officeCode = isTier01 ? (TIER0_1_CODES[plainSheetName] || plainSheetName) : plainSheetName;
  const newKpiRows = [];
  const newLinkRows = [];

  rows.forEach(function (r, i) {
    const kpiId = officeCode + '-' + String(i + 1).padStart(3, '0');
    const goalId = extractGoalId_(r[SOURCE_GOAL_COMBINED_HEADER]);
    const tier = isTier01 ? parseTierNumber_(r['Tier']) : parseTierNumber_(r['Tier']);
    const office = isTier01 ? officeCode : (r['Office'] || officeCode);

    const rowObj = {
      'KPI_ID': kpiId,
      'S.No': r['S.No'], 'Office': office, 'Tier': tier, 'Cluster': r['Cluster'],
      'KRA Type': r['KRA Type'], 'Strategic Pillar': r['Strategic Pillar'], 'Goal ID': goalId,
      'Goal Role': r['Goal Role'], 'Vertical': r['Vertical'], 'KRA': r['KRA'], 'KPI': r['KPI'],
      'Strategic Rationale (5-Yr Goals)': r['Strategic Rationale (5-Yr Goals)'] || '',
      'Definition of KPI': r['Definition of KPI'] || '',
      'How to Measure (KPI-Specific Formula)': r['How to Measure (KPI-Specific Formula)'],
      'Illustration': r['Illustration'] || '',
      'Unit': r['Unit'], 'Frequency': r['Frequency'],
      'Measurement Nature': r['Measurement Nature'] || '', 'Metric Type': r['Metric Type'] || '',
      'Indicator Nature': r['Indicator Nature'] || '', 'Wt (%)': r['Wt (%)'],
      'Annual Target (AY 26-27)': r['Annual Target (AY 26-27)'],
      'Q1 Milestone Target': r['Q1 Milestone Target'], 'Q2 Milestone Target': r['Q2 Milestone Target'],
      'Q3 Milestone Target': r['Q3 Milestone Target'], 'Q4 Milestone Target': r['Q4 Milestone Target'],
      'Q1 Actual': r['Q1 Actual'] || '', 'Q2 Actual': r['Q2 Actual'] || '',
      'Q3 Actual': r['Q3 Actual'] || '', 'Q4 Actual': r['Q4 Actual'] || '',
      'Q1 Status': r['Q1 Status'] || '', 'Q2 Status': r['Q2 Status'] || '',
      'Q3 Status': r['Q3 Status'] || '', 'Q4 Status': r['Q4 Status'] || '',
      'Q1 ATR': r['Q1 ATR'] || '', 'Q2 ATR': r['Q2 ATR'] || '',
      'Q3 ATR': r['Q3 ATR'] || '', 'Q4 ATR': r['Q4 ATR'] || '',
      'Score (Actual vs Target)': r['Score (Actual vs Target)'] || '', 'Achievement %': '',
      'Data Source': r['Data Source'] || '', 'KPI Owner': r['KPI Owner'] || '',
      'R (Responsible)': r['R (Responsible)'] || '', 'A (Accountable)': r['A (Accountable)'] || '',
      'C (Consulted)': r['C (Consulted)'] || '', 'I (Informed)': r['I (Informed)'] || '',
      'Responsible Officer': r['Responsible Officer'] || '',
      'Escalation / Remarks': r['Escalation / Remarks'] || '', 'Overlap / Remarks': r['Overlap / Remarks'] || '',
      'Rolls Up From': r['Rolls Up From'] || '', 'Owner Sign-off': r['Owner Sign-off'] || ''
    };
    newKpiRows.push(orderRow_(kpiMap, rowObj));

    // Co-ownership (section 4/12): the real workbook expresses this by
    // giving each co-owning office its OWN KPI row tagged Goal Role =
    // "Co-owner" for the same Goal ID (e.g. AiTI, QuTI and Research all
    // "Co-owner" on P3.1) rather than a single shared row with a % column.
    // Turn that into an explicit many-to-many link so roll-ups can find all
    // co-owners of a goal without re-parsing free text.
    if (String(r['Goal Role'] || '').trim().toLowerCase() === 'co-owner' && goalId) {
      newLinkRows.push(orderRow_(linkMap, {
        'KPI_ID': kpiId, 'Goal ID': goalId, 'Office Code': office, 'Role': 'Co-owner', 'Co-Ownership %': ''
      }));
    }
  });

  appendRows_(kpiSheet, newKpiRows);
  appendRows_(linkSheet, newLinkRows);
  logImport_(ss, sourceSheetName, isTier01 ? '' : indexRow.tier, isTier01 ? '' : indexRow.cluster,
    newKpiRows.length, 'OK', newKpiRows.length + ' KPI rows imported' + (newLinkRows.length ? (', ' + newLinkRows.length + ' co-ownership links') : ''));
}

// ---------------------------------------------------------------------------
// Derived tables: Clusters / Verticals / KRAs (unique combinations found)
// ---------------------------------------------------------------------------

function deriveClusters_(ss) {
  const officeSheet = ss.getSheetByName(SHEET.OFFICES);
  const officeMap = headerIndexMap_(officeSheet);
  if (officeSheet.getLastRow() < 2) return;
  const offices = existingRows_(officeSheet);

  const clusterSheet = ss.getSheetByName(SHEET.CLUSTERS);
  const clusterMap = headerIndexMap_(clusterSheet);
  const known = {};
  existingRows_(clusterSheet).forEach(function (r) { known[r['Cluster Name']] = r; });

  const membersByCluster = {};
  offices.forEach(function (o) {
    const cluster = o['Cluster'];
    if (!cluster) return;
    membersByCluster[cluster] = membersByCluster[cluster] || [];
    membersByCluster[cluster].push(o['Office Code']);
  });

  const newRows = [];
  Object.keys(membersByCluster).forEach(function (cluster, i) {
    if (known[cluster]) return;
    newRows.push(orderRow_(clusterMap, {
      'Cluster ID': 'CL' + (i + 1), 'Cluster Name': cluster,
      'Member Offices': membersByCluster[cluster].join(', '), 'Aggregate Weight (Check)': '', 'Achievement %': ''
    }));
  });
  appendRows_(clusterSheet, newRows);
}

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
    const officeKpis = kpiRows.filter(function (k) { return k[kpiMap['Office'] - 1] === code; });
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
