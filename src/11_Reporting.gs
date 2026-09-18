/**
 * 11_Reporting.gs
 *
 * Spec section 11's reports, plus regenerateAllScorecardTabs() which rebuilds
 * one tab per Tier 0/1/2 scorecard from the canonical KPIs sheet -- once
 * those tabs exist, File > Download > Microsoft Excel (.xlsx) in the Google
 * Sheets UI IS the "export back to the same tab structure as the source
 * workbook" the spec asks for; no custom export code is needed on top of
 * Sheets' own native download.
 *
 * These regenerated tabs use the CANONICAL column set (00_Constants.gs
 * COLS.KPIS), not a byte-for-byte reproduction of the original workbook's
 * exact header wording/newlines -- the continuity spec section 11 asks for
 * is "one tab per scorecard with the full field set", which this gives you;
 * it does not try to recreate the original file's cosmetic formatting.
 */

function generateQuarterlyScorecardReport(officeCode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  regenerateScorecardTab_(ss, officeCode, officeCode);
  ss.getSheetByName(officeCode).activate();
  SpreadsheetApp.getUi().alert('Quarterly Scorecard Report regenerated for ' + officeCode + '.');
}

function generateClusterVcReviewPack() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateBlank_(ss, 'Report_Cluster_VC_Review_Pack');
  let row = 1;
  sheet.getRange(row, 1).setValue('Cluster / VC Review Pack').setFontWeight('bold').setFontSize(14);
  row += 2;
  sheet.getRange(row, 1).setValue('Generated ' + formatDate_(new Date())).setFontStyle('italic');
  row += 2;

  const clusters = existingRows_(ss.getSheetByName(SHEET.CLUSTERS));
  row = writeTable_(sheet, row, 'Cluster-wise Achievement', ['Cluster Name', 'Achievement %', 'Member Offices'],
    clusters.map(function (c) { return [c['Cluster Name'], c['Achievement %'], c['Member Offices']]; }));

  const goals = existingRows_(ss.getSheetByName(SHEET.GOALS));
  row = writeTable_(sheet, row, 'Goal-wise Achievement', ['Goal ID', 'Parent Pillar', 'Owner Office', 'Contributing Offices', 'Achievement %'],
    goals.map(function (g) { return [g['Goal ID'], g['Parent Pillar'], g['Owner Office'], g['Contributing Offices'], g['Achievement %']]; }));

  const ragCounts = countInstitutionWideRag_(ss);
  row = writeTable_(sheet, row, 'RAG Summary (institution-wide)', ['Status', 'Count'],
    [[RAG.GREEN, ragCounts.Green], [RAG.AMBER, ragCounts.Amber], [RAG.RED, ragCounts.Red], [RAG.GREY, ragCounts.Grey]]);

  const openAtrs = topOpenAtrs_(ss, 100);
  writeTable_(sheet, row, 'Open ATRs (all Amber/Red quarters with a corrective-action note)',
    ['KPI_ID', 'Office', 'Quarter', 'Status', 'KPI', 'ATR / Corrective Action'],
    openAtrs.map(function (e) { return [e.kpiId, e.office, e.quarter, e.status, e.kpi, e.atr]; }));

  sheet.autoResizeColumns(1, 6);
  sheet.activate();
  SpreadsheetApp.getUi().alert('Cluster/VC Review Pack generated.');
}

function generateOwnerPerformanceReport(ownerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const safeName = String(ownerName).replace(/[^A-Za-z0-9 _-]/g, '').slice(0, 40);
  const sheet = getOrCreateBlank_(ss, 'Report_Owner_' + safeName);
  let row = 1;
  sheet.getRange(row, 1).setValue('Owner Performance Report: ' + ownerName).setFontWeight('bold').setFontSize(14);
  row += 2;

  const kpis = existingRows_(ss.getSheetByName(SHEET.KPIS))
    .filter(function (r) { return r['KPI Owner'] === ownerName || r['Responsible Officer'] === ownerName; });

  row = writeTable_(sheet, row, 'KPIs Owned / Overseen',
    ['KPI_ID', 'Office', 'KPI', 'Unit', 'Annual Target (AY 26-27)', 'Achievement %', 'Q4 Status'],
    kpis.map(function (k) { return [k['KPI_ID'], k['Office'], k['KPI'], k['Unit'], k['Annual Target (AY 26-27)'], k['Achievement %'], k['Q4 Status']]; }));

  const openAtrKpis = [];
  kpis.forEach(function (k) {
    QUARTERS.forEach(function (q) {
      if ((k[q + ' Status'] === RAG.AMBER || k[q + ' Status'] === RAG.RED) && k[q + ' ATR']) {
        openAtrKpis.push([k['KPI_ID'], q, k[q + ' Status'], k[q + ' ATR']]);
      }
    });
  });
  row = writeTable_(sheet, row, 'Open ATRs', ['KPI_ID', 'Quarter', 'Status', 'ATR / Corrective Action'], openAtrKpis);

  const closeDates = getQuarterCloseDates_(ss);
  const upcoming = [];
  kpis.forEach(function (k) {
    QUARTERS.forEach(function (q) {
      const actual = k[q + ' Actual'];
      const due = closeDates[q];
      if (due && due > new Date() && (actual === '' || actual === null || actual === undefined)) {
        upcoming.push([k['KPI_ID'], q, formatDate_(due), k['KPI']]);
      }
    });
  });
  upcoming.sort(function (a, b) { return new Date(a[2]) - new Date(b[2]); });
  writeTable_(sheet, row, 'Upcoming Quarterly Deadlines (no Actual logged yet)', ['KPI_ID', 'Quarter', 'Due', 'KPI'], upcoming);

  sheet.autoResizeColumns(1, 7);
  sheet.activate();
  SpreadsheetApp.getUi().alert('Owner Performance Report generated for ' + ownerName + '.');
}

// ---------------------------------------------------------------------------
// Regenerate original scorecard-shaped tabs (for native Excel export)
// ---------------------------------------------------------------------------

function regenerateAllScorecardTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  TIER0_1_SOURCE_SHEETS.forEach(function (name) {
    regenerateScorecardTab_(ss, name, TIER0_1_CODES[name]);
  });
  existingRows_(ss.getSheetByName(SHEET.OFFICES)).forEach(function (o) {
    if (o['Office Code']) regenerateScorecardTab_(ss, o['Office Code'], o['Office Code']);
  });
  SpreadsheetApp.getUi().alert(
    'Regenerated ' + (TIER0_1_SOURCE_SHEETS.length + existingRows_(ss.getSheetByName(SHEET.OFFICES)).length) +
    ' scorecard tabs from the canonical KPIs sheet.\n\n' +
    'Use File > Download > Microsoft Excel (.xlsx) now for a drop-in export.'
  );
}

function regenerateScorecardTab_(ss, sheetName, kpiIdPrefix) {
  const sheet = getOrCreateBlank_(ss, sheetName);
  const ay = getConfigValue_(ss, 'Current_AY', 'AY26-27');
  sheet.getRange(1, 1).setValue(sheetName + ' -- regenerated ' + formatDate_(new Date()) + ' (' + ay + ')');
  sheet.getRange(2, 1, 1, COLS.KPIS.length).setValues([COLS.KPIS])
    .setFontWeight('bold').setBackground('#1c4587').setFontColor('#ffffff');

  const rows = existingRows_(ss.getSheetByName(SHEET.KPIS))
    .filter(function (r) { return String(r['KPI_ID'] || '').indexOf(kpiIdPrefix + '-') === 0; });
  const dataRows = rows.map(function (r) { return COLS.KPIS.map(function (h) { return r[h]; }); });
  if (dataRows.length > 0) {
    sheet.getRange(3, 1, dataRows.length, COLS.KPIS.length).setValues(dataRows);
  }
  sheet.setFrozenRows(2);
}
