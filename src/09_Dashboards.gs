/**
 * 09_Dashboards.gs
 *
 * Spec section 8's five dashboard views, built the way that best fits a
 * Sheets-native system:
 *   - VC and Cluster dashboards are SCRIPT-COMPUTED SNAPSHOTS (buildDashboards()
 *     writes them). They involve aggregation logic (RAG counts, top-N
 *     escalations) that's clearer in code than in a spreadsheet formula.
 *   - Office, Goal and My KPIs dashboards are LIVE QUERY()-FORMULA SHEETS
 *     driven by a picker cell. They stay current automatically (no
 *     "refresh" step) because Sheets recalculates QUERY() whenever the
 *     underlying KPIs sheet or the picker cell changes -- the idiomatic
 *     Sheets way to build a filterable drill-down, rather than
 *     reimplementing a filter UI in Apps Script.
 *
 * For the fuller multi-dimension filtering spec section 8 asks for on the
 * Office dashboard (KRA Type / Metric Type / Indicator Nature / Status),
 * each dashboard's filtered table is left as a normal range so you can lay
 * a native Sheets filter view over it (Data > Create a filter) instead of a
 * second custom UI -- also idiomatic, and free.
 */

function buildDashboards() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  buildVcDashboard_(ss);
  buildClusterDashboard_(ss);
  buildOfficeDashboard_(ss);
  buildGoalDashboard_(ss);
  buildMyKpisDashboard_(ss);
  SpreadsheetApp.getUi().alert('Dashboards refreshed: ' + Object.values(DASHBOARD_SHEET).join(', '));
}

// ---------------------------------------------------------------------------
// VC / Institutional dashboard (script-computed snapshot)
// ---------------------------------------------------------------------------

function buildVcDashboard_(ss) {
  const sheet = getOrCreateBlank_(ss, DASHBOARD_SHEET.VC);
  let row = 1;
  sheet.getRange(row, 1).setValue('VC / Institutional Dashboard').setFontWeight('bold').setFontSize(14);
  row += 2;

  const pillars = existingRows_(ss.getSheetByName(SHEET.PILLARS));
  row = writeTable_(sheet, row, 'Pillar-wise Achievement', ['Pillar ID', 'Pillar Name', 'Achievement %'],
    pillars.map(function (p) { return [p['Pillar ID'], p['Pillar Name'], p['Achievement %']]; }));

  const goals = existingRows_(ss.getSheetByName(SHEET.GOALS));
  row = writeTable_(sheet, row, 'Goal-wise Achievement', ['Goal ID', 'Parent Pillar', 'Owner Office', 'Achievement %'],
    goals.map(function (g) { return [g['Goal ID'], g['Parent Pillar'], g['Owner Office'], g['Achievement %']]; }));

  const ragCounts = countInstitutionWideRag_(ss);
  row = writeTable_(sheet, row, 'KPI Status Counts (institution-wide, latest quarter with a status)',
    ['Status', 'Count'],
    [[RAG.GREEN, ragCounts.Green], [RAG.AMBER, ragCounts.Amber], [RAG.RED, ragCounts.Red], [RAG.GREY, ragCounts.Grey]]);

  const escalations = topOpenAtrs_(ss, 15);
  writeTable_(sheet, row, 'Top Open ATRs (corrective actions on Amber/Red KPIs)',
    ['KPI_ID', 'Office', 'Quarter', 'Status', 'KPI', 'ATR / Corrective Action'],
    escalations.map(function (e) { return [e.kpiId, e.office, e.quarter, e.status, e.kpi, e.atr]; }));

  sheet.autoResizeColumns(1, 6);
}

function countInstitutionWideRag_(ss) {
  const sheet = ss.getSheetByName(SHEET.KPIS);
  const map = headerIndexMap_(sheet);
  const counts = { Green: 0, Amber: 0, Red: 0, Grey: 0 };
  if (sheet.getLastRow() < 2) return counts;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  rows.forEach(function (row) {
    const status = latestQuarterStatus_(row, map);
    if (status && counts.hasOwnProperty(status)) counts[status]++;
  });
  return counts;
}

/** The most recent quarter (Q4 first) that has a Status value set on this KPI row. */
function latestQuarterStatus_(row, map) {
  for (let q = QUARTERS.length - 1; q >= 0; q--) {
    const status = row[map[QUARTERS[q] + ' Status'] - 1];
    if (status) return status;
  }
  return null;
}

function topOpenAtrs_(ss, limit) {
  const sheet = ss.getSheetByName(SHEET.KPIS);
  const map = headerIndexMap_(sheet);
  const results = [];
  if (sheet.getLastRow() < 2) return results;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  rows.forEach(function (row) {
    QUARTERS.forEach(function (q) {
      const status = row[map[q + ' Status'] - 1];
      const atr = row[map[q + ' ATR'] - 1];
      if ((status === RAG.AMBER || status === RAG.RED) && atr) {
        results.push({
          kpiId: row[map['KPI_ID'] - 1], office: row[map['Office'] - 1], quarter: q,
          status: status, kpi: row[map['KPI'] - 1], atr: atr
        });
      }
    });
  });
  results.sort(function (a, b) { return (a.status === RAG.RED ? 0 : 1) - (b.status === RAG.RED ? 0 : 1); });
  return results.slice(0, limit);
}

// ---------------------------------------------------------------------------
// Cluster dashboard (script-computed snapshot; all clusters stacked -- there
// are only 4 + VC Direct, so one scrollable sheet beats 5 separate tabs)
// ---------------------------------------------------------------------------

function buildClusterDashboard_(ss) {
  const sheet = getOrCreateBlank_(ss, DASHBOARD_SHEET.CLUSTER);
  let row = 1;
  sheet.getRange(row, 1).setValue('Cluster Dashboard').setFontWeight('bold').setFontSize(14);
  row += 2;

  const clusters = existingRows_(ss.getSheetByName(SHEET.CLUSTERS));
  const offices = existingRows_(ss.getSheetByName(SHEET.OFFICES));

  clusters.forEach(function (c) {
    sheet.getRange(row, 1).setValue(c['Cluster Name'] + '  --  Achievement: ' + (c['Achievement %'] || 'n/a') + '%')
      .setFontWeight('bold').setFontSize(12).setBackground('#d9e2f3');
    row++;
    const memberCodes = String(c['Member Offices'] || '').split(',').map(function (s) { return s.trim(); }).filter(String);
    const memberRows = offices.filter(function (o) { return memberCodes.indexOf(o['Office Code']) !== -1; })
      .map(function (o) { return [o['Office Code'], o['Full Name'], o['Achievement %'], o['Total Weight (Check)']]; });
    row = writeTable_(sheet, row, '', ['Office Code', 'Full Name', 'Achievement %', 'Total Weight (Check)'], memberRows);
  });

  sheet.autoResizeColumns(1, 4);
}

// ---------------------------------------------------------------------------
// Office dashboard (live QUERY formula, driven by a picker cell)
// ---------------------------------------------------------------------------

function buildOfficeDashboard_(ss) {
  const sheet = getOrCreateBlank_(ss, DASHBOARD_SHEET.OFFICE);
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  const lastCol = colLetter_(kpiSheet.getLastColumn());
  const lastRow = Math.max(kpiSheet.getLastRow(), 2);
  const officeCol = colLetter_(kpiMap['Office']);
  const wtCol = colLetter_(kpiMap['Wt (%)']);

  sheet.getRange(1, 1).setValue('Office Dashboard').setFontWeight('bold').setFontSize(14);
  sheet.getRange(3, 1).setValue('Select Office:').setFontWeight('bold');
  sheet.getRange(3, 2).setValue('AA');
  const officeCodes = existingRows_(ss.getSheetByName(SHEET.OFFICES)).map(function (o) { return o['Office Code']; }).filter(String);
  if (officeCodes.length > 0) {
    sheet.getRange(3, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(officeCodes, true).setAllowInvalid(true).build());
  }

  sheet.getRange(4, 1).setValue('Total Weight (should = 100):').setFontWeight('bold');
  sheet.getRange(4, 2).setFormula('=SUMIF(KPIs!' + officeCol + ':' + officeCol + ',B3,KPIs!' + wtCol + ':' + wtCol + ')');

  sheet.getRange(6, 1).setValue(
    'KPIs for the selected office (tip: select this range and use Data > Create a filter to filter by KRA Type / Metric Type / Indicator Nature / Status).'
  ).setFontStyle('italic');
  sheet.getRange(7, 1).setFormula(
    '=QUERY(KPIs!A1:' + lastCol + lastRow + ', "select * where ' + officeCol + ' = \'"&B3&"\'", 1)'
  );
}

// ---------------------------------------------------------------------------
// Goal dashboard (live QUERY formula, driven by a picker cell)
// ---------------------------------------------------------------------------

function buildGoalDashboard_(ss) {
  const sheet = getOrCreateBlank_(ss, DASHBOARD_SHEET.GOAL);
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  const lastCol = colLetter_(kpiSheet.getLastColumn());
  const lastRow = Math.max(kpiSheet.getLastRow(), 2);
  const goalCol = colLetter_(kpiMap['Goal ID']);

  sheet.getRange(1, 1).setValue('Goal Dashboard').setFontWeight('bold').setFontSize(14);
  sheet.getRange(3, 1).setValue('Select Goal ID:').setFontWeight('bold');
  const goalRows = existingRows_(ss.getSheetByName(SHEET.GOALS));
  const goalIds = goalRows.map(function (g) { return g['Goal ID']; }).filter(String);
  sheet.getRange(3, 2).setValue(goalIds[0] || '');
  if (goalIds.length > 0) {
    sheet.getRange(3, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(goalIds, true).setAllowInvalid(true).build());
  }

  const goalSheet = ss.getSheetByName(SHEET.GOALS);
  const goalMap = headerIndexMap_(goalSheet);
  const goalLastCol = colLetter_(goalSheet.getLastColumn());
  const ownerOfficeIdx = goalMap['Owner Office'];
  const contribIdx = goalMap['Contributing Offices'];
  const achievementIdx = goalMap['Achievement %'];

  sheet.getRange(4, 1).setValue('Owner Office:').setFontWeight('bold');
  sheet.getRange(4, 2).setFormula('=IFERROR(VLOOKUP(B3, Goals!A:' + goalLastCol + ', ' + ownerOfficeIdx + ', FALSE), "")');
  sheet.getRange(5, 1).setValue('Contributing Offices:').setFontWeight('bold');
  sheet.getRange(5, 2).setFormula('=IFERROR(VLOOKUP(B3, Goals!A:' + goalLastCol + ', ' + contribIdx + ', FALSE), "")');
  sheet.getRange(6, 1).setValue('Achievement %:').setFontWeight('bold');
  sheet.getRange(6, 2).setFormula('=IFERROR(VLOOKUP(B3, Goals!A:' + goalLastCol + ', ' + achievementIdx + ', FALSE), "")');

  sheet.getRange(8, 1).setValue('Every KPI institution-wide that feeds this goal (cross-office):').setFontStyle('italic');
  sheet.getRange(9, 1).setFormula(
    '=QUERY(KPIs!A1:' + lastCol + lastRow + ', "select * where ' + goalCol + ' = \'"&B3&"\'", 1)'
  );
}

// ---------------------------------------------------------------------------
// Individual accountability dashboard (live QUERY formula)
// ---------------------------------------------------------------------------

function buildMyKpisDashboard_(ss) {
  const sheet = getOrCreateBlank_(ss, DASHBOARD_SHEET.MY_KPIS);
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  const lastCol = colLetter_(kpiSheet.getLastColumn());
  const lastRow = Math.max(kpiSheet.getLastRow(), 2);
  const ownerCol = colLetter_(kpiMap['KPI Owner']);
  const officerCol = colLetter_(kpiMap['Responsible Officer']);

  sheet.getRange(1, 1).setValue('My KPIs (Individual Accountability)').setFontWeight('bold').setFontSize(14);
  sheet.getRange(3, 1).setValue('KPI Owner or Responsible Officer name (as it appears in the KPIs sheet):').setFontWeight('bold');
  sheet.getRange(3, 2).setValue('');

  sheet.getRange(5, 1).setValue('KPIs owned or overseen, with targets vs. actuals and open ATRs:').setFontStyle('italic');
  sheet.getRange(6, 1).setFormula(
    '=QUERY(KPIs!A1:' + lastCol + lastRow + ', "select * where ' + ownerCol + ' = \'"&B3&"\' or ' + officerCol + ' = \'"&B3&"\'", 1)'
  );
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

// getOrCreateBlank_() is shared from 02_SampleData.gs (clears an existing
// sheet's content or creates a new one -- identical need here).

/** Writes a titled table starting at `row`; returns the next free row (with a blank-line gap). */
function writeTable_(sheet, row, title, headers, dataRows) {
  if (title) {
    sheet.getRange(row, 1).setValue(title).setFontWeight('bold').setFontSize(12);
    row++;
  }
  if (headers && headers.length) {
    sheet.getRange(row, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8eef7');
    row++;
  }
  if (dataRows && dataRows.length) {
    sheet.getRange(row, 1, dataRows.length, headers.length).setValues(dataRows);
    row += dataRows.length;
  }
  return row + 1;
}

/** 1-based column index -> spreadsheet column letters (1 -> A, 27 -> AA). */
function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
