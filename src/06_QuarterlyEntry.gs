/**
 * 06_QuarterlyEntry.gs
 *
 * Server-side functions backing html/QuarterlyEntryForm.html -- the narrow,
 * functional screen from spec build-note #3: "Screens for KPI Owners to
 * enter Q1-Q4 actuals, trigger RAG status, and log ATR notes when a quarter
 * is Amber/Red."
 *
 * All writes go through submitQuarterlyActual_() so RAG computation, the
 * ATR-required-on-Amber/Red rule, and the audit log stay in one place rather
 * than being re-implemented per caller.
 */

function openQuarterlyEntrySidebar() {
  const html = HtmlService.createHtmlOutputFromFile('QuarterlyEntryForm')
    .setTitle('Quarterly KPI Entry');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Called by the sidebar on load: returns the KPI list for the picker, optionally filtered by office. */
function listKpisForEntry(officeFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.KPIS);
  const rows = existingRows_(sheet);
  return rows
    .filter(function (r) { return !officeFilter || r['Office'] === officeFilter; })
    .map(function (r) {
      return {
        kpiId: r['KPI_ID'], office: r['Office'], kra: r['KRA'], kpi: r['KPI'],
        unit: r['Unit'], frequency: r['Frequency'], owner: r['KPI Owner'],
        q1t: r['Q1 Milestone Target'], q2t: r['Q2 Milestone Target'],
        q3t: r['Q3 Milestone Target'], q4t: r['Q4 Milestone Target'],
        q1a: r['Q1 Actual'], q2a: r['Q2 Actual'], q3a: r['Q3 Actual'], q4a: r['Q4 Actual'],
        q1s: r['Q1 Status'], q2s: r['Q2 Status'], q3s: r['Q3 Status'], q4s: r['Q4 Status']
      };
    });
}

function listOfficesForEntry() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return existingRows_(ss.getSheetByName(SHEET.OFFICES)).map(function (r) { return r['Office Code']; }).filter(String);
}

/**
 * Writes one quarter's Actual (+ optional ATR) for one KPI, recomputes RAG
 * status, and enforces "ATR required whenever Status is Amber or Red"
 * (spec section 6). Returns { ok, status, message }.
 */
function submitQuarterlyActual(kpiId, quarter, actualValue, atrText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.KPIS);
  const map = headerIndexMap_(sheet);
  const rowNum = findKpiRow_(sheet, map, kpiId);
  if (!rowNum) return { ok: false, message: 'KPI not found: ' + kpiId };

  const oldActual = sheet.getRange(rowNum, map[quarter + ' Actual']).getValue();
  sheet.getRange(rowNum, map[quarter + ' Actual']).setValue(actualValue);
  logAudit_(ss, SHEET.KPIS, rowNum, quarter + ' Actual', oldActual, actualValue);

  // Clear any prior status so the recompute below isn't blocked by its own "don't overwrite" guard.
  sheet.getRange(rowNum, map[quarter + ' Status']).setValue('');
  const status = recomputeQuarterStatus_(ss, kpiId, quarter);

  if ((status === RAG.AMBER || status === RAG.RED) && !atrText) {
    return {
      ok: false,
      status: status,
      message: 'Status computed as ' + status + '. An ATR (corrective-action) note is required before this can be saved.'
    };
  }

  if (atrText) {
    const oldAtr = sheet.getRange(rowNum, map[quarter + ' ATR']).getValue();
    sheet.getRange(rowNum, map[quarter + ' ATR']).setValue(atrText);
    logAudit_(ss, SHEET.KPIS, rowNum, quarter + ' ATR', oldAtr, atrText);
  }

  return { ok: true, status: status || '(unit requires manual status)', message: 'Saved.' };
}

/** Lets a KPI Owner set Status manually for non-computable units (Band/Grade) -- still ATR-gated on Amber/Red. */
function submitManualStatus(kpiId, quarter, statusValue, atrText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET.KPIS);
  const map = headerIndexMap_(sheet);
  const rowNum = findKpiRow_(sheet, map, kpiId);
  if (!rowNum) return { ok: false, message: 'KPI not found: ' + kpiId };

  if ((statusValue === RAG.AMBER || statusValue === RAG.RED) && !atrText) {
    return { ok: false, message: 'Status is ' + statusValue + '. An ATR note is required before this can be saved.' };
  }

  const oldStatus = sheet.getRange(rowNum, map[quarter + ' Status']).getValue();
  sheet.getRange(rowNum, map[quarter + ' Status']).setValue(statusValue);
  logAudit_(ss, SHEET.KPIS, rowNum, quarter + ' Status', oldStatus, statusValue);

  if (atrText) {
    const oldAtr = sheet.getRange(rowNum, map[quarter + ' ATR']).getValue();
    sheet.getRange(rowNum, map[quarter + ' ATR']).setValue(atrText);
    logAudit_(ss, SHEET.KPIS, rowNum, quarter + ' ATR', oldAtr, atrText);
  }

  return { ok: true, message: 'Saved.' };
}

function logAudit_(ss, sheetName, row, field, oldValue, newValue) {
  if (String(oldValue) === String(newValue)) return;
  const sheet = ss.getSheetByName(SHEET.AUDIT_LOG);
  sheet.appendRow([new Date(), Session.getActiveUser().getEmail() || '(unknown)', sheetName, row, field, oldValue, newValue]);
}
