/**
 * 05_RAGEngine.gs
 *
 * Computes quarterly RAG (Red/Amber/Green/Grey) status from Actual vs.
 * Milestone Target, per spec section 6:
 *   Green >= 95% of milestone target (configurable)
 *   Amber  80-94%
 *   Red   < 80%
 *   Grey   no actual logged by the quarter's close
 *
 * Band/Grade-unit KPIs aren't numerically comparable, so they always require
 * manual status entry (the dropdown added in 01_SchemaSetup.gs still lets a
 * KPI Owner set Green/Amber/Red/Grey by hand for those rows).
 */

/**
 * Returns RAG.GREEN/AMBER/RED/GREY, or null if the KPI's unit isn't
 * computable and the caller should leave Status for manual entry.
 */
function computeRagStatus_(actual, milestoneTarget, unit, thresholds) {
  if (NON_COMPUTABLE_UNITS.indexOf(unit) !== -1) return null;
  if (actual === '' || actual === null || actual === undefined) return RAG.GREY;

  const actualNum = parseFloat(actual);
  const targetNum = parseFloat(milestoneTarget);
  if (isNaN(actualNum) || isNaN(targetNum) || targetNum === 0) return null;

  const ratio = actualNum / targetNum;
  if (ratio >= thresholds.GREEN_MIN) return RAG.GREEN;
  if (ratio >= thresholds.AMBER_MIN) return RAG.AMBER;
  return RAG.RED;
}

/**
 * Recomputes Status for every quarter on every KPI row where an Actual has
 * been entered and Status is currently blank or was last auto-computed
 * (tracked implicitly: we only overwrite blank cells so a manual override by
 * a KPI Owner is never clobbered -- see recomputeQuarterStatus_ below for the
 * single-cell version used by the quarterly entry form).
 */
function recomputeAllRagStatuses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const thresholds = getThresholds_(ss);
  const sheet = ss.getSheetByName(SHEET.KPIS);
  const map = headerIndexMap_(sheet);
  if (sheet.getLastRow() < 2) return;

  const numRows = sheet.getLastRow() - 1;
  const data = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();
  let updated = 0;

  QUARTERS.forEach(function (q) {
    const actualCol = map[q + ' Actual'];
    const targetCol = map[q + ' Milestone Target'];
    const statusCol = map[q + ' Status'];
    const unitCol = map['Unit'];
    if (!actualCol || !targetCol || !statusCol) return;

    data.forEach(function (row, i) {
      const currentStatus = row[statusCol - 1];
      if (currentStatus) return; // never overwrite an existing (manual or prior) status
      const computed = computeRagStatus_(row[actualCol - 1], row[targetCol - 1], row[unitCol - 1], thresholds);
      if (computed) {
        sheet.getRange(2 + i, statusCol).setValue(computed);
        updated++;
      }
    });
  });

  SpreadsheetApp.getUi().alert(updated + ' quarterly status cell(s) computed.');
}

/** Single-KPI, single-quarter recompute used by the quarterly entry form (06_QuarterlyEntry.gs). */
function recomputeQuarterStatus_(ss, kpiId, quarter) {
  const sheet = ss.getSheetByName(SHEET.KPIS);
  const map = headerIndexMap_(sheet);
  const rowNum = findKpiRow_(sheet, map, kpiId);
  if (!rowNum) return null;

  const thresholds = getThresholds_(ss);
  const rowValues = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  const actual = rowValues[map[quarter + ' Actual'] - 1];
  const target = rowValues[map[quarter + ' Milestone Target'] - 1];
  const unit = rowValues[map['Unit'] - 1];

  const status = computeRagStatus_(actual, target, unit, thresholds);
  if (status) sheet.getRange(rowNum, map[quarter + ' Status']).setValue(status);
  return status;
}

function findKpiRow_(sheet, map, kpiId) {
  if (sheet.getLastRow() < 2) return null;
  const ids = sheet.getRange(2, map['KPI_ID'], sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === kpiId) return i + 2;
  }
  return null;
}

function getThresholds_(ss) {
  return {
    GREEN_MIN: getConfigNumber_(ss, 'Green_Threshold', DEFAULT_THRESHOLDS.GREEN_MIN),
    AMBER_MIN: getConfigNumber_(ss, 'Amber_Threshold', DEFAULT_THRESHOLDS.AMBER_MIN)
  };
}
