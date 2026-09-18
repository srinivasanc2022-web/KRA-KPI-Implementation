/**
 * 99_Triggers.gs
 *
 * Simple onEdit trigger: audits direct in-sheet edits to the fields the spec
 * calls out explicitly ("audit trail of changes to targets, weights, owners
 * and status" -- section 7) and re-validates an office's weight total the
 * moment its Wt (%) column changes, surfacing a toast immediately rather than
 * waiting for the next manual "Validate Weights" run.
 *
 * Quarterly-entry-form edits are already audited in 06_QuarterlyEntry.gs;
 * this trigger exists for edits made directly on the sheet (e.g. an admin
 * bulk-correcting weights).
 */

const AUDITED_KPI_FIELDS = [
  'Wt (%)', 'Annual Target (AY 26-27)', 'KPI Owner',
  'R (Responsible)', 'A (Accountable)', 'C (Consulted)', 'I (Informed)', 'Responsible Officer',
  'Q1 Milestone Target', 'Q2 Milestone Target', 'Q3 Milestone Target', 'Q4 Milestone Target',
  'Q1 Status', 'Q2 Status', 'Q3 Status', 'Q4 Status'
];

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET.KPIS || e.range.getRow() === 1) return;

  const map = headerIndexMap_(sheet);
  const col = e.range.getColumn();
  const fieldName = Object.keys(map).filter(function (h) { return map[h] === col; })[0];
  if (!fieldName || AUDITED_KPI_FIELDS.indexOf(fieldName) === -1) return;

  const ss = e.source;
  logAudit_(ss, SHEET.KPIS, e.range.getRow(), fieldName, e.oldValue, e.value);

  if (fieldName === 'Wt (%)') {
    const office = sheet.getRange(e.range.getRow(), map['Office']).getValue();
    checkOfficeWeightNow_(ss, sheet, map, office);
  }
}

function checkOfficeWeightNow_(ss, sheet, map, office) {
  if (!office || sheet.getLastRow() < 2) return;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const total = rows
    .filter(function (r) { return r[map['Office'] - 1] === office; })
    .reduce(function (sum, r) { return sum + (parseFloat(r[map['Wt (%)'] - 1]) || 0); }, 0);

  const tolerance = getConfigNumber_(ss, 'Weight_Tolerance', 0.01);
  if (Math.abs(total - 100) > tolerance) {
    ss.toast(office + ' KPI weights now total ' + total.toFixed(2) + ' (expected 100). Publication should be blocked until this is fixed.', 'Weight check failed', 8);
  }
}
