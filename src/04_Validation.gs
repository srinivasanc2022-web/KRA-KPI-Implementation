/**
 * 04_Validation.gs
 *
 * Enforces "weights sum to 100" at Office, Cluster and Goal level (spec
 * section 4 / section 12: "Enforce weights sum to 100 ... as a hard
 * validation before publishing any scorecard"). Also flags publication
 * blockers other modules key off (see PUBLISH_BLOCKED in 06_QuarterlyEntry.gs
 * and the escalation rules described in NEXT_STEPS.md for phase 3).
 *
 * Office-level and Cluster-level weight sums are exact 100 checks.
 * Goal-level is different by design: the spec says the owner/contributor
 * split is admin-set per goal (Owner Weight % on the Goals sheet), not a
 * hard 100-across-offices rule, so goal-level validation only checks that an
 * Owner Weight % has been set once a goal has contributing offices.
 */

function validateAllWeights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearValidationLog_(ss);
  const tolerance = getConfigNumber_(ss, 'Weight_Tolerance', 0.01);

  const officeResults = validateOfficeWeights_(ss, tolerance);
  const clusterResults = validateClusterWeights_(ss, tolerance);
  const goalResults = validateGoalOwnerWeights_(ss);

  const all = officeResults.concat(clusterResults, goalResults);
  writeValidationLog_(ss, all);

  const failures = all.filter(function (r) { return r.status === 'FAIL'; });
  return { total: all.length, failures: failures.length, results: all };
}

function validateOfficeWeights_(ss, tolerance) {
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  if (kpiSheet.getLastRow() < 2) return [];
  const rows = kpiSheet.getRange(2, 1, kpiSheet.getLastRow() - 1, kpiSheet.getLastColumn()).getValues();

  // Only Tier 2 office scorecards must sum to exactly 100 (Tier 0/1 aggregate
  // KPIs are validated separately at cluster level).
  const byOffice = {};
  rows.forEach(function (r) {
    const tier = String(r[kpiMap['Tier'] - 1]);
    if (tier !== '2') return;
    const office = r[kpiMap['Office'] - 1];
    if (!office) return;
    const wt = parseFloat(r[kpiMap['Wt (%)'] - 1]) || 0;
    byOffice[office] = (byOffice[office] || 0) + wt;
  });

  return Object.keys(byOffice).map(function (office) {
    const total = byOffice[office];
    const ok = Math.abs(total - 100) <= tolerance;
    return {
      scope: 'Office', entityId: office, totalWeight: total, status: ok ? 'OK' : 'FAIL',
      details: ok ? 'Weights sum to 100' : ('Weights sum to ' + total.toFixed(2) + ', expected 100')
    };
  });
}

function validateClusterWeights_(ss, tolerance) {
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  if (kpiSheet.getLastRow() < 2) return [];
  const rows = kpiSheet.getRange(2, 1, kpiSheet.getLastRow() - 1, kpiSheet.getLastColumn()).getValues();

  // Tier 0/1 KPIs (cluster and institutional aggregates) must sum to 100 per cluster.
  const byCluster = {};
  rows.forEach(function (r) {
    const tier = String(r[kpiMap['Tier'] - 1]);
    if (tier !== '0' && tier !== '1') return;
    const cluster = r[kpiMap['Cluster'] - 1];
    if (!cluster) return;
    const wt = parseFloat(r[kpiMap['Wt (%)'] - 1]) || 0;
    byCluster[cluster] = (byCluster[cluster] || 0) + wt;
  });

  return Object.keys(byCluster).map(function (cluster) {
    const total = byCluster[cluster];
    const ok = Math.abs(total - 100) <= tolerance;
    return {
      scope: 'Cluster', entityId: cluster, totalWeight: total, status: ok ? 'OK' : 'FAIL',
      details: ok ? 'Weights sum to 100' : ('Weights sum to ' + total.toFixed(2) + ', expected 100')
    };
  });
}

function validateGoalOwnerWeights_(ss) {
  const goalSheet = ss.getSheetByName(SHEET.GOALS);
  if (goalSheet.getLastRow() < 2) return [];
  const rows = existingRows_(goalSheet);
  return rows.filter(function (r) { return String(r['Contributing Offices'] || '').trim().length > 0; })
    .map(function (r) {
      const hasSplit = r['Owner Weight %'] !== '' && r['Owner Weight %'] !== null && r['Owner Weight %'] !== undefined;
      return {
        scope: 'Goal', entityId: r['Goal ID'],
        totalWeight: hasSplit ? r['Owner Weight %'] : '',
        status: hasSplit ? 'OK' : 'FAIL',
        details: hasSplit
          ? ('Owner weight set to ' + r['Owner Weight %'] + '%')
          : 'Owner Weight % not set -- admin must set the owner/contributor roll-up split before publication (section 4)'
      };
    });
}

function clearValidationLog_(ss) {
  const sheet = ss.getSheetByName(SHEET.VALIDATION_LOG);
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
}

function writeValidationLog_(ss, results) {
  if (results.length === 0) return;
  const sheet = ss.getSheetByName(SHEET.VALIDATION_LOG);
  const now = new Date();
  const rows = results.map(function (r) { return [now, r.scope, r.entityId, r.totalWeight, r.status, r.details]; });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

  // Highlight failures red for quick scanning.
  const statusCol = headerIndexMap_(sheet)['Status'];
  rows.forEach(function (r, i) {
    if (r[4] === 'FAIL') {
      sheet.getRange(2 + i, 1, 1, rows[0].length).setBackground('#f4cccc');
    }
  });
}

function getConfigNumber_(ss, key, defaultValue) {
  const sheet = ss.getSheetByName(SHEET.CONFIG);
  const rows = existingRows_(sheet);
  const match = rows.filter(function (r) { return r['Key'] === key; })[0];
  const parsed = match ? parseFloat(match['Value']) : NaN;
  return isNaN(parsed) ? defaultValue : parsed;
}

/** Menu-facing wrapper: runs validation and shows a summary dialog. */
function runWeightValidationFromMenu() {
  const result = validateAllWeights();
  const message = result.failures === 0
    ? ('All ' + result.total + ' weight checks passed.')
    : (result.failures + ' of ' + result.total + ' weight checks FAILED. See the "' + SHEET.VALIDATION_LOG + '" sheet.');
  SpreadsheetApp.getUi().alert(message);
}
