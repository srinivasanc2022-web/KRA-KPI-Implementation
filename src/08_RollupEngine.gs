/**
 * 08_RollupEngine.gs
 *
 * KPI -> Office -> Cluster -> Goal -> Pillar achievement roll-up, per spec
 * section 4. Written as an isolated, single-entry-point calculation layer
 * (runFullRollup()) exactly as the build plan asks for phase 4, since this
 * is "the part most likely to need correcting once real people check the
 * numbers" -- every formula below is deliberately simple and documented so
 * it's easy to audit and adjust.
 *
 * ACHIEVEMENT %, PER KPI -- real-data caveat:
 * The real workbook's Milestone Target / Annual Target cells are frequently
 * NARRATIVE TEXT, not numbers (e.g. "Eligible programme list locked; SAR
 * co-ordinators appointed per programme."), even for KPIs whose Unit is
 * numeric (Count, %). A clean Actual-vs-Target ratio is therefore only
 * computable for a subset of rows. kpiAchievement_() falls back through
 * three sources, in order:
 *   1. 'Score (Actual vs Target)' if the KPI Owner has filled in a 0-100
 *      number there (this column exists for exactly this purpose).
 *   2. Q4 Actual / Q4 Milestone Target, when BOTH parse as plain numbers.
 *   3. The most recent quarter's RAG Status, mapped to a score
 *      (Green=100, Amber=70, Red=40, Grey=0) -- a coarse but defensible
 *      stand-in when no numeric comparison is possible.
 * A KPI with none of the above (nothing entered yet) contributes '' (blank)
 * and is excluded from weighted averages above it, not treated as 0.
 *
 * OFFICE roll-up: Σ(KPI achievement % x KPI Wt%) / Σ(KPI Wt%) over that
 * office's Tier 2 KPI rows -- the exact formula in spec section 4.
 *
 * CLUSTER roll-up: Tier 1 KPIs carry a 'Rolls Up From' office list (the
 * literal parent-child link named in spec section 2/4). Each Tier 1 KPI's
 * own achievement is the average of ITS named offices' Office Achievement %
 * (not a blind average across every office in the cluster), then the
 * cluster's achievement is the same Σ(achievement x Wt%) / Σ(Wt%) pattern
 * over its Tier 1 KPI rows.
 *
 * GOAL roll-up: Owner Office(s) (split on "/", e.g. "SEAS / PSB") weighted
 * by the Goals sheet's admin-set 'Owner Weight %'; the remaining weight is
 * split evenly across Contributing Offices (03_Strategic_Pillars) PLUS any
 * office linked as a KPI-level "Co-owner" of that goal (KPI_Office_Links).
 * If Owner Weight % is blank, owner and contributors are averaged evenly --
 * the explicit split spec section 4 asks admins to set is what overrides
 * that default.
 *
 * PILLAR roll-up: simple average of its Goals' achievement (the spec gives
 * no pillar-level weighting scheme; goals aren't weighted against each
 * other here).
 */

const RAG_SCORE = { Green: 100, Amber: 70, Red: 40, Grey: 0 };

function runFullRollup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpiAch = rollupKpis_(ss);           // KPI_ID -> achievement % (number) or null
  const officeAch = rollupOffices_(ss, kpiAch);   // Office Code -> achievement %
  const clusterAch = rollupClusters_(ss, kpiAch, officeAch);
  const goalAch = rollupGoals_(ss, officeAch);
  rollupPillars_(ss, goalAch);

  SpreadsheetApp.getUi().alert('Roll-up complete. Achievement % written to KPIs, Offices, Clusters, Goals and Pillars sheets.');
}

// ---------------------------------------------------------------------------
// KPI level
// ---------------------------------------------------------------------------

function rollupKpis_(ss) {
  const sheet = ss.getSheetByName(SHEET.KPIS);
  const map = headerIndexMap_(sheet);
  const result = {};
  if (sheet.getLastRow() < 2) return result;

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const achievementCol = map['Achievement %'];

  data.forEach(function (row, i) {
    const kpiId = row[map['KPI_ID'] - 1];
    const achievement = kpiAchievement_(row, map);
    result[kpiId] = achievement;
    sheet.getRange(2 + i, achievementCol).setValue(achievement === null ? '' : achievement);
  });
  return result;
}

function kpiAchievement_(row, map) {
  const score = parseFloat(row[map['Score (Actual vs Target)'] - 1]);
  if (!isNaN(score) && score >= 0 && score <= 100) return round1_(score);

  const q4Actual = parseFloat(row[map['Q4 Actual'] - 1]);
  const q4Target = parseFloat(row[map['Q4 Milestone Target'] - 1]);
  if (!isNaN(q4Actual) && !isNaN(q4Target) && q4Target !== 0) {
    return round1_((q4Actual / q4Target) * 100);
  }

  // Most recent quarter with a Status set, latest first.
  for (let q = QUARTERS.length - 1; q >= 0; q--) {
    const status = row[map[QUARTERS[q] + ' Status'] - 1];
    if (status && RAG_SCORE.hasOwnProperty(status)) return RAG_SCORE[status];
  }
  return null;
}

function round1_(n) { return Math.round(n * 10) / 10; }

// ---------------------------------------------------------------------------
// Office level
// ---------------------------------------------------------------------------

function rollupOffices_(ss, kpiAch) {
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  const officeSheet = ss.getSheetByName(SHEET.OFFICES);
  const officeMap = headerIndexMap_(officeSheet);
  const result = {};
  if (kpiSheet.getLastRow() < 2 || officeSheet.getLastRow() < 2) return result;

  const kpiRows = kpiSheet.getRange(2, 1, kpiSheet.getLastRow() - 1, kpiSheet.getLastColumn()).getValues();
  const byOffice = {};
  kpiRows.forEach(function (row) {
    if (parseTierNumber_(row[kpiMap['Tier'] - 1]) !== '2') return;
    const office = row[kpiMap['Office'] - 1];
    const kpiId = row[kpiMap['KPI_ID'] - 1];
    const wt = parseFloat(row[kpiMap['Wt (%)'] - 1]) || 0;
    const achievement = kpiAch[kpiId];
    if (achievement === null || achievement === undefined) return;
    byOffice[office] = byOffice[office] || { weightedSum: 0, totalWt: 0 };
    byOffice[office].weightedSum += achievement * wt;
    byOffice[office].totalWt += wt;
  });

  const officeRows = officeSheet.getRange(2, 1, officeSheet.getLastRow() - 1, officeSheet.getLastColumn()).getValues();
  officeRows.forEach(function (orow, i) {
    const code = orow[officeMap['Office Code'] - 1];
    const agg = byOffice[code];
    const achievement = (agg && agg.totalWt > 0) ? round1_(agg.weightedSum / agg.totalWt) : null;
    result[code] = achievement;
    officeSheet.getRange(2 + i, officeMap['Achievement %']).setValue(achievement === null ? '' : achievement);
  });
  return result;
}

// ---------------------------------------------------------------------------
// Cluster level
// ---------------------------------------------------------------------------

function rollupClusters_(ss, kpiAch, officeAch) {
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const kpiMap = headerIndexMap_(kpiSheet);
  const clusterSheet = ss.getSheetByName(SHEET.CLUSTERS);
  const clusterMap = headerIndexMap_(clusterSheet);
  const result = {};
  if (kpiSheet.getLastRow() < 2 || clusterSheet.getLastRow() < 2) return result;

  const kpiRows = kpiSheet.getRange(2, 1, kpiSheet.getLastRow() - 1, kpiSheet.getLastColumn()).getValues();
  const byCluster = {};
  kpiRows.forEach(function (row) {
    const tier = parseTierNumber_(row[kpiMap['Tier'] - 1]);
    if (tier !== '0' && tier !== '1') return;
    const cluster = row[kpiMap['Cluster'] - 1];
    const wt = parseFloat(row[kpiMap['Wt (%)'] - 1]) || 0;
    const rollsUpFrom = parseOfficeList_(row[kpiMap['Rolls Up From'] - 1]);

    const contributingAchievements = rollsUpFrom
      .map(function (code) { return officeAch[code]; })
      .filter(function (a) { return a !== null && a !== undefined; });
    if (contributingAchievements.length === 0) return;
    const kpiRollAchievement = average_(contributingAchievements);

    byCluster[cluster] = byCluster[cluster] || { weightedSum: 0, totalWt: 0 };
    byCluster[cluster].weightedSum += kpiRollAchievement * wt;
    byCluster[cluster].totalWt += wt;
  });

  const clusterRows = clusterSheet.getRange(2, 1, clusterSheet.getLastRow() - 1, clusterSheet.getLastColumn()).getValues();
  clusterRows.forEach(function (crow, i) {
    const name = crow[clusterMap['Cluster Name'] - 1];
    const agg = byCluster[name];
    const achievement = (agg && agg.totalWt > 0) ? round1_(agg.weightedSum / agg.totalWt) : null;
    result[name] = achievement;
    clusterSheet.getRange(2 + i, clusterMap['Achievement %']).setValue(achievement === null ? '' : achievement);
  });
  return result;
}

// ---------------------------------------------------------------------------
// Goal level
// ---------------------------------------------------------------------------

function rollupGoals_(ss, officeAch) {
  const goalSheet = ss.getSheetByName(SHEET.GOALS);
  const goalMap = headerIndexMap_(goalSheet);
  const result = {};
  if (goalSheet.getLastRow() < 2) return result;

  const coOwnersByGoal = coOwnersByGoal_(ss);
  const goalRows = goalSheet.getRange(2, 1, goalSheet.getLastRow() - 1, goalSheet.getLastColumn()).getValues();

  goalRows.forEach(function (row, i) {
    const goalId = row[goalMap['Goal ID'] - 1];
    const ownerOffices = parseOfficeList_(row[goalMap['Owner Office'] - 1]);
    const contributingOffices = parseOfficeList_(row[goalMap['Contributing Offices'] - 1])
      .concat(coOwnersByGoal[goalId] || []);
    const ownerWeightPct = parseFloat(row[goalMap['Owner Weight %'] - 1]);

    const ownerAchievements = ownerOffices.map(function (c) { return officeAch[c]; }).filter(isNumber_);
    const contribAchievements = contributingOffices.map(function (c) { return officeAch[c]; }).filter(isNumber_);

    let achievement = null;
    if (ownerAchievements.length > 0 && contribAchievements.length > 0 && !isNaN(ownerWeightPct)) {
      achievement = round1_(
        average_(ownerAchievements) * (ownerWeightPct / 100) +
        average_(contribAchievements) * (1 - ownerWeightPct / 100)
      );
    } else {
      const all = ownerAchievements.concat(contribAchievements);
      if (all.length > 0) achievement = round1_(average_(all));
    }

    result[goalId] = achievement;
    goalSheet.getRange(2 + i, goalMap['Achievement %']).setValue(achievement === null ? '' : achievement);
  });
  return result;
}

/** Goal ID -> [office codes] for every office whose KPI_Office_Links row has Role = 'Co-owner'. */
function coOwnersByGoal_(ss) {
  const linkSheet = ss.getSheetByName(SHEET.KPI_OFFICE_LINKS);
  const map = {};
  existingRows_(linkSheet).forEach(function (r) {
    if (String(r['Role'] || '').trim().toLowerCase() !== 'co-owner') return;
    const goalId = r['Goal ID'];
    if (!goalId) return;
    map[goalId] = map[goalId] || [];
    if (map[goalId].indexOf(r['Office Code']) === -1) map[goalId].push(r['Office Code']);
  });
  return map;
}

// ---------------------------------------------------------------------------
// Pillar level
// ---------------------------------------------------------------------------

function rollupPillars_(ss, goalAch) {
  const pillarSheet = ss.getSheetByName(SHEET.PILLARS);
  const pillarMap = headerIndexMap_(pillarSheet);
  const goalSheet = ss.getSheetByName(SHEET.GOALS);
  const goalMap = headerIndexMap_(goalSheet);
  if (pillarSheet.getLastRow() < 2) return;

  const goals = existingRows_(goalSheet);
  const byPillar = {};
  goals.forEach(function (g) {
    const a = goalAch[g['Goal ID']];
    if (a === null || a === undefined) return;
    byPillar[g['Parent Pillar']] = byPillar[g['Parent Pillar']] || [];
    byPillar[g['Parent Pillar']].push(a);
  });

  const pillarRows = pillarSheet.getRange(2, 1, pillarSheet.getLastRow() - 1, pillarSheet.getLastColumn()).getValues();
  pillarRows.forEach(function (prow, i) {
    const pillarId = prow[pillarMap['Pillar ID'] - 1];
    const values = byPillar[pillarId];
    const achievement = (values && values.length > 0) ? round1_(average_(values)) : null;
    if (pillarMap['Achievement %']) pillarSheet.getRange(2 + i, pillarMap['Achievement %']).setValue(achievement === null ? '' : achievement);
  });
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

function average_(arr) { return arr.reduce(function (s, v) { return s + v; }, 0) / arr.length; }
function isNumber_(v) { return typeof v === 'number' && !isNaN(v); }

/** Splits a free-text office list on "," or "/" (e.g. "SEAS / PSB", "AA, Admissions"). */
function parseOfficeList_(text) {
  return String(text || '')
    .split(/[,/]/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s.length > 0 && s !== '—'; });
}
