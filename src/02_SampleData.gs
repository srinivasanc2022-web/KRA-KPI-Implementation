/**
 * 02_SampleData.gs
 *
 * PLACEHOLDER / SAMPLE DATA ONLY.
 *
 * The real SRMAP_KRA-KPI_AY26-27 workbook was not available when this system
 * was built (see README.md / DEPLOY.md). This module fabricates a small
 * workbook-shaped dataset -- in the *same raw layout* the real workbook uses
 * (01_Index, 03_Strategic_Pillars, 04_Cascade_Map, a Tier 0/1 scorecard, and
 * two Tier 2 office scorecards) -- so the import pipeline (03_ImportPipeline.gs)
 * can be developed, demoed and unit-tested end to end.
 *
 * Every sample sheet is prefixed "SAMPLE_SRC_" so it can never be confused
 * with a real imported source tab. Delete these sheets (menu: "0. Remove
 * Sample Source Sheets") once the real workbook is imported.
 *
 * IMPORTANT: When the real workbook is available, do NOT rely on this file's
 * office list, pillar text, or KPI content -- re-run "2. Import Real
 * Workbook" against the actual file and treat this sample only as a smoke
 * test fixture.
 */

const SAMPLE_PREFIX = 'SAMPLE_SRC_';

function buildSampleSourceWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  buildSampleIndex_(ss);
  buildSamplePillars_(ss);
  buildSampleCascadeMap_(ss);
  buildSampleTier1Scorecard_(ss);
  buildSampleOfficeSheet_(ss, 'AA', 'Academic Affairs', 'Cluster 1 - Academic', 2);
  buildSampleOfficeSheet_(ss, 'SEAS', 'School of Engineering & Applied Sciences', 'Cluster 1 - Academic', 2);

  SpreadsheetApp.getUi().alert(
    'Sample source sheets created (prefixed "' + SAMPLE_PREFIX + '").\n\n' +
    'These are PLACEHOLDER data, not the real AY26-27 workbook. Use menu item ' +
    '"2. Import Real Workbook" once you have the actual file, pointing it at ' +
    'either this same spreadsheet or an external Sheet ID.'
  );
}

function removeSampleSourceSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(function (sheet) {
    if (sheet.getName().indexOf(SAMPLE_PREFIX) === 0) {
      ss.deleteSheet(sheet);
    }
  });
  SpreadsheetApp.getUi().alert('Sample source sheets removed.');
}

function buildSampleIndex_(ss) {
  const name = SAMPLE_PREFIX + SOURCE_SHEET.INDEX;
  const sheet = getOrCreateBlank_(ss, name);
  const headers = ['Sheet Name', 'Full Office Name', 'Tier', 'Cluster', 'KPI Count',
    'Strategic KPIs', 'Operational KPIs', 'Vertical Count', 'Total Wt'];
  const rows = [
    ['VC Scorecard', 'Vice-Chancellor Scorecard', 0, 'VC Direct', 3, 3, 0, 2, 100],
    ['C1 Academic Scorecard', 'Cluster 1 Academic Scorecard', 1, 'Cluster 1 - Academic', 2, 2, 0, 2, 100],
    ['AA', 'Academic Affairs', 2, 'Cluster 1 - Academic', 2, 1, 1, 1, 100],
    ['SEAS', 'School of Engineering & Applied Sciences', 2, 'Cluster 1 - Academic', 2, 1, 1, 1, 100]
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function buildSamplePillars_(ss) {
  const name = SAMPLE_PREFIX + SOURCE_SHEET.PILLARS;
  const sheet = getOrCreateBlank_(ss, name);
  const headers = ['Pillar ID', 'Pillar Name', 'Goal ID', 'Five-Year Goal / Target', 'Owner Office', 'Contributing Offices'];
  const rows = [
    ['P1', 'Institutional Excellence & Quality', 'P1.1', 'NBA/NAAC accreditation for all eligible programmes by 2030', 'AA', 'SEAS, QAR'],
    ['P1', 'Institutional Excellence & Quality', 'P1.4', 'AACSB/ABET accreditation readiness by 2029', 'SEAS', 'AA, QAR'],
    ['P3', 'Research & Innovation Leadership', 'P3.1', 'Top-decile research output and IP filings by 2031', 'Research', 'AiTI, QuTI']
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function buildSampleCascadeMap_(ss) {
  const name = SAMPLE_PREFIX + SOURCE_SHEET.CASCADE_MAP;
  const sheet = getOrCreateBlank_(ss, name);
  const headers = ['Office', 'Goals Owned', 'Goals Contributed To', 'Strategic KPI Count', 'Operational KPI Count'];
  const rows = [
    ['AA', 'P1.1', 'P1.4', 1, 1],
    ['SEAS', 'P1.4', 'P1.1', 1, 1],
    ['Research', 'P3.1', '', 1, 0],
    ['QAR', '', 'P1.1, P1.4', 0, 1]
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

/**
 * Reproduces the known real-workbook quirk deliberately: a blank row sits
 * above the real header row on some Tier 0/1 sheets (e.g. C1 Academic
 * Scorecard). findHeaderRow_() in 03_ImportPipeline.gs must skip it.
 */
function buildSampleTier1Scorecard_(ss) {
  const name = SAMPLE_PREFIX + 'C1 Academic Scorecard';
  const sheet = getOrCreateBlank_(ss, name);
  sheet.getRange(1, 1).setValue('Cluster 1 - Academic Scorecard (AY26-27)'); // decorative blank-ish row

  const headers = ['S.No', 'Office', 'Tier', 'Cluster', 'KRA Type', 'Strategic Pillar', 'Goal ID',
    'Goal Role', 'Vertical', 'KRA', 'KPI', 'Unit', 'Frequency', 'Annual Target', 'Wt (%)', 'Rolls up from'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);

  const rows = [
    [1, 'Cluster 1 - Academic', 1, 'Cluster 1 - Academic', 'Strategic', 'P1', 'P1.1',
      'Cluster aggregate', 'Accreditation & Ranking Readiness', 'Accreditation Readiness',
      'NBA/NAAC-Eligible Programmes Made SAR-Ready', 'Count', 'Quarterly', 12, 50, 'AA, SEAS'],
    [2, 'Cluster 1 - Academic', 1, 'Cluster 1 - Academic', 'Strategic', 'P1', 'P1.4',
      'Cluster aggregate', 'Accreditation & Ranking Readiness', 'International Accreditation',
      'AACSB/ABET SAR Milestones Completed', '%', 'Quarterly', 100, 50, 'SEAS, AA']
  ];
  sheet.getRange(3, 1, rows.length, headers.length).setValues(rows);
}

function buildSampleOfficeSheet_(ss, officeCode, officeFullName, cluster, tier) {
  const name = SAMPLE_PREFIX + officeCode;
  const sheet = getOrCreateBlank_(ss, name);
  sheet.getRange(1, 1, 1, COLS.KPIS.length - 2).setValues([COLS.KPIS.slice(1, -1)]); // omit system KPI_ID and Co-Owner cols

  const isAA = officeCode === 'AA';
  const rows = [
    buildSampleKpiRow_({
      sno: 1, office: officeFullName, tier: tier, cluster: cluster, kraType: 'Strategic',
      pillar: 'P1', goalId: isAA ? 'P1.1' : 'P1.4', goalRole: isAA ? 'Owner' : 'Contributor',
      vertical: 'Accreditation & Ranking Readiness', kra: 'Accreditation Readiness',
      kpi: isAA ? 'NBA-Eligible Programmes Made SAR-Ready' : 'ABET SAR Sections Drafted',
      rationale: 'Directly advances ' + (isAA ? 'P1.1 NBA/NAAC accreditation target' : 'P1.4 ABET readiness target'),
      definition: 'Count of programmes with a Self-Assessment Report submitted to the accrediting body',
      howToMeasure: 'Count of SAR-ready programmes / Count of NBA-eligible programmes',
      illustration: 'e.g. 9 of 12 eligible programmes SAR-ready = 75%',
      unit: 'Count', frequency: 'Quarterly', measurementNature: 'Quantitative', metricType: 'Strategic',
      indicatorNature: 'Leading', wt: 60, annualTarget: isAA ? 12 : 100,
      q1t: isAA ? 3 : 25, q2t: isAA ? 6 : 50, q3t: isAA ? 9 : 75, q4t: isAA ? 12 : 100,
      dataSource: 'Accreditation Cell register', owner: officeCode + ' Dean',
      r: officeCode, a: 'Dean - ' + officeFullName, c: 'QAR', i: 'Vice-Chancellor',
      responsibleOfficer: 'Associate Dean - Accreditation'
    }),
    buildSampleKpiRow_({
      sno: 2, office: officeFullName, tier: tier, cluster: cluster, kraType: 'Operational',
      pillar: 'P1', goalId: isAA ? 'P1.4' : 'P1.1', goalRole: 'Contributor',
      vertical: 'Quality Assurance', kra: 'Internal Audit Compliance',
      kpi: 'Internal QA Audit Findings Closed On Time',
      rationale: 'Operational hygiene supporting accreditation readiness',
      definition: 'Share of internal audit findings closed within the committed timeline',
      howToMeasure: 'Findings closed on time / Total findings raised',
      illustration: 'e.g. 18 of 20 findings closed on time = 90%',
      unit: '%', frequency: 'Monthly', measurementNature: 'Quantitative', metricType: 'Compliance',
      indicatorNature: 'Lagging', wt: 40, annualTarget: 95,
      q1t: 85, q2t: 90, q3t: 92, q4t: 95,
      dataSource: 'QA audit tracker', owner: officeCode + ' QA Lead',
      r: officeCode, a: 'Dean - ' + officeFullName, c: 'QAR', i: 'Pro Vice-Chancellor',
      responsibleOfficer: 'QA Lead'
    })
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function buildSampleKpiRow_(f) {
  // Order MUST match COLS.KPIS.slice(1, -1) (everything between KPI_ID and Co-Owner Offices).
  return [
    f.sno, f.office, f.tier, f.cluster, f.kraType, f.pillar, f.goalId, f.goalRole,
    f.vertical, f.kra, f.kpi, f.rationale, f.definition, f.howToMeasure, f.illustration,
    f.unit, f.frequency, f.measurementNature, f.metricType, f.indicatorNature, f.wt, f.annualTarget,
    f.q1t, f.q2t, f.q3t, f.q4t,
    '', '', '', '', // Q1-4 Actual -- blank, entered during the year (see spec section 6)
    '', '', '', '', // Q1-4 Status -- blank, computed or entered during the year
    '', '', '', '', // Q1-4 ATR -- blank, entered when Amber/Red
    '', // Score
    f.dataSource, f.owner, f.r, f.a, f.c, f.i, f.responsibleOfficer,
    '', '', // Escalation/Remarks, Overlap/Remarks
    '' // Rolls Up From (Tier 2 sheets don't populate this; Tier 0/1 sheets do)
  ];
}

function getOrCreateBlank_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}
