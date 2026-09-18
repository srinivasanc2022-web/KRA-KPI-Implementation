/**
 * 02_SampleData.gs
 *
 * PLACEHOLDER / SAMPLE DATA ONLY -- kept as a fast smoke test for the import
 * pipeline. The real SRMAP_KRA-KPI_AY26-27_Updated workbook (39 sheets) has
 * since been imported successfully (see NEXT_STEPS.md); use "3. Import Real
 * Workbook..." for real data. This fixture intentionally reproduces the real
 * workbook's header quirks -- an embedded newline in "S.\nNo", a blank title
 * row above the real header row, a compound "Goal ID + target" cell, and a
 * "Co-owner" Goal Role row -- so findHeaderRow_()/extractGoalId_() in
 * 03_ImportPipeline.gs stay exercised even without the real file.
 *
 * Every sample sheet is prefixed "SAMPLE_SRC_" so it can never be confused
 * with a real imported source tab. Delete these sheets (menu: "2c. Remove
 * Sample Source Sheets") once you're done smoke-testing.
 */

const SAMPLE_PREFIX = 'SAMPLE_SRC_';

function buildSampleSourceWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  buildSampleIndex_(ss);
  buildSamplePillars_(ss);
  buildSampleCascadeMap_(ss);
  buildSampleTier1Scorecard_(ss);
  buildSampleOfficeSheet_(ss, 'AA', 'Cluster 1 - Academic', 'Tier 2 - Office');
  buildSampleOfficeSheet_(ss, 'SEAS', 'Cluster 1 - Academic', 'Tier 2 - Office');
  buildSampleOfficeSheet_(ss, 'Research', 'Cluster 2 - Growth', 'Tier 2 - Office');

  SpreadsheetApp.getUi().alert(
    'Sample source sheets created (prefixed "' + SAMPLE_PREFIX + '").\n\n' +
    'These are PLACEHOLDER data, not the real AY26-27 workbook. Use menu item ' +
    '"3. Import Real Workbook..." once you have the actual file.'
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
  sheet.getRange(1, 1).setValue('SRM University AP - KRA/KPI framework AY26-27 (SAMPLE)');
  const headers = ['#', 'Sheet', 'Full name', 'Tier', 'Cluster', 'KPIs', 'Strategic', 'Operational', 'Verticals', 'Total Wt'];
  const rows = [
    [1, 'C1 Academic Scorecard', 'Cluster 1 Academic Scorecard', 'Tier 1', 'Cluster 1 - Academic', 2, 2, 0, 2, 100],
    [2, 'AA', 'Academic Affairs', 'Tier 2', 'Cluster 1 - Academic', 2, 1, 1, 1, 100],
    [3, 'SEAS', 'School of Engineering & Applied Sciences', 'Tier 2', 'Cluster 1 - Academic', 2, 1, 1, 1, 100],
    [4, 'Research', 'Office of Dean - Research', 'Tier 2', 'Cluster 2 - Growth', 2, 2, 0, 1, 100]
  ];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, rows.length, headers.length).setValues(rows);
}

function buildSamplePillars_(ss) {
  const name = SAMPLE_PREFIX + SOURCE_SHEET.PILLARS;
  const sheet = getOrCreateBlank_(ss, name);
  sheet.getRange(1, 1).setValue('Strategic pillars and five-year goals (SAMPLE)');
  const headers = ['Pillar', 'Pillar name', 'Goal ID', 'Five-year goal / target', 'Owner office', 'Contributing offices'];
  // Pillar / Pillar name deliberately blank on the 2nd+ row of each group,
  // exactly like the real workbook, to exercise the forward-fill logic.
  const rows = [
    ['P1', 'Institutional Excellence & Quality', 'P1.4', 'International Accreditation: AACSB and ABET', 'SEAS / PSB', 'AA, QAR'],
    [null, null, 'P1.5', 'Governance: 100% digitalisation and zero adverse findings', 'ITKM', 'AA'],
    ['P3', 'Research, Innovation & Knowledge Creation', 'P3.1', 'Sponsored Research: Rs.30 Cr+ per year', 'Research', 'AA']
  ];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, rows.length, headers.length).setValues(rows);
}

function buildSampleCascadeMap_(ss) {
  const name = SAMPLE_PREFIX + SOURCE_SHEET.CASCADE_MAP;
  const sheet = getOrCreateBlank_(ss, name);
  sheet.getRange(1, 1).setValue('Office responsibility map (SAMPLE)');
  const headers = ['Cluster', 'Office', 'Goals owned', 'Goals contributed to', 'KPIs', 'Strategic', 'Operational'];
  const rows = [
    ['Cluster 1 - Academic', 'AA', '—', 'P1.4', 2, 1, 1],
    ['Cluster 1 - Academic', 'SEAS', 'P1.4', '—', 2, 1, 1],
    ['Cluster 2 - Growth', 'Research', 'P3.1', '—', 2, 2, 0]
  ];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, rows.length, headers.length).setValues(rows);
}

/**
 * Reproduces the known real-workbook quirk deliberately: 2-3 blank/title
 * rows sit above the real header row (varies by sheet -- C2 Growth
 * Scorecard uses row 4, others use row 3 in the real workbook).
 * findHeaderRow_() must locate it generically rather than assuming a fixed
 * row number.
 */
function buildSampleTier1Scorecard_(ss) {
  const name = SAMPLE_PREFIX + 'C1 Academic Scorecard';
  const sheet = getOrCreateBlank_(ss, name);
  sheet.getRange(1, 1).setValue('C1 Academic - KRA/KPI scorecard AY26-27 (SAMPLE)');
  // row 2 intentionally blank

  const headers = ['S.No', 'Tier', 'Cluster', 'KRA Type', 'Pillar', '5-Yr Goal (ID + target)',
    'Goal Role', 'Vertical', 'KRA', 'KPI', 'How it is measured', 'Unit', 'Frequency',
    'Annual Target AY 26-27', 'Wt (%)', 'Rolls up from', 'Q1', 'Q2', 'Q3', 'Q4', 'Owner sign-off'];
  sheet.getRange(3, 1, 1, headers.length).setValues([headers]);

  const rows = [
    [1, 'Tier 1 - Cluster', 'Cluster 1 - Academic', 'Strategic', 'P1', 'P1.4 — International Accreditation: AACSB and ABET',
      'Cluster aggregate', 'Accreditation & Ranking Readiness', 'Accreditation Readiness',
      'NBA/NAAC-Eligible Programmes Made SAR-Ready', 'Milestones closed / planned', 'Count', 'Quarterly',
      '100% of eligible programmes', 50, 'AA, SEAS', '', '', '', '', ''],
    [2, 'Tier 1 - Cluster', 'Cluster 2 - Growth', 'Strategic', 'P3', 'P3.1 — Sponsored Research: Rs.30 Cr+ per year',
      'Cluster aggregate', 'Research funding', 'Grant capture', 'Sponsored research sanctioned',
      'Grants sanctioned across all agencies', '₹ Cr', 'Quarterly', 'Year-1 step toward ₹30 Cr', 50, 'Research', '', '', '', '', '']
  ];
  sheet.getRange(4, 1, rows.length, headers.length).setValues(rows);
}

function buildSampleOfficeSheet_(ss, officeCode, cluster, tierText) {
  const name = SAMPLE_PREFIX + officeCode;
  const sheet = getOrCreateBlank_(ss, name);
  sheet.getRange(1, 1).setValue(officeCode + ' (SAMPLE)');

  // Deliberately uses the embedded-newline variant on the S.No header, like
  // several real office tabs, to exercise header normalization.
  const headers = ['S.\nNo', 'Office', 'Tier', 'Cluster', 'KRA Type', 'Strategic Pillar',
    '5-Yr Goal (ID + target)', 'Goal Role', 'Vertical', 'KRA', 'KPI',
    'Strategic Rationale \n( 5-Yr Goals)', 'Definition of KPI', 'How to Measure \n(KPI-Specific Formula)',
    'Illustration', 'Unit', 'Frequency', 'Measurement Nature', 'Metric Type', 'Indicator Nature',
    'Wt (%)', 'Annual Target \n(AY 26-27)',
    'Q1 Milestone Target', 'Q2 Milestone Target', 'Q3 Milestone Target', 'Q4 Milestone Target',
    'Q1 Actual', 'Q2 Actual', 'Q3 Actual', 'Q4 Actual',
    'Q1 Status', 'Q2 Status', 'Q3 Status', 'Q4 Status',
    'Q1 ATR', 'Q2 ATR', 'Q3 ATR', 'Q4 ATR',
    'Score (Actual vs Target)', 'Data Source', 'KPI Owner',
    'R (Responsible)', 'A (Accountable)', 'C \n(Consulted)', 'I \n(Informed)',
    'Responsible Officer', 'Escalation / Remarks', 'Overlap / Remarks'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]);

  const isResearch = officeCode === 'Research';
  const rows = [
    buildSampleKpiRow_({
      sno: 1, office: officeCode, tier: tierText, cluster: cluster, kraType: 'Strategic',
      pillar: isResearch ? 'P3' : 'P1', goalText: isResearch ? 'P3.1 — Sponsored Research: Rs.30 Cr+ per year' : 'P1.4 — International Accreditation: AACSB and ABET',
      goalRole: isResearch ? 'Co-owner' : (officeCode === 'AA' ? 'Contributor' : 'Owner'),
      vertical: isResearch ? 'Research funding' : 'Accreditation & Ranking Readiness',
      kra: isResearch ? 'Grant capture' : 'Accreditation Readiness',
      kpi: isResearch ? 'Sponsored Research Sanctioned' : 'NBA-Eligible Programmes Made SAR-Ready',
      unit: isResearch ? '₹ Cr' : 'Count', frequency: 'Quarterly', measurementNature: 'Quantitative',
      metricType: 'Strategic', indicatorNature: 'Leading', wt: 60, annualTarget: isResearch ? 30 : 12,
      q1t: isResearch ? 5 : 3, q2t: isResearch ? 12 : 6, q3t: isResearch ? 20 : 9, q4t: isResearch ? 30 : 12,
      dataSource: 'Register', owner: officeCode + ' Dean', r: officeCode, a: 'Dean - ' + officeCode,
      c: 'QAR', i: 'Vice-Chancellor', responsibleOfficer: 'Associate Dean'
    }),
    buildSampleKpiRow_({
      sno: 2, office: officeCode, tier: tierText, cluster: cluster, kraType: 'Operational',
      pillar: 'P1', goalText: 'P1.5 — Governance: 100% digitalisation and zero adverse findings',
      goalRole: 'Contributor', vertical: 'Quality Assurance', kra: 'Internal Audit Compliance',
      kpi: 'Internal QA Audit Findings Closed On Time', unit: '%', frequency: 'Monthly',
      measurementNature: 'Quantitative', metricType: 'Compliance', indicatorNature: 'Lagging',
      wt: 40, annualTarget: 95, q1t: 85, q2t: 90, q3t: 92, q4t: 95,
      dataSource: 'QA audit tracker', owner: officeCode + ' QA Lead', r: officeCode,
      a: 'Dean - ' + officeCode, c: 'QAR', i: 'Pro Vice-Chancellor', responsibleOfficer: 'QA Lead'
    })
  ];
  sheet.getRange(3, 1, rows.length, rows[0].length).setValues(rows);
}

function buildSampleKpiRow_(f) {
  // Order MUST match the headers array in buildSampleOfficeSheet_.
  return [
    f.sno, f.office, f.tier, f.cluster, f.kraType, f.pillar, f.goalText, f.goalRole,
    f.vertical, f.kra, f.kpi,
    'Advances ' + f.pillar, 'Definition placeholder', 'Formula placeholder', 'Illustration placeholder',
    f.unit, f.frequency, f.measurementNature, f.metricType, f.indicatorNature, f.wt, f.annualTarget,
    f.q1t, f.q2t, f.q3t, f.q4t,
    '', '', '', '', // Q1-4 Actual -- blank, entered during the year (see spec section 6)
    '', '', '', '', // Q1-4 Status -- blank, computed or entered during the year
    '', '', '', '', // Q1-4 ATR -- blank, entered when Amber/Red
    '', // Score
    f.dataSource, f.owner, f.r, f.a, f.c, f.i, f.responsibleOfficer,
    '', '' // Escalation/Remarks, Overlap/Remarks
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
