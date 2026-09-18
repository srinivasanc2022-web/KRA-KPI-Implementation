/**
 * 00_Constants.gs
 *
 * Central schema definition for the SRM AP Strategic Goal -> Pillar -> Cluster ->
 * Office -> KRA/KPI Execution & Performance Management System.
 *
 * This file is the single source of truth for sheet names and column headers.
 * Every other module reads/writes sheets by HEADER NAME via the helpers in
 * 01_SchemaSetup.gs / 03_ImportPipeline.gs, never by fixed column index, so the
 * layout here can evolve without breaking downstream code.
 *
 * Column names and aliases below were reconciled against the real
 * SRMAP_KRA-KPI_AY26-27_Updated workbook (39 sheets). Two real-world quirks
 * that drove the design:
 *   1. Header text is inconsistent between sheets -- e.g. the S.No column is
 *      literally "S.\nNo" (an embedded newline) on some office tabs and
 *      "S.No" on others; several other headers wrap onto two lines too. The
 *      importer normalizes away ALL whitespace/punctuation before comparing
 *      headers (see normalizeHeaderText_ in 03_ImportPipeline.gs), so this
 *      needs no special-casing here.
 *   2. Some headers are genuinely WORDED differently between Tier 0/1
 *      scorecards and Tier 2 office scorecards (e.g. "Pillar" vs "Strategic
 *      Pillar", "How it is measured" vs "How to Measure (KPI-Specific
 *      Formula)"). Those need an explicit alias -> canonical mapping, in
 *      HEADER_ALIASES below.
 */

// ---- Canonical "system of record" sheets (the normalized data model) ----
const SHEET = {
  CONFIG: 'Config',
  PILLARS: 'Pillars',
  GOALS: 'Goals',
  CLUSTERS: 'Clusters',
  OFFICES: 'Offices',
  VERTICALS: 'Verticals',
  KRAS: 'KRAs',
  KPIS: 'KPIs',
  KPI_OFFICE_LINKS: 'KPI_Office_Links',
  ACTIONS: 'Actions',
  AUDIT_LOG: 'Audit_Log',
  VALIDATION_LOG: 'Validation_Log',
  IMPORT_LOG: 'Import_Log',
  ESCALATIONS: 'Escalations',
  NOTIFICATION_LOG: 'Notification_Log',
  OFFICE_CONTACTS: 'Office_Contacts'
};

// ---- Raw source workbook tab names this system knows how to read ----
// (matches SRMAP_KRA-KPI_AY26-27 workbook structure)
const SOURCE_SHEET = {
  INDEX: '01_Index',
  PILLARS: '03_Strategic_Pillars',
  CASCADE_MAP: '04_Cascade_Map'
};

const TIER0_1_SOURCE_SHEETS = [
  'VC Scorecard',
  'C1 Academic Scorecard',
  'C2 Growth Scorecard',
  'C3 Student Scorecard',
  'C4 Governance Scorecard'
];

// Short codes used as the KPI_ID prefix for each Tier 0/1 scorecard (these
// sheets don't have a per-row "Office" code the way Tier 2 sheets do).
const TIER0_1_CODES = {
  'VC Scorecard': 'VC',
  'C1 Academic Scorecard': 'C1',
  'C2 Growth Scorecard': 'C2',
  'C3 Student Scorecard': 'C3',
  'C4 Governance Scorecard': 'C4'
};

// ---- Column headers, exactly as the build spec (section 3) defines them ----

const COLS = {
  PILLARS: ['Pillar ID', 'Pillar Name', 'Achievement %'],

  GOALS: [
    'Goal ID', 'Parent Pillar', 'Five-Year Goal / Target', 'Owner Office',
    'Contributing Offices', 'Owner Weight %', 'Overall Progress', 'Achievement %'
  ],

  CLUSTERS: [
    'Cluster ID', 'Cluster Name', 'Member Offices', 'Aggregate Weight (Check)', 'Achievement %'
  ],

  OFFICES: [
    'Office Code', 'Full Name', 'Cluster', 'Tier', 'Goals Owned',
    'Goals Contributed To', 'KPI Count (Strategic)', 'KPI Count (Operational)',
    'Vertical Count', 'Total Weight (Check)', 'Achievement %'
  ],

  VERTICALS: ['Vertical ID', 'Vertical Name', 'Office Code', 'Linked KRAs'],

  KRAS: ['KRA ID', 'KRA Name', 'Parent Vertical ID', 'Office Code', 'Linked KPIs'],

  // Reproduces every column from the office-level sheets, section 3.7, plus
  // system fields needed for roll-up and cross-tier linkage: KPI_ID (join
  // key), Rolls Up From (Tier 0/1 only -- literal parent-child link),
  // Owner Sign-off (present in the real Tier 0/1 sheets), Achievement %
  // (written by the roll-up engine, 08_RollupEngine.gs).
  //
  // NOTE on co-ownership: the real workbook has NO per-KPI "co-owner %"
  // column. Co-ownership (e.g. AiTI/QuTI/Research all owning P3.1) is
  // expressed by each office having its OWN KPI rows with Goal Role =
  // "Co-owner" for that goal. The importer turns that into KPI_Office_Links
  // rows automatically (section 4/12) -- see importScorecardSheet_().
  KPIS: [
    'KPI_ID', 'S.No', 'Office', 'Tier', 'Cluster', 'KRA Type', 'Strategic Pillar',
    'Goal ID', 'Goal Role', 'Vertical', 'KRA', 'KPI',
    'Strategic Rationale (5-Yr Goals)', 'Definition of KPI',
    'How to Measure (KPI-Specific Formula)', 'Illustration', 'Unit', 'Frequency',
    'Measurement Nature', 'Metric Type', 'Indicator Nature', 'Wt (%)',
    'Annual Target (AY 26-27)',
    'Q1 Milestone Target', 'Q2 Milestone Target', 'Q3 Milestone Target', 'Q4 Milestone Target',
    'Q1 Actual', 'Q2 Actual', 'Q3 Actual', 'Q4 Actual',
    'Q1 Status', 'Q2 Status', 'Q3 Status', 'Q4 Status',
    'Q1 ATR', 'Q2 ATR', 'Q3 ATR', 'Q4 ATR',
    'Score (Actual vs Target)', 'Achievement %', 'Data Source', 'KPI Owner',
    'R (Responsible)', 'A (Accountable)', 'C (Consulted)', 'I (Informed)',
    'Responsible Officer', 'Escalation / Remarks', 'Overlap / Remarks',
    'Rolls Up From', 'Owner Sign-off'
  ],

  KPI_OFFICE_LINKS: ['KPI_ID', 'Goal ID', 'Office Code', 'Role', 'Co-Ownership %'],

  ACTIONS: [
    'Action ID', 'Description', 'Linked KPI_ID', 'Owner', 'Contributors',
    'Priority', 'Start Date', 'Due Date', 'Dependencies', 'Status',
    '% Complete', 'Closure Criteria', 'Remarks'
  ],

  AUDIT_LOG: ['Timestamp', 'User', 'Sheet', 'Row', 'Field', 'Old Value', 'New Value'],

  VALIDATION_LOG: ['Timestamp', 'Scope', 'Entity ID', 'Total Weight', 'Status', 'Details'],

  IMPORT_LOG: ['Timestamp', 'Source Sheet', 'Tier', 'Cluster', 'Rows Imported', 'Status', 'Details'],

  CONFIG: ['Key', 'Value', 'Notes'],

  ESCALATIONS: ['Timestamp', 'KPI_ID', 'Office', 'Reason', 'Detail', 'Status'],

  NOTIFICATION_LOG: ['Timestamp', 'KPI_ID', 'Recipient Role', 'Recipient', 'Reason', 'Sent'],

  OFFICE_CONTACTS: ['Office Code', 'KPI Owner Email', 'Responsible Officer Email', 'Escalation Email']
};

// Quarter fields, in the order they appear on a KPI row.
const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];

// AY26-27 quarter close dates (edit in Config sheet if the academic calendar
// changes; these are the fallback defaults). Used by escalation logic
// (10_Escalation.gs) to detect ATRs open past their quarter's close.
const DEFAULT_QUARTER_CLOSE_DATES = {
  Q1: '2026-09-30',
  Q2: '2026-12-31',
  Q3: '2027-03-31',
  Q4: '2027-06-30'
};

const RAG = {
  GREEN: 'Green',
  AMBER: 'Amber',
  RED: 'Red',
  GREY: 'Grey'
};

// Default RAG thresholds (% of milestone target achieved). Overridable via
// the 'Config' sheet.
const DEFAULT_THRESHOLDS = {
  GREEN_MIN: 0.95, // >= 95% of milestone target
  AMBER_MIN: 0.80  // 80-94% => Amber, < 80% => Red
};

// Units for which automatic Actual-vs-Target % computation is not meaningful;
// these require manual Status entry by the KPI Owner.
const NON_COMPUTABLE_UNITS = ['Band', 'Grade'];

const KRA_TYPES = ['Strategic', 'Operational'];
const METRIC_TYPES = ['Strategic', 'Operational', 'Compliance', 'Governance', 'Outcome', 'Output', 'Process'];
const INDICATOR_NATURES = ['Leading', 'Lagging'];
const MEASUREMENT_NATURES = ['Quantitative', 'Qualitative'];
// The real workbook uses many more cadences than a clean quarterly cycle
// (e.g. "Per cohort", "Per exam cycle", "Per intake") -- these are all real,
// intentional values, not typos, so the dropdown allows free text too (see
// applyKpiSheetValidation_ in 01_SchemaSetup.gs) rather than blocking them.
const FREQUENCIES = [
  'Annual', 'Half-Year', 'Semester', 'Quarterly', 'Monthly', 'Weekly',
  'Per cohort', 'Per cycle', 'Per event', 'Per exam cycle', 'Per intake',
  'Per meeting', 'Per project', 'Per submission', 'Per-Program'
];
// 'Co-owner' appears on Tier 2 rows where an office jointly owns a goal with
// another office (e.g. AiTI, QuTI and Research all "Co-owner" on P3.1).
// 'Enabler' appears on Tier 2 rows for the standard "University Impact" /
// operational-hygiene KRA every office carries in support of a goal it
// neither owns nor is a named contributor to.
const GOAL_ROLES = ['Owner', 'Contributor', 'Institutional', 'Cluster aggregate', 'Co-owner', 'Enabler'];

const DEFAULT_CONFIG = [
  ['Current_AY', 'AY26-27', 'Active academic year for milestone/actual entry'],
  ['Green_Threshold', '0.95', 'Actual/Milestone ratio at/above which Status = Green'],
  ['Amber_Threshold', '0.80', 'Actual/Milestone ratio at/above which Status = Amber (below Green)'],
  ['Weight_Tolerance', '0.01', 'Allowed rounding tolerance when checking Wt% sums to 100'],
  ['Q1_Close_Date', DEFAULT_QUARTER_CLOSE_DATES.Q1, 'Used by escalation logic to flag ATRs open past quarter close'],
  ['Q2_Close_Date', DEFAULT_QUARTER_CLOSE_DATES.Q2, ''],
  ['Q3_Close_Date', DEFAULT_QUARTER_CLOSE_DATES.Q3, ''],
  ['Q4_Close_Date', DEFAULT_QUARTER_CLOSE_DATES.Q4, ''],
  ['Reminder_Lead_Days', '7', 'How many days before a due date notifyKpiOwners() sends a reminder']
];

/**
 * Canonical field -> known alternate header wordings seen across the real
 * source sheets. normalizeHeaderText_() (03_ImportPipeline.gs) lowercases
 * and strips ALL whitespace/newlines/punctuation before comparing, so most
 * real-workbook quirks (embedded newlines, "S.\nNo" vs "S.No", extra spaces
 * around parentheses) already resolve without an entry here. Only entries
 * where the WORDS themselves differ between Tier 0/1 and Tier 2 sheets (or
 * where a Tier 0/1 sheet uses a bare "Q1" instead of "Q1 Milestone Target")
 * need to be listed.
 */
const HEADER_ALIASES = {
  'S.No': ['S.No', '#'],
  'Strategic Pillar': ['Strategic Pillar', 'Pillar'],
  'How to Measure (KPI-Specific Formula)': ['How to Measure (KPI-Specific Formula)', 'How it is measured'],
  'Q1 Milestone Target': ['Q1 Milestone Target', 'Q1'],
  'Q2 Milestone Target': ['Q2 Milestone Target', 'Q2'],
  'Q3 Milestone Target': ['Q3 Milestone Target', 'Q3'],
  'Q4 Milestone Target': ['Q4 Milestone Target', 'Q4']
};

// Dashboard sheet names (09_Dashboards.gs). VC/Cluster are script-computed
// snapshots (refreshed by "Refresh Dashboards"); Office/Goal/My KPIs are
// live QUERY()-formula sheets driven by a picker cell, so they stay current
// without needing a manual refresh.
const DASHBOARD_SHEET = {
  VC: 'Dashboard_VC',
  CLUSTER: 'Dashboard_Cluster',
  OFFICE: 'Dashboard_Office',
  GOAL: 'Dashboard_Goal',
  MY_KPIS: 'Dashboard_MyKPIs'
};

// The compound source field carrying both Goal ID and target text, e.g.
// "P1.4 — International Accreditation: AACSB and ABET". extractGoalId_() in
// 03_ImportPipeline.gs pulls the leading "P<n>.<n>" out of this.
const SOURCE_GOAL_COMBINED_HEADER = '5-Yr Goal (ID + target)';
