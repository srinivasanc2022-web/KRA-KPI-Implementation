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
  IMPORT_LOG: 'Import_Log'
};

// ---- Raw source workbook tab names this system knows how to read ----
// (matches SRMAP_KRA-KPI_AY26-27 workbook structure per the build spec)
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

// ---- Column headers, exactly as the build spec (section 3) defines them ----

const COLS = {
  PILLARS: ['Pillar ID', 'Pillar Name'],

  GOALS: [
    'Goal ID', 'Parent Pillar', 'Five-Year Goal / Target', 'Owner Office',
    'Contributing Offices', 'Owner Weight %', 'Overall Progress'
  ],

  CLUSTERS: [
    'Cluster ID', 'Cluster Name', 'Member Offices', 'Aggregate Weight (Check)'
  ],

  OFFICES: [
    'Office Code', 'Full Name', 'Cluster', 'Tier', 'Goals Owned',
    'Goals Contributed To', 'KPI Count (Strategic)', 'KPI Count (Operational)',
    'Vertical Count', 'Total Weight (Check)'
  ],

  VERTICALS: ['Vertical ID', 'Vertical Name', 'Office Code', 'Linked KRAs'],

  KRAS: ['KRA ID', 'KRA Name', 'Parent Vertical ID', 'Office Code', 'Linked KPIs'],

  // Reproduces every column from the office-level sheets, section 3.7, plus a
  // handful of system fields (KPI_ID, Rolls Up From, Co-Owner Offices) needed
  // for roll-up and co-ownership support (section 4).
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
    'Score (Actual vs Target)', 'Data Source', 'KPI Owner',
    'R (Responsible)', 'A (Accountable)', 'C (Consulted)', 'I (Informed)',
    'Responsible Officer', 'Escalation / Remarks', 'Overlap / Remarks',
    'Rolls Up From', 'Co-Owner Offices (Name:%)'
  ],

  KPI_OFFICE_LINKS: ['KPI_ID', 'Office Code', 'Role', 'Co-Ownership %'],

  ACTIONS: [
    'Action ID', 'Description', 'Linked KPI_ID', 'Owner', 'Contributors',
    'Priority', 'Start Date', 'Due Date', 'Dependencies', 'Status',
    '% Complete', 'Closure Criteria', 'Remarks'
  ],

  AUDIT_LOG: ['Timestamp', 'User', 'Sheet', 'Row', 'Field', 'Old Value', 'New Value'],

  VALIDATION_LOG: ['Timestamp', 'Scope', 'Entity ID', 'Total Weight', 'Status', 'Details'],

  IMPORT_LOG: ['Timestamp', 'Source Sheet', 'Tier', 'Cluster', 'Rows Imported', 'Status', 'Details'],

  CONFIG: ['Key', 'Value', 'Notes']
};

// Quarter fields, in the order they appear on a KPI row.
const QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4'];

const RAG = {
  GREEN: 'Green',
  AMBER: 'Amber',
  RED: 'Red',
  GREY: 'Grey'
};

// Default RAG thresholds (% of milestone target achieved). Overridable per
// KPI via the 'Config' sheet in a future iteration; see NEXT_STEPS.md.
const DEFAULT_THRESHOLDS = {
  GREEN_MIN: 0.95, // >= 95% of milestone target
  AMBER_MIN: 0.80  // 80-94% => Amber, < 80% => Red
};

// Units for which automatic Actual-vs-Target % computation is not meaningful;
// these require manual Status entry by the KPI Owner.
const NON_COMPUTABLE_UNITS = ['Band', 'Grade'];

const KRA_TYPES = ['Strategic', 'Operational'];
const METRIC_TYPES = ['Strategic', 'Operational', 'Compliance', 'Governance', 'Outcome'];
const INDICATOR_NATURES = ['Leading', 'Lagging'];
const MEASUREMENT_NATURES = ['Quantitative', 'Qualitative'];
const FREQUENCIES = ['Annual', 'Half-Year', 'Semester', 'Quarterly', 'Monthly'];
const GOAL_ROLES = ['Owner', 'Contributor', 'Institutional', 'Cluster aggregate'];

const DEFAULT_CONFIG = [
  ['Current_AY', 'AY26-27', 'Active academic year for milestone/actual entry'],
  ['Green_Threshold', '0.95', 'Actual/Milestone ratio at/above which Status = Green'],
  ['Amber_Threshold', '0.80', 'Actual/Milestone ratio at/above which Status = Amber (below Green)'],
  ['Weight_Tolerance', '0.01', 'Allowed rounding tolerance when checking Wt% sums to 100']
];
