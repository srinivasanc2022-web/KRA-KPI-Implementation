/**
 * 07_Menu.gs
 *
 * Wires every action built so far into a single custom menu so a non-technical
 * KPI Owner or admin never has to open the Apps Script editor to use the
 * system.
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('SRM AP Strategy System')
    .addItem('1. Setup Schema', 'setupSchema')
    .addSeparator()
    .addItem('2a. Build Sample Source Data (placeholder, for testing)', 'buildSampleSourceWorkbook')
    .addItem('2b. Import Sample Source Data -> Canonical Model', 'importSampleSourceData')
    .addItem('2c. Remove Sample Source Sheets', 'removeSampleSourceSheets')
    .addSeparator()
    .addItem('3. Import Real Workbook...', 'promptImportRealWorkbook_')
    .addSeparator()
    .addItem('4. Validate Weights (Office / Cluster / Goal)', 'runWeightValidationFromMenu')
    .addItem('5. Recompute RAG Status (bulk)', 'recomputeAllRagStatuses')
    .addItem('6. Open Quarterly Entry Sidebar', 'openQuarterlyEntrySidebar')
    .addSeparator()
    .addItem('7. Recompute Roll-up Achievement % (KPI -> Pillar)', 'runFullRollup')
    .addSeparator()
    .addItem('8a. Refresh Dashboards', 'buildDashboards')
    .addItem('8b. Open VC Dashboard', 'jumpToSheet_VcDashboard')
    .addItem('8c. Open Cluster Dashboard', 'jumpToSheet_ClusterDashboard')
    .addItem('8d. Open Office Dashboard (pick office in cell)', 'jumpToSheet_OfficeDashboard')
    .addItem('8e. Open Goal Dashboard (pick goal in cell)', 'jumpToSheet_GoalDashboard')
    .addItem('8f. Open My KPIs Dashboard', 'jumpToSheet_MyKpisDashboard')
    .addSeparator()
    .addItem('9. Scan for Escalations (Red / overdue ATR / weight drift)', 'runEscalationScanFromMenu')
    .addItem('10a. Send Due-Date Reminders to KPI Owners (run now)', 'notifyKpiOwners')
    .addItem('10b. Install Daily Escalation/Reminder Triggers', 'installEscalationTriggers')
    .addSeparator()
    .addItem('11a. Generate Quarterly Scorecard Report (pick office)', 'promptQuarterlyScorecardReport_')
    .addItem('11b. Generate Cluster/VC Review Pack', 'generateClusterVcReviewPack')
    .addItem('11c. Generate Owner Performance Report (pick owner)', 'promptOwnerPerformanceReport_')
    .addItem('11d. Regenerate All Scorecard Tabs (for Excel export)', 'regenerateAllScorecardTabs')
    .addSeparator()
    .addItem('View Audit Log', 'jumpToSheet_Audit')
    .addItem('View Import Log', 'jumpToSheet_ImportLog')
    .addItem('View Escalations', 'jumpToSheet_Escalations')
    .addToUi();
}

function promptImportRealWorkbook_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Import Real Workbook',
    'Paste the Google Sheet ID (or full URL) of the SRMAP_KRA-KPI_AY26-27 workbook.\n' +
    'Leave blank to import 01_Index / 03_Strategic_Pillars / 04_Cascade_Map / scorecard ' +
    'tabs from THIS spreadsheet instead (e.g. if you have pasted the real tabs in here directly).',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  let input = response.getResponseText().trim();
  let id = input;
  const match = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) id = match[1];

  importRealWorkbook(id);
}

function jumpToSheet_Audit() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.AUDIT_LOG).activate();
}

function jumpToSheet_ImportLog() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.IMPORT_LOG).activate();
}

function jumpToSheet_Escalations() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.ESCALATIONS).activate();
}

function jumpToSheet_VcDashboard() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET.VC).activate();
}

function jumpToSheet_ClusterDashboard() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET.CLUSTER).activate();
}

function jumpToSheet_OfficeDashboard() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET.OFFICE).activate();
}

function jumpToSheet_GoalDashboard() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET.GOAL).activate();
}

function jumpToSheet_MyKpisDashboard() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASHBOARD_SHEET.MY_KPIS).activate();
}

function promptQuarterlyScorecardReport_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Quarterly Scorecard Report', 'Office code (e.g. AA, SEAS):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const office = response.getResponseText().trim();
  if (!office) return;
  generateQuarterlyScorecardReport(office);
}

function promptOwnerPerformanceReport_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Owner Performance Report', 'KPI Owner name, exactly as it appears in the KPI Owner column:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const owner = response.getResponseText().trim();
  if (!owner) return;
  generateOwnerPerformanceReport(owner);
}
