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
    .addItem('View Audit Log', 'jumpToSheet_Audit')
    .addItem('View Import Log', 'jumpToSheet_ImportLog')
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
