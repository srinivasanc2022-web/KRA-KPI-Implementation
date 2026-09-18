/**
 * 10_Escalation.gs
 *
 * Spec section 10: auto-escalate when a KPI turns Red, an ATR is open past
 * its quarter's close, a milestone target is missed, or Wt(%) totals stop
 * summing to 100; notify KPI Owners ahead of their KPI's own Frequency-based
 * due date (not on a fixed schedule), and notify Responsible Officers /
 * Informed parties on Status changes.
 *
 * REAL-DATA CAVEAT: the workbook only ever carries FOUR milestone
 * checkpoints (Q1-Q4), regardless of what a KPI's own Frequency says
 * (Monthly, Weekly, "Per exam cycle", ...). There's no monthly/weekly slot
 * to be "due" against. So due-date reminders key off the same four
 * Config-configurable quarter-close dates for every KPI; a KPI whose
 * Frequency is finer than quarterly will just get reminded at each quarter
 * close like everything else. A true per-Frequency calendar would need the
 * workbook itself to carry more than 4 checkpoints -- noted in
 * NEXT_STEPS.md rather than faked here.
 *
 * CONTACT DATA CAVEAT: the workbook has no email addresses -- 'KPI Owner'
 * and 'Responsible Officer' are role/name text ("Dean AA"), not addresses.
 * Notifications resolve a recipient address via the Office_Contacts sheet
 * (Office Code -> email); until that's filled in, notifyKpiOwners() logs a
 * "SKIPPED - no contact email" row per KPI instead of silently doing
 * nothing, so the gap is visible rather than hidden.
 */

function runEscalationScanFromMenu() {
  const result = runEscalationScan();
  SpreadsheetApp.getUi().alert(
    'Escalation scan complete.\n' +
    'Red KPIs: ' + result.red + '\n' +
    'ATRs overdue: ' + result.overdueAtr + '\n' +
    'Milestones missed (no actual logged past due date): ' + result.missedMilestone + '\n' +
    'Weight drift: ' + result.weightDrift + '\n' +
    'New escalations logged: ' + result.newEscalations
  );
}

function runEscalationScan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const map = headerIndexMap_(kpiSheet);
  const counts = { red: 0, overdueAtr: 0, missedMilestone: 0, weightDrift: 0, newEscalations: 0 };
  if (kpiSheet.getLastRow() < 2) return counts;

  const openReasons = openEscalationKeys_(ss);
  const toLog = [];
  const now = new Date();
  const closeDates = getQuarterCloseDates_(ss);

  const rows = kpiSheet.getRange(2, 1, kpiSheet.getLastRow() - 1, kpiSheet.getLastColumn()).getValues();
  rows.forEach(function (row) {
    const kpiId = row[map['KPI_ID'] - 1];
    const office = row[map['Office'] - 1];

    QUARTERS.forEach(function (q) {
      const status = row[map[q + ' Status'] - 1];
      const atr = row[map[q + ' ATR'] - 1];
      const actual = row[map[q + ' Actual'] - 1];
      const closeDate = closeDates[q];
      const pastClose = closeDate && now > closeDate;

      if (status === RAG.RED) {
        counts.red++;
        pushEscalation_(toLog, openReasons, kpiId, office, 'Red status',
          q + ' status is Red.', now);
      }
      if ((status === RAG.AMBER || status === RAG.RED) && !atr && pastClose) {
        counts.overdueAtr++;
        pushEscalation_(toLog, openReasons, kpiId, office, 'ATR overdue',
          q + ' is ' + status + ' with no ATR logged, and ' + q + ' closed ' + formatDate_(closeDate) + '.', now);
      }
      if (pastClose && (actual === '' || actual === null || actual === undefined)) {
        counts.missedMilestone++;
        pushEscalation_(toLog, openReasons, kpiId, office, 'Milestone missed',
          q + ' closed ' + formatDate_(closeDate) + ' with no Actual logged.', now);
      }
    });
  });

  const weightFails = validateAllWeights().results.filter(function (r) { return r.status === 'FAIL'; });
  weightFails.forEach(function (r) {
    counts.weightDrift++;
    pushEscalation_(toLog, openReasons, r.entityId, r.scope, 'Weight drift', r.details, now);
  });

  counts.newEscalations = toLog.length;
  if (toLog.length > 0) {
    const sheet = ss.getSheetByName(SHEET.ESCALATIONS);
    sheet.getRange(sheet.getLastRow() + 1, 1, toLog.length, toLog[0].length).setValues(toLog);
  }
  return counts;
}

/** Skip re-logging an escalation that's already open for the same KPI/entity + reason. */
function openEscalationKeys_(ss) {
  const rows = existingRows_(ss.getSheetByName(SHEET.ESCALATIONS));
  const keys = {};
  rows.forEach(function (r) {
    if (String(r['Status'] || '').trim().toLowerCase() === 'open' || !r['Status']) {
      keys[r['KPI_ID'] + '|' + r['Reason']] = true;
    }
  });
  return keys;
}

function pushEscalation_(toLog, openReasons, entityId, office, reason, detail, now) {
  const key = entityId + '|' + reason;
  if (openReasons[key]) return;
  openReasons[key] = true;
  toLog.push([now, entityId, office, reason, detail, 'Open']);
}

function getQuarterCloseDates_(ss) {
  const dates = {};
  QUARTERS.forEach(function (q) {
    const raw = getConfigValue_(ss, q + '_Close_Date', DEFAULT_QUARTER_CLOSE_DATES[q]);
    const d = new Date(raw);
    dates[q] = isNaN(d.getTime()) ? null : d;
  });
  return dates;
}

function getConfigValue_(ss, key, defaultValue) {
  const rows = existingRows_(ss.getSheetByName(SHEET.CONFIG));
  const match = rows.filter(function (r) { return r['Key'] === key; })[0];
  return match ? match['Value'] : defaultValue;
}

function formatDate_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Kolkata', 'yyyy-MM-dd'); }

// ---------------------------------------------------------------------------
// Notifications
// ---------------------------------------------------------------------------

/**
 * Reminds KPI Owners ahead of the quarter close nearest their KPI's own
 * Frequency (see the real-data caveat above). Install this on a daily
 * time-driven trigger (Apps Script editor -> Triggers -> Add Trigger ->
 * notifyKpiOwners -> Time-driven -> Day timer) to run automatically; the
 * menu item runs it on demand too.
 */
function notifyKpiOwners() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kpiSheet = ss.getSheetByName(SHEET.KPIS);
  const map = headerIndexMap_(kpiSheet);
  if (kpiSheet.getLastRow() < 2) return;

  const leadDays = getConfigNumber_(ss, 'Reminder_Lead_Days', 7);
  const closeDates = getQuarterCloseDates_(ss);
  const now = new Date();
  const alreadySent = sentReminderKeys_(ss);
  const contacts = officeContacts_(ss);
  const logRows = [];

  const rows = kpiSheet.getRange(2, 1, kpiSheet.getLastRow() - 1, kpiSheet.getLastColumn()).getValues();
  rows.forEach(function (row) {
    const kpiId = row[map['KPI_ID'] - 1];
    const office = row[map['Office'] - 1];

    QUARTERS.forEach(function (q) {
      const closeDate = closeDates[q];
      if (!closeDate) return;
      const daysUntilDue = (closeDate.getTime() - now.getTime()) / 86400000;
      if (daysUntilDue < 0 || daysUntilDue > leadDays) return;

      const actual = row[map[q + ' Actual'] - 1];
      if (actual !== '' && actual !== null && actual !== undefined) return; // already logged

      const key = kpiId + '|' + q + '|reminder';
      if (alreadySent[key]) return;

      const email = contacts[office] && contacts[office].ownerEmail;
      const subject = 'KPI due-date reminder: ' + row[map['KPI'] - 1] + ' (' + q + ')';
      const body = q + ' closes ' + formatDate_(closeDate) + ' and no Actual has been logged yet for "' +
        row[map['KPI'] - 1] + '" (' + kpiId + '). Please update the Quarterly Entry sidebar.';

      if (email) {
        try {
          MailApp.sendEmail(email, subject, body);
          logRows.push([new Date(), kpiId, 'KPI Owner', email, 'Due-date reminder (' + q + ')', 'Yes']);
        } catch (err) {
          logRows.push([new Date(), kpiId, 'KPI Owner', email, 'Due-date reminder (' + q + ') - send failed: ' + err.message, 'No']);
        }
      } else {
        logRows.push([new Date(), kpiId, 'KPI Owner', row[map['KPI Owner'] - 1] || '(unknown)',
          'Due-date reminder (' + q + ') - SKIPPED: no contact email in ' + SHEET.OFFICE_CONTACTS, 'No']);
      }
    });
  });

  if (logRows.length > 0) {
    const logSheet = ss.getSheetByName(SHEET.NOTIFICATION_LOG);
    logSheet.getRange(logSheet.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
  }
}

/** Called from onEdit (99_Triggers.gs) when a quarter Status is edited directly to Red. */
function notifyOnStatusChange_(ss, kpiId, office, quarter, newStatus) {
  if (newStatus !== RAG.RED) return;
  const contacts = officeContacts_(ss);
  const contact = contacts[office];
  const logSheet = ss.getSheetByName(SHEET.NOTIFICATION_LOG);

  const recipients = [
    ['Responsible Officer', contact && contact.officerEmail],
    ['Escalation', contact && contact.escalationEmail]
  ];
  recipients.forEach(function (pair) {
    const role = pair[0], email = pair[1];
    if (email) {
      try {
        MailApp.sendEmail(email, 'KPI turned Red: ' + kpiId, quarter + ' status for ' + kpiId + ' (' + office + ') is now Red.');
        logSheet.appendRow([new Date(), kpiId, role, email, quarter + ' status changed to Red', 'Yes']);
      } catch (err) {
        logSheet.appendRow([new Date(), kpiId, role, email, quarter + ' status changed to Red - send failed: ' + err.message, 'No']);
      }
    } else {
      logSheet.appendRow([new Date(), kpiId, role, '(none)', quarter + ' status changed to Red - SKIPPED: no contact email', 'No']);
    }
  });
}

function sentReminderKeys_(ss) {
  const rows = existingRows_(ss.getSheetByName(SHEET.NOTIFICATION_LOG));
  const keys = {};
  rows.forEach(function (r) {
    if (r['Sent'] === 'Yes') keys[r['KPI_ID'] + '|' + String(r['Reason']).replace(/^.*\((Q\d)\).*$/, '$1') + '|reminder'] = true;
  });
  return keys;
}

/**
 * Installs a daily time-driven trigger for notifyKpiOwners() and a weekly
 * one for runEscalationScan(), so both run automatically instead of only on
 * menu click. Safe to re-run -- removes any trigger it previously installed
 * for these functions first, so it never stacks duplicates.
 */
function installEscalationTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'notifyKpiOwners' || fn === 'runEscalationScan') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('notifyKpiOwners').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('runEscalationScan').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('Daily triggers installed for notifyKpiOwners() and runEscalationScan() (~07:00, script timezone).');
}

function officeContacts_(ss) {
  const rows = existingRows_(ss.getSheetByName(SHEET.OFFICE_CONTACTS));
  const map = {};
  rows.forEach(function (r) {
    map[r['Office Code']] = {
      ownerEmail: r['KPI Owner Email'], officerEmail: r['Responsible Officer Email'], escalationEmail: r['Escalation Email']
    };
  });
  return map;
}
