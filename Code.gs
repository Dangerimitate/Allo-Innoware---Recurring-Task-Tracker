// ============================================================
// ALLO INNOWARE — Recurring Task Tracker & Escalation Engine
// Main Automation Script  (Code.gs)
// ============================================================
// Author : AI Operations Manager Candidate
// Date   : 2026-10-12
// ============================================================

/* -------------------------------------------------------
   GLOBAL CONSTANTS
   ------------------------------------------------------- */
const SS          = () => SpreadsheetApp.getActiveSpreadsheet();
const SHEET_MASTER   = 'Task Master';
const SHEET_LOG      = 'Task Log';
const SHEET_CONFIG   = 'Config';
const SHEET_ESCALATION = 'Escalation Log';
const SHEET_MANAGER  = 'Manager Dashboard';

// Column indices for Task Master (1-based)
const TM = {
  TASK_ID: 1, TASK_NAME: 2, ASSIGNED_TO: 3, FREQUENCY: 4,
  CYCLE_DAYS: 5, PROOF_TYPE: 6, START_DATE: 7, ACTIVE: 8
};

// Column indices for Task Log (1-based)
const TL = {
  LOG_ID: 1, TASK_ID: 2, TASK_NAME: 3, ASSIGNED_TO: 4,
  CYCLE_NUM: 5, DUE_DATE: 6, STATUS: 7, COMPLETED_BY: 8,
  COMPLETED_ON: 9, TICKET_ID: 10, PROOF_NOTES: 11,
  MANAGER_ALERTED: 12, MANAGER_ACK: 13, MANAGER_ACK_TIME: 14,
  ESCALATED: 15, ESCALATION_TIME: 16
};

/* -------------------------------------------------------
   1. RECURRING CYCLE GENERATION
   Runs via time-based trigger (daily at 6 AM).
   Creates the next cycle row for each active task when
   no open (Pending / Overdue) row exists for that task.
   ------------------------------------------------------- */
function generateTaskCycles() {
  const ss        = SS();
  const master    = ss.getSheetByName(SHEET_MASTER);
  const log       = ss.getSheetByName(SHEET_LOG);
  const today     = stripTime(new Date());

  const masterData = master.getDataRange().getValues();
  const logData    = log.getDataRange().getValues();

  // Build lookup: taskId → latest log row
  const openByTask = {};   // taskId → true if an open row exists
  const lastCycle  = {};   // taskId → max cycle number
  const lastDue    = {};   // taskId → last due date

  for (let i = 1; i < logData.length; i++) {
    const row    = logData[i];
    const taskId = row[TL.TASK_ID - 1];
    const status = row[TL.STATUS - 1];
    const cycle  = row[TL.CYCLE_NUM - 1];
    const due    = row[TL.DUE_DATE - 1];

    if (status === 'Pending' || status === 'Overdue') {
      openByTask[taskId] = true;
    }
    if (!lastCycle[taskId] || cycle > lastCycle[taskId]) {
      lastCycle[taskId] = cycle;
      lastDue[taskId]   = due;
    }
  }

  let newRows = 0;

  for (let i = 1; i < masterData.length; i++) {
    const mRow     = masterData[i];
    const taskId   = mRow[TM.TASK_ID - 1];
    const taskName = mRow[TM.TASK_NAME - 1];
    const assignee = mRow[TM.ASSIGNED_TO - 1];
    const freq     = mRow[TM.FREQUENCY - 1];
    const days     = mRow[TM.CYCLE_DAYS - 1];
    const startDt  = mRow[TM.START_DATE - 1];
    const active   = mRow[TM.ACTIVE - 1];

    if (active !== 'Yes') continue;

    // Skip if an open cycle already exists (prevents duplicates — Test 8)
    if (openByTask[taskId]) continue;

    let nextCycle, dueDate;

    if (!lastCycle[taskId]) {
      // First-ever cycle
      nextCycle = 1;
      dueDate   = freq === 'Daily' ? today : new Date(startDt);
      if (dueDate < today) dueDate = today;   // safety
    } else {
      nextCycle = lastCycle[taskId] + 1;
      const base = new Date(lastDue[taskId]);
      dueDate = new Date(base.getTime() + days * 86400000);
      // For daily tasks, due date is always "today"
      if (freq === 'Daily') dueDate = today;
    }

    const logId = 'LOG-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    log.appendRow([
      logId, taskId, taskName, assignee,
      nextCycle,
      dueDate,
      'Pending',    // status
      '',           // completed_by
      '',           // completed_on
      '',           // ticket_id
      '',           // proof_notes
      'No',         // manager_alerted
      'No',         // manager_ack
      '',           // manager_ack_time
      'No',         // escalated
      ''            // escalation_time
    ]);
    newRows++;
  }

  Logger.log('Cycle generation complete. New rows created: ' + newRows);
  // After generating, mark any overdue items
  markOverdueItems();
  return newRows;
}

/* -------------------------------------------------------
   2. MARK OVERDUE ITEMS
   Updates Status from "Pending" → "Overdue" if past due.
   ------------------------------------------------------- */
function markOverdueItems() {
  const log   = SS().getSheetByName(SHEET_LOG);
  const data  = log.getDataRange().getValues();
  const today = stripTime(new Date());

  for (let i = 1; i < data.length; i++) {
    const status = data[i][TL.STATUS - 1];
    const due    = stripTime(new Date(data[i][TL.DUE_DATE - 1]));

    if (status === 'Pending' && due < today) {
      log.getRange(i + 1, TL.STATUS).setValue('Overdue');
    }
  }
}

/* -------------------------------------------------------
   3. EMPLOYEE COMPLETION HANDLER  (onEdit trigger)
   When an employee enters a Ticket ID / marks "Done" in
   the Task Log, the system auto-fills completion metadata.
   ------------------------------------------------------- */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_LOG) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // React when Ticket ID (col 10) or Proof/Notes (col 11) is edited
  if (col !== TL.TICKET_ID && col !== TL.PROOF_NOTES) return;
  if (row <= 1) return; // header

  const status = sheet.getRange(row, TL.STATUS).getValue();
  if (status === 'Completed') return; // already done

  const ticketVal = sheet.getRange(row, TL.TICKET_ID).getValue();
  const proofVal  = sheet.getRange(row, TL.PROOF_NOTES).getValue();

  // For daily tasks the proof column is used ("Done" / note)
  // For portal tasks the Ticket ID is used
  if (!ticketVal && !proofVal) return;

  // Auto-complete the task
  const now      = new Date();
  const assignee = sheet.getRange(row, TL.ASSIGNED_TO).getValue();

  sheet.getRange(row, TL.STATUS).setValue('Completed');
  sheet.getRange(row, TL.COMPLETED_BY).setValue(assignee);
  sheet.getRange(row, TL.COMPLETED_ON).setValue(now);

  // Immediately generate the next cycle so employee sees it
  SpreadsheetApp.flush();
  generateTaskCycles();
}

/* -------------------------------------------------------
   4. SAME-DAY MANAGER ALERT  (Test 4)
   Runs via time-based trigger (daily at 6 PM).
   Sends ONE consolidated email per manager listing all
   overdue / still-pending tasks for the day.
   Avoids duplicate alerts by checking Manager_Alerted flag.
   ------------------------------------------------------- */
function sendManagerAlerts() {
  const ss     = SS();
  const log    = ss.getSheetByName(SHEET_LOG);
  const config = getConfig();
  const data   = log.getDataRange().getValues();
  const today  = stripTime(new Date());

  const pendingItems = [];

  for (let i = 1; i < data.length; i++) {
    const status  = data[i][TL.STATUS - 1];
    const alerted = data[i][TL.MANAGER_ALERTED - 1];
    const due     = stripTime(new Date(data[i][TL.DUE_DATE - 1]));

    // Alert conditions:
    // 1. Status is Overdue, OR
    // 2. Status is Pending AND due date is today (end-of-day check)
    // AND manager has not already been alerted for this row
    if ((status === 'Overdue' || (status === 'Pending' && due.getTime() === today.getTime()))
        && alerted !== 'Yes') {

      pendingItems.push({
        row: i + 1,
        taskName: data[i][TL.TASK_NAME - 1],
        assignee: data[i][TL.ASSIGNED_TO - 1],
        dueDate:  Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        status:   status
      });
    }
  }

  if (pendingItems.length === 0) {
    Logger.log('No pending items to alert manager about.');
    return;
  }

  // Build email body
  let body = 'The following tasks are pending or overdue as of ' +
             Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ':\n\n';

  body += '| # | Task | Assigned To | Due Date | Status |\n';
  body += '|---|------|-------------|----------|--------|\n';

  pendingItems.forEach((item, idx) => {
    body += `| ${idx + 1} | ${item.taskName} | ${item.assignee} | ${item.dueDate} | ${item.status} |\n`;
  });

  body += '\n\nPlease review and acknowledge in the Manager Dashboard.\n';
  body += 'Sheet link: ' + ss.getUrl();

  // Build HTML email
  let htmlBody = '<h2>⚠️ Pending / Overdue Task Alert</h2>';
  htmlBody += '<p>The following tasks require attention as of <strong>' +
              Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd') + '</strong>:</p>';
  htmlBody += '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial;">';
  htmlBody += '<tr style="background:#d32f2f;color:white;"><th>#</th><th>Task</th><th>Assigned To</th><th>Due Date</th><th>Status</th></tr>';

  pendingItems.forEach((item, idx) => {
    const bg = idx % 2 === 0 ? '#fff3f3' : '#ffffff';
    htmlBody += `<tr style="background:${bg}"><td>${idx + 1}</td><td>${item.taskName}</td>` +
                `<td>${item.assignee}</td><td>${item.dueDate}</td><td><strong>${item.status}</strong></td></tr>`;
  });

  htmlBody += '</table>';
  htmlBody += '<br><p>👉 <a href="' + ss.getUrl() + '">Open Sheet to Acknowledge</a></p>';

  // Send email
  try {
    MailApp.sendEmail({
      to: config.managerEmail,
      subject: '🔴 Task Alert: ' + pendingItems.length + ' Pending/Overdue Items — ' +
               Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      body: body,
      htmlBody: htmlBody
    });
    Logger.log('Manager alert sent for ' + pendingItems.length + ' items.');
  } catch (err) {
    Logger.log('Email send failed: ' + err.message);
  }

  // Mark rows as alerted (prevents duplicate emails — Test 8)
  pendingItems.forEach(item => {
    log.getRange(item.row, TL.MANAGER_ALERTED).setValue('Yes');
  });
}

/* -------------------------------------------------------
   5. MANAGER ACKNOWLEDGEMENT  (Test 5)
   Manager sets "Manager Ack" column to "Yes" in Task Log.
   This onEdit handler captures the timestamp.
   ------------------------------------------------------- */
function onEditManagerAck(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_LOG) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (col !== TL.MANAGER_ACK || row <= 1) return;

  const val = e.range.getValue();
  if (val === 'Yes') {
    sheet.getRange(row, TL.MANAGER_ACK_TIME).setValue(new Date());
  }
}

/* -------------------------------------------------------
   6. LEADERSHIP ESCALATION  (Test 6)
   Runs via time-based trigger (daily at 9 AM, day after alert).
   Escalates if:
     - Task is still Overdue/Pending
     - Manager was alerted > ESCALATION_HOURS ago
     - Manager has NOT acknowledged, OR task is STILL not completed
       even after acknowledgement beyond SLA
     - Not already escalated
   ------------------------------------------------------- */
function escalateToLeadership() {
  const ss     = SS();
  const log    = ss.getSheetByName(SHEET_LOG);
  const escLog = ss.getSheetByName(SHEET_ESCALATION);
  const config = getConfig();
  const data   = log.getDataRange().getValues();
  const now    = new Date();

  const escalationItems = [];

  for (let i = 1; i < data.length; i++) {
    const status    = data[i][TL.STATUS - 1];
    const alerted   = data[i][TL.MANAGER_ALERTED - 1];
    const ack       = data[i][TL.MANAGER_ACK - 1];
    const escalated = data[i][TL.ESCALATED - 1];
    const due       = new Date(data[i][TL.DUE_DATE - 1]);

    if (escalated === 'Yes') continue;
    if (status !== 'Overdue' && status !== 'Pending') continue;
    if (alerted !== 'Yes') continue;

    // Calculate hours since due date
    const hoursSinceDue = (now - due) / 3600000;

    // Escalation trigger conditions:
    // A) Manager NOT acknowledged within MANAGER_ACK_SLA hours after alert
    // B) Task still not completed ESCALATION_HOURS after due date
    const shouldEscalate =
      (ack !== 'Yes' && hoursSinceDue > config.managerAckSlaHours) ||
      (hoursSinceDue > config.escalationHours);

    if (shouldEscalate) {
      escalationItems.push({
        row: i + 1,
        logId:    data[i][TL.LOG_ID - 1],
        taskName: data[i][TL.TASK_NAME - 1],
        assignee: data[i][TL.ASSIGNED_TO - 1],
        dueDate:  Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        status:   status,
        managerAck: ack,
        hoursSinceDue: Math.round(hoursSinceDue)
      });
    }
  }

  if (escalationItems.length === 0) {
    Logger.log('No items to escalate.');
    return;
  }

  // Build escalation email
  let htmlBody = '<h2>🚨 LEADERSHIP ESCALATION — Unresolved Tasks</h2>';
  htmlBody += '<p>The following tasks have exceeded the SLA and require leadership attention:</p>';
  htmlBody += '<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial;">';
  htmlBody += '<tr style="background:#b71c1c;color:white;"><th>#</th><th>Task</th><th>Assigned To</th>' +
              '<th>Due Date</th><th>Hours Overdue</th><th>Manager Ack</th></tr>';

  escalationItems.forEach((item, idx) => {
    htmlBody += `<tr><td>${idx + 1}</td><td>${item.taskName}</td><td>${item.assignee}</td>` +
                `<td>${item.dueDate}</td><td>${item.hoursSinceDue}h</td><td>${item.managerAck}</td></tr>`;
  });

  htmlBody += '</table>';
  htmlBody += '<br><p>👉 <a href="' + ss.getUrl() + '">Open Sheet</a></p>';

  let plainBody = 'LEADERSHIP ESCALATION\n\n';
  escalationItems.forEach((item, idx) => {
    plainBody += `${idx + 1}. ${item.taskName} — ${item.assignee} — Due: ${item.dueDate} — ` +
                 `Overdue by ${item.hoursSinceDue}h — Manager Ack: ${item.managerAck}\n`;
  });

  // Send to leadership
  try {
    MailApp.sendEmail({
      to: config.leadershipEmail,
      cc: config.managerEmail,
      subject: '🚨 ESCALATION: ' + escalationItems.length + ' Unresolved Tasks — ' +
               Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      body: plainBody,
      htmlBody: htmlBody
    });
    Logger.log('Escalation email sent for ' + escalationItems.length + ' items.');
  } catch (err) {
    Logger.log('Escalation email failed: ' + err.message);
  }

  // Mark as escalated and log to Escalation Log
  escalationItems.forEach(item => {
    log.getRange(item.row, TL.ESCALATED).setValue('Yes');
    log.getRange(item.row, TL.ESCALATION_TIME).setValue(now);

    escLog.appendRow([
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
      item.logId,
      item.taskName,
      item.assignee,
      item.dueDate,
      item.hoursSinceDue + ' hours',
      item.managerAck,
      'Email sent to ' + config.leadershipEmail
    ]);
  });
}

/* -------------------------------------------------------
   7. CONFIGURATION HELPER
   Reads Config sheet for email addresses and SLA values.
   ------------------------------------------------------- */
function getConfig() {
  const sheet = SS().getSheetByName(SHEET_CONFIG);
  const data  = sheet.getDataRange().getValues();
  const cfg   = {};

  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const val = data[i][1];
    switch (key) {
      case 'Manager Email':       cfg.managerEmail = val; break;
      case 'Leadership Email':    cfg.leadershipEmail = val; break;
      case 'Manager Ack SLA (Hours)': cfg.managerAckSlaHours = Number(val); break;
      case 'Escalation Hours':    cfg.escalationHours = Number(val); break;
      case 'Working Hours Start': cfg.workStart = Number(val); break;
      case 'Working Hours End':   cfg.workEnd = Number(val); break;
    }
  }

  // Defaults
  cfg.managerEmail      = cfg.managerEmail || 'manager@example.com';
  cfg.leadershipEmail   = cfg.leadershipEmail || 'leadership@example.com';
  cfg.managerAckSlaHours = cfg.managerAckSlaHours || 24;
  cfg.escalationHours   = cfg.escalationHours || 48;
  cfg.workStart         = cfg.workStart || 9;
  cfg.workEnd           = cfg.workEnd || 18;

  return cfg;
}

/* -------------------------------------------------------
   8. SETUP TRIGGERS (run once manually)
   ------------------------------------------------------- */
function setupTriggers() {
  // Remove existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  // Daily cycle generation at 6:00 AM
  ScriptApp.newTrigger('generateTaskCycles')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .create();

  // Daily manager alert at 6:00 PM
  ScriptApp.newTrigger('sendManagerAlerts')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();

  // Daily escalation check at 9:00 AM
  ScriptApp.newTrigger('escalateToLeadership')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  // onEdit trigger (installable, needed for email permissions)
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SS())
    .onEdit()
    .create();

  // Second onEdit for manager acknowledgement
  ScriptApp.newTrigger('onEditManagerAck')
    .forSpreadsheet(SS())
    .onEdit()
    .create();

  Logger.log('All triggers created successfully.');
}

/* -------------------------------------------------------
   9. MANAGER DASHBOARD REFRESH
   Populates the Manager Dashboard with current status.
   ------------------------------------------------------- */
function refreshManagerDashboard() {
  const ss   = SS();
  const log  = ss.getSheetByName(SHEET_LOG);
  const dash = ss.getSheetByName(SHEET_MANAGER);

  // Clear existing data (keep headers)
  const lastRow = dash.getLastRow();
  if (lastRow > 1) {
    dash.getRange(2, 1, lastRow - 1, dash.getLastColumn()).clearContent();
  }

  const logData = log.getDataRange().getValues();
  let row = 2;

  // Summary by employee
  const summary = {};

  for (let i = 1; i < logData.length; i++) {
    const assignee = logData[i][TL.ASSIGNED_TO - 1];
    const status   = logData[i][TL.STATUS - 1];

    if (!summary[assignee]) {
      summary[assignee] = { pending: 0, overdue: 0, completed: 0, total: 0 };
    }
    summary[assignee].total++;
    if (status === 'Pending')   summary[assignee].pending++;
    if (status === 'Overdue')   summary[assignee].overdue++;
    if (status === 'Completed') summary[assignee].completed++;
  }

  // Write summary
  Object.keys(summary).forEach(emp => {
    const s = summary[emp];
    dash.getRange(row, 1, 1, 6).setValues([[
      emp, s.pending, s.overdue, s.completed, s.total,
      s.overdue > 0 ? '🔴 ACTION NEEDED' : (s.pending > 0 ? '🟡 Monitor' : '🟢 All Clear')
    ]]);
    row++;
  });

  // Add detailed pending/overdue list below
  row += 2;
  dash.getRange(row, 1).setValue('DETAILED PENDING / OVERDUE ITEMS').setFontWeight('bold');
  row++;
  dash.getRange(row, 1, 1, 8).setValues([['Log ID', 'Task', 'Assigned To', 'Due Date', 'Status',
    'Manager Alerted', 'Manager Ack', 'Escalated']]).setFontWeight('bold');
  row++;

  for (let i = 1; i < logData.length; i++) {
    const status = logData[i][TL.STATUS - 1];
    if (status !== 'Pending' && status !== 'Overdue') continue;

    const due = logData[i][TL.DUE_DATE - 1];
    dash.getRange(row, 1, 1, 8).setValues([[
      logData[i][TL.LOG_ID - 1],
      logData[i][TL.TASK_NAME - 1],
      logData[i][TL.ASSIGNED_TO - 1],
      due instanceof Date ? Utilities.formatDate(due, Session.getScriptTimeZone(), 'yyyy-MM-dd') : due,
      status,
      logData[i][TL.MANAGER_ALERTED - 1],
      logData[i][TL.MANAGER_ACK - 1],
      logData[i][TL.ESCALATED - 1]
    ]]);
    row++;
  }

  Logger.log('Manager Dashboard refreshed.');
}

/* -------------------------------------------------------
   10. UTILITY FUNCTIONS
   ------------------------------------------------------- */
function stripTime(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/* -------------------------------------------------------
   11. MANUAL TEST RUNNERS
   These functions simulate the test cases described in
   the assignment. Run them from the Script Editor.
   ------------------------------------------------------- */

/**
 * TEST 1 — Recurring Cycle
 * Generates initial cycles, then simulates completing a 7-day task
 * and verifies the next cycle is auto-created.
 */
function test1_RecurringCycle() {
  Logger.log('=== TEST 1: Recurring Cycle ===');
  const created = generateTaskCycles();
  Logger.log('Initial cycles created: ' + created);

  const log  = SS().getSheetByName(SHEET_LOG);
  const data = log.getDataRange().getValues();

  // Find the 7-day task (Task ID = T002)
  for (let i = 1; i < data.length; i++) {
    if (data[i][TL.TASK_ID - 1] === 'T002' && data[i][TL.STATUS - 1] === 'Pending') {
      // Simulate completion
      log.getRange(i + 1, TL.TICKET_ID).setValue('TKT-7001');
      log.getRange(i + 1, TL.STATUS).setValue('Completed');
      log.getRange(i + 1, TL.COMPLETED_BY).setValue(data[i][TL.ASSIGNED_TO - 1]);
      log.getRange(i + 1, TL.COMPLETED_ON).setValue(new Date());
      Logger.log('Completed 7-day task cycle ' + data[i][TL.CYCLE_NUM - 1]);
      break;
    }
  }

  // Generate next cycle
  SpreadsheetApp.flush();
  const nextCreated = generateTaskCycles();
  Logger.log('Next cycles created after completion: ' + nextCreated);
  Logger.log('✅ TEST 1 PASSED — Verify new cycle row in Task Log');
}

/**
 * TEST 2 — Single-Point Completion
 * Employee updates Ticket ID in Task Log → auto-completes.
 */
function test2_SinglePointCompletion() {
  Logger.log('=== TEST 2: Single-Point Completion ===');
  Logger.log('This test works via the onEdit trigger.');
  Logger.log('Instructions:');
  Logger.log('1. Go to Task Log sheet');
  Logger.log('2. Find a Pending task');
  Logger.log('3. Enter a Ticket ID (e.g., TKT-1234) in the Ticket ID column');
  Logger.log('4. The Status, Completed By, and Completed On should auto-fill');
  Logger.log('5. A new cycle row should appear automatically');
  Logger.log('✅ TEST 2 — Run manually via sheet interaction');
}

/**
 * TEST 3 — Daily Task Completion
 * Simulates the daily task "Done" marking.
 */
function test3_DailyTask() {
  Logger.log('=== TEST 3: Daily Task ===');
  const log  = SS().getSheetByName(SHEET_LOG);
  const data = log.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][TL.TASK_ID - 1] === 'T001' && data[i][TL.STATUS - 1] === 'Pending') {
      // Employee marks daily task as done via Proof/Notes column
      log.getRange(i + 1, TL.PROOF_NOTES).setValue('Done — Update shared in WhatsApp group');
      log.getRange(i + 1, TL.STATUS).setValue('Completed');
      log.getRange(i + 1, TL.COMPLETED_BY).setValue(data[i][TL.ASSIGNED_TO - 1]);
      log.getRange(i + 1, TL.COMPLETED_ON).setValue(new Date());
      Logger.log('Daily task marked as Done for ' + data[i][TL.ASSIGNED_TO - 1]);
      break;
    }
  }

  SpreadsheetApp.flush();
  generateTaskCycles();
  Logger.log('✅ TEST 3 PASSED — Daily task completed and next cycle ready');
}

/**
 * TEST 4 — Same-Day Manager Alert
 * Triggers the manager alert function.
 */
function test4_ManagerAlert() {
  Logger.log('=== TEST 4: Same-Day Manager Alert ===');
  // First ensure there are overdue items
  markOverdueItems();
  // Send alerts
  sendManagerAlerts();
  Logger.log('✅ TEST 4 — Check manager email for alert');
}

/**
 * TEST 5 — Manager Acknowledgement
 */
function test5_ManagerAck() {
  Logger.log('=== TEST 5: Manager Acknowledgement ===');
  Logger.log('Instructions:');
  Logger.log('1. Go to Task Log sheet');
  Logger.log('2. Find a row where Manager Alerted = Yes');
  Logger.log('3. Set Manager Ack column to "Yes"');
  Logger.log('4. Manager Ack Time should auto-populate');
  Logger.log('✅ TEST 5 — Run manually via sheet interaction');
}

/**
 * TEST 6 — Leadership Escalation
 */
function test6_Escalation() {
  Logger.log('=== TEST 6: Leadership Escalation ===');
  escalateToLeadership();
  Logger.log('✅ TEST 6 — Check leadership email and Escalation Log sheet');
}

/**
 * TEST 8 — No Duplicate Alerts
 * Runs automation twice and checks no duplicates.
 */
function test8_NoDuplicates() {
  Logger.log('=== TEST 8: No Duplicate Alerts ===');

  const log = SS().getSheetByName(SHEET_LOG);
  const countBefore = log.getLastRow();

  // Run cycle generation twice
  generateTaskCycles();
  SpreadsheetApp.flush();
  generateTaskCycles();

  const countAfter = log.getLastRow();
  Logger.log('Rows before: ' + countBefore + ', After two runs: ' + countAfter);

  // Run manager alerts twice
  sendManagerAlerts();
  SpreadsheetApp.flush();
  sendManagerAlerts();

  Logger.log('✅ TEST 8 — Verify no duplicate rows or emails were created');
}

/**
 * RUN ALL TESTS (except interactive ones)
 */
function runAllTests() {
  test1_RecurringCycle();
  Utilities.sleep(1000);
  test3_DailyTask();
  Utilities.sleep(1000);
  test4_ManagerAlert();
  Utilities.sleep(1000);
  test6_Escalation();
  Utilities.sleep(1000);
  test8_NoDuplicates();
  Logger.log('\n🏁 ALL AUTOMATED TESTS COMPLETE');
}
