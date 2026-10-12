// ============================================================
// ALLO INNOWARE — Sheet Structure Setup Script  (SheetSetup.gs)
// Run setupAllSheets() ONCE to create the complete structure.
// ============================================================

function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  createConfigSheet(ss);
  createTaskMasterSheet(ss);
  createTaskLogSheet(ss);
  createManagerDashboard(ss);
  createEscalationLogSheet(ss);
  createEmployeeDashboards(ss);

  Logger.log('✅ All sheets created successfully!');
}

/* -------------------------------------------------------
   CONFIG SHEET
   ------------------------------------------------------- */
function createConfigSheet(ss) {
  let sheet = ss.getSheetByName('Config');
  if (!sheet) sheet = ss.insertSheet('Config');
  sheet.clear();

  // Headers
  sheet.getRange('A1:B1').setValues([['Setting', 'Value']]).setFontWeight('bold')
    .setBackground('#1a237e').setFontColor('white');

  // Config data
  const configData = [
    ['Manager Email', 'manager@example.com'],
    ['Leadership Email', 'leadership@example.com'],
    ['Manager Ack SLA (Hours)', 24],
    ['Escalation Hours', 48],
    ['Working Hours Start', 9],
    ['Working Hours End', 18],
    ['Company Name', 'Allo Innoware'],
    ['System Timezone', 'Asia/Kolkata']
  ];

  sheet.getRange(2, 1, configData.length, 2).setValues(configData);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 300);
  sheet.getRange('A1:B' + (configData.length + 1)).setBorder(true, true, true, true, true, true);

  // Add instructions
  sheet.getRange('D1').setValue('⚙️ CONFIGURATION').setFontWeight('bold').setFontSize(14);
  sheet.getRange('D3').setValue('Update the values in column B to configure the system.');
  sheet.getRange('D4').setValue('• Manager Email: receives daily pending task alerts');
  sheet.getRange('D5').setValue('• Leadership Email: receives escalation notifications');
  sheet.getRange('D6').setValue('• Manager Ack SLA: hours manager has to acknowledge alert');
  sheet.getRange('D7').setValue('• Escalation Hours: hours after due date before leadership escalation');
}

/* -------------------------------------------------------
   TASK MASTER SHEET
   ------------------------------------------------------- */
function createTaskMasterSheet(ss) {
  let sheet = ss.getSheetByName('Task Master');
  if (!sheet) sheet = ss.insertSheet('Task Master');
  sheet.clear();

  // Headers
  const headers = [
    'Task ID', 'Task Name', 'Assigned To', 'Frequency',
    'Cycle Days', 'Proof Type', 'Start Date', 'Active'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#0d47a1').setFontColor('white')
    .setHorizontalAlignment('center');

  // Dummy task data
  const today = new Date();
  const tasks = [
    ['T001', 'Daily Operations Update', 'Employee A', 'Daily', 1,
     'Action Confirmation', today, 'Yes'],
    ['T002', 'Portal Task A', 'Employee B', 'Every 7 days', 7,
     'Ticket ID', today, 'Yes'],
    ['T003', 'Portal Task B', 'Employee C', 'Every 15 days', 15,
     'Ticket ID', today, 'Yes'],
    ['T004', 'Portal Task C', 'Employee D', 'Every 30 days', 30,
     'Ticket ID', today, 'Yes'],
    ['T005', 'Weekly Status Report', 'Employee A', 'Every 7 days', 7,
     'Ticket ID', today, 'Yes'],
    ['T006', 'Daily Client Check-in', 'Employee B', 'Daily', 1,
     'Action Confirmation', today, 'Yes']
  ];

  sheet.getRange(2, 1, tasks.length, headers.length).setValues(tasks);

  // Formatting
  sheet.setFrozenRows(1);
  headers.forEach((_, i) => sheet.setColumnWidth(i + 1, i === 1 ? 250 : 150));

  // Data validation for Frequency
  const freqRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Daily', 'Every 7 days', 'Every 15 days', 'Every 30 days'])
    .setAllowInvalid(false).build();
  sheet.getRange(2, 4, 100, 1).setDataValidation(freqRule);

  // Data validation for Active
  const activeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No']).setAllowInvalid(false).build();
  sheet.getRange(2, 8, 100, 1).setDataValidation(activeRule);

  // Data validation for Proof Type
  const proofRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Ticket ID', 'Action Confirmation']).setAllowInvalid(false).build();
  sheet.getRange(2, 6, 100, 1).setDataValidation(proofRule);

  // Alternating row colors
  for (let i = 2; i <= tasks.length + 1; i++) {
    sheet.getRange(i, 1, 1, headers.length).setBackground(i % 2 === 0 ? '#e3f2fd' : '#ffffff');
  }

  sheet.getRange(1, 1, tasks.length + 1, headers.length)
    .setBorder(true, true, true, true, true, true);
}

/* -------------------------------------------------------
   TASK LOG SHEET  (Primary Employee Working Area)
   ------------------------------------------------------- */
function createTaskLogSheet(ss) {
  let sheet = ss.getSheetByName('Task Log');
  if (!sheet) sheet = ss.insertSheet('Task Log');
  sheet.clear();

  const headers = [
    'Log ID', 'Task ID', 'Task Name', 'Assigned To',
    'Cycle #', 'Due Date', 'Status', 'Completed By',
    'Completed On', 'Ticket ID / Proof', 'Notes',
    'Manager Alerted', 'Manager Ack', 'Manager Ack Time',
    'Escalated', 'Escalation Time'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1b5e20').setFontColor('white')
    .setHorizontalAlignment('center').setWrap(true);

  // Set column widths
  const widths = [140, 80, 220, 120, 70, 110, 100, 120, 150, 150, 200, 120, 110, 150, 90, 150];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.setFrozenRows(1);

  // Data validation for Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Overdue', 'Completed']).setAllowInvalid(false).build();
  sheet.getRange(2, 7, 500, 1).setDataValidation(statusRule);

  // Data validation for Manager Ack
  const ackRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No']).setAllowInvalid(false).build();
  sheet.getRange(2, 13, 500, 1).setDataValidation(ackRule);

  // Conditional formatting — Overdue rows in red
  const overdueRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Overdue')
    .setBackground('#ffcdd2').setFontColor('#b71c1c')
    .setRanges([sheet.getRange('G2:G500')]).build();

  // Pending rows in yellow
  const pendingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending')
    .setBackground('#fff9c4').setFontColor('#f57f17')
    .setRanges([sheet.getRange('G2:G500')]).build();

  // Completed rows in green
  const completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Completed')
    .setBackground('#c8e6c9').setFontColor('#1b5e20')
    .setRanges([sheet.getRange('G2:G500')]).build();

  sheet.setConditionalFormatRules([overdueRule, pendingRule, completedRule]);

  // Add instruction note at top
  sheet.getRange('R1').setValue('📋 EMPLOYEE INSTRUCTIONS').setFontWeight('bold').setFontSize(12);
  sheet.getRange('R2').setValue('1. Find your name in "Assigned To" column');
  sheet.getRange('R3').setValue('2. Look for rows with Status = "Pending"');
  sheet.getRange('R4').setValue('3. For portal tasks: Enter your Ticket ID in column J');
  sheet.getRange('R5').setValue('4. For daily tasks: Type "Done" in column K (Notes)');
  sheet.getRange('R6').setValue('5. Status will auto-update to "Completed"');
  sheet.getRange('R7').setValue('6. Next cycle will be generated automatically');
  sheet.setColumnWidth(18, 350);
}

/* -------------------------------------------------------
   MANAGER DASHBOARD
   ------------------------------------------------------- */
function createManagerDashboard(ss) {
  let sheet = ss.getSheetByName('Manager Dashboard');
  if (!sheet) sheet = ss.insertSheet('Manager Dashboard');
  sheet.clear();

  // Title
  sheet.getRange('A1').setValue('📊 MANAGER DASHBOARD').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setValue('Last Refreshed:').setFontWeight('bold');
  sheet.getRange('B2').setFormula('=NOW()').setNumberFormat('yyyy-mm-dd hh:mm');

  // Summary headers
  const summaryHeaders = ['Employee', 'Pending', 'Overdue', 'Completed', 'Total', 'Status'];
  sheet.getRange(4, 1, 1, summaryHeaders.length).setValues([summaryHeaders])
    .setFontWeight('bold').setBackground('#e65100').setFontColor('white');

  sheet.setColumnWidths(1, 6, 150);
  sheet.setFrozenRows(4);

  // Add refresh button instruction
  sheet.getRange('H4').setValue('▶️ Run "refreshManagerDashboard" from Script Editor to update')
    .setFontColor('#666666');

  // Add manager acknowledgement instructions
  sheet.getRange('A' + 20).setValue('📌 MANAGER ACTIONS').setFontWeight('bold').setFontSize(12);
  sheet.getRange('A' + 21).setValue('After receiving an alert email:');
  sheet.getRange('A' + 22).setValue('1. Go to "Task Log" sheet');
  sheet.getRange('A' + 23).setValue('2. Find the alerted task row');
  sheet.getRange('A' + 24).setValue('3. Set "Manager Ack" column to "Yes"');
  sheet.getRange('A' + 25).setValue('4. Follow up with the employee to resolve the task');
  sheet.getRange('A' + 26).setValue('5. If not resolved within SLA, the system will escalate to leadership');
}

/* -------------------------------------------------------
   ESCALATION LOG
   ------------------------------------------------------- */
function createEscalationLogSheet(ss) {
  let sheet = ss.getSheetByName('Escalation Log');
  if (!sheet) sheet = ss.insertSheet('Escalation Log');
  sheet.clear();

  const headers = [
    'Escalation Time', 'Log ID', 'Task Name', 'Assigned To',
    'Original Due Date', 'Overdue Duration', 'Manager Ack Status', 'Action Taken'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#b71c1c').setFontColor('white');

  const widths = [180, 140, 220, 130, 130, 130, 140, 250];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  sheet.setFrozenRows(1);
}

/* -------------------------------------------------------
   EMPLOYEE FILTER VIEWS (per employee)
   Creates named filter views so each employee can quickly
   see only their own pending tasks.
   ------------------------------------------------------- */
function createEmployeeDashboards(ss) {
  const logSheet = ss.getSheetByName('Task Log');
  if (!logSheet) return;

  // We'll create a helper sheet with FILTER formulas per employee
  let sheet = ss.getSheetByName('My Pending Tasks');
  if (!sheet) sheet = ss.insertSheet('My Pending Tasks');
  sheet.clear();

  sheet.getRange('A1').setValue('👤 MY PENDING TASKS').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A3').setValue('Enter your name:').setFontWeight('bold');
  sheet.getRange('B3').setValue('Employee A');

  // Data validation for employee name
  const empRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Employee A', 'Employee B', 'Employee C', 'Employee D'])
    .setAllowInvalid(false).build();
  sheet.getRange('B3').setDataValidation(empRule).setBackground('#fff9c4').setFontWeight('bold');

  // Filter formula headers
  const headers = ['Task Name', 'Due Date', 'Status', 'Ticket ID / Proof', 'Notes', 'Cycle #'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1565c0').setFontColor('white');

  // FILTER formula that shows only this employee's pending/overdue tasks
  sheet.getRange('A6').setFormula(
    '=IFERROR(FILTER(\'Task Log\'!C2:C, \'Task Log\'!D2:D=B3, ' +
    'OR(\'Task Log\'!G2:G="Pending", \'Task Log\'!G2:G="Overdue")), "No pending tasks 🎉")'
  );
  sheet.getRange('B6').setFormula(
    '=IFERROR(FILTER(\'Task Log\'!F2:F, \'Task Log\'!D2:D=B3, ' +
    'OR(\'Task Log\'!G2:G="Pending", \'Task Log\'!G2:G="Overdue")), "")'
  );
  sheet.getRange('C6').setFormula(
    '=IFERROR(FILTER(\'Task Log\'!G2:G, \'Task Log\'!D2:D=B3, ' +
    'OR(\'Task Log\'!G2:G="Pending", \'Task Log\'!G2:G="Overdue")), "")'
  );
  sheet.getRange('D6').setFormula(
    '=IFERROR(FILTER(\'Task Log\'!J2:J, \'Task Log\'!D2:D=B3, ' +
    'OR(\'Task Log\'!G2:G="Pending", \'Task Log\'!G2:G="Overdue")), "")'
  );
  sheet.getRange('E6').setFormula(
    '=IFERROR(FILTER(\'Task Log\'!K2:K, \'Task Log\'!D2:D=B3, ' +
    'OR(\'Task Log\'!G2:G="Pending", \'Task Log\'!G2:G="Overdue")), "")'
  );
  sheet.getRange('F6').setFormula(
    '=IFERROR(FILTER(\'Task Log\'!E2:E, \'Task Log\'!D2:D=B3, ' +
    'OR(\'Task Log\'!G2:G="Pending", \'Task Log\'!G2:G="Overdue")), "")'
  );

  headers.forEach((_, i) => sheet.setColumnWidth(i + 1, i === 0 ? 250 : 150));

  sheet.getRange('A8').setValue('💡 TIP: Go to "Task Log" sheet to enter your Ticket ID or mark "Done"')
    .setFontColor('#1565c0');
  sheet.getRange('A9').setValue('Your completion will be recorded automatically — no need to update anywhere else.')
    .setFontColor('#666666');
}
