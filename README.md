# Allo Innoware - Recurring Task Tracker and Escalation System

## Complete Documentation and Submission

---

## Table of Contents

1. [System Overview](#1-system-overview)
2. [Google Sheet Structure](#2-google-sheet-structure)
3. [Automation and Apps Script Logic](#3-automation-and-apps-script-logic)
4. [Daily Task Completion Method](#4-daily-task-completion-tracking)
5. [Manager Alerts and Leadership Escalation (SLA)](#5-manager-alerts-and-leadership-escalation-sla)
6. [Test Cases and Working Demonstrations](#6-test-cases-and-working-demonstrations)
7. [Setup Instructions](#7-setup-instructions)

---

## 1. System Overview

This system converts Allo Innoware's recurring operational task management into a fully automated Google Sheets workflow. The design principles are:

| Principle | Implementation |
|-----------|---------------|
| **Zero-calculation for employees** | Due dates and cycles auto-generate |
| **Single-point update** | Employee enters Ticket ID or "Done" in ONE place |
| **No duplicate work** | Deduplication guards on cycles, alerts, and escalations |
| **Automated accountability** | Manager alerts, Manager acknowledgement, Leadership escalation |
| **Audit trail** | Every action is timestamped and logged |

### Flow Diagram

```
Task Becomes Due (auto-generated)
        |
        v
Employee Sees Pending Task (Task Log / My Pending Tasks)
        |
        |-- Portal Task? --> Enter Ticket ID in "Ticket ID / Proof" column
        |
        +-- Daily Task? --> Type "Done" in "Notes" column
                |
                v
        System Auto-Completes (Status -> Completed, timestamps filled)
                |
                v
        Next Cycle Auto-Generated
                |
                v
        +-- If NOT completed by due date --+
        |                                   |
        v                                   v
Status -> "Overdue"               Manager Alert Email (6 PM)
                                          |
                                          v
                                Manager Acknowledges (sets "Yes" in sheet)
                                          |
                                +-- If still unresolved --+
                                |   after SLA (48 hrs)    |
                                v                         v
                        Leadership Escalation       Logged in Escalation Log
```

---

## 2. Google Sheet Structure

### Sheet 1: Config

| Setting | Default Value | Purpose |
|---------|--------------|---------|
| Manager Email | manager@example.com | Receives daily pending task alerts |
| Leadership Email | leadership@example.com | Receives escalation notifications |
| Manager Ack SLA (Hours) | 24 | Hours manager has to acknowledge an alert |
| Escalation Hours | 48 | Hours after due date before leadership is notified |
| Working Hours Start | 9 | Business hours start |
| Working Hours End | 18 | Business hours end |

### Sheet 2: Task Master

The master registry of all recurring tasks.

| Column | Description | Example |
|--------|-------------|---------|
| Task ID | Unique identifier | T001, T002, etc. |
| Task Name | Descriptive task name | "Daily Operations Update" |
| Assigned To | Employee name | Employee A |
| Frequency | Recurrence pattern | Daily, Every 7 days, etc. |
| Cycle Days | Numeric frequency | 1, 7, 15, 30 |
| Proof Type | What proves completion | "Ticket ID" or "Action Confirmation" |
| Start Date | When tracking begins | 2026-10-12 |
| Active | Is the task active? | Yes / No |

**Pre-loaded Dummy Tasks:**

| Task ID | Task Name | Assigned To | Frequency | Proof Type |
|---------|-----------|-------------|-----------|------------|
| T001 | Daily Operations Update | Employee A | Daily | Action Confirmation |
| T002 | Portal Task A | Employee B | Every 7 days | Ticket ID |
| T003 | Portal Task B | Employee C | Every 15 days | Ticket ID |
| T004 | Portal Task C | Employee D | Every 30 days | Ticket ID |
| T005 | Weekly Status Report | Employee A | Every 7 days | Ticket ID |
| T006 | Daily Client Check-in | Employee B | Daily | Action Confirmation |

### Sheet 3: Task Log (PRIMARY EMPLOYEE WORKSPACE)

> [!IMPORTANT]
> This is the **single operating location** for employees. They view their pending tasks AND enter completion proof here -- nowhere else.

| Column | Who Fills It | Description |
|--------|-------------|-------------|
| Log ID | System | Auto-generated unique ID |
| Task ID | System | Links to Task Master |
| Task Name | System | Auto-populated |
| Assigned To | System | Auto-populated |
| Cycle # | System | Increments with each recurrence |
| Due Date | System | Auto-calculated |
| Status | System (auto) | Pending, Overdue, Completed |
| Completed By | System (auto) | Filled on completion |
| Completed On | System (auto) | Timestamp |
| **Ticket ID / Proof** | **Employee** | Employee enters Ticket ID here |
| **Notes** | **Employee** | Employee types "Done" here (daily tasks) |
| Manager Alerted | System | "Yes" after alert email sent |
| Manager Ack | **Manager** | Manager sets to "Yes" to acknowledge |
| Manager Ack Time | System (auto) | Timestamp of acknowledgement |
| Escalated | System | "Yes" after leadership escalation |
| Escalation Time | System (auto) | Timestamp |

**Color Coding (Conditional Formatting):**
- Yellow = Pending
- Red = Overdue
- Green = Completed

### Sheet 4: My Pending Tasks

A read-only filtered view for employees. They select their name from a dropdown, and the sheet shows only their pending/overdue tasks via FILTER formulas.

> [!NOTE]
> This sheet is for **viewing only**. Employees go to "Task Log" to enter their completion proof. This design keeps the process simple while giving employees a quick personal dashboard.

### Sheet 5: Manager Dashboard

Aggregated view showing:
- Per-employee summary (Pending / Overdue / Completed / Total counts)
- Status indicators (ACTION NEEDED / Monitor / All Clear)
- Detailed list of all currently pending/overdue items
- Manager action instructions

### Sheet 6: Escalation Log

Permanent audit trail of all leadership escalations with timestamps, overdue duration, and manager acknowledgement status.

---

## 3. Automation and Apps Script Logic

### 3.1 Time-Based Triggers (Automated)

| Trigger | Schedule | Function | Purpose |
|---------|----------|----------|---------|
| Cycle Generation | Daily at 6:00 AM | generateTaskCycles() | Creates new task cycle rows for due tasks |
| Manager Alert | Daily at 6:00 PM | sendManagerAlerts() | Emails manager about pending/overdue items |
| Escalation Check | Daily at 9:00 AM | escalateToLeadership() | Escalates unresolved items past SLA |

### 3.2 Event-Based Triggers (On Edit)

| Trigger | Event | Function | Purpose |
|---------|-------|----------|---------|
| Employee Completion | Edit in Ticket ID or Notes column | onEdit() | Auto-completes task, generates next cycle |
| Manager Ack | Manager sets "Manager Ack" to "Yes" | onEditManagerAck() | Captures acknowledgement timestamp |

### 3.3 Cycle Generation Logic

```
FOR each active task in Task Master:
    IF no open (Pending/Overdue) cycle exists for this task:
        IF first-ever cycle:
            Create cycle #1 with due_date = start_date (or today)
        ELSE:
            Create cycle #(last + 1) with due_date = last_due + cycle_days
            For daily tasks: due_date = today
        
        IMPORTANT: Only ONE open cycle per task at any time
                   --> Prevents duplicate rows (Test 8)
```

### 3.4 Completion Auto-Detection Logic

```
WHEN employee edits Ticket ID or Notes column in Task Log:
    IF current status is NOT "Completed":
        IF Ticket ID is non-empty OR Notes is non-empty:
            SET Status = "Completed"
            SET Completed By = Assigned To (from same row)
            SET Completed On = NOW()
            TRIGGER generateTaskCycles() --> creates next cycle
```

### 3.5 Deduplication Guards (Test 8)

| Guard | Where | How |
|-------|-------|-----|
| No duplicate cycles | generateTaskCycles() | Skips task if an open (Pending/Overdue) row already exists |
| No duplicate alerts | sendManagerAlerts() | Checks Manager Alerted flag before sending; marks "Yes" after send |
| No duplicate escalations | escalateToLeadership() | Checks Escalated flag before sending; marks "Yes" after send |
| No duplicate triggers | setupTriggers() | Deletes all existing triggers before creating new ones |

---

## 4. Daily Task Completion Tracking

> [!IMPORTANT]
> **Design Decision**: For daily tasks (like sharing an update in a WhatsApp group) that do not naturally generate a Ticket ID, I propose the following method:

### Method: "Done" Confirmation in Notes Column

**What the employee does:**
1. Complete the daily action (e.g., share update in WhatsApp group)
2. Go to the Task Log sheet
3. Find their pending daily task row
4. Type **"Done"** (or any brief note like "Done - shared client status at 3 PM") in the **Notes** column

**How the system records it:**
- The onEdit trigger detects the change in the Notes column
- System automatically sets:
  - Status to "Completed"
  - Completed By to employee name (from the same row)
  - Completed On to current timestamp (serves as proof of when they confirmed)
- The next day's cycle is auto-generated by the morning trigger

**Why this is the most practical approach:**

| Alternative Considered | Why Rejected |
|----------------------|--------------|
| Checkbox (TRUE/FALSE) | Too easy to accidentally check; no room for context |
| Separate Google Form | Adds another system for employees to manage |
| WhatsApp API integration | Over-engineered; requires external service |
| Email confirmation | Adds friction; yet another email in inbox |
| **Notes column (chosen)** | **Employee is already in the Task Log viewing their tasks. One cell edit = done. They can add optional context like "Shared revenue update at 3:15 PM" which serves as a soft audit trail.** |

**Minimum effort for employee:** The employee is already looking at the Task Log to see what is pending. They type "Done" in the adjacent cell. That is it -- one action, zero navigation.

---

## 5. Manager Alerts and Leadership Escalation (SLA)

### 5.1 Timing / SLA Logic (Test 7)

```
TIMELINE OF ESCALATION:

Day 0 (Due Date)
|-- 6:00 AM -- Task cycle generated (Status: Pending)
|-- During day -- Employee should complete the task
|-- 6:00 PM -- If not completed:
|             --> Manager receives email alert
|             --> Manager Alerted = "Yes"
|
Day 1 (Due + 24 hours)
|-- 9:00 AM -- Escalation check runs
|             --> If Manager has NOT acknowledged (Ack != "Yes")
|               AND 24+ hours since due date:
|               --> Leadership escalation triggered
|
Day 2 (Due + 48 hours)
|-- 9:00 AM -- Escalation check runs
|             --> If task STILL not completed
|               AND 48+ hours since due date:
|               --> Leadership escalation triggered
|               (regardless of manager acknowledgement)
```

### 5.2 Escalation Rules

| Rule | Trigger Condition | Action |
|------|-------------------|--------|
| **Manager Alert** | Task is Pending/Overdue on due date at 6 PM | Email to Manager |
| **Escalation (no ack)** | Manager has NOT acknowledged within 24 hours | Email to Leadership (CC Manager) |
| **Escalation (still open)** | Task still not completed 48 hours after due date | Email to Leadership (CC Manager) |
| **No repeat alerts** | Already alerted for this cycle | Skip (flag checked) |
| **No repeat escalations** | Already escalated for this cycle | Skip (flag checked) |

### 5.3 Reasoning Behind SLA Design

**Why 6 PM for manager alert?**
- Gives the employee the full working day to complete the task
- Manager sees the alert at end-of-day or next morning
- Avoids premature alerts that would create alert fatigue

**Why 24 hours for manager acknowledgement SLA?**
- Manager may be in meetings during the day
- 24 hours gives them a reasonable window to review and follow up
- If manager has not even looked at it in 24 hours, leadership should know

**Why 48 hours for resolution escalation?**
- Even if manager acknowledged, if the task is STILL not done after 2 full business days, it is a systemic issue
- Leadership needs visibility into chronic non-completion
- 48 hours is aggressive enough to maintain accountability but not so tight as to be unreasonable

**How are unnecessary repeat emails avoided?**
- Each task log row has Manager Alerted and Escalated flags
- Once set to "Yes", the system never sends another email for the same cycle
- A new cycle starts fresh with all flags reset to "No"

### 5.4 Manager Action After Receiving Alert

1. Manager receives email with list of pending/overdue tasks
2. Manager opens the Google Sheet (link in email)
3. Goes to Task Log sheet
4. Finds the alerted row(s)
5. Sets Manager Ack column to **"Yes"** -- system auto-captures timestamp
6. Follows up with the employee (Slack, call, etc.)
7. If the employee completes the task, Status auto-updates to "Completed"
8. If not resolved within 48 hours, the system escalates to leadership

---

## 6. Test Cases and Working Demonstrations

### Test 1 -- Recurring Cycle

**Scenario:** Complete a 7-day task and verify the next cycle auto-generates.

**Steps:**
1. Run generateTaskCycles() -- creates initial Pending rows
2. Find the "Portal Task A" (7-day) row in Task Log
3. Enter Ticket ID "TKT-7001" in the Ticket ID column
4. System auto-completes: Status to Completed, timestamps filled
5. generateTaskCycles() runs again and creates Cycle #2 with due date = original due + 7 days

**Verification:**
- Task Log shows Cycle 1 as "Completed" with Ticket ID
- Task Log shows Cycle 2 as "Pending" with new due date

**Run:** Execute test1_RecurringCycle() from Script Editor.

---

### Test 2 -- Single-Point Completion

**Scenario:** Employee sees a pending task and updates it from one location only.

**Steps:**
1. Employee goes to Task Log sheet
2. Finds their Pending row
3. Types Ticket ID (e.g., "TKT-1234") in column J
4. onEdit trigger fires automatically
5. Status, Completed By, and Completed On all auto-fill

**Verification:**
- Employee did NOT go to any other sheet to update
- No manual copy-paste between sheets
- All metadata captured from a single cell edit

**Run:** This is an interactive test -- manually edit a cell in Task Log.

---

### Test 3 -- Daily Task

**Scenario:** Employee A completes the "Daily Operations Update" (WhatsApp group share).

**Steps:**
1. Employee A shares update in WhatsApp group
2. Opens Task Log, finds today's "Daily Operations Update" row
3. Types "Done -- Update shared in WhatsApp group" in the Notes column
4. System auto-completes the task
5. Next day at 6 AM, a new Pending row appears for the next day

**Verification:**
- No Ticket ID required for daily tasks
- "Done" note serves as completion proof
- Timestamp captures when the employee confirmed

**Run:** Execute test3_DailyTask() from Script Editor.

---

### Test 4 -- Same-Day Manager Alert

**Scenario:** Employee B does not complete "Portal Task A" by end of day.

**Steps:**
1. A Pending task for Employee B exists with today's due date
2. Employee B does NOT enter a Ticket ID during the day
3. At 6:00 PM, sendManagerAlerts() runs
4. Manager receives email: "Task Alert: 1 Pending/Overdue Items"
5. Email contains: Task name, Employee name, Due date, Status

**Verification:**
- Manager received exactly ONE email (not one per task)
- Email is consolidated with all pending items in a table
- Manager Alerted column is set to "Yes" for those rows

**Run:** Execute test4_ManagerAlert() from Script Editor.

---

### Test 5 -- Manager Acknowledgement

**Scenario:** Manager reviews the alert and acknowledges it.

**Steps:**
1. Manager opens Google Sheet from email link
2. Goes to Task Log sheet
3. Finds the row where Manager Alerted = "Yes"
4. Sets Manager Ack column to "Yes"
5. onEditManagerAck() trigger fires
6. Manager Ack Time auto-populates with current timestamp

**Verification:**
- System knows the manager has seen and acknowledged the issue
- Timestamp proves when the acknowledgement happened
- This resets the escalation timer (leadership will not be notified if task is resolved within SLA)

**Run:** This is an interactive test -- manually set Manager Ack to "Yes".

---

### Test 6 -- Leadership Escalation

**Scenario:** Task remains unresolved 48+ hours after due date.

**Steps:**
1. Task is overdue, manager was alerted
2. Manager either did NOT acknowledge, OR acknowledged but task is still not completed
3. 48 hours pass
4. escalateToLeadership() runs at 9:00 AM
5. Leadership receives email: "ESCALATION: 1 Unresolved Tasks"
6. Email shows task details, hours overdue, and manager ack status
7. Entry added to Escalation Log sheet

**Verification:**
- Leadership email contains all relevant details
- Manager is CC'd on the escalation email
- Escalation Log has a permanent record
- Escalated column is set to "Yes" preventing re-escalation

**Run:** Execute test6_Escalation() from Script Editor.

---

### Test 7 -- Escalation Timing Logic

**SLA Design (detailed above in Section 5):**

| Event | When | Condition |
|-------|------|-----------|
| Task becomes due | 6:00 AM daily trigger | Auto-generated for active tasks |
| Manager alert | 6:00 PM same day | Task still Pending/Overdue |
| Manager ack expected | Within 24 hours | Manager sets "Yes" in sheet |
| Leadership escalation (no ack) | 9:00 AM next day | Manager has not acknowledged in 24h |
| Leadership escalation (unresolved) | 9:00 AM day after | Task still open after 48h |

---

### Test 8 -- No Duplicate Alerts or Work

**Scenario:** Run all automation functions multiple times.

**Steps:**
1. Run generateTaskCycles() twice consecutively
2. Verify no duplicate task rows created (guard: open-cycle check)
3. Run sendManagerAlerts() twice consecutively
4. Verify only one email sent (guard: Manager Alerted flag)
5. Run escalateToLeadership() twice consecutively
6. Verify only one escalation (guard: Escalated flag)

**Verification:**
- Row count in Task Log remains the same after second run
- No duplicate emails in manager's inbox
- No duplicate entries in Escalation Log

**Run:** Execute test8_NoDuplicates() from Script Editor.

---

## 7. Setup Instructions

### Step 1: Create a New Google Sheet
1. Go to sheets.google.com and create a new blank spreadsheet
2. Name it: **"Allo Innoware - Recurring Task Tracker"**

### Step 2: Open Apps Script Editor
1. In the Google Sheet, go to **Extensions then Apps Script**
2. Delete the default Code.gs content

### Step 3: Add the Scripts
1. Paste the contents of Code.gs into the script editor
2. Click **+ then Script** to create a new file, name it SheetSetup
3. Paste the contents of SheetSetup.gs

### Step 4: Initialize the Sheets
1. In the script editor, select setupAllSheets from the function dropdown
2. Click **Run**
3. Grant permissions when prompted (the script needs access to your spreadsheet and email)
4. All 6 sheets will be created with proper formatting and dummy data

### Step 5: Generate Initial Task Cycles
1. Select generateTaskCycles from the dropdown
2. Click **Run**
3. Check the Task Log sheet -- you should see Pending rows for each task

### Step 6: Set Up Automated Triggers
1. Select setupTriggers from the dropdown
2. Click **Run**
3. This creates the daily scheduled triggers and edit-based triggers

### Step 7: Update Configuration
1. Go to the Config sheet
2. Replace manager@example.com with the actual manager's email
3. Replace leadership@example.com with the leadership team's email
4. Adjust SLA hours if needed

### Step 8: Test the System
1. Run runAllTests() to execute all automated test cases
2. Check the Execution Log (View then Execution log) for test results
3. Perform the interactive tests (Tests 2 and 5) manually

---

## Files Included

| File | Purpose |
|------|---------|
| Code.gs | Main automation script -- all business logic |
| SheetSetup.gs | One-time sheet structure creator |
| README.md | This documentation file |

---

*Built for Allo Innoware -- AI Operations Manager Assignment*
