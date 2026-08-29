/**
 * Sri Narasu's Coffee Company Pvt Ltd
 * Stores Department - Staff Working Register, KPI & HR Appraisal System
 * Backend: Google Apps Script (API layer) + Google Sheets (database)
 *
 * ROLE MODEL (audit item 24 - updated from the older "2 roles only" note
 * below, which was left stale after the HR role was added): 3 roles ->
 * Manager, Staff, HR.
 *   Staff  : create/submit own Working Register, Attendance, Leave; view own
 *            KPI/Score/HR Appraisal only - never another staff member's.
 *   Manager: everything Staff can do + view all staff, Approve/Reject
 *            Working Register entries, manage Master data (Staff/Workflow/
 *            Monthly Target), view Reports/Dashboard/HR Summary for the
 *            whole department - but CANNOT edit HR Appraisal data itself
 *            (Manual Scores, Recognition, Memo) - see requireHR() below.
 *   HR     : owns the HR Appraisal module - Manual Scores, Recognition,
 *            Memo/Disciplinary lifecycle, HR Management Summary, HR
 *            Dashboard - and can view (read-only) everything Manager can
 *            view for context, but cannot approve Working Register entries
 *            or edit Masters (that stays Manager-only).
 * Every role boundary above is enforced server-side (requireManager() /
 * requireHR() / the session-based staffId scoping in handle()) - the
 * frontend hiding a button is convenience only, never the real gate.
 *
 * Deploy as a Web App (Execute as: Me, Access: Anyone with the link)
 * doPost/doGet both route through the same JSON API dispatcher.
 */

// ============================================================
// CONFIG
// ============================================================
const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEETS = {
  COMPANY: 'Company Master',
  SECTION: 'Section Master',
  STAFF: 'Staff Master',
  ACTIVITY: 'Activity Master',
  WORKFLOW: 'Workflow Master',
  REGISTER: 'Working Register',
  ATTENDANCE: 'Attendance Register',
  LEAVE: 'Leave Register',
  USERS: 'Users',
  AUDIT: 'Audit Log',
  TARGETS: 'Monthly Targets',
  // ---- HR KPI Appraisal module (8-criteria / 100-mark manual+auto system) ----
  RECOGNITION: 'Recognition Register',       // bonus marks (Employee of the Month etc.)
  DISCIPLINARY: 'Disciplinary Register',     // memo / warning deductions
  MANUAL_SCORES: 'HR Manual Scores',         // manager-entered scores for the 6 manual criteria
  APPRAISAL_HISTORY: 'HR Appraisal History',  // saved per-period snapshots, used for trend charts
  // ---- Automatic Training Score module (Aug 2026) ----
  TRAINING_MASTER: 'Training_Master',     // catalogue of trainings - single source of truth, HR-managed
  TRAINING_REGISTER: 'Training_Register',  // per-staff training completion records - HR-managed, drives the automatic Training Score
  // ---- Notification module (Aug 2026 audit fix - P0) ----
  NOTIFICATIONS: 'Notification Register'  // in-app bell notifications; each row targets either a specific To User ID or an entire To Role
};

// Single source of truth for valid Attendance Status values - used by:
//   1. markAttendance() to reject anything else with a clear error
//   2. buildSheet()'s dropdown on the Attendance Register sheet (setup())
//   3. applyAttendanceStatusValidation() for re-applying the dropdown on an
//      existing spreadsheet (see migrateAttendanceStatusDropdown() below)
// Must always match the <select id="attStatus"> options in index.html.
const ATTENDANCE_STATUS_VALUES = ['Present', 'Absent', 'Half Day', 'Permission', 'Late'];

// Workflow Master fields editable ONLY through the Edit Workflow screen
// (api.updateWorkflow) or the onEdit guardrail below - never meant to be
// hand-edited in the sheet without going through one of those two paths.
const EDITABLE_WORKFLOW_FIELDS = ['Target', 'Weightage %', 'Max Score', 'Expected Output', 'Status'];

// ============================================================
// HR KPI APPRAISAL POLICY (edit these to match your real company policy)
// ============================================================
// The 100-mark appraisal = 2 AUTO criteria (computed from existing Attendance
// Register + Working Register data) + 6 MANUAL criteria (entered by the
// Manager per staff per period, 0-10 each):
//   AUTO   : Attendance & Punctuality (10), Work Performance & Productivity (30)
//   MANUAL : Quality & Accuracy (10), Teamwork & Communication (10),
//            Discipline & Compliance (10), Initiative & Problem Solving (10),
//            Training & Skill Development (10), Customer/Internal Service &
//            Professional Behaviour (10)
// ------------------------------------------------------------------
// HR APPRAISAL v2 (Aug 2026): pure HR/discipline/attendance module,
// fully separated from Stores work-performance (Working Register KPI).
// "Work Performance & Productivity" and the old Quality/Teamwork/
// Initiative/Service manual criteria were REMOVED from this module
// because they duplicated the Stores KPI Module - Stores performance
// is combined in separately via FINAL_SCORE_WEIGHTS below, not scored
// twice. The 100 marks are now 11 pure-HR KPIs; Overtime Hours is
// tracked/reported but does NOT contribute marks (info-only, optional).
// Every KPI here is entered monthly by HR on the HR Appraisal screen
// (Working Days, Present Days, Leave Days, Permission Hours, Late
// Entries, Early Leaving Count, Shift Compliance %, Overtime Hours,
// Discipline/Policy/Behaviour/Training ratings 0-10, HR Remarks 0-5).
// ------------------------------------------------------------------
const HR_APPRAISAL_POLICY = {
  WEIGHTS: {           // must sum to 100 - edit to match real company policy
    attendance: 15, leaveManagement: 8, permissionHours: 7, lateComing: 8,
    earlyLeaving: 7, shiftCompliance: 10, discipline: 10, policyCompliance: 10,
    behaviour: 10, training: 10, hrRemarks: 5
  },
  RATING_MAX: { discipline: 10, policyCompliance: 10, behaviour: 10, training: 10, hrRemarks: 5 },
  // Monthly allowances before a penalty kicks in, and the penalty per
  // excess unit (in percentage points of that KPI's Achievement %).
  SANCTIONED_LEAVE_DAYS: 1, LEAVE_PENALTY_PER_EXCESS_DAY: 20,
  ALLOWED_PERMISSION_HOURS: 2, PERMISSION_PENALTY_PER_EXCESS_HOUR: 10,
  LATE_COMING_PENALTY_PER_ENTRY: 10,
  EARLY_LEAVING_PENALTY_PER_ENTRY: 10,
  // Combined Final Employee Score = Stores KPI % x storesWeight + HR Score x hrWeight
  FINAL_SCORE_WEIGHTS: { storesKpi: 70, hrAppraisal: 30 },
  GRADE_BANDS: [ // combined Final Score, highest first
    { min: 90, grade: 'A+' }, { min: 80, grade: 'A' }, { min: 70, grade: 'B' },
    { min: 60, grade: 'C' }, { min: 0, grade: 'D' }
  ],
  // HR Appraisal's own 0-100 score rating bands (6-band, per HR spec)
  RATING_BANDS: [
    { min: 90, rating: 'Outstanding' }, { min: 80, rating: 'Excellent' },
    { min: 70, rating: 'Good' }, { min: 60, rating: 'Satisfactory' },
    { min: 50, rating: 'Needs Improvement' }, { min: 0, rating: 'Unsatisfactory' }
  ],
  // ------------------------------------------------------------------
  // AUTOMATED WORKFLOW KPI RATING (per-KPI Achievement % -> qualitative
  // Rating + star Score) - DISPLAY/AUDIT LABEL ONLY. The numeric KPI
  // Contribution actually applied to the score always stays the existing
  // continuous "capped Achievement % x Weightage" formula (see
  // computeKpiScore()) - this table never changes that arithmetic, it only
  // gives each KPI row a human-readable Rating alongside its Contribution
  // (see getMyScore()/resolveWorkflowKpiRating()), same as RATING_BANDS/
  // GRADE_BANDS above. Edit freely; highest min first.
  WORKFLOW_KPI_RATING_BANDS: [
    { min: 100, rating: 'Excellent', ratingScore: 5 },
    { min: 90, rating: 'Very Good', ratingScore: 4 },
    { min: 80, rating: 'Good', ratingScore: 3 },
    { min: 70, rating: 'Average', ratingScore: 2 },
    { min: 0, rating: 'Needs Improvement', ratingScore: 1 }
  ],
  WORKFLOW_KPI_MAX_RATING_SCORE: 5,
  // Over-achievement control (>100% Achievement): the % actually applied to
  // the KPI Contribution formula is capped here before the Weightage is
  // applied, so a huge overachievement can never blow past this workflow's
  // own max possible contribution. Configurable, not hard-coded - raise this
  // only if company policy explicitly wants over-achievement credit.
  MAX_ACHIEVEMENT_PCT_FOR_SCORING: 100
};

// ============================================================
// AUTOMATIC TRAINING SCORE POLICY (Aug 2026)
// ------------------------------------------------------------------
// Replaces the old manual 0-10 "Training Rating" HR typed in on the
// Manual Scores screen. Training is now driven entirely by
// Training_Master + Training_Register (see calculateTrainingScore_(),
// calculateTrainingEligibility_(), calculateTrainingExpiry_(),
// getTrainingAppraisalScore_() below) and feeds into computeHRScores()
// via the SAME 'Training Rating' field name (0-10 scale) it always
// used - computeHRScores() itself is intentionally left unchanged;
// only the value now comes from an automatic calculation instead of
// an HR-typed number. HR_APPRAISAL_POLICY.WEIGHTS.training (10) and
// RATING_MAX.training (10) stay as the scoring cap - unchanged.
// ------------------------------------------------------------------
const TRAINING_POLICY = {
  // Per-record score bands, applied to Assessment Score (0-100) of an
  // ELIGIBLE record only. Ineligible/invalid records always score 0.
  SCORE_BANDS: [ // highest min first
    { min: 80, score: 10 },
    { min: 60, score: 8 },
    { min: 40, score: 6 },
    { min: 0, score: 0 }
  ],
  MAX_TRAINING_SCORE: 10, // final per-staff-per-period Training Score cap (sum of eligible records, capped here)
  ATTENDANCE_OK: ['Present'],
  COMPLETION_OK: ['Completed']
};

const HR_POLICY = {
  // LEGACY CLEANUP (audit item 25): the old 8-criteria HR model's
  // MANUAL_CRITERIA_MAX / ATTENDANCE_MAX / PRODUCTIVITY_MAX constants were
  // removed here - searched every file, confirmed zero references left
  // anywhere (they were superseded by HR_APPRAISAL_POLICY.WEIGHTS /
  // RATING_MAX above when the module moved to the 11-pure-HR-KPI model).
  // Bonus marks for Special Recognition. "Special Appreciation" has a spec'd
  // range (+2 to +5); the value actually applied is whatever the Manager
  // enters on addRecognition (capped to this range), the rest below are fixed.
  RECOGNITION_BONUS: {
    'Employee of the Month': 5,
    'Appreciation Letter': 3,
    'Special Appreciation': 3,   // default within the +2..+5 range - overridable
    'Outstanding Achievement': 5,
    'Management Recognition': 3
  },
  SPECIAL_APPRECIATION_RANGE: [2, 5],
  // Deduction marks per Memo/Disciplinary Action type - edit to match company policy.
  MEMO_DEDUCTION: {
    'Verbal Warning': 2,
    'Written Warning': 5,
    'Final Warning': 10,
    'Disciplinary Memo': 8
  },
  // Memo categories (what the issue was actually about) - independent of the
  // Action Type above (which drives the deduction/severity). Edit freely.
  MEMO_CATEGORIES: [
    'Late Coming', 'Poor Performance', 'Absenteeism', 'Safety Violation',
    'Process Violation', 'Misconduct', 'Quality Issue', 'Customer Complaint', 'Others'
  ],
  // Repeated-memo escalation: if a staff member receives thresholdCount or
  // more memos in the SAME Category within one appraisal Period, an extra
  // one-time deduction is added on top of the individual memo deductions
  // (applied per repeated category, in getFullAppraisal). Set thresholdCount
  // very high (e.g. 999) to effectively disable this rule.
  MEMO_REPEAT_RULE: { thresholdCount: 3, extraDeduction: 5 },
  // Final Decision options for a memo, and the Effective Deduction policy
  // applied for each (see decideMemo()) - edit freely, but decideMemo()'s
  // if/else and MEMO_DECISION_NO_PENALTY below must be updated to match if
  // you rename or add an option.
  // MEMO -> REPLY -> HR REVIEW -> FINAL DECISION flow: every option here is
  // the FINAL DECISION step only. Issuing a memo, the employee replying, or
  // HR reviewing that reply NEVER by themselves change Effective Deduction
  // (see addDisciplinaryAction/replyToMemo/reviewMemo) - only decideMemo()
  // (this step) sets it, exactly once, from whichever of these options HR
  // actually confirms.
  MEMO_DECISION_OPTIONS: [
    'Reply Accepted / Satisfactory', 'Warning Only / No Penalty',
    'Memo Closed / No Action', 'Memo Withdrawn / Cancelled',
    'Penalty Confirmed', 'Deduction Adjusted'
  ],
  // Every decision above that resolves to "0 - no penalty". Also lists the
  // OLD (pre-rename) option labels so historical memo rows decided before
  // this rename still resolve/display/exclude-from-repeat-escalation
  // identically - old Closed records are never rewritten (see migration
  // notes), only read consistently.
  MEMO_DECISION_NO_PENALTY: [
    'Reply Accepted / Satisfactory', 'Warning Only / No Penalty',
    'Memo Closed / No Action', 'Memo Withdrawn / Cancelled',
    'Warning Withdrawn', 'Closed - No Action' // legacy labels, pre-rename
  ],
  // Thresholds used to build the HR Management Summary buckets.
  PROMOTION_MIN_SCORE: 85,
  INCREMENT_MIN_SCORE: 80,
  TRAINING_BELOW_SCORE: 70,     // Final Score below this -> recommend training
  IMPROVEMENT_BELOW_SCORE: 70,  // Final Score below this -> "Needs Improvement" bucket
  TOP_PERFORMER_MIN_SCORE: 90
};

// Specific WORKFLOWS (not whole Activities) where several staff work the SAME
// physical truck / bill together. Only a workflow whose ID is listed here
// shows the Co-Staff checklist and gets its score/weightage split; every other
// workflow under the same Activity is treated as an individual (solo) task.
// The staff member who fills the form ("Logged By") ticks the Co-Staff who
// also worked on it; the backend auto-generates a shared "Team Ref No" and
// creates one Working Register row per participant (see addWorkingEntries).
// See withTeamSplit() below for how the shared Team Ref No is used to divide
// score/weightage among participants at read time.
//
// EDIT THIS LIST to match your real process - add or remove Workflow IDs
// (see the Workflow Master sheet / seedWorkflows() below for the full list
// and what each ID corresponds to). This starter list is a guess based on
// which tasks are physically done together (Unloading, Picking, Packing,
// Loading, Handover) - it is NOT verified against your actual practice.
const TEAM_WORKFLOW_IDS = [
  'WF0001', // ACT001 REC-PACK - Unloading & Bill Confirmation
  'WF0005', // ACT001 REC-PACK - Weight / QC / Packing Sample Report
  'WF0006', // ACT001 REC-PACK - Bin & Lot Marking
  'WF0009', // ACT002 REC-NONPACK - Unloading
  'WF0015', // ACT002 REC-NONPACK - Material Delivery
  'WF0017', // ACT002 REC-NONPACK - Bin / Lot Movement
  'WF0018', // ACT003 BR-DESPATCH - Branch / Distributor / Super Stockist
  'WF0019', // ACT003 BR-DESPATCH - Picking
  'WF0020', // ACT003 BR-DESPATCH - Checking
  'WF0021', // ACT003 BR-DESPATCH - Packing
  'WF0023', // ACT003 BR-DESPATCH - Handover to Despatch
  'WF0026', // ACT004 HO-DESPATCH - Section Selection
  'WF0027', // ACT004 HO-DESPATCH - Picking
  'WF0029', // ACT004 HO-DESPATCH - Material Delivery
  'WF0032', // ACT005 DESPATCH-PACKMAT - Sales Indent Preparation
  'WF0033', // ACT005 DESPATCH-PACKMAT - Picking
  'WF0034', // ACT005 DESPATCH-PACKMAT - Arrangement
  'WF0040', // ACT006 DESPATCH-JOBWORK - Checking
  'WF0041'  // ACT006 DESPATCH-JOBWORK - Truck Loading
  // WF0039 (DESPATCH-JOBWORK - Picking) reverted to Individual per the
  // final complete Staff/Co-Staff table.
];

// ============================================================
// NOTE: Truck No is no longer scoped to a fixed list of "Receiving"
// Activities. It is scoped WORKFLOW-WISE via TEAM_WORKFLOW_IDS above -
// any line whose Workflow ID is in TEAM_WORKFLOW_IDS can generate/carry
// a Truck No, for ANY Activity (REC-PACK, REC-NONPACK, BR-DESPATCH,
// HO-DESPATCH, DESPATCH-PACKMAT, DESPATCH-JOBWORK, ...). See
// generateTruckNo() / addWorkingEntries(). (The separate legacy Truck
// Tracker module - getTruckTracker/getPendingTrucksByActivity/
// assignTruckWorkflow/getTruckAssignmentHistory - was removed in the Aug
// 2026 audit: it was dead, unreferenced by the frontend, and duplicated
// tracking that Working Register + Truck No already provide.)
// ============================================================

// ============================================================
// ACHIEVEMENT % / KPI SCORE - single source of truth
// ============================================================
// Every place in the system that turns an Achieved/Approved Qty into an
// Achievement % or a KPI Score MUST go through these two functions, so
// there is exactly one formula in the whole app and the required
// guarantees always hold:
//   - Achievement % can only increase or stay the same when Achieved/
//     Approved Qty increases (Target unchanged) - plain division, nothing
//     else feeds into it.
//   - Remaining Target, Team Split, Co-Staff entries, or workflow count
//     never enter this calculation, so none of them can reduce it.
//   - KPI Score is Achievement % (capped at 100%) applied to Weightage %,
//     and is itself capped at Weightage % - never exceeds its own ceiling.
// TARGET = 0 AMBIGUITY FIX (P1): a Target of 0 can mean two different
// things - "nobody has configured a target for this workflow/month yet"
// (should never silently score as if 0% or 100% were achieved) or "a
// Manager deliberately set the target to 0" (equally meaningless to divide
// by). Both cases are business-undefined, so BOTH now return null
// ("N/A") instead of the old 0 - a real 0% (target set, nothing achieved)
// only ever comes from computeAchievementPct(0, target>0) i.e. an actual
// Actual/Approved Qty of 0 against a real positive target. Callers must
// treat a null return as "N/A" for display and as "not yet measurable"
// for scoring (write/aggregate as 0 contribution, never as a penalty).
function computeAchievementPct(achievedQty, targetQty) {
  const t = Number(targetQty) || 0;
  const a = Number(achievedQty) || 0;
  if (t <= 0) return null; // N/A - target not set OR explicitly zero, never divide by it
  return Math.round((a / t) * 10000) / 100; // NOT capped here - see computeKpiScore for the 100% cap
}
function computeKpiScore(achievementPct, weightagePct) {
  if (achievementPct === null || achievementPct === undefined) return null; // N/A propagates - no score can be derived
  const w = Number(weightagePct) || 0;
  // OVER-ACHIEVEMENT CONTROL: capped at HR_APPRAISAL_POLICY.MAX_ACHIEVEMENT_PCT_FOR_SCORING
  // (100 by default = no overachievement credit) - configurable, see policy.
  const capped = Math.min(Number(achievementPct) || 0, HR_APPRAISAL_POLICY.MAX_ACHIEVEMENT_PCT_FOR_SCORING);
  return Math.round((capped * w / 100) * 100) / 100; // = MIN((Achieved/Target) x Weightage, Weightage x cap/100)
}

// ============================================================
// PASSWORD SECURITY (audit item 15)
// ============================================================
// PREVIOUSLY: the Users sheet stored the real password in plain text in the
// 'Password' column, and login()/changePassword() compared it with a plain
// '==='. Anyone with view access to the spreadsheet (or a copy/backup/export
// of it) could read every user's real password outright - passwords never
// belonged in a readable column at all.
// NOW: 'Password Hash' + 'Salt' columns hold, respectively,
// base64(SHA-256(salt + password)) and a random per-user salt - the actual
// password is never written anywhere (not the sheet, not Logger, not the
// Audit Log, not any API response). hashPassword() is the only place that
// combines a plaintext password with a digest; nothing else ever sees it
// after login()/changePassword() finish executing.
// SAFE MIGRATION: existing rows created before this fix still have their old
// plaintext value sitting in 'Password' and an empty 'Password Hash'. login()
// below detects that case (hash columns empty) for ONLY that one user, and if
// the plaintext still matches, hashes it right then, saves the hash+salt,
// and blanks out the plaintext column - all inside the same request, so nyone/
// nothing else has to change and no existing user's password is invalidated.
// A user who never logs in again keeps their old plaintext row until they do
// - unavoidable without either resetting everyone's password by force or
// reading it without their input, both worse outcomes - but every active
// login migrates that row, so the plaintext footprint only shrinks over time.
function makeSalt() {
  return Utilities.getUuid().replace(/-/g, '');
}
function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + String(password));
  return Utilities.base64Encode(bytes);
}
// Constant-shape comparison (both inputs are always same-length base64 SHA-
// 256 digests here) - not strictly required in Apps Script's single-threaded
// execution model, but avoids relying on '===' short-circuiting as the only
// guard between "password checking" and "password checking safely".
function hashesMatch(a, b) {
  if (!a || !b || a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= (a.charCodeAt(i) ^ b.charCodeAt(i));
  return diff === 0;
}

// ============================================================
// LOGIN LOCKOUT (audit fix - brute-force protection)
// ============================================================
// PREVIOUSLY: login() had no memory of failed attempts at all - a wrong
// password only ever produced a LOGIN_FAILED audit line, so a User ID could
// be guessed against indefinitely with no slowdown or lockout.
// NOW: every failed attempt increments 'Failed Login Count' on that user's
// Users row and stamps 'Last Failed Login'. Once the count reaches
// MAX_FAILED_ATTEMPTS, 'Locked Until' is set LOCKOUT_MINUTES into the
// future and every login attempt (even a correct password) is rejected
// with a generic locked-out message until that time passes. A SUCCESSFUL
// login always clears Failed Login Count back to 0 and blanks Locked
// Until - the counter is "attempts since the last success", not lifetime.
// Requires the 'Failed Login Count' / 'Locked Until' / 'Last Failed Login'
// columns on Users - run migrateAddLoginLockoutFields() once to add them to
// an existing spreadsheet (see that function for details).
const LOGIN_LOCKOUT_POLICY = {
  MAX_FAILED_ATTEMPTS: 5,
  LOCKOUT_MINUTES: 10
};

// ============================================================
// ENTRY POINTS
// ============================================================
function doGet(e) {
  return handle(e);
}
function doPost(e) {
  return handle(e);
}

// Adds a "HR Appraisal" menu to the spreadsheet's own UI (Sheets menu bar)
// so a Manager can open the appraisal panel without any external app -
// runs automatically whenever the spreadsheet is opened in a browser.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('HR Appraisal')
    .addItem('Open Appraisal Panel', 'showHRAppraisalSidebar')
    .addToUi();
}

function showHRAppraisalSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('HRAppraisalSidebar')
    .setTitle('HR KPI Appraisal')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// GUARDRAIL: Workflow Master editable-field edits need confirmation
// Covers: Target, Weightage %, Max Score, Expected Output, Status
// (EDITABLE_WORKFLOW_FIELDS) - the same fields the in-app Edit Workflow
// screen (api.updateWorkflow) manages. This is a backstop for anyone who
// edits the sheet directly instead of using that screen. Handles
// single-cell AND bulk (multi-row paste/drag) edits. Simple trigger -
// runs automatically, no separate deployment needed.
// ============================================================
function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEETS.WORKFLOW) return;

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const workflowIdCol = headers.indexOf('Workflow ID') + 1;
  const workflowNameCol = headers.indexOf('Workflow Name') + 1;
  if (!workflowIdCol) return;

  const editStartCol = e.range.getColumn();
  const editEndCol = editStartCol + e.range.getNumColumns() - 1;
  const editedFieldCols = EDITABLE_WORKFLOW_FIELDS
    .map(f => ({ field: f, col: headers.indexOf(f) + 1 }))
    .filter(fc => fc.col && fc.col >= editStartCol && fc.col <= editEndCol);
  if (editedFieldCols.length === 0) return; // none of the guarded fields were touched

  const startRow = e.range.getRow();
  const numRows = e.range.getNumRows();
  const snapshot = getWorkflowFieldSnapshot();
  const changes = [];

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row === 1) continue; // header row
    const wfId = sheet.getRange(row, workflowIdCol).getValue();
    if (!wfId) continue; // blank row
    const wfName = workflowNameCol ? sheet.getRange(row, workflowNameCol).getValue() : wfId;
    editedFieldCols.forEach(fc => {
      const newVal = sheet.getRange(row, fc.col).getValue();
      const key = wfId + '__' + fc.field;
      const oldVal = snapshot[key] !== undefined ? snapshot[key] : '(unknown)';
      if (String(oldVal) !== String(newVal)) {
        changes.push({ row: row, col: fc.col, workflowId: wfId, workflowName: wfName, field: fc.field, oldVal: oldVal, newVal: newVal, key: key });
      }
    });
  }
  if (changes.length === 0) return;

  const ui = SpreadsheetApp.getUi();
  let msg = changes.length + ' Workflow Master மாற்றம்(கள்):\n\n';
  changes.forEach(c => { msg += c.workflowName + ' (' + c.workflowId + ') — ' + c.field + ': ' + c.oldVal + ' → ' + c.newVal + '\n'; });
  msg += '\nஇந்த மதிப்புகள் KPI Score calculation-ஐ (Working Register, Dashboard, Appraisal) direct-ஆ affect பண்ணும். ' +
    'Edit Workflow screen-ல் பண்ணுவது தான் பரிந்துரைக்கப்படுகிறது. இப்படியே தொடரலாமா?';

  const resp = ui.alert('Confirm Workflow Master changes', msg, ui.ButtonSet.YES_NO);

  if (resp !== ui.Button.YES) {
    // PRE-LIVE FIX #3 (audit): a change whose baseline was never captured in
    // the snapshot (oldVal === '(unknown)' - e.g. a brand-new workflow that
    // was never touched via updateWorkflow()/refreshWorkflowFieldSnapshotAfterApiEdit_)
    // must NEVER be reverted to '' on "No" - that would silently blank out
    // Target/Weightage/Status and corrupt live KPI scoring. Only revert
    // changes we actually have a confirmed prior value for; leave unresolved
    // ones exactly as the user typed them and separately warn that they must
    // be fixed manually, instead of guessing.
    const knownChanges = changes.filter(c => c.oldVal !== '(unknown)');
    const unresolvedChanges = changes.filter(c => c.oldVal === '(unknown)');
    knownChanges.forEach(c => sheet.getRange(c.row, c.col).setValue(c.oldVal));
    if (unresolvedChanges.length > 0) {
      let warnMsg = 'இந்த மாற்றங்களுக்கு பழைய மதிப்பு (baseline) system-ல் இல்லாததால், தானாக revert செய்ய முடியவில்லை. ' +
        'கீழே உள்ளவை தற்போதைய (புதிய) மதிப்பிலேயே உள்ளன — தயவுசெய்து கைமுறையாக சரிபார்த்து சரி செய்யவும்:\n\n';
      unresolvedChanges.forEach(c => { warnMsg += c.workflowName + ' (' + c.workflowId + ') — ' + c.field + ': ' + c.newVal + '\n'; });
      ui.alert('கைமுறை திருத்தம் தேவை (Manual correction needed)', warnMsg, ui.ButtonSet.OK);
    }
    return;
  }

  const userEmail = Session.getActiveUser().getEmail() || 'unknown';
  changes.forEach(c => {
    snapshot[c.key] = c.newVal;
    logAudit(userEmail, 'WORKFLOW_FIELD_CHANGED',
      'Workflow ' + c.workflowId + ' (' + c.workflowName + ') | ' + c.field + ': "' + c.oldVal + '" -> "' + c.newVal + '" (direct sheet edit)');
  });
  saveWorkflowFieldSnapshot(snapshot);
}

function getWorkflowFieldSnapshot() {
  const raw = PropertiesService.getScriptProperties().getProperty('WORKFLOW_FIELD_SNAPSHOT');
  return raw ? JSON.parse(raw) : buildWorkflowFieldSnapshotFromSheet();
}
function saveWorkflowFieldSnapshot(snapshot) {
  PropertiesService.getScriptProperties().setProperty('WORKFLOW_FIELD_SNAPSHOT', JSON.stringify(snapshot));
}
function buildWorkflowFieldSnapshotFromSheet() {
  const snap = {};
  readAll(SHEETS.WORKFLOW).forEach(r => {
    EDITABLE_WORKFLOW_FIELDS.forEach(f => { snap[r['Workflow ID'] + '__' + f] = r[f]; });
  });
  saveWorkflowFieldSnapshot(snap);
  return snap;
}
// Call this once (from api.updateWorkflow's caller side is automatic; this
// manual entry point is only needed if you ever bypass updateWorkflow) to
// keep the guardrail's snapshot in sync after an in-app edit.
function refreshWorkflowFieldSnapshotAfterApiEdit(workflowId, field, newVal) {
  const snapshot = getWorkflowFieldSnapshot();
  snapshot[workflowId + '__' + field] = newVal;
  saveWorkflowFieldSnapshot(snapshot);
}

// ============================================================
// SIDEBAR WRAPPERS
// google.script.run can only call top-level functions (not api.xxx methods
// directly), so these thin wrappers expose exactly what the sidebar needs.
// Every write action still goes through the same api.* function, the same
// requireManager() check, and the same Audit Log entry as the JSON API.
// ============================================================
function ui_login(userId, password) {
  return api.login({ userId: userId, password: password });
}

// Companion to resolveUiActor_() below - the sidebar previously only ever
// cleared its own local SESSION variable on logout, leaving the real
// server-side session token alive in CacheService until it naturally
// expired. Explicitly destroying it server-side closes that window.
function ui_logout(token) {
  if (token) destroySession(token);
  return { success: true };
}

// SIDEBAR AUTH FIX (post-live audit): these five reference-data getters
// used to take no token and never called resolveUiActor_(), unlike every
// other read path in the system (the JSON API requires a valid session for
// every action except login, and every other ui_* wrapper below already
// resolves+verifies a token). ui_getStaffList() in particular returned the
// full active staff roster (name/ID/designation) to any caller able to
// invoke google.script.run against this sidebar, logged in or not - a
// real inconsistency with the app's zero-trust pattern even though actually
// reaching the sidebar already requires Sheet access. Each now requires
// params.token and resolves it the same way resolveUiActor_() does,
// throwing before any data is returned if the session is missing/expired.
function requireUiSession_(token) {
  const session = resolveSession(token);
  if (!session) {
    throw new Error('Session expired. Please log in again.');
  }
  return session;
}

function ui_getStaffList(params) {
  requireUiSession_(params && params.token);
  return readAll(SHEETS.STAFF).filter(s => s['Status'] === 'Active')
    .map(s => ({ staffId: s['Staff ID'], staffName: s['Staff Name'], designation: s['Designation'] }));
}

// TRUSTED-INTERNAL HELPERS (getMasters() runtime fix): the two functions
// below contain the actual Recognition/Memo type lookup logic and require
// no session/token of their own. They exist so that getMasters() - which
// already runs behind handle()'s central session check (see PUBLIC_ACTIONS)
// - can read this same data without needing to fabricate or forward a
// google.script.run-style token that only the Sidebar's ui_* wrappers use.
// The public ui_getRecognitionTypes()/ui_getMemoTypes() wrappers below
// remain the only way to reach this data from the Sidebar, and still
// enforce requireUiSession_() exactly as before - nothing about their
// authentication changed, they just delegate the actual lookup here.
function _getRecognitionTypesInternal_() {
  return Object.keys(HR_POLICY.RECOGNITION_BONUS).map(type => ({
    type: type, bonus: HR_POLICY.RECOGNITION_BONUS[type],
    isRange: type === 'Special Appreciation', range: HR_POLICY.SPECIAL_APPRECIATION_RANGE
  }));
}

function _getMemoTypesInternal_() {
  return Object.keys(HR_POLICY.MEMO_DEDUCTION).map(type => ({ type: type, deduction: HR_POLICY.MEMO_DEDUCTION[type] }));
}

function ui_getRecognitionTypes(params) {
  requireUiSession_(params && params.token);
  return _getRecognitionTypesInternal_();
}

function ui_getMemoTypes(params) {
  requireUiSession_(params && params.token);
  return _getMemoTypesInternal_();
}

// SINGLE SOURCE OF TRUTH (audit item 18): added so the Sidebar (and any
// other caller) never needs its own hardcoded copy of HR_POLICY.MEMO_
// CATEGORIES - previously the Sidebar had no getter for this and kept a
// manually-synced duplicate list with a comment asking whoever edits
// HR_POLICY to remember to edit the Sidebar too.
function ui_getMemoCategories(params) {
  requireUiSession_(params && params.token);
  return HR_POLICY.MEMO_CATEGORIES.slice();
}

// NOT session-gated on purpose: called from the Sidebar's window.onload,
// before login, to prefill the Period Label field on the login screen
// itself. It returns nothing but the current calendar month/year string -
// no staff or business data - so gating it would break that pre-login
// prefill for no security benefit.
function ui_currentMonthLabel() {
  return currentMonthLabel();
}

// SIDEBAR AUTH FIX (pre-live audit, P1): every ui_* wrapper below used to
// call api.xxx(params) with params.actorRole/actorUserId taken straight
// from the SIDEBAR'S OWN CLIENT-SIDE `SESSION` object - unlike handle()
// (the JSON API used by index.html), nothing here ever verified who was
// actually calling. Since google.script.run functions are directly
// invocable from the browser console of anyone who can open this sidebar,
// a caller could set params = {actorRole:'HR', actorUserId:'anything', ...}
// by hand and reach requireHR()-gated actions (submitManualScores,
// addRecognition, addDisciplinaryAction, saveAppraisalSnapshot,
// lockAppraisalSnapshot, reopenAppraisalSnapshot, the HR dashboards) with
// zero real authentication - the Sidebar's own login screen was entirely
// bypassable.
//
// Fix: reuse the EXACT SAME session store the JSON API already trusts
// (createSession()/resolveSession() - ui_login() below already returns a
// real token via api.login()). Every params-taking wrapper now requires
// params.token, resolves it server-side, and OVERWRITES
// actorRole/actorUserId/actorStaffId/actorStaffName from the verified
// session - any client-supplied actorRole/actorUserId in params is
// discarded, exactly as handle() already does for the JSON API. A
// missing/expired/forged token throws before api.xxx() is ever reached.
function resolveUiActor_(params) {
  params = params || {};
  const session = resolveSession(params.token);
  if (!session) {
    throw new Error('Session expired. Please log in again.');
  }
  params.actorRole = session.role;
  params.actorUserId = session.userId;
  params.actorStaffId = session.staffId;
  params.actorStaffName = session.staffName;
  return params;
}

function ui_submitManualScores(params) {
  return api.submitManualScores(resolveUiActor_(params));
}

function ui_addRecognition(params) {
  return api.addRecognition(resolveUiActor_(params));
}

function ui_addDisciplinaryAction(params) {
  return api.addDisciplinaryAction(resolveUiActor_(params));
}

function ui_getFullAppraisal(params) {
  const actor = resolveUiActor_(params);
  // SIDEBAR AUTH FIX (audit finding, Aug 2026): api.getFullAppraisal() has no
  // role/scope check of its own - the JSON API path (index.html -> handle())
  // is safe only because STAFF_SCOPED_READ_ACTIONS force-sets p.staffId to
  // the session's own staffId whenever the trusted role is 'Staff'. This
  // Sidebar wrapper never applied that same rule - ui_login() allows ANY
  // role (including Staff) to log into the Sidebar, and resolveUiActor_()
  // above only verifies identity, it does not scope staffId. A Staff user
  // calling ui_getFullAppraisal with no staffId therefore got the WHOLE
  // department's appraisal data (combinedFinalScore, HR remarks, memos,
  // criteria breakdown) - the exact class of leak already fixed for
  // getAppraisal()/getPendingApprovals() on the JSON API side. Mirror the
  // same force-set rule here so the Sidebar path can never be more
  // permissive than the JSON API for the same data.
  if (actor.actorRole === 'Staff') {
    actor.staffId = actor.actorStaffId;
  }
  return api.getFullAppraisal(actor);
}

function ui_saveAppraisalSnapshot(params) {
  return api.saveAppraisalSnapshot(resolveUiActor_(params));
}
function ui_lockAppraisalSnapshot(params) {
  return api.lockAppraisalSnapshot(resolveUiActor_(params));
}
function ui_reopenAppraisalSnapshot(params) {
  return api.reopenAppraisalSnapshot(resolveUiActor_(params));
}

function ui_getHRManagementSummary(params) {
  return api.getHRManagementSummary(resolveUiActor_(params));
}

function ui_getHRDashboard(params) {
  return api.getHRDashboard(resolveUiActor_(params));
}

function ui_getHRLiveDashboard(params) {
  return api.getHRLiveDashboard(resolveUiActor_(params));
}

// ============================================================
// SESSION / SERVER-SIDE AUTHORIZATION
// ============================================================
// SECURITY FIX (audit item 4): the client used to send {actorRole,
// actorUserId, staffId, ...} as plain JSON params, and every api.*
// function trusted them directly - a Staff user could edit those values
// in the browser/devtools/any HTTP client and act as Manager/HR, or as
// a different staff member. That is now impossible: login() issues an
// opaque, server-only session token; every OTHER action must present it;
// handle() looks the token up in CacheService (server-side, never
// visible to or derivable by the client) and OVERWRITES whatever
// actorRole/actorUserId/actorStaffId/actorStaffName the client sent with
// the server-verified values before the action runs. For actions a Staff
// user should only ever perform on their OWN record (see SELF_SCOPED_
// ACTIONS below), the trusted staffId/staffName also overwrite the
// client-supplied ones - a Staff user can no longer submit/edit/view
// data under someone else's Staff ID by changing the request body.
// Manager/HR are left free to pass an explicit target staffId (e.g. to
// view another staff's appraisal); that path is still separately gated
// by requireManager()/requireHR() inside the relevant api.* function,
// now using the trusted actorRole so it can't be spoofed either.
// SLIDING SESSION, EXTENDED IDLE TIMEOUT (audit follow-up, Aug 2026):
// CacheService's OWN entries cannot outlive 6 hours no matter how often you
// re-put them with a fresh TTL - that's a hard Apps Script platform limit
// (SESSION_TTL_SECONDS below), not a choice. For a normal working day
// (~8-10h) where staff genuinely go quiet for a few hours (e.g. a long
// floor shift with no system use), that 6h cap alone would still log them
// out even though the sliding refresh is in place.
//
// Fix: keep CacheService as the FAST path (checked first on every request),
// but back every session with a slower, non-expiring PropertiesService
// entry that we manage by hand using our OWN idle clock
// (IDLE_TIMEOUT_SECONDS, configurable independently of the 6h cache cap).
// If the cache entry has fallen off (idle 6h+ but still under
// IDLE_TIMEOUT_SECONDS), resolveSession() transparently revives it from
// PropertiesService and re-seeds the cache - the user never sees a
// "Session expired" in between, right up until IDLE_TIMEOUT_SECONDS of
// TRUE inactivity.
const SESSION_TTL_SECONDS = 6 * 60 * 60;          // 6h - CacheService hard max, fast path only
const IDLE_TIMEOUT_SECONDS = 10 * 60 * 60;        // 10h - real idle-logout threshold (tune as needed)
const PROPS_WRITE_THROTTLE_SECONDS = 5 * 60;      // only refresh the PropertiesService backup every 5 min of active use, not on every single request (quota + latency)

function createSession(user) {
  const token = Utilities.getUuid();
  const payload = JSON.stringify({
    userId: user['User ID'], staffId: user['Staff ID'],
    staffName: user['Staff Name'], role: user['Role']
  });
  CacheService.getScriptCache().put('sess_' + token, payload, SESSION_TTL_SECONDS);
  // Slow-path backup, keyed the same way, plus our own idle clock. Stored
  // as {payload, lastActivity} so resolveSession() can decide staleness
  // itself instead of relying on any platform-side expiry.
  PropertiesService.getScriptProperties().setProperty('sess_' + token,
    JSON.stringify({ payload: payload, lastActivity: Date.now() }));
  return token;
}

function resolveSession(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();
  const key = 'sess_' + token;

  const cached = cache.get(key);
  if (cached) {
    // Fast path hit - refresh the cache TTL on EVERY request (cheap, no
    // quota concern). The PropertiesService backup only needs to stay
    // roughly current, not exact-to-the-second - writing it on every single
    // request (including the 60s notification poll, multiplied across every
    // active staff member) needlessly burns PropertiesService's daily
    // read/write quota and adds latency to every call. Throttle using a
    // second, short-lived CacheService marker (cheap to check) instead of
    // reading PropertiesService itself on every request.
    cache.put(key, cached, SESSION_TTL_SECONDS);
    const throttleKey = 'proswrite_' + token;
    if (!cache.get(throttleKey)) {
      try {
        props.setProperty(key, JSON.stringify({ payload: cached, lastActivity: Date.now() }));
        cache.put(throttleKey, '1', PROPS_WRITE_THROTTLE_SECONDS);
      } catch (e) { /* best-effort; cache hit is enough to serve this request */ }
    }
    return JSON.parse(cached);
  }

  // Cache entry aged out (6h+ since last hit) - check the slow path before
  // declaring the session dead.
  const raw = props.getProperty(key);
  if (!raw) {
    // TRUE miss: neither CacheService nor PropertiesService has this token
    // - either genuinely wrong/stale, or from an old deployment.
    return null;
  }
  let rec;
  try { rec = JSON.parse(raw); } catch (e) { props.deleteProperty(key); return null; }
  const idleMs = Date.now() - (rec.lastActivity || 0);
  if (idleMs > IDLE_TIMEOUT_SECONDS * 1000) {
    props.deleteProperty(key); // truly stale - clean up so it doesn't linger forever
    return null;
  }
  // Still within the idle window - revive the fast-path cache entry and
  // refresh the idle clock, so the user's next request goes back to the
  // cheap cache-hit path.
  cache.put(key, rec.payload, SESSION_TTL_SECONDS);
  props.setProperty(key, JSON.stringify({ payload: rec.payload, lastActivity: Date.now() }));
  return JSON.parse(rec.payload);
}

function destroySession(token) {
  if (!token) return;
  CacheService.getScriptCache().remove('sess_' + token);
  try { PropertiesService.getScriptProperties().deleteProperty('sess_' + token); } catch (e) {}
}

// Housekeeping: ScriptProperties has a 500-key / 500KB total ceiling, so
// abandoned sessions (idle timeout reached, but the user never explicitly
// logged out to trigger destroySession()) must eventually be swept out
// rather than left forever. Run this from a time-driven trigger (e.g. daily,
// via Apps Script's Triggers UI: Edit > Current project's triggers > Add
// Trigger > cleanupExpiredSessions_ > Time-driven > Day timer). Safe to run
// as often as you like; it only ever removes entries already past
// IDLE_TIMEOUT_SECONDS.
function cleanupExpiredSessions_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  const now = Date.now();
  Object.keys(all).forEach(key => {
    if (key.indexOf('sess_') !== 0) return;
    try {
      const rec = JSON.parse(all[key]);
      if (now - (rec.lastActivity || 0) > IDLE_TIMEOUT_SECONDS * 1000) {
        props.deleteProperty(key);
      }
    } catch (e) {
      props.deleteProperty(key); // corrupt entry - discard
    }
  });
}

// Public wrapper for cleanupExpiredSessions_() - functions ending in "_"
// are treated as private by Apps Script, so they never appear in the
// Triggers page's "Select function to run" dropdown. Select THIS function
// (runCleanupExpiredSessions) when adding the daily time-driven trigger;
// it simply calls the private cleanup function above.
function runCleanupExpiredSessions() {
  cleanupExpiredSessions_();
}

// Actions where the ACTING user must equal the TARGET staffId - the
// trusted session's staffId/staffName always overwrite whatever the
// client sent for these, regardless of role, so nobody (including a
// Manager) accidentally/maliciously submits a working entry, attendance
// mark, or leave application "as" a different staff member.
const SELF_SCOPED_ACTIONS = {
  addWorkingEntries: true, markAttendance: true, applyLeave: true
};
// Read actions that take an optional p.staffId filter ("show me just this
// staff member's records"). For a Manager/HR that filter is legitimately
// optional (blank = everyone). For a Staff user it must NEVER be optional
// or spoofable - without this, a Staff user could either omit staffId to
// see every employee's data, or pass someone else's Staff ID to see theirs.
// handle() force-sets p.staffId = the trusted session staffId for these
// actions whenever the trusted role is 'Staff'.
const STAFF_SCOPED_READ_ACTIONS = {
  getWorkingRegister: true, getEntryGroups: true, getAttendance: true,
  getLeaves: true, getDashboard: true, getRecognitions: true,
  getDisciplinaryActions: true, getManualScores: true, getMyScore: true,
  getFullAppraisal: true,
  // Per-staff HR Appraisal trend (Aug 2026) - same rule: a Staff user's
  // p.staffId is force-set to their own session staffId here, so they can
  // only ever pull their own Combined Final Score trend, never another
  // employee's. Manager/HR pass an explicit staffId to see anyone's.
  getStaffAppraisalTrend: true,
  // REPORTS MODULE (Aug 2026): same rule as every other read action above -
  // a Staff user's p.staffId is force-set to their own session staffId here,
  // so a Staff can only ever pull these 8 reports scoped to themself, never
  // another employee's Approved Actual/KPI Score. Manager/HR pass staffId
  // blank (or any staffId) and see the full department, same as today.
  getStaffPerformanceReport: true, getActivityPerformanceReport: true,
  getKPIWorkflowPerformanceReport: true, getSectionPerformanceReport: true,
  getDateRangePerformanceReport: true, getMonthlyTargetAchievementReport: true,
  getStaffAppraisalReport: true, getApprovalWorkingRegisterReport: true,
  // AUTOMATIC TRAINING (Aug 2026): a Staff user may only ever see their own
  // Training Register history - same force-set-staffId rule as every other
  // read action in this list (item 15 of the spec: "Staff can view only
  // authorized training records").
  getTrainingRecords: true
};
// Actions that don't require a prior login (only the login call itself).
const PUBLIC_ACTIONS = { login: true };

// FORCED PASSWORD CHANGE ENFORCEMENT (server-side hardening): when a user's
// 'Must Change Password' flag is Yes, the ONLY actions handle() will still
// dispatch for them are the ones needed to actually clear that flag (or to
// log out/re-authenticate). Every other protected business API is blocked
// server-side, regardless of what the frontend does or doesn't show - the
// existing frontend modal (openChangePasswordModal) was already doing this
// visually, but a direct API call bypassed it. login stays unaffected since
// it's already in PUBLIC_ACTIONS above and runs before any of this.
const FORCED_PASSWORD_CHANGE_ALLOWED_ACTIONS = { changePassword: true, logout: true };

// Server-side source of truth for "does this user currently have to change
// their password" - always re-read from the Users sheet by the trusted
// session userId (never from params/session cache), so a flag flipped by an
// admin mid-session takes effect on this user's very next request, and a
// client can never claim mustChangePassword=false to skip the gate.
function userMustChangePassword(userId) {
  const u = readAll(SHEETS.USERS).find(x => x['User ID'] === userId);
  if (!u) return false; // unknown/removed user - downstream checks (e.g. session no longer valid) handle this
  return String(u['Must Change Password'] || '').trim().toLowerCase() === 'yes';
}

function handle(e) {
  let action = '', params = {};
  try {
    if (e.postData && e.postData.contents) {
      const body = JSON.parse(e.postData.contents);
      action = body.action;
      params = body.params || {};
    } else {
      action = e.parameter.action;
      params = e.parameter.params ? JSON.parse(e.parameter.params) : e.parameter;
    }
    // Internal-only helpers (e.g. _addWorkingEntriesLocked, only meant to be
    // called from inside addWorkingEntries while already holding the script
    // lock) must never be reachable directly through the API dispatcher -
    // that would let a client skip the locking/validation of the public
    // action that normally wraps them.
    if (!action || action.charAt(0) === '_' || typeof api[action] !== 'function') {
      return json({ ok: false, error: 'Unknown action: ' + action });
    }

    if (!PUBLIC_ACTIONS[action]) {
      const session = resolveSession(params.token);
      if (!session) {
        return json({ ok: false, error: 'Session expired. Please log in again.', sessionExpired: true });
      }
      // Server-verified identity - never trust the client's own copy of
      // these fields for authorization decisions from this point on.
      params.actorRole = session.role;
      params.actorUserId = session.userId;
      params.actorStaffId = session.staffId;
      params.actorStaffName = session.staffName;

      // FORCED PASSWORD CHANGE ENFORCEMENT (server-side) - decided purely
      // from the trusted session's userId + the Users sheet, never from
      // params.mustChangePassword or any other client-supplied value.
      if (!FORCED_PASSWORD_CHANGE_ALLOWED_ACTIONS[action] && userMustChangePassword(session.userId)) {
        try { logAudit(session.userId, 'BLOCKED_MUST_CHANGE_PASSWORD', 'action=' + action); } catch (logErr) {}
        return json({ ok: false, error: 'Password change required before using the application.', mustChangePassword: true });
      }

      if (SELF_SCOPED_ACTIONS[action]) {
        params.staffId = session.staffId;
        params.staffName = session.staffName;
      }
      if (STAFF_SCOPED_READ_ACTIONS[action] && session.role === 'Staff') {
        params.staffId = session.staffId;
      }
    }

    const result = api[action](params);
    return json({ ok: true, data: result });
  } catch (err) {
    // ERROR HANDLING (audit item 33): a deliberately-thrown validation/
    // authorization message (new Error('...'), name === 'Error') is safe and
    // useful to show as-is (e.g. "Duplicate entry...", "Not authorized...").
    // An unexpected runtime error (TypeError, ReferenceError, a Sheets API
    // exception, etc.) can leak internal details ("Cannot read properties of
    // undefined...", sheet/column names, function/stack info) that aren't
    // meaningful to a normal user and must NEVER reach the browser in
    // production - the full technical detail is only ever written to the
    // Apps Script Logger and the Audit Log. The API response itself carries
    // nothing beyond the safe, generic (or deliberately-thrown) `error`
    // message - no separate technical/debug field.
    const isExpected = err && err.name === 'Error';
    const userMessage = isExpected ? err.message : 'Unable to complete the request. Please verify the entered data and try again.';
    if (!isExpected) {
      Logger.log('Unexpected error in action "' + action + '": ' + err.message + '\n' + err.stack);
      try { logAudit(params && (params.actorUserId || params.staffId) || '(unknown)', 'SERVER_ERROR', action + ': ' + err.message); } catch (logErr) {}
    }
    return json({ ok: false, error: userMessage });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SHEET HELPERS
// ============================================================
function sh(name) {
  const s = SS.getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name);
  return s;
}

function readAll(name) {
  const s = sh(name);
  const values = s.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].join('') === '') continue;
    const obj = {};
    headers.forEach((h, idx) => obj[h] = values[i][idx]);
    obj._row = i + 1;
    rows.push(obj);
  }
  return rows;
}

// FORMULA-INJECTION GUARD (final live-audit finding, Aug 2026): every free-
// text field a Staff/Manager/HR ever types (Staff Name, Memo Reply, Notes,
// Activity/Section Name, Training remarks, etc.) flows through appendRow()/
// updateRow() below via Range.setValue(). Sheets treats any STRING value
// starting with =, +, -, or @ as a formula/expression when the sheet is
// opened directly (which this app's own onOpen() menu explicitly invites
// Manager/HR to do) - so an untrusted "=HYPERLINK(...)" or "=IMPORTXML(...)"
// typed into any text field would silently become a live formula, not
// stored text. Prefixing a straight apostrophe is the same trick Sheets'
// own UI uses to force literal text, and only ever touches actual JS
// strings - numbers/Dates/booleans pass through unchanged, so no legitimate
// numeric or date field is affected.
function sheetSafeValue_(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

function appendRow(name, obj) {
  const s = sh(name);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const row = headers.map(h => sheetSafeValue_(obj[h] !== undefined ? obj[h] : ''));
  s.appendRow(row);
  return s.getLastRow();
}

function updateRow(name, rowNum, obj) {
  const s = sh(name);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  headers.forEach((h, idx) => {
    if (obj[h] !== undefined) s.getRange(rowNum, idx + 1).setValue(sheetSafeValue_(obj[h]));
  });
}

// Case/whitespace-insensitive Period Label comparison. Period Labels
// ("Aug-2026") are free-text on both the HR page's Period field and every
// place that writes/filters by Period, so a stray space or different
// casing ("aug-2026", " Aug-2026 ") produced a byte-for-byte mismatch
// against the exact stored value and silently hid otherwise-correct rows
// from whichever screen was filtering by Period - see
// getDisciplinaryActions() for the reported case (a Staff's Employee Reply
// saved correctly, but didn't show up in HR's Memo Management table
// because the two Period strings didn't match exactly). This normalizes
// both sides before comparing; it never changes what's actually stored.
function normalizePeriodLabel(label) {
  // Sheets can silently auto-convert a plain-text Period label like
  // "Aug-2026" into a real Date value if the column isn't (or is no
  // longer, e.g. after a manual paste) forced to plain-text format (see
  // buildSheet()'s textColumnNames). String(dateObject) then produces
  // something like "Tue Aug 01 2026 00:00:00 GMT+0530..." which never
  // matches the typed label "aug-2026" - silently breaking every
  // downstream exact-match comparison (Recognition/Memo not reaching the
  // Scorecard's Bonus/Memo columns, etc). Reconstruct the same "Mon-YYYY"
  // text the app itself generates (see defaultPeriodLabel()) so a Period
  // stored as a Date still matches the typed label.
  if (label instanceof Date) {
    const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
    return months[label.getMonth()] + '-' + label.getFullYear();
  }
  return String(label || '').trim().toLowerCase();
}

// SINGLE SOURCE OF TRUTH for chronological ordering of free-text Period
// Labels ("Jul-2026") - used by getHRDashboard()'s monthlyTrend AND
// getStaffAppraisalTrend()'s per-staff monthlyTrend, so both trend charts
// sort the same way and can never drift out of sync with each other.
// Previously this was a local function nested inside getHRDashboard() only
// - hoisted here (Aug 2026) so the new per-staff trend endpoint can reuse
// it instead of duplicating the parsing/sort logic a second time.
// Unparseable labels sort last (Infinity), after every real month.
const PERIOD_MONTH_INDEX_ = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
function periodSortKey_(label) {
  const m = String(label).trim().match(/^([A-Za-z]{3})[a-z]*[-\s]?(\d{4})$/);
  if (m && PERIOD_MONTH_INDEX_.hasOwnProperty(m[1].toLowerCase())) {
    return Number(m[2]) * 12 + PERIOD_MONTH_INDEX_[m[1].toLowerCase()];
  }
  return Infinity;
}
// Stable comparator built on periodSortKey_() - chronological first, then
// alphabetical fallback for ties/unparseable labels so sort order never
// depends on insertion order (Object.keys() iteration).
function periodSortCompare_(a, b) {
  const ka = periodSortKey_(a), kb = periodSortKey_(b);
  if (ka !== kb) return ka - kb;
  return String(a).localeCompare(String(b));
}

// Same class of bug as normalizePeriodLabel() above, applied to Staff ID:
// Recognition/Disciplinary rows are entered via HR forms that can pick up a
// stray leading/trailing space (copy-paste, autocomplete) even though the
// value LOOKS identical on screen. A byte-for-byte '===' match against
// Staff Master's Staff ID then silently drops the row from
// getFullAppraisal()'s bonus/memo totals - the row still exists and still
// displays correctly in Recognition/Memo Management (which don't cross-
// match against Staff Master), so it looks like Bonus/Memo just "isn't
// applying" on the Scorecard. Normalize both sides before comparing; never
// changes what's actually stored.
function normalizeStaffId(id) {
  return String(id || '').trim().toUpperCase();
}

function nextId(prefix, existingIds) {
  let max = 0;
  existingIds.forEach(id => {
    const n = parseInt(String(id).replace(prefix, ''), 10);
    if (!isNaN(n) && n > max) max = n;
  });
  return prefix + String(max + 1).padStart(4, '0');
}

function deleteRow(name, rowNum) {
  sh(name).deleteRow(rowNum);
}

// Monthly Targets use a "MMM-yyyy" label (e.g. "Jul-2026"), same style as the
// spec's example sheet. This checks whether a Working Register row's actual
// Date value falls inside that labelled month.
const MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
function monthLabelOf(date) {
  const d = new Date(date);
  return MONTH_NAMES[d.getMonth()] + '-' + d.getFullYear();
}
function currentMonthLabel() {
  return monthLabelOf(new Date());
}
// Google Sheets sometimes auto-detects a plain-text label like "Jul-2026"
// typed/written into a cell and silently stores it as a real Date instead
// (the cell still *displays* "Jul-2026", so this is easy to miss). Any code
// comparing Month values must go through this so both cases match correctly.
function normalizeMonthLabel(val) {
  if (val instanceof Date) return MONTH_NAMES[val.getMonth()] + '-' + val.getFullYear();
  return String(val || '').trim();
}

// DATE RANGE FIX (audit item 22): a "To" date filter like "2026-08-08" used
// to be compared as new Date("2026-08-08"), i.e. midnight at the START of
// that day - any row whose Date/timestamp fell later in the day was wrongly
// excluded from the range. endOfDay() normalizes a "To" bound to 23:59:59.999
// of that calendar day so the whole selected end date is always included.
// Used everywhere a date range's upper bound is applied (Working Register,
// Attendance, Leave, HR Appraisal, Reports, Dashboard, KPI reports).
function endOfDay(dateVal) {
  const d = new Date(dateVal);
  d.setHours(23, 59, 59, 999);
  return d;
}
function startOfDay(dateVal) {
  const d = new Date(dateVal);
  d.setHours(0, 0, 0, 0);
  return d;
}

function logAudit(userId, action, details) {
  appendRow(SHEETS.AUDIT, {
    'Timestamp': new Date(),
    'User ID': userId,
    'Action': action,
    'Details': details
  });
}

// ============================================================
// NOTIFICATIONS (bell icon) - audit fix P0
// ============================================================
// Each row targets EITHER one specific user (toUserId) OR an entire role
// (toRole, e.g. every HR user) - never both. getNotifications() below does
// the server-side matching against the trusted session identity (actorUserId
// /actorRole from handle()) so a client can never read another user's
// notifications by guessing/forging an ID.
// Call sites so far: replyToMemo (Staff -> HR), decideMemo (HR -> Staff),
// approveEntryLine/approveEntryLines/approveEntryGroup on Rejected
// (Manager -> Staff). Add more call sites the same way as new events need
// to notify a role/user - never write directly to the sheet elsewhere.
function addNotification_(opts) {
  // opts: { toUserId (optional), toRole (optional - required if no toUserId),
  //         type, message, refId (optional) }
  if (!opts || (!opts.toUserId && !opts.toRole)) return; // nothing to target - fail silently, never block the calling action
  if (!opts.message) return;
  const existing = readAll(SHEETS.NOTIFICATIONS).map(r => r['Notification ID']);
  const id = nextId('NT', existing);
  appendRow(SHEETS.NOTIFICATIONS, {
    'Notification ID': id,
    'To User ID': opts.toUserId || '',
    'To Role': opts.toRole || '',
    'Type': opts.type || '',
    'Message': String(opts.message),
    'Ref ID': opts.refId || '',
    'Read': false,
    'Created On': new Date(),
    'Read On': ''
  });
  // EMAIL BACKEND (previously missing - audit item flagged Aug 2026): every
  // in-app bell notification also attempts a best-effort email to the same
  // target(s), through the single sendEmailNotification_() helper below.
  // Deliberately routed through THIS one function (never a second, separate
  // call site) so email can never drift out of sync with what the bell
  // shows, and a missing/quota-exceeded/misconfigured email can never
  // block or fail the notification (or the action that triggered it) - see
  // sendEmailNotification_()'s own try/catch.
  sendEmailNotification_(opts.toUserId, opts.toRole, opts.type, String(opts.message), opts.refId);
  return id;
}

// ============================================================
// EMAIL NOTIFICATION BACKEND (Aug 2026 - fills the "missing email/
// notification backend" gap flagged in the security/functionality audit)
// ============================================================
// Best-effort only: this must NEVER throw back into the caller and must
// NEVER block the in-app action it was triggered from (approval, memo
// reply/decision, etc.) - wrapped entirely in try/catch, and simply does
// nothing if MailApp quota is exhausted, no recipient email is on file, or
// any other failure occurs. The in-app Notification Register (bell icon)
// remains the authoritative record either way; email is a convenience
// layer on top of it, not a replacement.
//
// Recipient resolution:
//   toUserId  -> that Users row's Staff ID -> that Staff Master row's Email
//   toRole    -> every Users row with that Role (Status != Inactive) ->
//                each one's Staff Master Email
// A recipient with no Email on file (Staff Master 'Email' column blank, or
// the column itself not yet migrated in - see migrateAddStaffEmailField())
// is silently skipped, never an error.
function sendEmailNotification_(toUserId, toRole, type, message, refId) {
  try {
    const staff = readAll(SHEETS.STAFF);
    const users = readAll(SHEETS.USERS);
    const emailOf = (staffId) => {
      const s = staff.find(x => x['Staff ID'] === staffId);
      const e = s && s['Email'] ? String(s['Email']).trim() : '';
      return e && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e) ? e : '';
    };
    let recipients = [];
    if (toUserId) {
      const u = users.find(x => x['User ID'] === toUserId);
      if (u) { const e = emailOf(u['Staff ID']); if (e) recipients.push(e); }
    } else if (toRole) {
      users.filter(u => u['Role'] === toRole && u['Status'] !== 'Inactive').forEach(u => {
        const e = emailOf(u['Staff ID']);
        if (e && recipients.indexOf(e) === -1) recipients.push(e);
      });
    }
    if (!recipients.length) return; // nothing to send to - not an error
    // Stay well inside Apps Script's daily MailApp quota - this is a
    // best-effort convenience layer, never worth risking the account's
    // quota for other legitimate mail. If this ever needs to scale beyond
    // a handful of recipients per event, batch/digest instead of raising
    // this further.
    if (recipients.length > 20) recipients = recipients.slice(0, 20);
    const subject = "Sri Narasu's Stores KPI - " + (type || 'Notification');
    const body = message + (refId ? ('\n\nReference: ' + refId) : '') +
      '\n\n(This is an automated message from the Stores KPI & Appraisal system. Please do not reply to this email.)';
    MailApp.sendEmail({ to: recipients.join(','), subject: subject, body: body });
  } catch (e) {
    // Swallow every failure (quota exceeded, invalid address, etc.) - see
    // function comment above. Nothing is logged to Audit Log here since a
    // failed best-effort email is not a security/business event.
  }
}

// A notification row belongs to the caller if it was addressed to their
// exact User ID, OR broadcast to their Role with no specific User ID set.
function notificationBelongsToActor_(row, actorUserId, actorRole) {
  if (row['To User ID']) return row['To User ID'] === actorUserId;
  return !!row['To Role'] && row['To Role'] === actorRole;
}

// Looks up the Users-sheet User ID for a given Staff ID (Notification "To
// User ID" is a login User ID, not a Staff ID - they're different sheets).
// Returns '' if no matching active user account is found, so a lookup
// failure just means "no notification sent", never a thrown error.
function userIdForStaff_(staffId) {
  if (!staffId) return '';
  const u = readAll(SHEETS.USERS).find(x => x['Staff ID'] === staffId);
  return u ? u['User ID'] : '';
}

// ============================================================
// EFFECTIVE KPI TARGET (single source of truth for every page)
// ============================================================
// Priority, exactly as specified (FIXED-3-1 FINAL audit item 3):
//   1) Manager-set Monthly Target for that Month + Workflow ID
//      (Monthly Targets sheet) - even if it was explicitly set to 0.
//   2) Else, the Workflow/KPI Master's numeric "Target" column.
// "Expected Output" is a DESCRIPTIVE field only (e.g. "100% bills received
// without shortage") and is NEVER used as, or coerced into, a numeric KPI
// Target - not even when it happens to contain digits. Numeric Target and
// Expected Output are two clearly separate fields; only "Target" feeds the
// Achievement %/Score/Remaining/Progress calculations below.
// PRIOR BEHAVIOUR (removed): this used to try "Expected Output" first
// whenever it parsed as a number, falling back to "Target" only when
// Expected Output was blank/non-numeric. That could silently let a
// numeric-looking Expected Output override the real Target. Removed per
// explicit policy requirement; harmless for all seeded/default data since
// every seeded Expected Output is descriptive text, never a bare number.
// The Target is NEVER left blank: this only returns 0 if the Workflow ID
// itself cannot be found in the Workflow Master at all.
//
// Used everywhere a KPI Target is needed, so KPI entry, Approval, Dashboard,
// KPI Score, Achievement %, Remaining Target and Progress % all agree:
//   - addWorkingEntries / updateEntryGroup (Working Register + Approval page)
//   - getTargetStatus / getTargetDashboard (Monthly Target Status + Dashboard)
//
// opts (optional, all optional) lets callers that already loaded these sheets
// pass them in, instead of re-reading the sheet on every single call:
//   { monthlyTargets: [...rows from Monthly Targets sheet],
//     workflows:      [...rows from Workflow Master sheet] }
// Returns { value, source, configured }. `configured` (P1 fix) tells
// callers whether a target actually EXISTS for this workflow/month -
// distinct from `value`, which can legitimately be 0 either way:
//   configured=true  -> a Manager explicitly set this (Monthly Target row,
//                        even if set to 0, OR a non-blank Workflow Master
//                        Target, even if that value is 0).
//   configured=false -> nothing has been set at all (blank Workflow Master
//                        Target and no Monthly Target row, or the workflow
//                        itself doesn't exist) - `value` is just a safe 0
//                        fallback for arithmetic, NEVER to be read as "the
//                        real target is zero".
// `source` gains a new 'not_set' value (alongside the existing 'monthly'/
// 'default'/'none') so the frontend can show "Not Set" distinctly from a
// real 0 target anywhere Target Source is already displayed.
function getEffectiveTarget(workflowId, month, opts) {
  opts = opts || {};
  const monthlyTargets = opts.monthlyTargets || readAll(SHEETS.TARGETS);
  const workflows = opts.workflows || readAll(SHEETS.WORKFLOW);
  const monthLabel = normalizeMonthLabel(month);

  // 1) Manager-set Monthly Target takes priority whenever one exists for
  // this exact Month + Workflow ID - even if it was explicitly set to 0
  // (explicit means configured=true; it is a real business decision).
  const mt = monthlyTargets.find(t =>
    t['Workflow ID'] === workflowId && normalizeMonthLabel(t['Month']) === monthLabel);
  if (mt) {
    return { value: Number(mt['Monthly Target']) || 0, source: 'monthly', configured: true };
  }

  const wf = workflows.find(w => w['Workflow ID'] === workflowId);
  if (!wf) return { value: 0, source: 'none', configured: false };

  // 2) Workflow/KPI Master's numeric "Target" column. Expected Output is
  // description-only and is intentionally never consulted here. A blank/
  // empty cell means nobody has configured a target at all (Not Set); a
  // cell literally containing 0 is a deliberate, configured, valid zero.
  const raw = wf['Target'];
  const isBlank = raw === '' || raw === null || raw === undefined;
  if (isBlank) return { value: 0, source: 'not_set', configured: false };
  return { value: Number(raw) || 0, source: 'default', configured: true };
}

// ============================================================
// MEMO EFFECTIVE DEDUCTION - SINGLE AUTHORITATIVE RESOLVER
// ============================================================
// Every place that needs "how much does this memo actually cost this
// employee" (score calculation, dashboard/board display, snapshot save,
// migrations) MUST call this - never read/derive Effective Deduction any
// other way. This is what guarantees a memo can only ever be deducted once,
// however many times the row is read: Memo Issued / Employee Reply / HR
// Review never write to Effective Deduction at all (see
// addDisciplinaryAction/replyToMemo/reviewMemo) - only decideMemo() writes
// it, exactly once per decision, and this function just reads that one
// field back consistently.
//
// r: a Disciplinary/Memo row (from readAll(SHEETS.DISCIPLINARY)).
function resolveMemoEffectiveDeduction(r) {
  const stored = r['Effective Deduction'];
  if (stored !== '' && stored !== undefined && stored !== null) {
    // Normal case (every memo created after this fix, and every migrated
    // legacy row - see migrateAddEffectiveDeductionField/
    // migrateZeroPendingMemoEffectiveDeduction): the column already holds
    // the correct, final answer. Read it as-is.
    return Number(stored) || 0;
  }
  // Only reachable for a legacy row that predates the 'Effective Deduction'
  // column entirely and has NOT yet been migrated. Never default to the
  // full original Deduction Marks here - that would deduct for a memo with
  // no confirmed penalty. Only a Final Decision that actually confirms a
  // penalty may deduct; a memo still open (no decision yet) or explicitly
  // decided as no-penalty is always 0.
  const decision = r['Final Decision'];
  if (!decision || HR_POLICY.MEMO_DECISION_NO_PENALTY.indexOf(decision) !== -1) {
    return 0;
  }
  // 'Penalty Confirmed' / 'Deduction Adjusted' / legacy 'Warning Upheld' /
  // 'Escalated' - Deduction Marks already holds the right figure (the
  // original for a plain confirmed penalty, or the adjusted value for a
  // pre-fix "Deduction Adjusted" row that overwrote Deduction Marks itself).
  return Number(r['Deduction Marks']) || 0;
}

// Friendly display label for a memo's lifecycle stage - DISPLAY ONLY. The
// underlying 'Status' column (Issued/Replied/Reviewed/Closed) is never
// renamed/changed anywhere else in the app; this just maps it to the
// wording the Memo -> Reply -> HR Review -> Final Decision flow uses on
// the employee/HR board.
function memoStatusLabel(status) {
  switch (status) {
    case 'Replied': return 'Reply Submitted';
    case 'Reviewed': return 'Reply Reviewed';
    case 'Closed': return 'Decision Finalized';
    case 'Issued':
    default: return 'Memo Issued';
  }
}

// ============================================================
// TRUCK NO (auto-generated, month-scoped, TEAM WORKFLOWS ONLY)
// ============================================================
// Truck No resets daily: Truck-0001, Truck-0002... starting fresh each
// calendar day (based on the entry's Date, not "today", so backfilled
// entries for an older date still get that date's own sequence).
// Scoped WORKFLOW-WISE to TEAM_WORKFLOW_IDS (see hasTeamWorkflowLine in
// addWorkingEntries) - across ALL Activities, not just Receiving. A
// submission only consumes/increments the sequence if it contains at
// least one team-workflow line. The row scan below additionally filters
// to Workflow ID in TEAM_WORKFLOW_IDS as a safety net, so an individual-
// workflow row (blank Truck No) can never inflate the sequence. One
// Truck No per addWorkingEntries()
// call - shared across every line/participant in that submission. Purely
// a reference/tracking field - Team Ref No (above) is still what drives
// team-based score splitting; Truck No never affects scoring.
// NOTE: no LockService call in here anymore - see addWorkingEntries(), which
// now wraps the duplicate check + this sequence read + the row inserts in
// ONE lock so the two never run interleaved for two concurrent submissions
// (was previously two separate locked sections with a gap between them,
// where two simultaneous users could both read the same "next" sequence).
//
// MONTH-BASED SEQUENCE (changed from day-based): the sequence now resets
// only when the calendar MONTH changes, not every day, and the month is
// embedded in the number itself - 'Truck-2608-0001' means truck #1 of
// August 2026 (yyMM). This is deliberate, not cosmetic: a pending truck
// from earlier in the month must never collide with a different physical
// truck later in the SAME month (which a pure day-reset could allow), and
// 'Truck-0001' in August and 'Truck-0001' in September must never be
// treated as the same truck. Embedding the month directly in the string
// guarantees both, because every place in this file that matches an
// existing Truck No (the existing-truck lookup in _addWorkingEntriesLocked)
// already matches on the canonical Truck No STRING - once that string is
// unique per month, all of that matching is correct automatically, with no
// further changes needed anywhere else.
// Old truck numbers already in the sheet from before this change (plain
// 'Truck-0001', no month) are left exactly as they are - never renamed -
// and stay perfectly safe: a bare old-style number can never collide with
// a new month-qualified one, since the two formats are textually distinct.
function generateTruckNo(dateVal) {
  const tz = Session.getScriptTimeZone();
  const monthLabel = Utilities.formatDate(new Date(dateVal), tz, 'yyMM'); // e.g. '2608' = Aug 2026
  const rows = readAll(SHEETS.REGISTER);
  const trackedActivityIds = getTruckTrackedActivityIds();
  let maxSeq = 0;
  const seqPattern = new RegExp('^Truck-' + monthLabel + '-(\\d+)$');
  rows.forEach(r => {
    if (!r['Truck No'] || !r['Date']) return;
    if (!trackedActivityIds.has(r['Activity ID'])) return;
    if (Utilities.formatDate(new Date(r['Date']), tz, 'yyMM') !== monthLabel) return;
    const canonical = normalizeTruckNo(r['Truck No']);
    const m = canonical.match(seqPattern);
    if (m) maxSeq = Math.max(maxSeq, parseInt(m[1], 10));
  });
  // 4-digit zero-padded sequence within the month (Truck-2608-0001,
  // Truck-2608-0002, ...). September then starts its own fresh
  // Truck-2609-0001 automatically, since monthLabel changes.
  return 'Truck-' + monthLabel + '-' + ('0000' + (maxSeq + 1)).slice(-4);
}


// ============================================================
// TRUCK NUMBER NORMALIZATION
// ============================================================
// Receiving trucks are displayed/stored in one canonical form:
//   001 / 0001 / Truck-001 / Truck-0001  -> Truck-0001
// This lets old 3-digit records and new 4-digit records refer to the same
// physical truck while preserving the existing display convention.
// Non-generated truck identifiers are only trimmed, not rewritten.
function normalizeTruckNo(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  const m = raw.match(/^truck[\s-]*(\d+)$/i);
  if (m) return 'Truck-' + ('0000' + String(parseInt(m[1], 10))).slice(-4);
  if (/^\d+$/.test(raw)) return 'Truck-' + ('0000' + String(parseInt(raw, 10))).slice(-4);
  return raw;
}

// ============================================================
// TRUCK-TRACKED ACTIVITIES (derived, not hardcoded)
// ============================================================
// An Activity is "truck-tracked" if at least one of its active workflows
// is in TEAM_WORKFLOW_IDS (i.e. it has a co-staff/team entry point, such
// as Unloading & Bill Confirmation). Once an Activity is truck-tracked,
// the Truck No applies to EVERY workflow step under that Activity for
// that submission - including individual (non-team) steps like GRN
// Entry, Flow Report Preparation, In-Charge Signature, Bin Card Update,
// Purchase Copy Submission - so the full receive-to-close chain for one
// physical truck can be followed end-to-end, even though most of those
// steps are done solo and never split points. Team Ref No / score-
// splitting is completely unaffected by this - that still depends only
// on TEAM_WORKFLOW_IDS membership per workflow, exactly as before.
function getTruckTrackedActivityIds(workflows) {
  const wfs = workflows || readAll(SHEETS.WORKFLOW);
  const ids = new Set();
  wfs.forEach(w => {
    if (w['Status'] !== 'Inactive' && TEAM_WORKFLOW_IDS.indexOf(w['Workflow ID']) !== -1) {
      ids.add(w['Activity ID']);
    }
  });
  return ids;
}

// ============================================================
// DUPLICATE ENTRIES CONTROL
// ============================================================
// Prevents a duplicate Working Register record for the same physical job.
// For truck-tracked work the identity is Date + Activity + Workflow + Truck No,
// plus the exact same set of people. Therefore the same Team can log the same
// Workflow multiple times on one day when each entry is for a different truck.
// For non-truck work the physical-job identity remains Date + Activity + Workflow,
// plus the exact same set of people
// (Staff ID who logged it + any Co-Staff), that already exists as a Truck
// instance (a Team Ref No group for team-based Workflows, or a single solo
// row for individual ones). Rejected instances are excluded, so a staff
// member can always resubmit after a rejection.
//
// Runs BEFORE generateTruckNo()/row creation in addWorkingEntries, so a
// rejected duplicate never consumes a Truck No or leaves partial rows.
//
// date         : the entry's Date (string/Date, same as p.date)
// activityId   : line.activityId
// workflowId   : line.workflowId
// participants : Staff IDs working this line (Logged By + ticked Co-Staff),
//                NOT yet deduped/sorted - this function normalizes them.
// truckNo      : selected existing Truck No for truck-tracked work; blank for a new truck
// returns      : the existing Truck No if a duplicate instance is found, otherwise null.
function findDuplicateTruckEntry(date, activityId, workflowId, participants, truckNo) {
  const tz = Session.getScriptTimeZone();
  const dayLabel = Utilities.formatDate(new Date(date), tz, 'yyyy-MM-dd');
  const participantKey = Array.from(new Set((participants || []).filter(Boolean))).sort().join(',');
  const rows = readAll(SHEETS.REGISTER).filter(r =>
    r['Activity ID'] === activityId &&
    r['Workflow ID'] === workflowId &&
    r['Approval Status'] !== 'Rejected' &&
    r['Date'] && Utilities.formatDate(new Date(r['Date']), tz, 'yyyy-MM-dd') === dayLabel
  );
  if (!rows.length) return null;
  const groups = {};
  rows.forEach(r => {
    const key = String(r['Team Ref No'] || r['Line ID'] || r['Staff ID'] || r['Logged By ID'] || '');
    (groups[key] = groups[key] || []).push(r);
  });
  for (const key in groups) {
    const p = Array.from(new Set(groups[key].map(r => r['Staff ID'] || r['Logged By ID']).filter(Boolean))).sort().join(',');
    if (p === participantKey) return groups[key][0]['Truck No'] || groups[key][0]['Line ID'] || 'existing entry';
  }
  return null;
}


// ============================================================
// STAFF-WISE STORES KPI - CORRECT PER-WORKFLOW-MONTH AGGREGATION
// (Stores KPI Team-Split Dilution Fix, Aug 2026)
// ============================================================
// ROOT CAUSE OF THE BUG: 'KPI Score'/'Achievement %' are stored on EACH
// Working Register row at entry time, computed against that row's OWN
// Actual vs the workflow's full Monthly Target. That is correct when a
// staff has only ONE Approved row for a given Workflow+Month, but WRONG
// the instant they have more than one (e.g. a Team job + a later
// Individual job on the same truck workflow, in the same month): each row
// gets scored as if it alone had to satisfy the WHOLE month's target, so
// two partial rows (5/10=50%, 3/10=30%) sum to a far lower score than one
// combined row would (5+3=8/10=80%). Team Split then divides that
// (already too low) per-row score again, compounding the error. Actual
// participation itself was NEVER divided by team size (see withTeamSplit
// below) - only the derived Score/Weightage was - but computing
// Achievement %/Score per row instead of per Workflow+Month silently
// re-applied the full Target multiple times, which produced the same
// end result as if participation HAD been split.
//
// THE FIX: aggregate raw (undiluted) Actual across every Approved row a
// staff has for the same Workflow ID + Month FIRST (Step 2 of the
// business rule), compute Achievement % and KPI Score ONCE from that
// combined total (Step 6), and only then apply the staff's own
// Team-Split share of the Weightage (Section 5 - weightage MAY be shared,
// participation MUST NOT be). This is the single authoritative grouping
// used by both computeStoresKPIPct() and getMyScore() below, so the
// Stores % shown to a Manager and the KPI breakdown shown to a Staff
// member are always derived the same way.
//
// regRowsForStaff : this staff's Approved Working Register rows (any
//                    Workflow/Month - grouped internally by this function).
// approvedInRange : ALL Approved rows in the same date range (any staff) -
//                    needed by withTeamSplit() to size each Team Ref No
//                    group correctly (same contract as every other
//                    withTeamSplit() call in this file).
function computeStaffWorkflowKpiGroups_(regRowsForStaff, approvedInRange, workflows, monthlyTargets) {
  const splitRows = withTeamSplit(regRowsForStaff, approvedInRange);
  const groups = {};
  splitRows.forEach(r => {
    const month = monthLabelOf(r['Date']);
    const key = r['Workflow ID'] + '__' + month;
    if (!groups[key]) {
      groups[key] = { workflowId: r['Workflow ID'], month, teamRefNos: [], recordCount: 0, totalActual: 0, weightagePct: 0, frozenTargetRows: [] };
    }
    const g = groups[key];
    g.recordCount += 1;
    g.totalActual += Number(r['Actual']) || 0; // RAW, undiluted - Actual/participation is NEVER divided
    g.weightagePct += Number(r['Effective Weightage %']) || 0; // this staff's own Team-Split share only
    if (r['Team Ref No'] && g.teamRefNos.indexOf(r['Team Ref No']) === -1) g.teamRefNos.push(r['Team Ref No']);
    // HISTORICAL FREEZE FIX (audit item - Monthly Target retroactive change,
    // spec section 4/21): every Working Register row already stores the
    // Target/Target Source that was resolved via getEffectiveTarget() at
    // the moment it was submitted (see _addWorkingEntriesLocked) - approval
    // never rewrites it (see _buildApprovalFields). That stored value is
    // this row's frozen snapshot and must be reused here so an already-
    // Approved transaction's KPI Score cannot silently change if a Manager
    // edits the Monthly Target for that Workflow+Month afterwards. Only
    // rows genuinely missing a stored Target (legacy pre-snapshot data)
    // fall back to a live lookup - see 'legacy' handling below.
    const hasFrozenTarget = r['Target'] !== '' && r['Target'] !== undefined && r['Target'] !== null;
    if (hasFrozenTarget) {
      g.frozenTargetRows.push({
        value: Number(r['Target']) || 0,
        source: r['Target Source'] || 'frozen',
        submittedOn: r['Submitted On'] ? new Date(r['Submitted On']).getTime() : 0,
        row: r._row || 0
      });
    }
  });
  return Object.values(groups).map(g => {
    const wf = workflows.find(w => w['Workflow ID'] === g.workflowId);
    let resolved, legacy = false, targetMismatch = false;
    if (g.frozenTargetRows.length === g.recordCount) {
      // Every row in this Workflow+Month group has its own frozen Target -
      // use the EARLIEST-submitted row's value as the group's historical
      // baseline (the target that was actually in effect when work for
      // this Workflow+Month first began). If a later row in the same
      // group was submitted after the Monthly Target changed mid-month,
      // its frozen Target will differ - flagged via targetMismatch for
      // Manager visibility rather than silently averaged or overwritten.
      const sorted = g.frozenTargetRows.slice().sort((a, b) => (a.submittedOn - b.submittedOn) || (a.row - b.row));
      resolved = { value: sorted[0].value, source: sorted[0].source, configured: true };
      targetMismatch = sorted.some(x => x.value !== sorted[0].value);
    } else {
      // LEGACY / SNAPSHOT NOT AVAILABLE (spec section 24): at least one row
      // in this group predates the frozen-Target column being populated -
      // cannot safely fabricate its historical basis, so fall back to the
      // old live-resolution behaviour for this group only (unchanged from
      // before this fix) and mark it so callers/UI can show a legacy notice.
      resolved = getEffectiveTarget(g.workflowId, g.month, { monthlyTargets, workflows });
      legacy = true;
    }
    // Achievement % computed ONCE from the combined Total Actual - this is
    // the fix. achievementPctRaw preserves the true (uncapped) value for
    // audit/over-achievement visibility; achievementPct is the capped
    // (max 100%) figure used for display.
    const achievementPctRaw = computeAchievementPct(g.totalActual, resolved.value);
    const achievementPct = achievementPctRaw === null ? null : Math.min(achievementPctRaw, 100);
    const weightagePct = Math.round(g.weightagePct * 100) / 100;
    const kpiScoreRaw = computeKpiScore(achievementPctRaw, weightagePct);
    return Object.assign({}, g, {
      legacySnapshot: legacy, targetMismatch: targetMismatch,
      workflowName: wf ? wf['Workflow Name'] : g.workflowId,
      target: resolved.value, targetConfigured: resolved.configured,
      achievementPct, achievementPctRaw, weightagePct,
      kpiScore: kpiScoreRaw === null ? 0 : kpiScoreRaw
    });
  });
}

function withTeamSplit(rows, groupWithin) {
  const groups = {};
  groupWithin.forEach(r => {
    if (TEAM_WORKFLOW_IDS.indexOf(r['Workflow ID']) === -1) return;
    if (!r['Team Ref No']) return;
    const key = r['Team Ref No'] + '__' + r['Workflow ID'];
    (groups[key] = groups[key] || []).push(r);
  });
  return rows.map(r => {
    // Rejected (or any non-Approved) rows never earn KPI credit - skip Team
    // Split entirely for them and force their Effective figures to 0, even
    // though their Raw KPI Score/Weightage may still show a divided value
    // elsewhere in the sheet for reference.
    if (r['Approval Status'] !== 'Approved') {
      return Object.assign({}, r, {
        'Effective KPI Score': 0,
        'Effective Weightage %': 0,
        'Team Split Size': 1
      });
    }
    const rawScore = Number(r['KPI Score']) || 0;
    const rawWeightage = (r['Weightage %'] !== '' && r['Weightage %'] !== undefined) ? Number(r['Weightage %']) : Number(r['Max Score']) || 0;
    let size = 1;
    if (TEAM_WORKFLOW_IDS.indexOf(r['Workflow ID']) !== -1 && r['Team Ref No']) {
      const key = r['Team Ref No'] + '__' + r['Workflow ID'];
      size = groups[key] ? groups[key].length : 1;
    }
    return Object.assign({}, r, {
      'Effective KPI Score': Math.round((rawScore / size) * 100) / 100,
      'Effective Weightage %': Math.round((rawWeightage / size) * 100) / 100,
      'Team Split Size': size
    });
  });
}

// P2 FIX #7 (audit): Monthly Target add/update previously did
// `Number(p.target) || 0`, which silently turns invalid input (NaN, i.e.
// non-numeric text) into 0 - indistinguishable from a deliberately-set
// zero target - and let Infinity/negative values through untouched (since
// both are truthy, `|| 0` never catches them). This helper rejects
// anything that isn't a valid, finite, non-negative number, with target
// left completely untouched (still 0) if that's genuinely what the
// Manager entered - Target = 0 remains a valid, distinguishable value
// (Target Not Set stays a separate, unrelated state - a row simply not
// existing in the Targets sheet - so this doesn't affect that at all).
function _validateMonthlyTargetValue(raw) {
  if (raw === undefined || raw === null || String(raw).trim() === '') {
    throw new Error('Monthly Target is required.');
  }
  const n = Number(raw);
  if (!isFinite(n) || isNaN(n)) {
    throw new Error('Monthly Target must be a valid number.');
  }
  if (n < 0) {
    throw new Error('Monthly Target cannot be negative.');
  }
  return n;
}

// Shared Approved-base TOTAL KPI % calculator for the Dashboard card (Aug
// 2026) - factored out of getDashboard() so the same exact methodology
// (Approved rows only, Team-Split-safe, fixed denominator, ACT009 excluded
// from the numerator) can be run twice per getDashboard() call: once for
// the requested period and once for the immediately preceding period of
// the same length, to power the "uptrend"/"downtrend" arrow on the card.
// See getDashboard() for the FIX history this logic was carried over from.
function computeDashboardKpiForRange_(from, to, staffId) {
  const approvedInRange = readAll(SHEETS.REGISTER).filter(r => {
    const d = new Date(r['Date']);
    return d >= from && d <= to && r['Approval Status'] === 'Approved';
  });
  let rows = approvedInRange;
  if (staffId) rows = rows.filter(r => r['Staff ID'] === staffId);

  const workflows = readAll(SHEETS.WORKFLOW);
  const monthlyTargets = readAll(SHEETS.TARGETS);
  const staffMaster = readAll(SHEETS.STAFF);
  const rowsByStaffId = {};
  rows.forEach(r => { (rowsByStaffId[r['Staff ID']] = rowsByStaffId[r['Staff ID']] || []).push(r); });

  const byStaff = {}, byActivity = {}, byWorkflow = {}, bySection = {};
  let totalKpiScore = 0, totalMaxScore = 0;
  Object.keys(rowsByStaffId).forEach(sId => {
    const staffRows = rowsByStaffId[sId];
    // groupWithin MUST be the full (not staffId-filtered) approvedInRange,
    // or a personal Dashboard (staffId given) would only ever see its own
    // row per team job and wrongly resolve Team Split Size to 1 - same
    // contract as every other withTeamSplit()/computeStaffWorkflowKpiGroups_ call.
    const groups = computeStaffWorkflowKpiGroups_(staffRows, approvedInRange, workflows, monthlyTargets);
    const sm = staffMaster.find(s => s['Staff ID'] === sId);
    const staffName = staffRows[0]['Staff Name'];
    const section = sm ? sm['Section'] : 'Unknown';
    groups.forEach(g => {
      const wf = workflows.find(w => w['Workflow ID'] === g.workflowId);
      const activityName = wf ? wf['Activity Name'] : g.workflowId;
      byStaff[staffName] = (byStaff[staffName] || 0) + g.kpiScore;
      byActivity[activityName] = (byActivity[activityName] || 0) + g.kpiScore;
      byWorkflow[g.workflowName] = (byWorkflow[g.workflowName] || 0) + g.kpiScore;
      bySection[section] = (bySection[section] || 0) + g.kpiScore;
      // TOTAL KPI % card - FIXED DENOMINATOR (Option B, Aug 2026) + single
      // source with computeStoresKPIPct()/getMyScore(): exclude Attendance
      // (ACT009) groups from the numerator, to match the fixed
      // (ACT009-excluded) denominator added once per staff below.
      if (STORES_KPI_EXCLUDED_ACTIVITY_IDS.indexOf(wf ? wf['Activity ID'] : '') === -1) {
        totalKpiScore += g.kpiScore;
      }
    });
    // Fixed Max/Possible Score added ONCE per staff who appears in this
    // range (not once per Workflow+Month group) - see
    // computeStaffFixedActiveWorkflowMaxScore_().
    totalMaxScore += computeStaffFixedActiveWorkflowMaxScore_(sId, workflows);
  });

  // CAP AT 100 (Aug 2026 fix - Dashboard parity): mirrors the same clamp
  // already applied in computeStoresKPIPct(). 'totalKpiScore' is built from
  // each Register row's FROZEN historical Weightage % (Historical Freeze
  // policy), while 'totalMaxScore' is the CURRENT/live sum of Weightage %
  // across today's Active workflows (computeStaffFixedActiveWorkflowMaxScore_).
  // If a Workflow's Weightage % was reduced after entries were Approved
  // (e.g. the Physical:System 60:40 recalibration), 'totalKpiScore' can
  // legitimately exceed 'totalMaxScore' and the raw % can exceed 100 -
  // without this clamp the Dashboard's "Total KPI % (MTD)" card could show
  // e.g. 100.6% even though every other KPI %/headline-score consumer
  // (computeStoresKPIPct(), getMyScore()) is already clamped to 0-100.
  // This keeps the Dashboard card consistent with those single-source-of-
  // truth callers of computeStaffFixedActiveWorkflowMaxScore_().
  const totalKpiPctRaw = totalMaxScore > 0 ? (totalKpiScore / totalMaxScore) * 100 : 0;
  const totalKpiPct = Math.round(clamp(totalKpiPctRaw, 0, 100) * 100) / 100;

  return {
    totalEntries: rows.length,
    totalKpiScore: Math.round(totalKpiScore * 100) / 100,
    totalMaxScore,
    totalKpiPct,
    byStaff, byActivity, byWorkflow, bySection
  };
}

// ============================================================
// API
// ============================================================
const api = {

  // ---------- AUTH ----------
  login: function (p) {
    const users = readAll(SHEETS.USERS);
    const u = users.find(x => x['User ID'] === p.userId && x['Status'] !== 'Inactive');
    // Same generic message whether the User ID doesn't exist, is inactive,
    // or the password is wrong - never reveal which (avoids account
    // enumeration).
    const fail = () => {
      logAudit(p.userId || '(unknown)', 'LOGIN_FAILED', 'Invalid credentials');
      throw new Error('Invalid User ID or Password');
    };
    // LOGIN LOCKOUT (audit fix - brute-force protection): record a failed
    // attempt against this specific user's row and lock them out once too
    // many pile up. Kept separate from fail() because fail() also covers
    // "no such user" (nothing to increment there) and because a lockout hit
    // needs its own distinct audit action/message rather than the generic
    // "Invalid credentials". Silently a no-op if the lockout columns don't
    // exist yet (pre-migration spreadsheet) - see migrateAddLoginLockoutFields().
    const failWithLockoutTracking = () => {
      if (u && u.hasOwnProperty('Failed Login Count')) {
        const nextCount = (Number(u['Failed Login Count']) || 0) + 1;
        const update = { 'Failed Login Count': nextCount, 'Last Failed Login': new Date() };
        if (nextCount >= LOGIN_LOCKOUT_POLICY.MAX_FAILED_ATTEMPTS) {
          update['Locked Until'] = new Date(Date.now() + LOGIN_LOCKOUT_POLICY.LOCKOUT_MINUTES * 60 * 1000);
          logAudit(p.userId, 'ACCOUNT_LOCKED',
            nextCount + ' failed attempts - locked for ' + LOGIN_LOCKOUT_POLICY.LOCKOUT_MINUTES + ' minute(s)');
        }
        updateRow(SHEETS.USERS, u._row, update);
      }
      fail();
    };
    if (!u) fail();

    // Check lockout BEFORE verifying the password - a locked account stays
    // locked even if the correct password is supplied, otherwise the lock
    // would be pointless.
    if (u.hasOwnProperty('Locked Until') && u['Locked Until']) {
      const lockedUntil = new Date(u['Locked Until']);
      if (!isNaN(lockedUntil.getTime()) && lockedUntil.getTime() > Date.now()) {
        logAudit(p.userId, 'LOGIN_BLOCKED_LOCKED', 'Account locked until ' + lockedUntil.toISOString());
        throw new Error('This account is temporarily locked due to too many failed login attempts. Please try again later.');
      }
    }

    const storedHash = u['Password Hash'];
    if (storedHash) {
      // Normal path: verify against the hash, plaintext password never
      // touches the sheet at all.
      if (!hashesMatch(hashPassword(p.password, u['Salt']), storedHash)) failWithLockoutTracking();
    } else {
      // SAFE MIGRATION path (audit item 15): this row predates the hash
      // columns and still has a plaintext 'Password'. Verify against that
      // exactly as before, then immediately hash+salt it and blank the
      // plaintext column - this user's password itself is unchanged, only
      // how it's stored is. Runs at most once per user (storedHash is set
      // on every subsequent login from here on).
      if (String(u['Password'] || '') !== String(p.password) || !p.password) failWithLockoutTracking();
      const salt = makeSalt();
      updateRow(SHEETS.USERS, u._row, {
        'Password Hash': hashPassword(p.password, salt), 'Salt': salt, 'Password': ''
      });
      logAudit(p.userId, 'PASSWORD_MIGRATED_TO_HASH', '');
    }

    logAudit(p.userId, 'LOGIN', 'Success');
    // Successful login - reset the failed-attempt counter and any lock, so
    // the count reflects "attempts since the last success", not a lifetime
    // total. No-op if the lockout columns don't exist yet.
    if (u.hasOwnProperty('Failed Login Count')) {
      updateRow(SHEETS.USERS, u._row, { 'Failed Login Count': 0, 'Locked Until': '' });
    }
    // FORCED PASSWORD CHANGE (audit item 31): a starter/reset account is
    // flagged 'Must Change Password' = Yes; the frontend blocks entry into
    // the app (beyond the change-password screen) until it's cleared.
    const mustChangePassword = String(u['Must Change Password'] || '').trim().toLowerCase() === 'yes';
    return {
      token: createSession(u),   // opaque session token - the ONLY thing the
                                  // client needs to send back on every later call
      userId: u['User ID'],
      staffId: u['Staff ID'],
      staffName: u['Staff Name'],
      role: u['Role'], // 'Manager' or 'Staff' or 'HR'
      mustChangePassword: mustChangePassword
    };
  },

  logout: function (p) {
    logAudit(p.actorUserId, 'LOGOUT', '');
    destroySession(p.token);
    return { success: true };
  },

  // FRONTEND IDLE-WARNING SUPPORT: a deliberately trivial, cheap action for
  // the "Continue working?" modal's Continue button to call. It does no
  // real work - handle() has already run resolveSession(params.token) by
  // the time this executes (ping is not in PUBLIC_ACTIONS), and
  // resolveSession() itself re-puts the session into CacheService with a
  // fresh SESSION_TTL_SECONDS TTL. So simply reaching this function at all
  // IS the session refresh; returning ok:true just lets the client confirm
  // it succeeded and dismiss the warning.
  ping: function (p) {
    return { ok: true, serverTime: new Date().toISOString() };
  },

  // p.actorUserId is now server-verified (handle() overwrote it from the
  // session) - a user can only ever change THEIR OWN password, never
  // someone else's, even though the old code technically accepted any
  // p.userId the client sent.
  changePassword: function (p) {
    const users = readAll(SHEETS.USERS);
    const u = users.find(x => x['User ID'] === p.actorUserId);
    if (!u) throw new Error('Current password incorrect');
    const oldOk = u['Password Hash']
      ? hashesMatch(hashPassword(p.oldPassword, u['Salt']), u['Password Hash'])
      : String(u['Password'] || '') === String(p.oldPassword) && !!p.oldPassword;
    if (!oldOk) throw new Error('Current password incorrect');
    if (!p.newPassword || String(p.newPassword).length < 6) {
      throw new Error('New password must be at least 6 characters');
    }
    const salt = makeSalt();
    updateRow(SHEETS.USERS, u._row, {
      'Password Hash': hashPassword(p.newPassword, salt), 'Salt': salt, 'Password': '',
      'Must Change Password': 'No'
    });
    logAudit(p.actorUserId, 'CHANGE_PASSWORD', '');
    return { success: true };
  },

  // ---------- MASTERS (read) ----------
  getMasters: function () {
    return {
      company: readAll(SHEETS.COMPANY),
      sections: readAll(SHEETS.SECTION),
      staff: readAll(SHEETS.STAFF),
      activities: readAll(SHEETS.ACTIVITY),
      workflows: readAll(SHEETS.WORKFLOW),
      // SINGLE SOURCE OF TRUTH (audit item 18): Recognition Types, Memo
      // Types and Memo Categories used to also be hardcoded a second time
      // in index.html (RECOGNITION_TYPES/MEMO_TYPES/MEMO_CATEGORIES consts),
      // duplicating HR_POLICY above - editing HR_POLICY alone would silently
      // stop matching what the dropdowns showed. index.html now reads these
      // three lists from here instead of keeping its own copy.
      //
      // GETMASTERS RUNTIME FIX: this used to call ui_getRecognitionTypes()/
      // ui_getMemoTypes() directly, but those require params.token and
      // throw via requireUiSession_() when called with no params - which is
      // exactly what happened here, since getMasters() is invoked with no
      // params at all. getMasters() already only runs after handle() has
      // independently verified the caller's session (see PUBLIC_ACTIONS),
      // so it reads the same underlying data through the session-free
      // internal helpers instead. The ui_* wrappers (used by the Sidebar,
      // which calls them directly via google.script.run and bypasses
      // handle() entirely) are unchanged and still require a valid token.
      recognitionTypes: _getRecognitionTypesInternal_(),
      memoTypes: _getMemoTypesInternal_(),
      memoCategories: HR_POLICY.MEMO_CATEGORIES.slice(),
      memoDecisionOptions: HR_POLICY.MEMO_DECISION_OPTIONS.slice(),
      // SINGLE SOURCE OF TRUTH: TEAM_WORKFLOW_IDS used to also be hardcoded
      // a second time in index.html, with only a comment asking whoever
      // edits one list to remember to edit the other. index.html now reads
      // this from getMasters() instead of keeping its own copy.
      teamWorkflowIds: TEAM_WORKFLOW_IDS.slice(),
      // AUTOMATIC TRAINING (Aug 2026): single source of truth for the
      // Training dropdown on the Training Management screen - never
      // hard-code training names in index.html. Only Active trainings
      // should be offered for new records (frontend filters Status).
      trainingMaster: readAll(SHEETS.TRAINING_MASTER)
    };
  },

  // Always returns the CURRENT Workflow Master state straight off the
  // sheet - never the session-cached MASTERS.workflows snapshot taken at
  // login. Used by the Edit Workflow screen and by the Working Register's
  // Activity/Workflow selection, so a Manager's Target/Weightage/Expected
  // Output/Status change is visible immediately, without logging out/in.
  getWorkflowsFresh: function () {
    return readAll(SHEETS.WORKFLOW);
  },

  // WEIGHTAGE VALIDATION (audit item 7 - REVISED Aug 2026 for the
  // Physical:System effort-based model): this used to require every
  // Activity's ACTIVE workflows to individually sum to 100. That assumption
  // no longer holds - the whole point of the Physical:System 60:40
  // recalibration is that different Activities intentionally carry
  // different total weight (e.g. Housekeeping's single workflow alone is
  // ~14.4, MIS Reports' four workflows total ~10.4). What must sum to 100
  // is the GRAND TOTAL across every Active, KPI-scored workflow -
  // ATTENDANCE/ACT009 is excluded here exactly as it is in
  // computeStaffFixedActiveWorkflowMaxScore_() (single source of truth kept
  // in sync with that function, so this check can never drift from what
  // Stores KPI % actually uses as its denominator). Read-only report (does
  // not block anything by itself) so a Manager can see and fix drift;
  // surfaced as a warning banner on the Masters/Edit Workflow screen. Only
  // Active workflows count (an Inactive one is retired and shouldn't count
  // against the total). Small tolerance (0.5) allows for the deliberate
  // 3-decimal Physical:System weightages rounding to ~99.995-100.005.
  validateWeightages: function () {
    const workflows = readAll(SHEETS.WORKFLOW).filter(w => w['Status'] !== 'Inactive' &&
      STORES_KPI_EXCLUDED_ACTIVITY_IDS.indexOf(w['Activity ID']) === -1);
    const grandTotal = Math.round(workflows.reduce((sum, w) => sum + (Number(w['Weightage %']) || 0), 0) * 100) / 100;
    const valid = Math.abs(grandTotal - 100) <= 0.5;
    return { valid: valid, grandTotal: grandTotal, workflowCount: workflows.length, issues: valid ? [] : [{ total: grandTotal }] };
  },

  // ---------- EDIT WORKFLOW (Manager only) ----------
  // The ONLY sanctioned way to change Target / Weightage % / Expected
  // Output / Status - never edit the Workflow Master sheet directly. Max
  // Score is intentionally excluded (read-only, see below). Every changed
  // field gets its own Audit Log entry (old -> new).
  updateWorkflow: function (p) {
    // p: {actorRole, actorUserId, workflowId, fields: {Target, 'Weightage %', 'Max Score', 'Expected Output', Status}}
    requireManager(p.actorRole);
    const workflows = readAll(SHEETS.WORKFLOW);
    const row = workflows.find(w => w['Workflow ID'] === p.workflowId);
    if (!row) throw new Error('Workflow not found: ' + p.workflowId);

    // VALIDATION (audit item 7) - reject an out-of-range Weightage % outright.
    if (p.fields && Object.prototype.hasOwnProperty.call(p.fields, 'Weightage %')) {
      const wPct = Number(p.fields['Weightage %']);
      if (isNaN(wPct) || wPct < 0 || wPct > 100) {
        throw new Error('Weightage % must be between 0 and 100');
      }
    }
    if (p.fields && Object.prototype.hasOwnProperty.call(p.fields, 'Target')) {
      const tVal = Number(p.fields['Target']);
      if (isNaN(tVal) || tVal < 0) throw new Error('Target cannot be negative');
    }

    const updates = {};
    const changes = [];
    EDITABLE_WORKFLOW_FIELDS.forEach(field => {
      if (field === 'Max Score') return; // never independently editable - see below
      if (!p.fields || !Object.prototype.hasOwnProperty.call(p.fields, field)) return;
      const oldVal = row[field];
      const newVal = p.fields[field];
      if (String(oldVal) !== String(newVal)) {
        updates[field] = newVal;
        changes.push({ field: field, oldVal: oldVal, newVal: newVal });
      }
    });
    // MAX SCORE CLEANUP (audit item 27): Weightage % is the single
    // authoritative scoring value (see computeKpiScore()) - Max Score is
    // NEVER used in any KPI Score calculation anymore, it is kept only as a
    // read-only, always-synchronized display/legacy field. Any 'Max Score'
    // the client sends is ignored outright; it is always forced equal to
    // the (possibly just-updated) Weightage %, so it can never independently
    // drift into a second, conflicting scoring number.
    const effectiveWeightage = updates['Weightage %'] !== undefined ? updates['Weightage %'] : row['Weightage %'];
    if (String(row['Max Score']) !== String(effectiveWeightage)) {
      updates['Max Score'] = effectiveWeightage;
      changes.push({ field: 'Max Score', oldVal: row['Max Score'], newVal: effectiveWeightage });
    }
    if (changes.length === 0) return { success: true, changed: 0 };

    updateRow(SHEETS.WORKFLOW, row._row, updates);

    // One Audit Log row per changed field - Timestamp/User ID come from
    // logAudit() itself; Workflow ID + field + old/new go into Details.
    changes.forEach(c => {
      logAudit(p.actorUserId, 'WORKFLOW_FIELD_CHANGED',
        'Workflow ' + p.workflowId + ' (' + row['Workflow Name'] + ') | ' +
        c.field + ': "' + c.oldVal + '" -> "' + c.newVal + '"');
      // Keep the onEdit guardrail's snapshot in sync so a later direct-sheet
      // edit compares against this new value, not the pre-update one.
      refreshWorkflowFieldSnapshotAfterApiEdit(p.workflowId, c.field, c.newVal);
    });

    return { success: true, changed: changes.length };
  },

  // ---------- MASTERS (write - Manager only, enforced client-side + here) ----------
  addStaff: function (p) {
    requireManager(p.actorRole);
    if (!p.staffName || !String(p.staffName).trim()) throw new Error('Staff Name is required');
    // CONCURRENCY FIX (live audit, Aug 2026): nextId() reads the current max
    // ID and appendRow() writes the new one - without a lock, two concurrent
    // addStaff calls (double-click, two Manager tabs/users) could both read
    // the same max and both generate/save the SAME auto Staff ID, silently
    // creating two Staff Master rows sharing one ID and corrupting every
    // staffId-keyed lookup (Working Register, Attendance, Leave, Appraisal).
    // Same LockService pattern as addWorkingEntries()/approveEntryGroup().
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const ids = readAll(SHEETS.STAFF).map(s => s['Staff ID']);
      const id = p.staffId || nextId('EMP', ids);
      // VALIDATION (audit item 23): reject a manually-supplied Staff ID that
      // already exists - previously an auto-generated ID (blank p.staffId) was
      // always unique, but a Manager typing in an existing ID silently created
      // a second Staff Master row sharing it, which then corrupts every
      // staffId-keyed lookup (Working Register, Attendance, Leave, Appraisal).
      if (p.staffId && ids.indexOf(p.staffId) !== -1) {
        throw new Error('Staff ID "' + p.staffId + '" already exists. Staff ID must be unique.');
      }
      appendRow(SHEETS.STAFF, {
        'Staff ID': id, 'Staff Name': p.staffName, 'Designation': p.designation,
        'Section': p.section, 'Reporting Manager': p.reportingManager,
        'Date of Joining': p.doj, 'Status': p.status || 'Active'
      });
      logAudit(p.actorUserId, 'ADD_STAFF', id);
      return { staffId: id };
    } finally {
      lock.releaseLock();
    }
  },

  updateStaff: function (p) {
    requireManager(p.actorRole);
    const staff = readAll(SHEETS.STAFF);
    const s = staff.find(x => x['Staff ID'] === p.staffId);
    if (!s) throw new Error('Staff not found');
    updateRow(SHEETS.STAFF, s._row, p.fields);
    logAudit(p.actorUserId, 'UPDATE_STAFF', p.staffId);
    return { success: true };
  },

  // ---------- SECTION MASTER (audit fix: was Google-Sheet-only editing) ----------
  // Same validated pattern as addStaff/updateStaff above - Manager-only,
  // server-side duplicate ID/name prevention, audit-logged. Sections are
  // referenced by Staff Master ('Section' column, free text match) and by
  // getMasters().sections for the Add/Edit Staff dropdown, so a stray
  // duplicate name here would silently split one section into two in
  // every Section-based report/filter.
  addSection: function (p) {
    requireManager(p.actorRole);
    if (!p.sectionName || !String(p.sectionName).trim()) throw new Error('Section Name is required');
    // CONCURRENCY FIX (live audit, Aug 2026) - same ID-collision race as
    // addStaff() above, wrapped the same way.
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sections = readAll(SHEETS.SECTION);
      const name = String(p.sectionName).trim();
      if (sections.some(s => String(s['Section Name']).trim().toLowerCase() === name.toLowerCase())) {
        throw new Error('Section "' + name + '" already exists.');
      }
      const ids = sections.map(s => s['Section ID']);
      const id = nextId('SEC', ids);
      appendRow(SHEETS.SECTION, { 'Section ID': id, 'Section Name': name });
      logAudit(p.actorUserId, 'ADD_SECTION', id + ' | ' + name);
      return { sectionId: id };
    } finally {
      lock.releaseLock();
    }
  },

  updateSection: function (p) {
    requireManager(p.actorRole);
    if (!p.sectionName || !String(p.sectionName).trim()) throw new Error('Section Name is required');
    // CONCURRENCY FIX (audit finding, Aug 2026): this had the same
    // duplicate-name check-then-write race that addSection() was already
    // locked against - two concurrent renames could both pass the
    // uniqueness check before either wrote, leaving two Sections with the
    // same name. Wrapped with the same LockService pattern used by
    // addSection/addActivity/addMonthlyTarget above.
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sections = readAll(SHEETS.SECTION);
      const s = sections.find(x => x['Section ID'] === p.sectionId);
      if (!s) throw new Error('Section not found');
      const name = String(p.sectionName).trim();
      if (sections.some(x => x['Section ID'] !== p.sectionId && String(x['Section Name']).trim().toLowerCase() === name.toLowerCase())) {
        throw new Error('Section "' + name + '" already exists.');
      }
      const oldName = s['Section Name'];
      updateRow(SHEETS.SECTION, s._row, { 'Section Name': name });
      logAudit(p.actorUserId, 'UPDATE_SECTION', p.sectionId + ' | "' + oldName + '" -> "' + name + '"');
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  },

  // ---------- ACTIVITY MASTER (audit fix: was Google-Sheet-only editing) ----------
  // Same pattern again. Activity ID is referenced by Workflow Master
  // ('Activity ID' column) and Monthly Targets, so duplicate-ID prevention
  // here matters even more than for Staff/Section.
  addActivity: function (p) {
    requireManager(p.actorRole);
    if (!p.activityName || !String(p.activityName).trim()) throw new Error('Activity Name is required');
    if (!p.activityCode || !String(p.activityCode).trim()) throw new Error('Activity Code is required');
    // CONCURRENCY FIX (live audit, Aug 2026) - same ID-collision race as
    // addStaff() above, wrapped the same way.
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const activities = readAll(SHEETS.ACTIVITY);
      const ids = activities.map(a => a['Activity ID']);
      const id = p.activityId || nextId('ACT', ids);
      if (p.activityId && ids.indexOf(p.activityId) !== -1) {
        throw new Error('Activity ID "' + p.activityId + '" already exists. Activity ID must be unique.');
      }
      if (activities.some(a => String(a['Activity Code']).trim().toLowerCase() === String(p.activityCode).trim().toLowerCase())) {
        throw new Error('Activity Code "' + p.activityCode + '" already exists.');
      }
      appendRow(SHEETS.ACTIVITY, {
        'Activity ID': id, 'Activity Code': String(p.activityCode).trim(), 'Activity Name': String(p.activityName).trim()
      });
      logAudit(p.actorUserId, 'ADD_ACTIVITY', id + ' | ' + p.activityName);
      return { activityId: id };
    } finally {
      lock.releaseLock();
    }
  },

  updateActivity: function (p) {
    requireManager(p.actorRole);
    if (!p.activityName || !String(p.activityName).trim()) throw new Error('Activity Name is required');
    // CONCURRENCY FIX (audit finding, Aug 2026): this had the same
    // duplicate-name check-then-write race that addActivity()/updateSection()
    // were already locked against - two concurrent renames could both pass
    // the uniqueness check before either wrote, leaving two Activities with
    // the same name. Wrapped with the same LockService pattern used by
    // addStaff/addSection/addActivity/updateSection above.
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const activities = readAll(SHEETS.ACTIVITY);
      const a = activities.find(x => x['Activity ID'] === p.activityId);
      if (!a) throw new Error('Activity not found');
      const newName = String(p.activityName).trim();
      // VALIDATION (audit finding, Aug 2026): updateSection() already
      // rejected a rename that collides with another row's name - updateActivity()
      // was missing this same check, so two Activities could silently end up
      // sharing one Activity Name.
      if (activities.some(x => x['Activity ID'] !== p.activityId && String(x['Activity Name']).trim().toLowerCase() === newName.toLowerCase())) {
        throw new Error('Activity "' + newName + '" already exists.');
      }
      const oldName = a['Activity Name'];
      updateRow(SHEETS.ACTIVITY, a._row, { 'Activity Name': newName });
      logAudit(p.actorUserId, 'UPDATE_ACTIVITY', p.activityId + ' | "' + oldName + '" -> "' + newName + '"');
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  },

  // ---------- MONTHLY TARGETS (Manager only for writes) ----------
  // A row = one Activity->Workflow target for one Month. Reuses the existing
  // Workflow Master and existing Working Register (Approved rows) - no parallel
  // "Daily Register" sheet is introduced, so nothing about the current KPI/
  // Appraisal system is touched.
  addMonthlyTarget: function (p) {
    // p: {actorRole, actorUserId, month, activityId, workflowId, target}
    requireManager(p.actorRole);
    const validatedTarget = _validateMonthlyTargetValue(p.target);
    const wf = readAll(SHEETS.WORKFLOW).find(w => w['Workflow ID'] === p.workflowId);
    if (!wf) throw new Error('Unknown workflow: ' + p.workflowId);
    // CONCURRENCY FIX (live audit, Aug 2026): both the existing-row lookup
    // (existing/new decision) and the ID-generation below were previously
    // unlocked - two concurrent addMonthlyTarget calls for the same
    // Workflow+Month could each see "no existing row" and both append,
    // creating two Target rows for the same Workflow+Month (silently
    // double-counting/ambiguous for every getEffectiveTarget() lookup), or
    // generate a colliding auto Target ID. Same LockService pattern as
    // addStaff()/addSection()/addActivity() above.
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const existing = readAll(SHEETS.TARGETS).find(t => normalizeMonthLabel(t['Month']) === p.month && t['Workflow ID'] === p.workflowId);
      if (existing) {
        const oldTarget = existing['Monthly Target'];
        updateRow(SHEETS.TARGETS, existing._row, { 'Monthly Target': validatedTarget });
        logAudit(p.actorUserId, 'UPDATE_MONTHLY_TARGET',
          existing['Target ID'] + ' | ' + p.workflowId + ' ' + p.month + ' | ' + oldTarget + ' -> ' + validatedTarget);
        return { targetId: existing['Target ID'], updated: true };
      }
      const ids = readAll(SHEETS.TARGETS).map(t => t['Target ID']);
      const id = nextId('MT', ids);
      appendRow(SHEETS.TARGETS, {
        'Target ID': id,
        'Month': p.month,
        'Activity ID': p.activityId,
        'Activity Name': wf['Activity Name'],
        'Unit': wf['Unit'],
        'Workflow ID': p.workflowId,
        'Workflow Name': wf['Workflow Name'],
        'KPI Name': wf['KPI Name'],
        'Monthly Target': validatedTarget,
        'Created By': p.actorUserId,
        'Created On': new Date()
      });
      logAudit(p.actorUserId, 'ADD_MONTHLY_TARGET', id);
      return { targetId: id, updated: false };
    } finally {
      lock.releaseLock();
    }
  },

  updateMonthlyTarget: function (p) {
    // p: {actorRole, actorUserId, targetId, target}
    requireManager(p.actorRole);
    const validatedTarget = _validateMonthlyTargetValue(p.target);
    const t = readAll(SHEETS.TARGETS).find(x => x['Target ID'] === p.targetId);
    if (!t) throw new Error('Monthly target not found');
    const oldTarget = t['Monthly Target'];
    updateRow(SHEETS.TARGETS, t._row, { 'Monthly Target': validatedTarget });
    // AUDIT DETAIL (spec section 23 - "Monthly Target update" must be
    // audited): old value + new value are recorded, not just the Target
    // ID, so an HR/Manager reviewing the Audit Log can see exactly what a
    // retroactive change would have shifted the live target to - even
    // though computeStaffWorkflowKpiGroups_() now protects already-
    // Approved rows from it (see the frozen-Target fix above), a Manager
    // reading the log still needs the before/after to judge intent.
    logAudit(p.actorUserId, 'UPDATE_MONTHLY_TARGET',
      p.targetId + ' | ' + t['Workflow ID'] + ' ' + t['Month'] + ' | ' + oldTarget + ' -> ' + validatedTarget);
    return { success: true };
  },

  deleteMonthlyTarget: function (p) {
    // p: {actorRole, actorUserId, targetId}
    requireManager(p.actorRole);
    const t = readAll(SHEETS.TARGETS).find(x => x['Target ID'] === p.targetId);
    if (!t) throw new Error('Monthly target not found');
    deleteRow(SHEETS.TARGETS, t._row);
    logAudit(p.actorUserId, 'DELETE_MONTHLY_TARGET', p.targetId);
    return { success: true };
  },

  getMonthlyTargets: function (p) {
    // p.month (optional, e.g. "Jul-2026")
    let rows = readAll(SHEETS.TARGETS);
    if (p && p.month) rows = rows.filter(r => normalizeMonthLabel(r['Month']) === p.month);
    return rows;
  },

  // Working Register KPI Balance:
  // Reuses getTargetStatus() as the single source of truth.
  // IMPORTANT: Monthly Target is never changed. This only exposes the
  // calculated Approved Qty and Remaining Target for the selected month.
  // Existing rules are preserved:
  //   - Approved rows only
  //   - Pending/Rejected/Draft do not reduce the balance
  //   - Team jobs are counted once using Team Ref No
  //   - Remaining is never negative
  getMonthlyWorkflowBalances: function (p) {
    const month = (p && p.month) || currentMonthLabel();
    return api.getTargetStatus({ month }).map(s => ({
      workflowId: s.workflowId,
      workflowName: s.workflowName,
      unit: s.unit,
      monthlyTarget: Number(s.monthlyTarget) || 0,
      approvedQty: Number(s.approvedQty) || 0,
      remaining: Math.max(Number(s.remaining) || 0, 0),
      targetSource: s.targetSource || 'default',
      progressPct: Number(s.progressPct) || 0
    }));
  },

  // Per-workflow status for a month: Monthly Target, Approved Qty (sum of Actual
  // from Approved Working Register rows that month, counting each team job once
  // via Team Ref No - see below), Remaining Target, Progress %.
  // Only Approved entries count - Pending/Rejected/Draft never affect the target,
  // matching the existing Working Register approval flow.
  getTargetStatus: function (p) {
    // p.month (optional, e.g. "Jul-2026")
    // p.onlyExplicit (optional, default false) - pass true to restore the
    // OLD behaviour of only listing workflows that already have a Manager-set
    // Monthly Target row for this month. Default is now every Active workflow,
    // each resolved through getEffectiveTarget (Monthly Target, else Default
    // KPI Target from "Expected Output"), so Target/Remaining/Progress % is
    // never blank for any workflow, matching every other page.
    const month = (p && p.month) || currentMonthLabel();
    const monthlyTargets = readAll(SHEETS.TARGETS).filter(t => normalizeMonthLabel(t['Month']) === month);
    const allWorkflows = readAll(SHEETS.WORKFLOW);
    const workflowsToShow = (p && p.onlyExplicit)
      ? allWorkflows.filter(wf => monthlyTargets.some(t => t['Workflow ID'] === wf['Workflow ID']))
      : allWorkflows.filter(wf => wf['Status'] !== 'Inactive');
    const approvedRows = readAll(SHEETS.REGISTER).filter(r =>
      r['Approval Status'] === 'Approved' && monthLabelOf(r['Date']) === month);

    return workflowsToShow.map(wf => {
      const mtRow = monthlyTargets.find(t => t['Workflow ID'] === wf['Workflow ID']);
      // Same Monthly Target -> Default KPI Target resolution used everywhere
      // else (Working Register entry/edit, Approval, Dashboard).
      const resolved = getEffectiveTarget(wf['Workflow ID'], month, { monthlyTargets, workflows: allWorkflows });

      // Team-based entries create one Working Register row PER PARTICIPANT (see
      // addWorkingEntries), all sharing the same Team Ref No + Workflow ID and
      // the SAME Actual quantity - it's one physical job, not a separate
      // quantity per person. Summing every participant's row would multiply
      // that job's quantity by the team size (e.g. 1 truck, 5 participants ->
      // wrongly counted as 5 toward the Monthly Target). Dedupe by Team Ref No
      // so each team job is counted exactly once.
      const seenTeamRefs = {};
      const approvedQty = approvedRows
        .filter(r => r['Workflow ID'] === wf['Workflow ID'])
        .reduce((sum, r) => {
          const ref = r['Team Ref No'];
          if (ref) {
            if (seenTeamRefs[ref]) return sum; // this job's quantity already counted
            seenTeamRefs[ref] = true;
          }
          return sum + (Number(r['Actual']) || 0);
        }, 0);
      const monthlyTarget = resolved.value;
      const remaining = Math.max(monthlyTarget - approvedQty, 0);
      // Capped at 100% (audit requirement) - Approved Qty exceeding Target
      // still shows Remaining = 0 above, but must not push Progress % past 100.
      const progressPct = monthlyTarget > 0 ? Math.min(100, Math.round((approvedQty / monthlyTarget) * 10000) / 100) : 0;
      // TARGET=0/NOT-SET FIX: an unconfigured target shows a distinct ⚪ "Not
      // Set" status instead of the misleading 🔴 (which reads as "badly
      // behind on a real target" - not the same thing as "no target exists").
      let status = resolved.configured ? '🔴' : '⚪';
      if (resolved.configured) {
        if (progressPct >= 75) status = '🟢';
        else if (progressPct >= 40) status = '🟡';
      }
      return {
        targetId: mtRow ? mtRow['Target ID'] : '', month,
        activityId: wf['Activity ID'], activityName: wf['Activity Name'], unit: wf['Unit'],
        workflowId: wf['Workflow ID'], workflowName: wf['Workflow Name'], kpiName: wf['KPI Name'],
        monthlyTarget, targetSource: resolved.source, targetConfigured: resolved.configured,
        approvedQty, remaining, progressPct, status
      };
    });
  },

  // Dashboard KPI cards for the Monthly Target module (section 6 of spec) - kept
  // entirely separate from getDashboard so the existing KPI/Appraisal dashboard
  // is untouched.
  getTargetDashboard: function (p) {
    const month = (p && p.month) || currentMonthLabel();
    const statuses = api.getTargetStatus({ month, onlyExplicit: p && p.onlyExplicit });
    const totalMonthlyTarget = statuses.reduce((s, x) => s + x.monthlyTarget, 0);
    const totalApprovedQty = statuses.reduce((s, x) => s + x.approvedQty, 0);
    const remainingTarget = Math.max(totalMonthlyTarget - totalApprovedQty, 0);
    const overallProgressPct = totalMonthlyTarget > 0 ? Math.round((totalApprovedQty / totalMonthlyTarget) * 10000) / 100 : 0;
    const completedWorkflows = statuses.filter(x => x.progressPct >= 100).length;
    const pendingWorkflows = statuses.length - completedWorkflows;
    return {
      month, totalMonthlyTarget, totalApprovedQty, remainingTarget, overallProgressPct,
      totalWorkflows: statuses.length, completedWorkflows, pendingWorkflows, statuses
    };
  },

  // ---------- WORKING REGISTER ----------
  // A "day's submission" is a set of Working Register line rows sharing EntryGroupId.
  // Normally one row per Workflow the staff member performed that day. For a
  // TEAM-based Workflow (its ID is in TEAM_WORKFLOW_IDS) where the staff member
  // ticks Co-Staff who worked the same job with them, ONE row per
  // participant is created instead (all sharing the same auto-generated Team
  // Ref No + Workflow ID), so the Manager can Approve/Reject each participant
  // individually - see approveEntryLine. Other lines in the same submission,
  // for workflows NOT in TEAM_WORKFLOW_IDS, stay solo even if they belong to
  // the same Activity.
  addWorkingEntries: function (p) {
    // p.staffId, p.staffName (the "Logged By" person, i.e. who is filling this form),
    // p.date,
    // p.lines: [{activityId, workflowId, unit, target, actual, remarks, coStaffIds}]
    // coStaffIds (optional, team-based Workflows only) = Staff IDs of OTHER staff
    // who worked the same job, in addition to p.staffId.

    // CONCURRENCY FIX (audit items 10 & 11): duplicate-check, Truck No
    // sequence generation, and row insertion now all happen inside ONE
    // LockService transaction. Previously the duplicate check ran unlocked
    // and generateTruckNo() took its own separate lock - two submissions
    // arriving at (almost) the same instant could both pass the duplicate
    // check before either had written a row, or both read the same "next"
    // Truck sequence number. Holding a single lock across all three steps
    // makes both races impossible; the lock is only released after every
    // row for this submission has been saved.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      return api._addWorkingEntriesLocked(p);
    } finally {
      lock.releaseLock();
    }
  },

  // Internal - only ever called from addWorkingEntries() above, already
  // holding the script lock. Do not call this directly / do not expose it
  // as an action name.
  _addWorkingEntriesLocked: function (p) {
    // DUPLICATE CONTROL: Team workflows intentionally allow multiple entries
    // on the same date. Since Truck No/Pending Truck tracking is no longer part
    // of Working Register, each submission gets its own Team Ref No and is a
    // separate job. Individual workflows retain the existing same-day duplicate
    // protection for the same employee.
    p.lines.forEach(line => {
      const isTeamCheck = TEAM_WORKFLOW_IDS.indexOf(line.workflowId) !== -1;
      if (isTeamCheck) return; // multiple same-day team jobs are valid
      const assignedCheck = (line.assignedStaffId && p.actorRole === 'Manager') ? line.assignedStaffId : p.staffId;
      const dup = findDuplicateTruckEntry(p.date, line.activityId, line.workflowId, [assignedCheck], '');
      if (dup) {
        throw new Error('Duplicate entry: a record already exists for the same Employee, Workflow, and Activity on this date.');
      }
    });

    const groupId = 'WR' + new Date().getTime();
    // Truck tracking is intentionally disabled in this version. Historical
    // Truck No values remain untouched, but all new Working Register rows leave
    // Truck No blank. Team grouping/KPI splitting uses Team Ref No only.
    const truckNo = '';
    const workflows = readAll(SHEETS.WORKFLOW);
    const monthlyTargets = readAll(SHEETS.TARGETS);
    const staffMaster = readAll(SHEETS.STAFF);
    const staffNameOf = (id) => {
      const s = staffMaster.find(x => x['Staff ID'] === id);
      return s ? s['Staff Name'] : id;
    };
    let lineSeq = 0;
    const newLineId = () => 'WL' + new Date().getTime() + '_' + (lineSeq++);
    const entryMonth = monthLabelOf(p.date);

    p.lines.forEach(line => {
      const wf = workflows.find(w => w['Workflow ID'] === line.workflowId);
      if (!wf) throw new Error('Unknown workflow: ' + line.workflowId);
      // VALIDATION (audit item 8) - server-side, never relies on the HTML
      // number input's min="0" alone. Target is re-resolved below, so it is
      // validated after that; Actual is client-supplied, validated here.
      if (line.actual !== undefined && line.actual !== '' && Number(line.actual) < 0) {
        throw new Error('Actual/Approved Quantity cannot be negative (Workflow ' + line.workflowId + ')');
      }
      if (!p.date) throw new Error('Date is required');
      // KPI Target = Monthly Target (if the Manager set one for this Month +
      // Workflow) else the Workflow/KPI Master's Default KPI Target
      // ("Expected Output") - see getEffectiveTarget(). SECURITY: the
      // client-supplied line.target is NEVER used, even as a fallback - a
      // manipulated browser request must not be able to inject a target
      // into the backend calculation. If neither a Monthly Target nor a
      // configured Workflow Master target exists, resolvedTarget.value is a
      // safe 0 and resolvedTarget.source is 'not_set' - see getEffectiveTarget().
      const resolvedTarget = getEffectiveTarget(line.workflowId, entryMonth, { monthlyTargets, workflows });
      const target = resolvedTarget.value;
      const actual = Number(line.actual) || 0;
      const achievementPct = computeAchievementPct(actual, target);
      const maxScore = Number(wf['Max Score']) || 0;
      const weightagePct = Number(wf['Weightage %']) || 0;
      // KPI Score = capped Achievement % applied to the workflow's Weightage %
      // (the share of its Activity's 100% that this workflow is worth).
      // Both may be null ("N/A") when the target isn't configured/is zero -
      // stored below as 'N/A' / 0 respectively (see append below).
      const kpiScore = computeKpiScore(achievementPct, weightagePct);

      const isTeam = TEAM_WORKFLOW_IDS.indexOf(line.workflowId) !== -1;
      const coStaffIds = isTeam && Array.isArray(line.coStaffIds) ? line.coStaffIds.filter(id => id && id !== p.staffId) : [];
      // Dedupe, always including the person who logged the entry.
      let participants = isTeam ? Array.from(new Set([p.staffId].concat(coStaffIds))) : [p.staffId];

      // ASSIGNED STAFF (individual/non-team workflows only) - lets a Manager
      // credit the KPI Score to a different staff member than whoever is
      // logged in and filling the form (e.g. Manager logging on behalf of a
      // staff member who does not use the system). This is intentionally
      // separate from 'Logged By ID'/'Logged By Name' (still always the
      // trusted session user - see SELF_SCOPED_ACTIONS in handle()) so the
      // audit trail always shows who actually submitted the entry, while
      // 'Staff ID'/'Staff Name' (who gets the KPI credit) can differ.
      // Only honoured when the trusted session role is Manager; a Staff
      // user's assignedStaffId is silently ignored so nobody can credit (or
      // steal credit from) another staff member via a crafted request.
      if (!isTeam && line.assignedStaffId && p.actorRole === 'Manager') {
        const assignedExists = staffMaster.some(s => s['Staff ID'] === line.assignedStaffId);
        if (!assignedExists) throw new Error('Unknown Staff ID: ' + line.assignedStaffId);
        participants = [line.assignedStaffId];
      }
      // Auto-generated shared reference - staff no longer type this in, which
      // removes the risk of a typo/capitalisation mismatch splitting a truck's
      // team into two separate (wrong) groups.
      const teamRefNo = (isTeam && participants.length > 1) ? ('TS' + new Date().getTime() + '_' + lineSeq) : '';

      participants.forEach(pid => {
        appendRow(SHEETS.REGISTER, {
          'Line ID': newLineId(),
          'Entry Group ID': groupId,
          'Date': p.date,
          'Truck No': '',
          'Staff ID': pid,
          'Staff Name': staffNameOf(pid),
          'Logged By ID': p.staffId,
          'Logged By Name': p.staffName,
          'Activity ID': line.activityId,
          'Activity Name': wf['Activity Name'],
          'Team Ref No': teamRefNo,
          'Workflow ID': line.workflowId,
          'Workflow Name': wf['Workflow Name'],
          'KPI Name': wf['KPI Name'],
          'Unit': line.unit || wf['Unit'],
          'Target': target,
          'Target Source': resolvedTarget.source,
          'Actual': actual,
          'Achievement %': achievementPct === null ? 'N/A' : achievementPct,
          'Weightage %': weightagePct,
          'Max Score': maxScore,
          'KPI Score': kpiScore === null ? 0 : kpiScore,
          'Remarks': line.remarks || '',
          'Approval Status': 'Draft',
          'Submitted On': '',
          'Approved By': '',
          'Approved On': '',
          'Rejected By': '',
          'Rejected On': '',
          'Rejection Reason': ''
        });
      });
    });
    logAudit(p.staffId, 'ADD_WORKING_ENTRIES', groupId + ' (' + p.lines.length + ' line(s)' + (truckNo ? ', Truck ' + truckNo : '') + ')');
    return { groupId: groupId, truckNo: truckNo };
  },

  // ============================================================
  // APPROVAL WORKFLOW STATE MACHINE (P1 fix)
  // ============================================================
  // Allowed transitions ONLY:
  //   Draft     --submit-->            Submitted
  //   Submitted --approve/reject-->    Approved / Rejected
  // Everything else (Draft->Approve/Reject directly, Approved/Rejected->
  // anything, re-approving/re-rejecting an already-decided line, etc.) is
  // blocked HERE in the backend - not just by disabling a frontend button.
  // A disabled button stops an honest user in the UI; it does nothing to
  // stop a direct API call. Every write path below re-checks the row's
  // CURRENT status immediately before writing, so an invalid transition
  // is impossible however the request was made.
  APPROVAL_TRANSITIONS: { Draft: ['Submitted'], Submitted: ['Approved', 'Rejected'] },
  // FINAL LIVE AUDIT FIX: this and the two helpers below (_buildApprovalFields,
  // _notifyRejectedStaff) are internal-only - never meant to be reachable as a
  // top-level API action. They previously had no leading underscore, so
  // handle()'s `action.charAt(0) === '_'` guard did NOT block them; any
  // authenticated session (any role) could call them directly by action name.
  // Renamed with a leading underscore, matching every other internal helper
  // in this file (_approveEntryGroupLocked, _addWorkingEntriesLocked, etc.),
  // so they are now correctly blocked by that same guard. No behavior change
  // for legitimate callers - only api.* call sites, which were updated too.
  _assertTransition: function (currentStatus, targetStatus, contextLabel) {
    const allowed = (api.APPROVAL_TRANSITIONS[currentStatus] || []).indexOf(targetStatus) !== -1;
    if (!allowed) {
      throw new Error('Invalid transition' + (contextLabel ? ' for ' + contextLabel : '') +
        ': cannot move from "' + currentStatus + '" to "' + targetStatus + '".');
    }
  },

  submitEntryGroup: function (p) {
    // CONCURRENCY FIX (live audit, Aug 2026): this read-then-write state
    // transition (Draft -> Submitted) previously ran with no lock, unlike
    // every sibling entry-group action right below it (approveEntryGroup,
    // updateEntryGroup, deleteEntryGroup) which all wrap the same
    // read-check-write pattern in LockService. A submitEntryGroup racing a
    // concurrent updateEntryGroup/deleteEntryGroup call on the same group
    // (e.g. double-click, or a staff editing a draft in one tab while
    // submitting from another) could interleave and leave rows in an
    // inconsistent Approval Status. Wrapped the same way.
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      // moves all rows for a groupId from Draft -> Submitted
      const rows = readAll(SHEETS.REGISTER).filter(r => r['Entry Group ID'] === p.groupId);
      if (rows.length === 0) throw new Error('Entry group not found');
      // OWNERSHIP CHECK (audit item 4) - p.actorStaffId is the server-verified
      // logged-in staff, never the client-editable p.staffId. A Staff user may
      // only submit their own Draft; Manager may submit on anyone's behalf.
      if (p.actorRole !== 'Manager' && rows[0]['Logged By ID'] !== p.actorStaffId) {
        throw new Error('Not authorized to submit this entry');
      }
      // STATE VALIDATION: every row in the group must currently be Draft -
      // blocks re-submitting an already-Submitted/Approved/Rejected group.
      rows.forEach(r => api._assertTransition(r['Approval Status'], 'Submitted', 'Entry Group ' + p.groupId));
      rows.forEach(r => updateRow(SHEETS.REGISTER, r._row, {
        'Approval Status': 'Submitted', 'Submitted On': new Date()
      }));
      logAudit(p.actorStaffId, 'SUBMIT_ENTRY_GROUP', p.groupId);
      return { success: true, count: rows.length };
    } finally {
      lock.releaseLock();
    }
  },

  // Same whole-transaction locking pattern as approveEntryLine/approveEntryLines
  // above - the group is re-read fresh AFTER the lock is acquired, so the
  // Submitted-state check can never run against stale pre-lock data.
  approveEntryGroup: function (p) {
    // p.decision: 'Approved' | 'Rejected' - decides EVERY line in the group the
    // same way. Fine for a solo (non-team) submission; for a team submission
    // with several participants, prefer approveEntryLine/approveEntryLines so a
    // single wrong participant can be rejected without affecting the others.
    // p.reason (optional): stored as Rejection Reason when decision='Rejected'.
    requireManager(p.actorRole);
    if (p.decision !== 'Approved' && p.decision !== 'Rejected') throw new Error('decision must be "Approved" or "Rejected"');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      return api._approveEntryGroupLocked(p);
    } finally {
      lock.releaseLock();
    }
  },

  // Internal - only ever called from approveEntryGroup() above, already
  // holding the script lock. Do not call this directly / do not expose it
  // as an action name.
  _approveEntryGroupLocked: function (p) {
    const rows = readAll(SHEETS.REGISTER).filter(r => r['Entry Group ID'] === p.groupId); // fresh read, inside the lock
    if (rows.length === 0) throw new Error('Entry group not found');
    // STATE VALIDATION: only rows currently "Submitted" may be decided -
    // skip (never silently re-decide) any row already Approved/Rejected or
    // still Draft, and report exactly how many were actually acted on.
    const decidable = rows.filter(r => r['Approval Status'] === 'Submitted');
    if (decidable.length === 0) {
      throw new Error('No Submitted lines to ' + p.decision.toLowerCase() + ' in this entry group (already decided, or still Draft).');
    }
    decidable.forEach(r => updateRow(SHEETS.REGISTER, r._row, api._buildApprovalFields(p.decision, p.actorUserId, p.reason)));
    logAudit(p.actorUserId, p.decision.toUpperCase() + '_ENTRY_GROUP', p.groupId +
      (rows.length !== decidable.length ? ' (' + (rows.length - decidable.length) + ' line(s) skipped - not in Submitted state)' : ''));
    if (p.decision === 'Rejected') api._notifyRejectedStaff(decidable, p.reason);
    return { success: true, count: decidable.length, skipped: rows.length - decidable.length };
  },

  // Notifies each distinct staff member credited on a set of just-rejected
  // Working Register rows. Best-effort/non-blocking - a lookup failure for
  // one staff member must never fail the approval action itself. reason is
  // passed explicitly (not re-read from the row) since these rows were read
  // BEFORE the update that just wrote Rejection Reason to the sheet.
  _notifyRejectedStaff: function (rejectedRows, reason) {
    const notified = {};
    rejectedRows.forEach(r => {
      const staffId = r['Staff ID'];
      if (!staffId || notified[staffId]) return;
      notified[staffId] = true;
      try {
        const uid = userIdForStaff_(staffId);
        if (uid) {
          addNotification_({ toUserId: uid, type: 'ENTRY_REJECTED', refId: r['Entry Group ID'] || r['Line ID'],
            message: 'Your working register entry (' + (r['Workflow Name'] || r['Line ID']) + ') was rejected' +
              (reason ? ': ' + reason : '') });
        }
      } catch (e) { /* never block the approval action on a notification failure */ }
    });
  },

  // Builds the field set written for an Approve/Reject decision. Approved
  // and Rejected each get their OWN By/On columns (Rejected By/On/Reason
  // are new - see migrateAddApprovalAuditFields) instead of both decisions
  // sharing 'Approved By'/'Approved On', which previously mislabeled a
  // rejection as if a Manager had "approved" it into a rejected state.
  _buildApprovalFields: function (decision, actorUserId, reason) {
    const now = new Date();
    if (decision === 'Rejected') {
      return { 'Approval Status': 'Rejected', 'Rejected By': actorUserId, 'Rejected On': now, 'Rejection Reason': reason || '' };
    }
    return { 'Approval Status': 'Approved', 'Approved By': actorUserId, 'Approved On': now };
  },

  // Decide ONE participant's line within a submission (identified by its
  // unique Line ID) without touching any other participant's line - this is
  // what lets a Manager reject one wrong entry (e.g. one co-staff on a 5-person
  // truck) while approving the other four normally.
  // ATOMIC LOCKING (audit hardening): acquires the script-wide lock BEFORE
  // reading/validating/writing anything, so two concurrent Approve/Reject
  // requests for the same line can never both pass the Submitted-state
  // check before either has written its decision. Re-reads the row fresh
  // AFTER the lock is held (never trusts data read before acquiring it),
  // so the state check inside the lock always sees the latest value.
  approveEntryLine: function (p) {
    // p: {actorRole, actorUserId, lineId, decision, reason (optional, for Rejected)}
    requireManager(p.actorRole);
    if (p.decision !== 'Approved' && p.decision !== 'Rejected') throw new Error('decision must be "Approved" or "Rejected"');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      return api._approveEntryLineLocked(p);
    } finally {
      lock.releaseLock();
    }
  },

  // Internal - only ever called from approveEntryLine() above, already
  // holding the script lock. Do not call this directly / do not expose it
  // as an action name.
  _approveEntryLineLocked: function (p) {
    const all = readAll(SHEETS.REGISTER); // fresh read, inside the lock
    const row = all.find(r => r['Line ID'] === p.lineId);
    if (!row) throw new Error('Entry line not found');
    // STATE VALIDATION: this exact line must currently be Submitted.
    api._assertTransition(row['Approval Status'], p.decision, 'Line ' + p.lineId);
    updateRow(SHEETS.REGISTER, row._row, api._buildApprovalFields(p.decision, p.actorUserId, p.reason));
    logAudit(p.actorUserId, p.decision.toUpperCase() + '_ENTRY_LINE', p.lineId);
    if (p.decision === 'Rejected') api._notifyRejectedStaff([row], p.reason);
    return { success: true };
  },

  // Same as approveEntryLine but for several Line IDs at once with the same
  // decision - used for the "Approve All" convenience button, which applies to
  // every line currently shown (i.e. still Submitted) in that card. Same
  // whole-transaction locking pattern as approveEntryLine above.
  approveEntryLines: function (p) {
    // p: {actorRole, actorUserId, lineIds: [...], decision, reason (optional)}
    requireManager(p.actorRole);
    if (p.decision !== 'Approved' && p.decision !== 'Rejected') throw new Error('decision must be "Approved" or "Rejected"');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      return api._approveEntryLinesLocked(p);
    } finally {
      lock.releaseLock();
    }
  },

  // Internal - only ever called from approveEntryLines() above, already
  // holding the script lock. Do not call this directly / do not expose it
  // as an action name.
  _approveEntryLinesLocked: function (p) {
    const all = readAll(SHEETS.REGISTER); // fresh read, inside the lock
    let count = 0, skipped = 0;
    const decidedRows = [];
    (p.lineIds || []).forEach(lineId => {
      const row = all.find(r => r['Line ID'] === lineId);
      if (!row) return;
      // STATE VALIDATION: silently skip (never force) any line that isn't
      // currently Submitted, so a stale "Approve All" click on a card that
      // changed underneath it can't re-decide an already-decided line.
      if (row['Approval Status'] !== 'Submitted') { skipped++; return; }
      updateRow(SHEETS.REGISTER, row._row, api._buildApprovalFields(p.decision, p.actorUserId, p.reason));
      count++;
      if (p.decision === 'Rejected') decidedRows.push(row);
    });
    logAudit(p.actorUserId, p.decision.toUpperCase() + '_ENTRY_LINES', (p.lineIds || []).join(',') +
      (skipped ? ' (' + skipped + ' skipped - not Submitted)' : ''));
    if (decidedRows.length) api._notifyRejectedStaff(decidedRows, p.reason);
    return { success: true, count, skipped };
  },

  getWorkingRegister: function (p) {
    // p.staffId (optional - Staff role only sees own), p.from, p.to, p.status
    const all = readAll(SHEETS.REGISTER);
    // Team split group size must be computed against the FULL Approved set,
    // not the filtered slice being returned, or a teammate's Approved row
    // outside this filter would be missed and the split would be wrong.
    const approvedAll = all.filter(r => r['Approval Status'] === 'Approved');
    let rows = all;
    if (p.staffId) rows = rows.filter(r => r['Staff ID'] === p.staffId);
    if (p.from) rows = rows.filter(r => new Date(r['Date']) >= startOfDay(p.from));
    if (p.to) rows = rows.filter(r => new Date(r['Date']) <= endOfDay(p.to));
    if (p.status) rows = rows.filter(r => r['Approval Status'] === p.status);
    return withTeamSplit(rows, approvedAll);
  },

  getPendingApprovals: function (p) {
    // AUTH FIX (live audit, Aug 2026): previously took no params at all, so
    // actorRole was never even available to check, and its call to
    // getEntryGroups({status:'Submitted'}) carried no staffId - a direct
    // action:'getPendingApprovals' call by ANY logged-in Staff user (this
    // action was never added to STAFF_SCOPED_READ_ACTIONS, since it's meant
    // to be Manager-only) returned every staff member's Submitted entries
    // department-wide. Gated the same as the 'approvals' nav tab
    // (NAV_ITEMS: roles:['Manager']) this exists to serve.
    requireManager(p.actorRole);
    return api.getEntryGroups({ status: 'Submitted' });
  },

  // Generic grouped fetch, reused for the Manager's "Draft Entries" / "Submitted
  // Entries" tabs and for a Staff member's own "My Draft Entries" list.
  // p.status ('Draft'|'Submitted'|'Approved'|'Rejected', optional), p.staffId (optional
  // - matches who LOGGED the entry, i.e. filled the form, not necessarily every staff
  // credited on it - a team submission credits several staff but only one logs it).
  getEntryGroups: function (p) {
    const all = readAll(SHEETS.REGISTER);
    const approvedAll = all.filter(r => r['Approval Status'] === 'Approved');
    let rows = all;
    if (p && p.status) rows = rows.filter(r => r['Approval Status'] === p.status);
    if (p && p.staffId) rows = rows.filter(r => r['Logged By ID'] === p.staffId);
    rows = withTeamSplit(rows, approvedAll);
    const groups = {};
    rows.forEach(r => {
      if (!groups[r['Entry Group ID']]) {
        groups[r['Entry Group ID']] = {
          groupId: r['Entry Group ID'],
          loggedById: r['Logged By ID'], loggedByName: r['Logged By Name'],
          date: r['Date'], status: r['Approval Status'], truckNo: r['Truck No'], lines: []
        };
      }
      groups[r['Entry Group ID']].lines.push(r);
    });
    return Object.values(groups);
  },

  // Edit the Actual/Remarks of an existing Draft or Submitted group (Manager
  // "Edit" action, or a Staff member correcting their own Draft before submitting).
  // Recomputes Achievement % and KPI Score the same way addWorkingEntries does.
  // p: {actorRole, actorUserId, staffId, groupId, lines: [{workflowId, actual, remarks}]}
  updateEntryGroup: function (p) {
    const rows = readAll(SHEETS.REGISTER).filter(r => r['Entry Group ID'] === p.groupId);
    if (rows.length === 0) throw new Error('Entry group not found');
    const status = rows[0]['Approval Status'];
    if (status === 'Approved' || status === 'Rejected') {
      throw new Error('Cannot edit an entry that has already been ' + status.toLowerCase());
    }
    if (p.actorRole !== 'Manager' && rows[0]['Logged By ID'] !== p.actorStaffId) {
      throw new Error('Not authorized to edit this entry');
    }
    const workflows = readAll(SHEETS.WORKFLOW);
    const monthlyTargets = readAll(SHEETS.TARGETS);
    (p.lines || []).forEach(line => {
      // A team-based workflow has ONE row per participant sharing the same job -
      // apply the edited Actual/Remarks to all of them, not just the first match,
      // so every participant's row stays consistent with the real outcome.
      const matchingRows = rows.filter(r => r['Workflow ID'] === line.workflowId);
      if (matchingRows.length === 0) return;
      const wf = workflows.find(w => w['Workflow ID'] === line.workflowId);
      // Re-resolve the Target the same way addWorkingEntries does (Monthly
      // Target for that row's Month + Workflow ID, else the Default KPI
      // Target), instead of trusting the value stored when the row was first
      // created - this keeps an edited entry in sync if the Monthly Target
      // was added/changed/removed afterwards, and matches every other page.
      const entryMonth = monthLabelOf(matchingRows[0]['Date']);
      const resolvedTarget = getEffectiveTarget(line.workflowId, entryMonth, { monthlyTargets, workflows });
      // SECURITY: same resolution as addWorkingEntries - see fix note there.
      // Never fall back to the previously-stored row value; always
      // re-resolve from the Monthly Target / Workflow Master so an edit
      // stays in sync if targets changed, and a not-configured target is
      // always 0/'not_set', never a stale number.
      const target = resolvedTarget.value;
      const actual = Number(line.actual) || 0;
      const achievementPct = computeAchievementPct(actual, target);
      const weightagePct = Number(wf ? wf['Weightage %'] : matchingRows[0]['Weightage %']) || 0;
      const kpiScore = computeKpiScore(achievementPct, weightagePct);
      matchingRows.forEach(row => {
        updateRow(SHEETS.REGISTER, row._row, {
          'Target': target,
          'Target Source': resolvedTarget.source,
          'Actual': actual,
          'Achievement %': achievementPct === null ? 'N/A' : achievementPct,
          'Weightage %': weightagePct,
          'KPI Score': kpiScore === null ? 0 : kpiScore,
          'Remarks': line.remarks !== undefined ? line.remarks : row['Remarks']
        });
      });
    });
    logAudit(p.actorUserId || p.staffId, 'UPDATE_ENTRY_GROUP', p.groupId);
    return { success: true };
  },

  // Delete a whole Draft (own staff, or Manager). Never allowed once Submitted/
  // Approved/Rejected, to keep the audit trail intact.
  // CONCURRENCY FIX (audit follow-up): wrapped in the same LockService
  // transaction pattern as addWorkingEntries. Without this, a delete
  // running at (almost) the same instant as an updateEntryGroup edit on the
  // same groupId could read a stale row snapshot and delete/keep rows based
  // on it - both actions now serialize against each other.
  deleteEntryGroup: function (p) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      return api._deleteEntryGroupLocked(p);
    } finally {
      lock.releaseLock();
    }
  },

  // Internal - only ever called from deleteEntryGroup() above, already
  // holding the script lock. Do not call this directly / do not expose it
  // as an action name.
  _deleteEntryGroupLocked: function (p) {
    const rows = readAll(SHEETS.REGISTER).filter(r => r['Entry Group ID'] === p.groupId);
    if (rows.length === 0) throw new Error('Entry group not found');
    if (rows[0]['Approval Status'] !== 'Draft') throw new Error('Only Draft entries can be deleted');
    if (p.actorRole !== 'Manager' && rows[0]['Logged By ID'] !== p.actorStaffId) {
      throw new Error('Not authorized to delete this entry');
    }
    rows.sort((a, b) => b._row - a._row).forEach(r => deleteRow(SHEETS.REGISTER, r._row));
    logAudit(p.actorUserId || p.staffId, 'DELETE_ENTRY_GROUP', p.groupId);
    return { success: true };
  },

  // ---------- ATTENDANCE ----------
  // DUPLICATE CONTROL (audit item 13): unique key is Date + Staff ID. If a
  // record already exists for that day, this UPDATES it in place instead of
  // appending a second row - never two active attendance rows for the same
  // employee on the same date. Locked so two near-simultaneous submissions
  // for the same day can't both pass the "does it exist" check and both
  // append a fresh row.
  markAttendance: function (p) {
    // p: {staffId, staffName, date, status: Present/Absent/Half Day/Permission/Late, permissionHours, remarks}
    if (!p.date) throw new Error('Date is required');
    // STATUS VALIDATION FIX: previously any string was accepted here and
    // written straight to the sheet. If the sheet's Status column has a
    // strict dropdown (data validation) that doesn't include every value
    // the frontend can send, the save could be silently rejected by
    // Sheets with no clear error reaching the user. Validating here first
    // means the user always gets a clear message instead of a mystery
    // failure, regardless of the sheet's dropdown state.
    if (ATTENDANCE_STATUS_VALUES.indexOf(p.status) === -1) {
      throw new Error('Status must be one of: ' + ATTENDANCE_STATUS_VALUES.join(', '));
    }
    if (p.permissionHours !== undefined && p.permissionHours !== '' && Number(p.permissionHours) < 0) {
      throw new Error('Permission hours cannot be negative');
    }
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const tz = Session.getScriptTimeZone();
      const dayLabel = Utilities.formatDate(new Date(p.date), tz, 'yyyy-MM-dd');
      const existing = readAll(SHEETS.ATTENDANCE).find(r =>
        r['Staff ID'] === p.staffId && r['Date'] &&
        Utilities.formatDate(new Date(r['Date']), tz, 'yyyy-MM-dd') === dayLabel);
      if (existing) {
        updateRow(SHEETS.ATTENDANCE, existing._row, {
          'Status': p.status, 'Permission Hours': p.permissionHours || 0, 'Remarks': p.remarks || ''
        });
        logAudit(p.staffId, 'UPDATE_ATTENDANCE', p.date + ' ' + p.status + ' (corrected existing record)');
        return { success: true, updated: true };
      }
      appendRow(SHEETS.ATTENDANCE, {
        'Date': p.date, 'Staff ID': p.staffId, 'Staff Name': p.staffName,
        'Status': p.status, 'Permission Hours': p.permissionHours || 0, 'Remarks': p.remarks || ''
      });
      logAudit(p.staffId, 'MARK_ATTENDANCE', p.date + ' ' + p.status);
      return { success: true, updated: false };
    } finally {
      lock.releaseLock();
    }
  },

  getAttendance: function (p) {
    let rows = readAll(SHEETS.ATTENDANCE);
    if (p.staffId) rows = rows.filter(r => r['Staff ID'] === p.staffId);
    if (p.from) rows = rows.filter(r => new Date(r['Date']) >= startOfDay(p.from));
    if (p.to) rows = rows.filter(r => new Date(r['Date']) <= endOfDay(p.to));
    return rows;
  },

  // ---------- LEAVE ----------
  // OVERLAP CONTROL (audit item 14): a new leave request is rejected if its
  // From-To range overlaps any of that staff member's existing Pending or
  // Approved leave records (a Rejected/Cancelled one is ignored, so a
  // resubmission after rejection is unaffected). Locked so two
  // near-simultaneous requests for overlapping ranges can't both pass the
  // check before either is saved.
  applyLeave: function (p) {
    if (!p.fromDate || !p.toDate) throw new Error('From Date and To Date are required');
    const from = new Date(p.fromDate), to = new Date(p.toDate);
    if (to < from) throw new Error('To Date cannot be before From Date');
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const overlap = readAll(SHEETS.LEAVE).find(l =>
        l['Staff ID'] === p.staffId &&
        (l['Status'] === 'Pending' || l['Status'] === 'Approved') &&
        new Date(l['From Date']) <= to && new Date(l['To Date']) >= from);
      if (overlap) {
        throw new Error('This overlaps an existing ' + overlap['Status'].toLowerCase() +
          ' leave (' + overlap['Leave ID'] + ') for this employee. Please review the existing record instead.');
      }
      const ids = readAll(SHEETS.LEAVE).map(l => l['Leave ID']);
      const id = nextId('LV', ids);
      appendRow(SHEETS.LEAVE, {
        'Leave ID': id, 'Staff ID': p.staffId, 'Staff Name': p.staffName,
        'From Date': p.fromDate, 'To Date': p.toDate, 'Leave Type': p.leaveType,
        'Reason': p.reason, 'Status': 'Pending', 'Applied On': new Date()
      });
      logAudit(p.staffId, 'APPLY_LEAVE', id);
      return { leaveId: id };
    } finally {
      lock.releaseLock();
    }
  },

  // P2 FIX #5 (audit): backend state validation was missing entirely - any
  // decision string was accepted and written regardless of the leave's
  // current status, with no re-check for a concurrent decision. State
  // machine: Pending -> Approved / Rejected only. Approved/Rejected are
  // terminal - re-deciding an already-decided leave (in either direction,
  // or the same direction again) is rejected. Locked + fresh re-read
  // inside the lock, same pattern as applyLeave()/approveEntryGroup(), so
  // two near-simultaneous decisions on the same leave can't both succeed.
  decideLeave: function (p) {
    requireManager(p.actorRole);
    if (p.decision !== 'Approved' && p.decision !== 'Rejected') {
      throw new Error('decision must be "Approved" or "Rejected"');
    }
    if (!p.leaveId) throw new Error('leaveId is required');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const l = readAll(SHEETS.LEAVE).find(x => x['Leave ID'] === p.leaveId); // fresh read, inside the lock
      if (!l) throw new Error('Leave not found');
      if (l['Status'] !== 'Pending') {
        throw new Error('This leave request has already been ' + String(l['Status'] || '').toLowerCase() +
          ' and cannot be decided again.');
      }
      updateRow(SHEETS.LEAVE, l._row, { 'Status': p.decision });
      logAudit(p.actorUserId, p.decision.toUpperCase() + '_LEAVE', p.leaveId);
      return { success: true };
    } finally {
      lock.releaseLock();
    }
  },

  getLeaves: function (p) {
    let rows = readAll(SHEETS.LEAVE);
    if (p.staffId) rows = rows.filter(r => r['Staff ID'] === p.staffId);
    return rows;
  },

  // ---------- DASHBOARD / APPRAISAL ----------
  getDashboard: function (p) {
    const from = p.from ? startOfDay(p.from) : new Date(new Date().getFullYear(), new Date().getMonth(), 1);
    const to = p.to ? endOfDay(p.to) : new Date();

    // Current-period breakdown (byStaff/byActivity/byWorkflow/bySection) +
    // TOTAL KPI % (MTD) card, via the shared Approved-base helper below.
    const current = computeDashboardKpiForRange_(from, to, p.staffId);

    // UPTREND (Aug 2026): compare this period's Approved-base TOTAL KPI %
    // against the immediately preceding period of the SAME length (e.g. if
    // "from..to" is 1-21 Aug, the previous period is 11-31 Jul - 21 days
    // ending the instant before "from"). Uses the exact same Approved-only,
    // fixed-denominator computeDashboardKpiForRange_() helper as the current
    // period, so the comparison is apples-to-apples (same "approved base"
    // methodology on both sides) - this is what makes the arrow trustworthy
    // rather than comparing two differently-scoped numbers.
    const durationMs = to.getTime() - from.getTime();
    const prevTo = endOfDay(new Date(from.getTime() - 1));
    const prevFrom = startOfDay(new Date(prevTo.getTime() - durationMs));
    const previous = computeDashboardKpiForRange_(prevFrom, prevTo, p.staffId);

    // No trend if there's no Approved-base activity in the prior period to
    // compare against (Max Score 0 = staff had no active workflows / no
    // Approved rows then) - showing "▼ -12%" against a true zero would be
    // misleading, so the frontend gets trend:null and hides the arrow.
    let trend = null;
    if (previous.totalMaxScore > 0) {
      const change = Math.round((current.totalKpiPct - previous.totalKpiPct) * 100) / 100;
      trend = {
        direction: change > 0 ? 'up' : (change < 0 ? 'down' : 'flat'),
        change,                                   // percentage-point change, e.g. +3.4
        previousKpiPct: previous.totalKpiPct,
        previousFrom: prevFrom.toISOString(),
        previousTo: prevTo.toISOString()
      };
    }

    const pending = readAll(SHEETS.REGISTER).filter(r => r['Approval Status'] === 'Submitted').length;

    return {
      totalEntries: current.totalEntries,
      totalKpiScore: current.totalKpiScore,      // kept for any code/report still reading it
      totalKpiPct: current.totalKpiPct,        // TOTAL KPI % (MTD) - what the Dashboard card should show
      trend,                                   // { direction, change, previousKpiPct } or null
      pendingApprovals: pending,
      byStaff: current.byStaff, byActivity: current.byActivity, byWorkflow: current.byWorkflow, bySection: current.bySection
    };
  },

  getAppraisal: function (p) {
    // AUTH FIX (live audit, Aug 2026): this returns EVERY staff member's
    // aggregated KPI score/rating/section for the requested range - the
    // frontend only ever reaches it from the Manager-only 'appraisal' nav
    // tab (NAV_ITEMS: roles:['Manager']), but that is UI convenience only,
    // same as every other cross-staff aggregate in this file. Before this
    // fix there was NO server-side check at all, so any logged-in Staff
    // user could call action:'getAppraisal' directly (valid session token,
    // forged request) and pull the whole department's appraisal data -
    // exactly the class of bug already fixed for getHRLiveDashboard/
    // getHRManagementSummary/getHRDashboard above. Matches the UI's
    // Manager-only gating.
    requireManager(p.actorRole);
    // p.from, p.to (custom range only, per spec). Returns per-staff aggregated appraisal.
    // SINGLE-SOURCE FIX (Option B audit, Aug 2026): this used to sum each
    // Approved row's own 'Effective KPI Score'/'Effective Weightage %'
    // straight off the Working Register - the exact per-row dilution bug
    // already fixed everywhere else (Dashboard/Stores KPI/My Score/HR Final
    // Score) via computeStaffWorkflowKpiGroups_() + the Fixed Denominator
    // (computeStaffFixedActiveWorkflowMaxScore_()). Left un-migrated, a
    // staff with 2+ Approved rows for the same Workflow+Month (or a
    // Team-Split workflow) got double-scored against the full Target here,
    // silently drifting this Appraisal tab away from every other KPI %
    // shown in the app for the same Staff/period. Now routed through the
    // same two authoritative helpers as everything else, so
    // Appraisal % = Dashboard % = Stores KPI % = My Score % = HR Final
    // Score's Stores KPI input, for the same Staff/period, with no second
    // formula anywhere.
    const approvedInRange = readAll(SHEETS.REGISTER).filter(r => {
      const d = new Date(r['Date']);
      const inRange = (!p.from || d >= startOfDay(p.from)) && (!p.to || d <= endOfDay(p.to));
      return inRange && r['Approval Status'] === 'Approved' &&
        STORES_KPI_EXCLUDED_ACTIVITY_IDS.indexOf(r['Activity ID']) === -1;
    });
    const staffMaster = readAll(SHEETS.STAFF);
    const workflows = readAll(SHEETS.WORKFLOW);
    const monthlyTargets = readAll(SHEETS.TARGETS);

    const rowsByStaffId = {};
    approvedInRange.forEach(r => { (rowsByStaffId[r['Staff ID']] = rowsByStaffId[r['Staff ID']] || []).push(r); });

    const agg = {};
    Object.keys(rowsByStaffId).forEach(staffId => {
      const staffRows = rowsByStaffId[staffId];
      const sm = staffMaster.find(s => s['Staff ID'] === staffId);
      // groupWithin MUST be the full approvedInRange (not staffId-filtered),
      // same contract as every other computeStaffWorkflowKpiGroups_ call, so
      // Team-Split Size resolves correctly regardless of which staff this
      // loop iteration is on.
      const groups = computeStaffWorkflowKpiGroups_(staffRows, approvedInRange, workflows, monthlyTargets);
      const totalKpiScoreRaw = Math.round(groups.reduce((s, g) => s + g.kpiScore, 0) * 100) / 100;
      // FIXED DENOMINATOR (Option B) - identical helper/inputs used by
      // computeStoresKPIPct()/getDashboard()/getMyScore()/getStoresKpiDebug(),
      // not a per-row/per-entered-Workflow sum that used to grow (and could
      // make Overall % drop) the moment a new Workflow got its first Approval.
      const totalMaxScore = computeStaffFixedActiveWorkflowMaxScore_(staffId, workflows);
      // CAP AT MAX SCORE (Aug 2026 fix - sidebar Staff Appraisal tab parity
      // with getMyScore()/computeStoresKPIPct()): 'totalKpiScoreRaw' is built
      // from each Register row's FROZEN historical Weightage % while
      // 'totalMaxScore' is the CURRENT/live sum of Weightage % across
      // today's Active workflows - the exact same Historical Freeze
      // mismatch already clamped in getMyScore() and computeStoresKPIPct(),
      // but this Manager-only "Staff Appraisal" nav tab (action:'getAppraisal',
      // a THIRD, separate code path from both of those) had never received
      // the same fix, so it kept showing e.g. 122.7/197.1 - a different,
      // uncapped number from the "Reports > Staff Appraisal" screen for the
      // identical staff/period. Clamping here brings all three appraisal
      // surfaces into agreement.
      const totalKpiScore = Math.round(clamp(totalKpiScoreRaw, 0, totalMaxScore) * 100) / 100;
      agg[staffId] = {
        staffId, staffName: staffRows[0]['Staff Name'],
        designation: sm ? sm['Designation'] : '', section: sm ? sm['Section'] : '',
        entries: staffRows.length, totalKpiScore, totalMaxScore
      };
    });
    return Object.values(agg).map(a => {
      const pct = a.totalMaxScore > 0 ? Math.round((a.totalKpiScore / a.totalMaxScore) * 10000) / 100 : 0;
      return Object.assign(a, { overallPct: pct, rating: ratingFor(pct) });
    });
  },

  // ============================================================
  // MY SCORE / SCORE BREAKDOWN
  // ============================================================
  // Per-KPI (Workflow + Month) breakdown for one staff member over a date
  // range: Target, Achieved (every logged entry, any status - visibility
  // only), Approved (Approved only - what's actually credited), Achievement
  // %, Weightage %, KPI Score, Max Score. Grouped by Workflow + Month
  // because a Monthly Target is one figure per Workflow per Month - see
  // getEffectiveTarget().
  //
  // Achievement % is computed from the RAW (undiluted) Approved Qty, NOT
  // the Team Split share - a shared job still shows its true
  // completion % regardless of how many people worked it. Only the
  // Weightage %/KPI Score actually credited to this person is divided
  // (via withTeamSplit), per the existing Team Split Dilution Fix - this
  // satisfies "Team Split, Co-Staff entries, or workflow count must not
  // reduce the Achievement Percentage" while still crediting each person
  // their fair share of the score.
  //
  // A Staff caller can only ever see their own score (p.actorRole +
  // p.actorStaffId are compared against p.staffId below) - Manager/HR/
  // Admin may request any staffId, for the Manager-facing breakdown
  // screen. UI hides the option for a Staff user; this is the real gate.
  getMyScore: function (p) {
    let staffId = p.staffId;
    if (p.actorRole === 'Staff') {
      if (staffId && p.actorStaffId && staffId !== p.actorStaffId) {
        throw new Error('Staff can only view their own Score');
      }
      staffId = p.actorStaffId || staffId;
    }
    if (!staffId) throw new Error('staffId is required');
    const staff = readAll(SHEETS.STAFF).find(s => s['Staff ID'] === staffId);
    if (!staff) throw new Error('Staff not found: ' + staffId);

    const from = p.from ? startOfDay(p.from) : null;
    const to = p.to ? endOfDay(p.to) : null;
    const allInRange = readAll(SHEETS.REGISTER).filter(r => {
      const d = new Date(r['Date']);
      return (!from || d >= from) && (!to || d <= to);
    });
    const myRows = allInRange.filter(r => r['Staff ID'] === staffId);
    const approvedAll = allInRange.filter(r => r['Approval Status'] === 'Approved');
    const myApproved = myRows.filter(r => r['Approval Status'] === 'Approved');

    const workflows = readAll(SHEETS.WORKFLOW);
    const monthlyTargets = readAll(SHEETS.TARGETS);

    // Achieved Qty - every logged entry regardless of status, for visibility
    // of what's pending vs credited. Kept separate from the Approved-only
    // aggregation below since it intentionally includes Draft/Submitted/
    // Rejected rows too.
    const achievedByGroup = {};
    myRows.forEach(r => {
      const key = r['Workflow ID'] + '__' + monthLabelOf(r['Date']);
      achievedByGroup[key] = (achievedByGroup[key] || 0) + (Number(r['Actual']) || 0);
    });

    // Approved Qty / Achievement % / KPI Score / Weightage - single
    // authoritative aggregation, shared with computeStoresKPIPct() (see
    // computeStaffWorkflowKpiGroups_() for why per-row scoring was wrong
    // whenever a staff had more than one Approved row for the same
    // Workflow+Month - the Stores KPI Team-Split Dilution Fix, Aug 2026).
    const groups = computeStaffWorkflowKpiGroups_(myApproved, approvedAll, workflows, monthlyTargets);

    const kpis = groups.map(g => {
      const key = g.workflowId + '__' + g.month;
      const wf = workflows.find(w => w['Workflow ID'] === g.workflowId);
      // AUTOMATED RATING (see resolveWorkflowKpiRating()) - a qualitative
      // label + star score alongside the numeric Contribution, based on the
      // TRUE (uncapped) achievement so a genuine 110% still reads as the
      // top band, not silently downgraded by the capped display value.
      const ratingInfo = resolveWorkflowKpiRating(g.achievementPctRaw);
      return {
        workflowId: g.workflowId, workflowName: g.workflowName,
        kpiName: wf ? wf['KPI Name'] : '', activityName: wf ? wf['Activity Name'] : '',
        month: g.month, target: g.target, targetConfigured: g.targetConfigured,
        achievedQty: achievedByGroup[key] || 0, approvedQty: g.totalActual, sourceEntryCount: g.recordCount,
        achievementPct: g.achievementPct, achievementPctRaw: g.achievementPctRaw,
        weightagePct: g.weightagePct, kpiScore: g.kpiScore, maxScore: g.weightagePct,
        rating: ratingInfo.rating, ratingScore: ratingInfo.ratingScore, maxRatingScore: ratingInfo.maxRatingScore
      };
    }).sort((a, b) => a.month === b.month ? a.workflowName.localeCompare(b.workflowName) : a.month.localeCompare(b.month));

    const totalScoreRaw = Math.round(
      kpis.filter(k => STORES_KPI_EXCLUDED_ACTIVITY_IDS.indexOf(
        (workflows.find(w => w['Workflow ID'] === k.workflowId) || {})['Activity ID']
      ) === -1).reduce((s, k) => s + k.kpiScore, 0) * 100
    ) / 100;
    // FIXED DENOMINATOR (Option B, Aug 2026), single-source with
    // computeStoresKPIPct()/getDashboard() - see
    // computeStaffFixedActiveWorkflowMaxScore_(). Replaces
    // "kpis.reduce((s,k)=>s+k.maxScore,0)", which only summed the
    // Weightage % of Workflow+Month groups this staff already had an
    // Approved entry for, and so grew (and could make Overall % drop) every
    // time a new Workflow received its first Approval. The kpis[] list
    // itself is left untouched above (still shows every Workflow+Month this
    // staff has entries for, Attendance included, for visibility) - only
    // this headline total/percentage is scoped to match "Stores KPI".
    const totalMaxScore = computeStaffFixedActiveWorkflowMaxScore_(staffId, workflows);
    // CAP AT MAX SCORE (Aug 2026 fix - Reports/My Score parity with
    // Dashboard): each individual kpiScore is already capped at its own
    // Weightage % (computeKpiScore() caps Achievement % at 100 before
    // applying Weightage), but 'totalScoreRaw' is built from each Register
    // row's FROZEN historical Weightage % (Historical Freeze policy) while
    // 'totalMaxScore' is the CURRENT/live sum of Weightage % across today's
    // Active workflows. If a Workflow's Weightage % was reduced after
    // entries were Approved, totalScoreRaw can legitimately exceed
    // totalMaxScore and Overall % can exceed 100 (e.g. the 122.72% /
    // 197.07% seen in the Staff Performance report) even though
    // getDashboard()'s Total KPI % (MTD) card is already clamped. This
    // brings getMyScore() - and every report built on it (Staff
    // Performance, Staff Appraisal, the Staff's own My Score screen) -
    // into line with that same 0-100 cap, so Overall %/KPI Score can never
    // show above the fixed 100-point ceiling anywhere in the app.
    const totalScore = Math.round(clamp(totalScoreRaw, 0, totalMaxScore) * 100) / 100;
    const overallPct = totalMaxScore > 0 ? Math.round((totalScore / totalMaxScore) * 10000) / 100 : 0;

    return {
      staffId, staffName: staff['Staff Name'], designation: staff['Designation'], section: staff['Section'],
      from: p.from || '', to: p.to || '',
      kpis, totalScore, totalMaxScore, overallPct,
      grade: gradeForCombined(overallPct), rating: ratingFor(overallPct),
      // AUDIT TRAIL: proves this is calculated fresh from live operational
      // data on every call, never a cached/stored figure a client could
      // have influenced.
      calculatedOn: new Date()
    };
  },

  // ---------- STORES KPI DEBUG REPORT (staff-wise, per Workflow+Month) ----------
  // Manager/HR only. Proves exactly how each staff's Stores KPI % is
  // derived - Team Actual vs Individual Actual (for legibility only; the
  // scoring itself always uses their combined Total Actual, never split),
  // Applicable Target, Achievement %, this staff's KPI Weightage share
  // (Team-Split applied per the business rule), Earned KPI Score, and the
  // resulting Final Staff Stores KPI %. Built on the same
  // computeStaffWorkflowKpiGroups_() helper as computeStoresKPIPct()/
  // getMyScore(), so this debug view can never drift from what the app
  // actually scores.
  getStoresKpiDebug: function (p) {
    if (p.actorRole !== 'Manager' && p.actorRole !== 'HR') throw new Error('Not authorized: Manager/HR only');
    if (!p || !p.from || !p.to) throw new Error('From Date and To Date are required');
    const from = startOfDay(p.from), to = endOfDay(p.to);
    const approvedInRange = readAll(SHEETS.REGISTER).filter(r => {
      const d = new Date(r['Date']);
      return d >= from && d <= to && r['Approval Status'] === 'Approved' &&
        STORES_KPI_EXCLUDED_ACTIVITY_IDS.indexOf(r['Activity ID']) === -1;
    });
    const workflows = readAll(SHEETS.WORKFLOW);
    const monthlyTargets = readAll(SHEETS.TARGETS);
    let staffList = readAll(SHEETS.STAFF).filter(s => s['Status'] === 'Active');
    if (p.staffId) staffList = staffList.filter(s => s['Staff ID'] === p.staffId);

    const rows = [];
    staffList.forEach(staff => {
      const myApproved = approvedInRange.filter(r => r['Staff ID'] === staff['Staff ID']);
      if (!myApproved.length) return;
      const groups = computeStaffWorkflowKpiGroups_(myApproved, approvedInRange, workflows, monthlyTargets);
      // FIXED DENOMINATOR (Option B, Aug 2026) - see
      // computeStaffFixedActiveWorkflowMaxScore_(). This is a fixed total
      // per staff (all Active, non-Attendance Workflows), not a sum of only
      // the Workflow+Month groups this staff happens to have an Approved
      // entry for - so it no longer changes just because a new Workflow
      // gets its first Approval.
      const fixedMax = computeStaffFixedActiveWorkflowMaxScore_(staff['Staff ID'], workflows);
      let earned = 0;
      const staffRows = groups.map(g => {
        // Team Actual vs Individual Actual - split back out purely so the
        // debug table matches the spec's requested columns; Total Actual
        // (used for scoring) is always their sum, computed once above.
        const teamActual = myApproved
          .filter(r => r['Workflow ID'] === g.workflowId && monthLabelOf(r['Date']) === g.month && r['Team Ref No'])
          .reduce((s, r) => s + (Number(r['Actual']) || 0), 0);
        const individualActual = Math.round((g.totalActual - teamActual) * 100) / 100;
        earned += g.kpiScore;
        return {
          staffId: staff['Staff ID'], staffName: staff['Staff Name'],
          workflowId: g.workflowId, workflowName: g.workflowName, month: g.month,
          teamRefNos: g.teamRefNos, numberOfRecords: g.recordCount,
          teamActual, individualActual, totalActual: g.totalActual,
          applicableTarget: g.target, targetConfigured: g.targetConfigured,
          achievementPct: g.achievementPct, achievementPctRaw: g.achievementPctRaw,
          kpiWeightage: g.weightagePct, earnedKpiScore: g.kpiScore
        };
      });
      const finalStaffStoresKpiPct = fixedMax > 0 ? Math.round((earned / fixedMax) * 10000) / 100 : 0;
      staffRows.forEach(r => {
        r.teamSharedWeightage = r.kpiWeightage;
        r.fixedMaxScore = fixedMax; // full company-wide Active/non-Attendance total, for audit visibility
        r.finalStaffStoresKpiPct = finalStaffStoresKpiPct;
      });
      rows.push.apply(rows, staffRows);
    });
    return rows;
  },

  // ==========================================================
  // REPORTS MODULE (Aug 2026 rebuild)
  // ==========================================================
  // SINGLE CALCULATION PATH (spec item 18): every report below is built
  // from exactly two authoritative helpers -
  //   _buildWorkflowPerfRows() - physical, per-Workflow Monthly Target ->
  //     Approved Actual (Team Ref No deduped) -> Balance -> Achievement % ->
  //     KPI Score, for Activity / KPI-Workflow / Section / Date-Range /
  //     Monthly Target & Achievement reports.
  //   api.getMyScore() - the EXISTING, already-trusted per-staff KPI engine
  //     (unchanged, just reused/looped) for Staff Performance / Staff
  //     Appraisal reports, so there is still only ONE Staff KPI formula in
  //     the whole app, not a second competing one.
  // NEITHER path ever sums "Target" off individual Working Register rows
  // (the old, removed "Total Per-Entry Target" bug) and NEITHER counts
  // Draft/Submitted/Rejected rows as Approved Actual - both read
  // getEffectiveTarget() for Target and filter Approval Status === 'Approved'
  // for Actual, exactly like every other trusted page in this app
  // (Dashboard, Target Status, My Score).

  getStaffPerformanceReport: function (p) {
    _validateReportRange(p);
    let staffList = readAll(SHEETS.STAFF).filter(s => s['Status'] !== 'Inactive');
    if (p.staffId) staffList = staffList.filter(s => s['Staff ID'] === p.staffId);
    if (p.section) staffList = staffList.filter(s => s['Section'] === p.section);
    const rows = staffList.map(s => {
      // Reuses the existing, trusted getMyScore() engine (actorRole
      // 'Manager' so it does not force-scope to a different staffId) -
      // this is the SAME formula the Staff's own "My Score" screen shows,
      // never a second/duplicate calculation.
      const score = api.getMyScore({ staffId: s['Staff ID'], from: p.from, to: p.to, actorRole: 'Manager' });
      const monthlyTarget = Math.round(score.kpis.reduce((sum, k) => sum + (k.target || 0), 0) * 100) / 100;
      const approvedActual = Math.round(score.kpis.reduce((sum, k) => sum + (k.approvedQty || 0), 0) * 100) / 100;
      const balance = Math.round((monthlyTarget - approvedActual) * 100) / 100;
      const achievementPct = computeAchievementPct(approvedActual, monthlyTarget);
      return {
        staffId: s['Staff ID'], staffName: s['Staff Name'], section: s['Section'], designation: s['Designation'],
        monthlyTarget, approvedActual, balance, achievementPct,
        kpiScore: score.totalScore, maxScore: score.totalMaxScore, overallPct: score.overallPct, rating: score.rating
      };
    }).sort((a, b) => a.staffName.localeCompare(b.staffName));

    // BUG FIX (Aug 2026 - "Staff Performance report stuck on Loading..."):
    // the frontend's REPORT_DEFS.staffPerf entry is marked
    // isSummaryPlusRows:true and runReport() does
    // `renderReportSummaryCards(data.summary); rows = data.rows;` -
    // it has ALWAYS expected { rows, summary } back from this action (see
    // that entry's own comment - "the server now returns { rows, summary
    // }"). This function was still returning a bare array, so
    // data.summary was undefined and renderReportSummaryCards(undefined)
    // threw reading s[key] - an UNCAUGHT exception outside runReport()'s
    // try/catch (which only wraps the callApi call), leaving tbody stuck
    // on the "Loading..." HTML that had already been written before the
    // throw. Every other isSummaryPlusRows report
    // (getKPIWorkflowPerformanceReport) already returns this shape - only
    // this one function had drifted from its own frontend contract.
    //
    // summary = DEDUPED, physical department/section-wide Monthly Target /
    // Approved Actual / Balance / Achievement % (Team Ref No counted once
    // via _dedupApprovedQty() inside _buildWorkflowPerfRows()) - reuses the
    // SAME shared aggregator every other Performance Report uses (single
    // source of truth, spec section 10), scoped to p.section if given but
    // never to a single p.staffId (a company/section-wide total, not one
    // person's).
    // ENTRYCOUNT>0 FILTER (Aug 2026 fix - full Report consistency): applies
    // the same filter as Section/Activity/KPI-Workflow Performance's
    // aggregators above, for the same reason - without it, a Workflow
    // nobody logged an Approved entry against this period still added its
    // full Monthly Target into this summary card, so Staff Performance's
    // "Monthly Target" card could differ by a small amount from Section
    // Performance's for the identical date range (e.g. 32,603 vs 32,602).
    const wfRows = _buildWorkflowPerfRows({ from: p.from, to: p.to, section: p.section }).filter(r => r.entryCount > 0);
    const summaryTarget = Math.round(wfRows.reduce((s, r) => s + r.monthlyTarget, 0) * 100) / 100;
    const summaryActual = Math.round(wfRows.reduce((s, r) => s + r.approvedActual, 0) * 100) / 100;
    const summaryBalance = Math.round((summaryTarget - summaryActual) * 100) / 100;
    const summaryAchievementPct = computeAchievementPct(summaryActual, summaryTarget);

    return {
      rows,
      summary: {
        monthlyTarget: summaryTarget, approvedActual: summaryActual,
        balance: summaryBalance, achievementPct: summaryAchievementPct
      }
    };
  },

  getActivityPerformanceReport: function (p) {
    _validateReportRange(p);
    // ENTRYCOUNT>0 FILTER (Aug 2026 fix - Section Performance parity): only
    // Workflows that actually have an Approved entry in this date range
    // should count toward Max Score/Overall % here - otherwise every OTHER
    // Active Workflow nobody touched this period (which getSectionPerformanceReport()
    // already excludes via this same filter) inflates the Max Score
    // denominator for no reason, making the same underlying data show two
    // different Max Score/Overall % totals depending on which report you
    // ran (e.g. Max Score 100 here vs 98.43 on Section Performance for the
    // identical KPI Score 73.46). Matches getSectionPerformanceReport()'s
    // wfRows/wfRowsAll filter exactly, so both reports agree.
    const wfRows = _buildWorkflowPerfRows(p).filter(r => r.entryCount > 0);
    const byActivity = {};
    wfRows.forEach(r => {
      const g = byActivity[r.activityId] || (byActivity[r.activityId] = {
        activityId: r.activityId, activityName: r.activityName,
        monthlyTarget: 0, approvedActual: 0, kpiScore: 0, maxScore: 0
      });
      g.monthlyTarget += r.monthlyTarget;
      g.approvedActual += r.approvedActual;
      g.kpiScore += (r.kpiScore || 0);
      g.maxScore += r.maxScore;
    });
    return Object.values(byActivity).map(g => {
      const monthlyTarget = Math.round(g.monthlyTarget * 100) / 100;
      const approvedActual = Math.round(g.approvedActual * 100) / 100;
      const kpiScore = Math.round(g.kpiScore * 100) / 100;
      const maxScore = Math.round(g.maxScore * 100) / 100;
      const balance = Math.round((monthlyTarget - approvedActual) * 100) / 100;
      const achievementPct = computeAchievementPct(approvedActual, monthlyTarget);
      const overallPct = maxScore > 0 ? Math.round((kpiScore / maxScore) * 10000) / 100 : 0;
      return { activityId: g.activityId, activityName: g.activityName, monthlyTarget, approvedActual, balance, achievementPct, kpiScore, maxScore, overallPct };
    }).sort((a, b) => a.activityName.localeCompare(b.activityName));
  },

  getKPIWorkflowPerformanceReport: function (p) {
    _validateReportRange(p);
    // ENTRYCOUNT>0 FILTER (Aug 2026 fix - Section/Activity Performance
    // parity): the frontend totals this report's Max Score column too
    // (REPORT_DEFS.kpiWorkflowPerf.totals includes 'maxScore'), so without
    // this filter it inherited the exact same bug just fixed on Activity
    // Performance - every OTHER Active Workflow nobody logged an Approved
    // entry against this period was still adding its full Weightage % to
    // the Total Max Score row, making this report's Total Max Score/
    // Overall % disagree with Section Performance (and now Activity
    // Performance) for the identical date range/data.
    const wfRows = _buildWorkflowPerfRows(p).filter(r => r.entryCount > 0);
    return wfRows.map(r => ({
      activityName: r.activityName, workflowName: r.workflowName, kpiName: r.kpiName,
      section: p.section || '', monthlyTarget: r.monthlyTarget, approvedActual: r.approvedActual,
      balance: r.balance, achievementPct: r.achievementPct, weightagePct: r.weightagePct,
      kpiScore: r.kpiScore, maxScore: r.maxScore,
      status: _statusForAchievement(r.achievementPct, r.targetConfigured)
    })).sort((a, b) => a.activityName === b.activityName ? a.workflowName.localeCompare(b.workflowName) : a.activityName.localeCompare(b.activityName));
  },

  getSectionPerformanceReport: function (p) {
    _validateReportRange(p);
    const staff = readAll(SHEETS.STAFF).filter(s => s['Status'] !== 'Inactive');
    let sections = Array.from(new Set(staff.map(s => s['Section']).filter(Boolean)));
    if (p.section) sections = sections.filter(sec => sec === p.section);
    const rows = sections.map(sec => {
      const staffCount = staff.filter(s => s['Section'] === sec).length;
      // Only workflows this Section actually has an Approved entry for in
      // range count toward the Section's Monthly Target/Actual - otherwise
      // every OTHER active workflow's target (which this section never
      // touches at all) would inflate the Section total for no reason.
      const wfRows = _buildWorkflowPerfRows({ from: p.from, to: p.to, section: sec }).filter(r => r.entryCount > 0);
      const monthlyTarget = Math.round(wfRows.reduce((s, r) => s + r.monthlyTarget, 0) * 100) / 100;
      const approvedActual = Math.round(wfRows.reduce((s, r) => s + r.approvedActual, 0) * 100) / 100;
      const kpiScore = Math.round(wfRows.reduce((s, r) => s + (r.kpiScore || 0), 0) * 100) / 100;
      const maxScore = Math.round(wfRows.reduce((s, r) => s + r.maxScore, 0) * 100) / 100;
      const balance = Math.round((monthlyTarget - approvedActual) * 100) / 100;
      const achievementPct = computeAchievementPct(approvedActual, monthlyTarget);
      const overallPct = maxScore > 0 ? Math.round((kpiScore / maxScore) * 10000) / 100 : 0;
      return { section: sec, staffCount, monthlyTarget, approvedActual, balance, achievementPct, kpiScore, maxScore, overallPct, rating: ratingFor(overallPct) };
    }).sort((a, b) => a.section.localeCompare(b.section));

    // BUG FIX (Aug 2026, same root cause as getStaffPerformanceReport() -
    // see its comment): REPORT_DEFS.sectionPerf on the frontend is also
    // isSummaryPlusRows:true and expects { rows, summary }; this function
    // was returning a bare array too, which would hit the exact same
    // uncaught renderReportSummaryCards(undefined) exception and leave the
    // table stuck on "Loading...". summary = company-wide total across
    // every Section in scope (all Sections, or just p.section if the
    // report was filtered to one) - same _buildWorkflowPerfRows() single
    // source of truth aggregator, entryCount>0 filter kept consistent with
    // each per-Section row above.
    const wfRowsAll = _buildWorkflowPerfRows({ from: p.from, to: p.to, section: p.section }).filter(r => r.entryCount > 0);
    const summaryTarget = Math.round(wfRowsAll.reduce((s, r) => s + r.monthlyTarget, 0) * 100) / 100;
    const summaryActual = Math.round(wfRowsAll.reduce((s, r) => s + r.approvedActual, 0) * 100) / 100;
    const summaryKpiScore = Math.round(wfRowsAll.reduce((s, r) => s + (r.kpiScore || 0), 0) * 100) / 100;
    const summaryMaxScore = Math.round(wfRowsAll.reduce((s, r) => s + r.maxScore, 0) * 100) / 100;
    const summaryBalance = Math.round((summaryTarget - summaryActual) * 100) / 100;
    const summaryAchievementPct = computeAchievementPct(summaryActual, summaryTarget);

    return {
      rows,
      summary: {
        monthlyTarget: summaryTarget, approvedActual: summaryActual, balance: summaryBalance,
        achievementPct: summaryAchievementPct, kpiScore: summaryKpiScore, maxScore: summaryMaxScore
      }
    };
  },

  getDateRangePerformanceReport: function (p) {
    _validateReportRange(p);
    const wfRows = _buildWorkflowPerfRows(p).filter(r => r.entryCount > 0 || r.targetConfigured);
    const monthlyTarget = Math.round(wfRows.reduce((s, r) => s + r.monthlyTarget, 0) * 100) / 100;
    const approvedActual = Math.round(wfRows.reduce((s, r) => s + r.approvedActual, 0) * 100) / 100;
    const kpiScore = Math.round(wfRows.reduce((s, r) => s + (r.kpiScore || 0), 0) * 100) / 100;
    const maxScore = Math.round(wfRows.reduce((s, r) => s + r.maxScore, 0) * 100) / 100;
    const balance = Math.round((monthlyTarget - approvedActual) * 100) / 100;
    const achievementPct = computeAchievementPct(approvedActual, monthlyTarget);
    return {
      from: p.from, to: p.to,
      summary: { monthlyTarget, approvedActual, balance, achievementPct, kpiScore, maxScore },
      rows: wfRows.map(r => ({
        activityName: r.activityName, workflowName: r.workflowName,
        monthlyTarget: r.monthlyTarget, approvedActual: r.approvedActual, balance: r.balance,
        achievementPct: r.achievementPct, kpiScore: r.kpiScore
      })).sort((a, b) => a.activityName === b.activityName ? a.workflowName.localeCompare(b.workflowName) : a.activityName.localeCompare(b.activityName))
    };
  },

  // Month-by-month breakdown (each row = one Workflow within one calendar
  // Month that falls inside From/To, clipped to the selected range for that
  // month) - NOT a per-entry list. Staff/Section filters (if given) scope
  // Approved Actual to that slice, same as every other report here; Monthly
  // Target itself always stays the full department Target for that
  // Workflow+Month (Monthly Target is never a per-staff value).
  getMonthlyTargetAchievementReport: function (p) {
    _validateReportRange(p);
    const tz = Session.getScriptTimeZone();
    const monthLabels = _monthLabelsInRange(p.from, p.to);
    const rangeFrom = new Date(p.from), rangeTo = new Date(p.to);
    const out = [];
    monthLabels.forEach(m => {
      const parts = m.split('-');
      const monIdx = MONTH_NAMES.indexOf(parts[0]);
      const year = Number(parts[1]);
      const monthStart = new Date(year, monIdx, 1);
      const monthEnd = new Date(year, monIdx + 1, 0);
      const sliceFrom = monthStart > rangeFrom ? monthStart : rangeFrom;
      const sliceTo = monthEnd < rangeTo ? monthEnd : rangeTo;
      const wfRows = _buildWorkflowPerfRows({
        from: Utilities.formatDate(sliceFrom, tz, 'yyyy-MM-dd'),
        to: Utilities.formatDate(sliceTo, tz, 'yyyy-MM-dd'),
        section: p.section, staffId: p.staffId, activityId: p.activityId, workflowId: p.workflowId
      });
      wfRows.forEach(r => {
        if (!r.targetConfigured && r.entryCount === 0) return; // nothing to show for this Workflow this Month
        out.push({
          month: m, section: p.section || 'All', staff: p.staffId || 'All',
          activityName: r.activityName, workflowName: r.workflowName, kpiName: r.kpiName,
          monthlyTarget: r.monthlyTarget, approvedActual: r.approvedActual, balance: r.balance,
          achievementPct: r.achievementPct, kpiScore: r.kpiScore,
          status: _statusForAchievement(r.achievementPct, r.targetConfigured)
        });
      });
    });
    return out;
  },

  getStaffAppraisalReport: function (p) {
    _validateReportRange(p);
    let staffList = readAll(SHEETS.STAFF).filter(s => s['Status'] !== 'Inactive');
    if (p.staffId) staffList = staffList.filter(s => s['Staff ID'] === p.staffId);
    if (p.section) staffList = staffList.filter(s => s['Section'] === p.section);
    // ONE trusted KPI scoring engine (spec item 10/18): this is the exact
    // same api.getMyScore() the Staff Performance report and the Staff's own
    // "My Score" screen use - no second/independent appraisal formula.
    return staffList.map(s => {
      const score = api.getMyScore({ staffId: s['Staff ID'], from: p.from, to: p.to, actorRole: 'Manager' });
      return {
        staffId: s['Staff ID'], staffName: s['Staff Name'], section: s['Section'], designation: s['Designation'],
        totalKpiScore: score.totalScore, maxScore: score.totalMaxScore, overallPct: score.overallPct, rating: score.rating
      };
    }).sort((a, b) => a.staffName.localeCompare(b.staffName));
  },

  // Workflow/status report ONLY - deliberately never rolls into Achievement
  // %/KPI Score anywhere (spec item 11). Includes every status (Draft/
  // Submitted/Approved/Rejected) so Approvals can be audited, unlike every
  // other report above which is Approved-only.
  getApprovalWorkingRegisterReport: function (p) {
    _validateReportRange(p);
    const from = startOfDay(p.from), to = endOfDay(p.to);
    const staffSectionOf = {};
    readAll(SHEETS.STAFF).forEach(s => staffSectionOf[s['Staff ID']] = s['Section']);
    let rows = readAll(SHEETS.REGISTER).filter(r => {
      const d = new Date(r['Date']);
      return d >= from && d <= to;
    });
    if (p.staffId) rows = rows.filter(r => r['Staff ID'] === p.staffId);
    if (p.activityId) rows = rows.filter(r => r['Activity ID'] === p.activityId);
    if (p.workflowId) rows = rows.filter(r => r['Workflow ID'] === p.workflowId);
    if (p.section) rows = rows.filter(r => staffSectionOf[r['Staff ID']] === p.section);
    if (p.status) rows = rows.filter(r => r['Approval Status'] === p.status);
    return rows.map(r => ({
      date: r['Date'], staffId: r['Staff ID'], staffName: r['Staff Name'], section: staffSectionOf[r['Staff ID']] || '',
      activityName: r['Activity Name'], workflowName: r['Workflow Name'], kpiName: r['KPI Name'],
      actual: r['Actual'], status: r['Approval Status'],
      submittedOn: r['Submitted On'], approvedOn: r['Approved On'], approvedBy: r['Approved By'],
      rejectionReason: r['Rejection Reason'] || ''
    })).sort((a, b) => new Date(b.date) - new Date(a.date));
  },

  // ==========================================================
  // HR KPI APPRAISAL MODULE (8-criteria / 100-mark system)
  // p.periodLabel is a free-text label (e.g. "Jul-2026" or "Q1 FY2026-27")
  // used only to tag/match records; p.from/p.to are the actual dates used
  // to compute the two AUTO criteria and to filter Recognition/Disciplinary
  // entries for this period. Manager-only for all writes.
  // ==========================================================

  // ---------- Special Recognition (bonus marks) ----------
  addRecognition: function (p) {
    // p: {actorRole, actorUserId, staffId, periodLabel, type, bonusMarks, reason}
    requireHR(p.actorRole);
    _requireAppraisalPeriodUnlocked(p.periodLabel);
    const staff = readAll(SHEETS.STAFF).find(s => s['Staff ID'] === p.staffId);
    if (!staff) throw new Error('Staff not found: ' + p.staffId);
    if (!p.reason || !String(p.reason).trim()) throw new Error('A reason is required for Special Recognition');
    let bonus = HR_POLICY.RECOGNITION_BONUS[p.type];
    if (bonus === undefined) throw new Error('Unknown recognition type: ' + p.type);
    if (p.type === 'Special Appreciation' && p.bonusMarks !== undefined) {
      bonus = clamp(Number(p.bonusMarks) || 0, HR_POLICY.SPECIAL_APPRECIATION_RANGE[0], HR_POLICY.SPECIAL_APPRECIATION_RANGE[1]);
    }
    // CONCURRENCY FIX (audit item 20): id-read + append held under one
    // script lock, same pattern as addWorkingEntries() - two HR users
    // saving a Recognition at the same instant could otherwise both read
    // the same "max existing ID" and both write RC0007, etc.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    let id;
    try {
      const ids = readAll(SHEETS.RECOGNITION).map(r => r['Recognition ID']);
      id = nextId('RC', ids);
      appendRow(SHEETS.RECOGNITION, {
        'Recognition ID': id, 'Staff ID': p.staffId, 'Staff Name': staff['Staff Name'],
        'Period': p.periodLabel, 'Type': p.type, 'Bonus Marks': bonus, 'Reason': p.reason,
        'Awarded By': p.actorUserId, 'Awarded On': new Date()
      });
    } finally {
      lock.releaseLock();
    }
    logAudit(p.actorUserId, 'ADD_RECOGNITION', id + ' | ' + p.staffId + ' | +' + bonus);
    return { recognitionId: id, bonusMarks: bonus };
  },

  getRecognitions: function (p) {
    // p: {staffId?, periodLabel?}
    // PERIOD LABEL FIX: same root cause/fix as getDisciplinaryActions() (see
    // normalizePeriodLabel() above) - a Recognition row saved correctly could
    // fail to show up here (and so never reach the HR Appraisal bonus total)
    // if the stored 'Period' and the queried periodLabel differed only by
    // whitespace/case, or if the sheet cell auto-formatted the text into a
    // Date. Comparing normalized values fixes both without changing what is
    // actually stored.
    let rows = readAll(SHEETS.RECOGNITION);
    if (p && p.staffId) rows = rows.filter(r => r['Staff ID'] === p.staffId);
    if (p && p.periodLabel) {
      const wanted = normalizePeriodLabel(p.periodLabel);
      rows = rows.filter(r => normalizePeriodLabel(r['Period']) === wanted);
    }
    return rows;
  },

  deleteRecognition: function (p) {
    requireHR(p.actorRole);
    const r = readAll(SHEETS.RECOGNITION).find(x => x['Recognition ID'] === p.recognitionId);
    if (!r) throw new Error('Recognition record not found');
    _requireAppraisalPeriodUnlocked(r['Period']);
    deleteRow(SHEETS.RECOGNITION, r._row);
    logAudit(p.actorUserId, 'DELETE_RECOGNITION', p.recognitionId);
    return { success: true };
  },

  // ---------- Memo / Disciplinary Action (deductions) ----------
  addDisciplinaryAction: function (p) {
    // p: {actorRole, actorUserId, staffId, periodLabel, type, reason, deductionMarks?,
    //     memoDate?, category?, description?, attachmentUrl?}
    requireHR(p.actorRole);
    _requireAppraisalPeriodUnlocked(p.periodLabel);
    const staff = readAll(SHEETS.STAFF).find(s => s['Staff ID'] === p.staffId);
    if (!staff) throw new Error('Staff not found: ' + p.staffId);
    if (!p.reason || !String(p.reason).trim()) throw new Error('A reason is required for a Memo/Disciplinary Action');
    const defaultDeduction = HR_POLICY.MEMO_DEDUCTION[p.type];
    if (defaultDeduction === undefined) throw new Error('Unknown memo/action type: ' + p.type);
    const deduction = p.deductionMarks !== undefined ? Math.abs(Number(p.deductionMarks) || 0) : defaultDeduction;
    // CONCURRENCY FIX (audit item 20): see addRecognition() above for why.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    let id;
    try {
      const ids = readAll(SHEETS.DISCIPLINARY).map(r => r['Memo ID']);
      id = nextId('DM', ids);
      appendRow(SHEETS.DISCIPLINARY, {
        'Memo ID': id, 'Memo Date': p.memoDate ? new Date(p.memoDate) : new Date(),
        'Staff ID': p.staffId, 'Staff Name': staff['Staff Name'], 'Department': staff['Section'],
        'Period': p.periodLabel, 'Type': p.type, 'Category': p.category || '',
        // Deduction Marks = original/configured penalty amount, never
        // changed after issue (audit trail/history).
        // Effective Deduction = the ONLY value the final appraisal score
        // ever subtracts (see resolveMemoEffectiveDeduction()). BUSINESS
        // RULE FIX: this starts at 0, NOT the original deduction - issuing
        // a memo must never by itself cost the employee marks. It is only
        // ever changed by decideMemo(), and only once HR confirms a Final
        // Decision (Memo Issued -> Employee Reply -> HR Review -> Final
        // Decision -> Effective Deduction -> Final Score).
        'Deduction Marks': deduction, 'Effective Deduction': 0,
        'Reason': p.reason, 'Description': p.description || '', 'Status': 'Issued',
        'Issued By': p.actorUserId, 'Issued On': new Date(), 'Attachment URL': p.attachmentUrl || ''
      });
    } finally {
      lock.releaseLock();
    }
    logAudit(p.actorUserId, 'ADD_DISCIPLINARY_ACTION', id + ' | ' + p.staffId + ' | proposed -' + deduction + ' (pending Final Decision, not yet applied)');
    return { memoId: id, deductionMarks: deduction };
  },

  // ---------- Memo lifecycle: Employee Reply -> Manager Review -> Final Decision ----------
  replyToMemo: function (p) {
    // p: {actorRole, actorUserId, actorStaffId (server-verified), memoId, reply}
    const r = readAll(SHEETS.DISCIPLINARY).find(x => x['Memo ID'] === p.memoId);
    if (!r) throw new Error('Memo not found: ' + p.memoId);
    // OWNERSHIP CHECK: compare against p.actorStaffId - the server-verified
    // logged-in staff (set by handle() from the session token) - never the
    // client-editable p.staffId. Same reasoning as the ownership check in
    // submitEntryGroup(): a client-supplied staffId could otherwise be set
    // to match the memo's real owner while the actual logged-in user is
    // someone else entirely, letting a Staff user reply to (and inspect)
    // another employee's memo.
    if (p.actorRole === 'Staff' && r['Staff ID'] !== p.actorStaffId) throw new Error('You can only reply to your own memo');
    if (!p.reply || !String(p.reply).trim()) throw new Error('Reply text is required');
    updateRow(SHEETS.DISCIPLINARY, r._row, {
      'Employee Reply': p.reply, 'Employee Replied On': new Date(),
      'Status': (r['Status'] === 'Issued' || !r['Status']) ? 'Replied' : r['Status']
    });
    logAudit(p.actorUserId, 'MEMO_EMPLOYEE_REPLY', p.memoId);
    addNotification_({ toRole: 'HR', type: 'MEMO_REPLY', refId: p.memoId,
      message: (p.actorStaffName || r['Staff Name'] || 'A staff member') + ' replied to memo ' + p.memoId });
    return { success: true };
  },

  reviewMemo: function (p) {
    // p: {actorRole, actorUserId, memoId, managerReview, hrReview?}
    // MEMO LIFECYCLE (Aug 2026 update): Employee Reply -> Manager Review ->
    // HR Final Decision. Both Manager and HR can call this action now, but
    // each role is confined to its own field so a Manager can never write
    // an HR Review and cannot bypass HR's Final Decision step (that stays
    // gated behind requireHR() in decideMemo() below - unchanged).
    if (p.actorRole !== 'Manager' && p.actorRole !== 'HR') {
      throw new Error('Manager or HR role required for this action');
    }
    const r = readAll(SHEETS.DISCIPLINARY).find(x => x['Memo ID'] === p.memoId);
    if (!r) throw new Error('Memo not found: ' + p.memoId);
    const update = {};
    if (p.actorRole === 'Manager') {
      // Manager can only write their own review - never touch HR Review or
      // Status beyond moving it forward one step (Issued/Replied ->
      // Reviewed). If HR already finalized (Closed) or already reviewed,
      // a Manager editing their note afterwards should not re-open/regress
      // the record's Status.
      update['Manager Review'] = p.managerReview || '';
      update['Manager Reviewed By'] = p.actorUserId;
      update['Manager Reviewed On'] = new Date();
      if (r['Status'] === 'Issued' || r['Status'] === 'Replied' || !r['Status']) {
        update['Status'] = 'Reviewed';
      }
    } else {
      // HR: can see/edit both fields (e.g. to correct a Manager's note
      // before deciding) and always sets HR Review explicitly.
      update['Manager Review'] = p.managerReview !== undefined ? p.managerReview : r['Manager Review'];
      if (p.managerReview !== undefined && !r['Manager Reviewed By']) {
        update['Manager Reviewed By'] = p.actorUserId;
        update['Manager Reviewed On'] = new Date();
      }
      update['HR Review'] = p.hrReview !== undefined ? p.hrReview : r['HR Review'];
      if (r['Status'] === 'Issued' || r['Status'] === 'Replied' || !r['Status']) {
        update['Status'] = 'Reviewed';
      }
    }
    updateRow(SHEETS.DISCIPLINARY, r._row, update);
    logAudit(p.actorUserId, p.actorRole === 'Manager' ? 'MEMO_MANAGER_REVIEW' : 'MEMO_HR_REVIEW', p.memoId);
    return { success: true };
  },

  decideMemo: function (p) {
    // p: {actorRole, actorUserId, memoId, finalDecision, adjustedDeductionMarks?}
    requireHR(p.actorRole);
    if (!p.memoId) throw new Error('memoId is required');
    if (!p.finalDecision) throw new Error('Final Decision is required');
    if (HR_POLICY.MEMO_DECISION_OPTIONS.indexOf(p.finalDecision) === -1) {
      throw new Error('Invalid Final Decision: ' + p.finalDecision);
    }
    // P1 FIX #2 / #3 (audit): a Final Decision must be exactly-once and
    // race-free. Two HR users deciding the same memo at nearly the same
    // moment must never both succeed (last-write-silently-wins). Acquire
    // the script lock BEFORE re-reading the row, so the "already decided"
    // check below always sees the latest state, not a stale read taken
    // before either request started.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      return api._decideMemoLocked(p);
    } finally {
      lock.releaseLock();
    }
  },

  // Internal - only ever called from decideMemo() above, already holding
  // the script lock. Do not call this directly / do not expose it as an
  // action name.
  _decideMemoLocked: function (p) {
    const r = readAll(SHEETS.DISCIPLINARY).find(x => x['Memo ID'] === p.memoId); // fresh read, inside the lock
    if (!r) throw new Error('Memo not found: ' + p.memoId);
    // IMMUTABILITY CHECK (audit item 3): once a memo has a Final Decision
    // (Status = 'Closed' or Final Decision already set), it can never be
    // changed by a second decideMemo call - no re-decide, no overwrite.
    if (r['Status'] === 'Closed' || r['Final Decision']) {
      throw new Error('This memo has already received a final decision and cannot be changed.');
    }
    // HR LOCK GUARD (spec section 12): a Final Decision writes Effective
    // Deduction, which directly changes hrFinalScore/combinedFinalScore for
    // this memo's Period - must not be allowed once that Period is
    // Approved/locked in HR Appraisal History (see _requireAppraisalPeriodUnlocked).
    _requireAppraisalPeriodUnlocked(r['Period']);
    const originalDeduction = Number(r['Deduction Marks']) || 0;
    // MEMO DEDUCTION POLICY: Deduction Marks is the original/configured
    // penalty and is never edited here (kept for audit history). Effective
    // Deduction is the ONLY value getFullAppraisal() ever counts (via
    // resolveMemoEffectiveDeduction()), and is set HERE - exactly once per
    // decision - from the Final Decision:
    //   Reply Accepted / Satisfactory  -> 0 (no penalty)
    //   Warning Only / No Penalty      -> 0 (no penalty)
    //   Memo Closed / No Action        -> 0 (no penalty)
    //   Memo Withdrawn / Cancelled     -> 0 (no penalty)
    //   Deduction Adjusted             -> the adjusted value (Deduction
    //                                     Marks itself stays at the
    //                                     original, so the before/after is
    //                                     always visible)
    //   Penalty Confirmed              -> the original configured deduction
    let effectiveDeduction;
    if (HR_POLICY.MEMO_DECISION_NO_PENALTY.indexOf(p.finalDecision) !== -1) {
      effectiveDeduction = 0;
    } else if (p.finalDecision === 'Deduction Adjusted') {
      if (p.adjustedDeductionMarks === undefined || String(p.adjustedDeductionMarks).trim() === '') {
        throw new Error('Adjusted Deduction is required when Final Decision is "Deduction Adjusted"');
      }
      effectiveDeduction = Math.abs(Number(p.adjustedDeductionMarks) || 0);
    } else {
      // Penalty Confirmed - the only remaining option once no-penalty and
      // Deduction Adjusted are ruled out (validation above already rejects
      // any Final Decision not in HR_POLICY.MEMO_DECISION_OPTIONS).
      effectiveDeduction = originalDeduction;
    }
    const update = {
      'Final Decision': p.finalDecision, 'Decided On': new Date(), 'Status': 'Closed',
      'Effective Deduction': effectiveDeduction
    };
    updateRow(SHEETS.DISCIPLINARY, r._row, update);
    logAudit(p.actorUserId, 'MEMO_FINAL_DECISION', p.memoId + ' | ' + p.finalDecision + ' | effective -' + effectiveDeduction);
    const staffUserId = userIdForStaff_(r['Staff ID']);
    if (staffUserId) {
      addNotification_({ toUserId: staffUserId, type: 'MEMO_DECISION', refId: p.memoId,
        message: 'Memo ' + p.memoId + ' final decision: ' + p.finalDecision });
    }
    return { success: true };
  },

  // Adds the same board-display fields used in getFullAppraisal (section
  // 6) to a raw Disciplinary row, without changing/renaming any stored
  // column - purely additive, computed fields for the UI.
  _withMemoDisplayFields: function (r) {
    const effectiveMarks = resolveMemoEffectiveDeduction(r);
    return Object.assign({}, r, {
      memoStatus: memoStatusLabel(r['Status']),
      replyStatus: r['Employee Reply'] ? 'Submitted' : 'Pending',
      hrDecision: r['Final Decision'] || 'Pending',
      originalDeduction: Number(r['Deduction Marks']) || 0,
      effectiveDeduction: effectiveMarks,
      appraisalImpact: -effectiveMarks
    });
  },

  getMemoDetail: function (p) {
    const r = readAll(SHEETS.DISCIPLINARY).find(x => x['Memo ID'] === p.memoId);
    if (!r) throw new Error('Memo not found: ' + p.memoId);
    // OWNERSHIP CHECK: a Staff user may only view their own memo - getMemoDetail
    // is looked up by Memo ID (not staffId), so it is NOT covered by
    // STAFF_SCOPED_READ_ACTIONS' automatic staffId injection and needs its
    // own explicit check here (same rule as replyToMemo).
    if (p.actorRole === 'Staff' && r['Staff ID'] !== p.actorStaffId) {
      throw new Error('Not authorized to view this memo');
    }
    return api._withMemoDisplayFields(r);
  },

  getDisciplinaryActions: function (p) {
    // p: {staffId?, periodLabel?}
    let rows = readAll(SHEETS.DISCIPLINARY);
    if (p && p.staffId) rows = rows.filter(r => r['Staff ID'] === p.staffId);
    // PERIOD MATCH FIX (Aug 2026): compare Period Labels via
    // normalizePeriodLabel() (trim + case-insensitive) instead of a strict
    // === match. Root cause of the reported bug: a memo's Employee Reply
    // was saved correctly (Status: Issued -> Replied), but HR's Memo
    // Management table filters by the same free-text Period field used to
    // create the memo - any stray whitespace/case difference between the
    // two meant an exact === match silently returned zero rows for that
    // memo, even though nothing was actually lost.
    if (p && p.periodLabel) {
      const wanted = normalizePeriodLabel(p.periodLabel);
      rows = rows.filter(r => normalizePeriodLabel(r['Period']) === wanted);
    }
    return rows.map(api._withMemoDisplayFields);
  },

  deleteDisciplinaryAction: function (p) {
    requireHR(p.actorRole);
    const r = readAll(SHEETS.DISCIPLINARY).find(x => x['Memo ID'] === p.memoId);
    if (!r) throw new Error('Memo/Disciplinary record not found');
    // P2 FIX #4 (audit): a finalized memo (Status = 'Closed', i.e. it has
    // already been through decideMemo()) must never be hard-deleted - it's
    // the audit record of a completed disciplinary decision that already
    // affected (or deliberately did not affect) the staff member's
    // appraisal. Only a not-yet-finalized record (Issued/Replied/Reviewed)
    // may still be deleted.
    if (r['Status'] === 'Closed') {
      throw new Error('Finalized disciplinary records cannot be deleted.');
    }
    deleteRow(SHEETS.DISCIPLINARY, r._row);
    logAudit(p.actorUserId, 'DELETE_DISCIPLINARY_ACTION', p.memoId);
    return { success: true };
  },

  // ---------- HR Appraisal entry (11 pure-HR KPIs, entered monthly by HR) ----------
  // AUTO-DERIVE FROM SOURCE REGISTERS (audit item 15): lets HR pull suggested
  // Working/Present/Leave Days, Permission Hours, Late/Early counts straight
  // from the existing Attendance Register + Leave Register for this staff +
  // period, instead of re-typing numbers HR already recorded elsewhere. This
  // is a SUGGESTION only - submitManualScores() still saves whatever the HR
  // user has on the form after they review/adjust it (clearly an override,
  // never silently auto-submitted).
  getAttendanceAutoFields: function (p) {
    requireHR(p.actorRole);
    if (!p.staffId || !p.from || !p.to) throw new Error('staffId, from and to are required');
    const from = startOfDay(p.from), to = endOfDay(p.to);
    const attRows = readAll(SHEETS.ATTENDANCE).filter(r =>
      r['Staff ID'] === p.staffId && r['Date'] && new Date(r['Date']) >= from && new Date(r['Date']) <= to);
    // ATTENDANCE FORMULA FIX (P1): attendance status and performance/
    // discipline deduction are two different things - Late, Permission,
    // and Early Leaving all mean the employee DID work that day, so each
    // counts as a full present day toward Attendance %/Effective Attendance
    // Days (Present/Late/Permission = 1.0, Half Day = 0.5, Absent = 0.0).
    // Repeated lateness/permission usage is penalised separately via
    // lateEntries/permissionHours feeding their OWN HR criteria
    // (lateComing/permissionHours in computeHRScores) - so a single late
    // arrival or a single permission request never unfairly drags down the
    // attendance figure itself, but a pattern of them still lowers the
    // employee's discipline score. Previously 'Permission' fell through to
    // the else-branch (0 credit) and was wrongly scored as if absent.
    let presentDays = 0, permissionHours = 0, lateEntries = 0, earlyLeavingCount = 0;
    attRows.forEach(r => {
      const status = String(r['Status'] || '');
      if (status === 'Present') presentDays += 1;
      else if (status === 'Half Day') presentDays += 0.5;
      else if (status === 'Late') { presentDays += 1; lateEntries += 1; } // present, but logged late
      else if (status === 'Permission') presentDays += 1; // present, worked the day - permission hours penalised separately below
      else if (status === 'Early Leaving') { presentDays += 1; earlyLeavingCount += 1; }
      // 'Absent' (and any unrecognised status) intentionally falls through
      // with 0 day credit - never added to presentDays.
      permissionHours += Number(r['Permission Hours']) || 0;
    });
    const leaveDays = readAll(SHEETS.LEAVE)
      .filter(l => l['Staff ID'] === p.staffId && l['Status'] === 'Approved')
      .reduce((sum, l) => {
        const lf = new Date(l['From Date']), lt = new Date(l['To Date']);
        const overlapFrom = lf > from ? lf : from;
        const overlapTo = lt < to ? lt : to;
        if (overlapTo < overlapFrom) return sum;
        const days = Math.floor((endOfDay(overlapTo) - startOfDay(overlapFrom)) / 86400000) + 1;
        return sum + Math.max(0, days);
      }, 0);
    return {
      workingDays: attRows.length, // days actually tracked in Attendance Register for this staff/period
      presentDays: Math.round(presentDays * 100) / 100,
      leaveDays,
      permissionHours: Math.round(permissionHours * 100) / 100,
      lateEntries, earlyLeavingCount,
      source: 'Attendance Register + Leave Register (auto-derived - review before saving)'
    };
  },

  // ---------- AUTO HR APPRAISAL (source-register driven) ----------
  // Safely derives every KPI for which the application has an authoritative
  // source. Existing manual scores remain untouched; when an employee has no
  // manual row, getFullAppraisal() can use this auto result as the base.
  // Training and HR Remarks are intentionally NOT invented: without a source
  // register they remain Not Set (0) and are flagged for HR review.
  getAutoHRAppraisal: function (p) {
    requireHR(p.actorRole);
    if (!p || !p.periodLabel || !p.from || !p.to) throw new Error('periodLabel, from and to are required');
    const from = startOfDay(p.from), to = endOfDay(p.to);
    let staffList = readAll(SHEETS.STAFF).filter(s => s['Status'] === 'Active');
    if (p.staffId) staffList = staffList.filter(s => s['Staff ID'] === p.staffId);
    const attAll = readAll(SHEETS.ATTENDANCE);
    const leaveAll = readAll(SHEETS.LEAVE);
    // PERIOD LABEL FIX: normalize, same reason as getRecognitions() above.
    const wantedPeriodAuto = normalizePeriodLabel(p.periodLabel);
    const memosAll = readAll(SHEETS.DISCIPLINARY).filter(r => normalizePeriodLabel(r['Period']) === wantedPeriodAuto);
    const existing = readAll(SHEETS.MANUAL_SCORES).filter(r => normalizePeriodLabel(r['Period']) === wantedPeriodAuto);
    const P = HR_APPRAISAL_POLICY;
    const clamp10 = v => clamp(Number(v) || 0, 0, 10);
    const approvedLeaveDays = (staffId) => leaveAll.filter(l => {
      if (l['Staff ID'] !== staffId || l['Status'] !== 'Approved' || !l['From Date'] || !l['To Date']) return false;
      const lf = new Date(l['From Date']), lt = new Date(l['To Date']);
      return lt >= from && lf <= to;
    }).reduce((sum, l) => {
      const lf = new Date(l['From Date']), lt = new Date(l['To Date']);
      const a = lf > from ? lf : from, b = lt < to ? lt : to;
      return sum + Math.max(0, Math.floor((endOfDay(b) - startOfDay(a)) / 86400000) + 1);
    }, 0);

    return staffList.map(staff => {
      const id = staff['Staff ID'];
      const rows = attAll.filter(r => r['Staff ID'] === id && r['Date'] && new Date(r['Date']) >= from && new Date(r['Date']) <= to);
      let presentDays = 0, permissionHours = 0, late = 0, early = 0, shiftGood = 0;
      rows.forEach(r => {
        const st = String(r['Status'] || '');
        if (st === 'Present') { presentDays += 1; shiftGood += 1; }
        else if (st === 'Half Day') presentDays += 0.5;
        else if (st === 'Late') { presentDays += 1; late += 1; }
        else if (st === 'Permission') { presentDays += 1; permissionHours += Number(r['Permission Hours']) || 0; }
        else if (st === 'Early Leaving') { presentDays += 1; early += 1; }
        permissionHours += st === 'Permission' ? 0 : (Number(r['Permission Hours']) || 0);
      });
      const leaveDays = approvedLeaveDays(id);
      const wd = rows.length;
      const shiftPct = wd > 0 ? (shiftGood / wd) * 100 : 0;
      const leavePct = 100 - Math.max(0, leaveDays - P.SANCTIONED_LEAVE_DAYS) * P.LEAVE_PENALTY_PER_EXCESS_DAY;
      const permPct = 100 - Math.max(0, permissionHours - P.ALLOWED_PERMISSION_HOURS) * P.PERMISSION_PENALTY_PER_EXCESS_HOUR;
      const latePct = 100 - late * P.LATE_COMING_PENALTY_PER_ENTRY;
      const earlyPct = 100 - early * P.EARLY_LEAVING_PENALTY_PER_ENTRY;
      const myMemos = memosAll.filter(r => r['Staff ID'] === id);
      // IMPORTANT: Memo/disciplinary deductions are applied exactly once in
      // getFullAppraisal() as an explicit HR Final adjustment. Do NOT bake
      // memo deductions into Discipline / Policy / Behaviour here; doing so
      // would double-count the same memo for employees using Auto HR data.
      // Discipline/Policy/Behaviour have no independent authoritative source
      // register, so Auto mode uses the neutral full rating until HR
      // supplies manual ratings. Training (Aug 2026) is DIFFERENT - it now
      // has an authoritative source (Training_Register) and is ALWAYS
      // computed automatically here, manual row or not - see
      // getTrainingAppraisalScore_(). It is intentionally never in
      // `missing` below since it is never "missing", only ever 0 if the
      // staff has no eligible training on record for this period.
      const existingRow = existing.filter(r => r['Staff ID'] === id).sort((a,b)=>new Date(b['Entered On']||0)-new Date(a['Entered On']||0))[0];
      const training = getTrainingAppraisalScore_(id, from, to);
      const remarks = existingRow ? Number(existingRow['HR Remarks Rating']) || 0 : 0;
      const criteriaRatings = {
        discipline: P.RATING_MAX.discipline,
        policyCompliance: P.RATING_MAX.policyCompliance,
        behaviour: P.RATING_MAX.behaviour,
        training: clamp(training, 0, P.RATING_MAX.training),
        hrRemarks: clamp(remarks, 0, P.RATING_MAX.hrRemarks)
      };
      const manualEquivalent = {
        'Working Days': wd, 'Present Days': Math.round(presentDays*100)/100,
        'Leave Days': leaveDays, 'Permission Hours': Math.round(permissionHours*100)/100,
        'Late Entries': late, 'Early Leaving Count': early,
        'Shift Compliance %': Math.round(shiftPct*100)/100,
        'Overtime Hours': existingRow ? Number(existingRow['Overtime Hours']) || 0 : 0,
        'Discipline Rating': criteriaRatings.discipline,
        'Policy Compliance Rating': criteriaRatings.policyCompliance,
        'Behaviour Rating': criteriaRatings.behaviour,
        'Training Rating': criteriaRatings.training,
        'HR Remarks Rating': criteriaRatings.hrRemarks
      };
      const score = computeHRScores(manualEquivalent);
      const missing = [];
      if (!existingRow) missing.push('HR Remarks Rating');
      return {
        staffId:id, staffName:staff['Staff Name'], designation:staff['Designation'], department:staff['Section'],
        appraisalPeriod:p.periodLabel, auto:true, source:'Attendance Register + Leave Register + Disciplinary Register',
        manualEquivalent, criteria:score.criteria, hrScore:score.hrScore,
        manualScoresEntered:!!existingRow, autoReviewRequired:missing.length>0, missingCriteria:missing
      };
    });
  },


  // ---------- AUTOMATIC TRAINING SCORE MODULE (Aug 2026) ----------
  // Read-only Training Master list. All roles may call this (Staff needs it
  // to label their own "View Training Details" records); the sheet itself
  // is HR-managed (only api.addTrainingRecord/updateTrainingRecord write to
  // it indirectly via Training_Register - the Master list itself is edited
  // directly in the sheet by HR, same governance level as other Masters).
  getTrainingMaster: function () {
    return readAll(SHEETS.TRAINING_MASTER).filter(t => t['Status'] === 'Active');
  },

  // p: {staffId?, from?, to?} - Manager/HR pass staffId blank for everyone;
  // a Staff caller has p.staffId force-set to their own session staffId by
  // handle() (see STAFF_SCOPED_READ_ACTIONS above) - never spoofable.
  getTrainingRecords: function (p) {
    let rows = readAll(SHEETS.TRAINING_REGISTER);
    if (p && p.staffId) rows = rows.filter(r => r['Staff ID'] === p.staffId);
    if (p && p.from) rows = rows.filter(r => r['Training Date'] && new Date(r['Training Date']) >= startOfDay(p.from));
    if (p && p.to) rows = rows.filter(r => r['Training Date'] && new Date(r['Training Date']) <= endOfDay(p.to));
    return rows.sort((a, b) => new Date(b['Training Date']) - new Date(a['Training Date']));
  },

  // p: {actorRole, actorUserId, staffId, trainingId, trainingDate, trainer,
  //     mode, attendance, assessmentScore, completionStatus, certificateUrl}
  // HR-only, same governance level as Recognition/Memo/Manual Scores.
  // Eligible/Expiry Date/Auto Score are NEVER accepted from the client -
  // always computed here, server-side (spec item 4/7/15).
  addTrainingRecord: function (p) {
    requireHR(p.actorRole);
    if (!p.staffId || !p.trainingId || !p.trainingDate) {
      throw new Error('Staff, Training and Training Date are required');
    }
    const staff = readAll(SHEETS.STAFF).find(s => s['Staff ID'] === p.staffId);
    if (!staff || staff['Status'] !== 'Active') throw new Error('Staff not found or not active: ' + p.staffId);
    const master = readAll(SHEETS.TRAINING_MASTER).find(t => t['Training ID'] === p.trainingId);
    if (!master) throw new Error('Training not found: ' + p.trainingId);
    if (master['Status'] !== 'Active') throw new Error('This training is not Active in Training Master');
    const assessmentScore = clamp(Number(p.assessmentScore) || 0, 0, 100);
    const trainingDate = new Date(p.trainingDate);
    if (isNaN(trainingDate.getTime())) throw new Error('Invalid Training Date');
    const validAttendance = ['Present', 'Absent'];
    if (validAttendance.indexOf(p.attendance) === -1) throw new Error('Attendance must be Present or Absent');
    const validCompletion = ['Completed', 'Not Completed', 'Failed'];
    if (validCompletion.indexOf(p.completionStatus) === -1) throw new Error('Invalid Completion Status');

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      // DUPLICATE PROTECTION: Staff ID + Training ID + Training Date must be
      // unique (spec item 4.8 / UAT test 8). Compared as date-only (ignores
      // time-of-day) so the same calendar date always collides correctly.
      const dateKey = d => d ? Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
      const dup = readAll(SHEETS.TRAINING_REGISTER).find(r =>
        r['Staff ID'] === p.staffId && r['Training ID'] === p.trainingId && dateKey(r['Training Date']) === dateKey(trainingDate));
      if (dup) throw new Error('Duplicate: this staff already has a record for this Training on this date (' + dup['Training Record ID'] + ')');

      const expiryDate = calculateTrainingExpiry_(trainingDate, master['Validity Months']);
      const record = {
        'Staff ID': p.staffId, 'Training ID': p.trainingId, 'Training Name': master['Training Name'],
        'Training Date': trainingDate, 'Trainer': p.trainer || '', 'Mode': p.mode || '',
        'Attendance': p.attendance, 'Assessment Score': assessmentScore, 'Completion Status': p.completionStatus,
        'Certificate / Evidence': p.certificateUrl || '', 'Expiry Date': expiryDate
      };
      // NOTE: Eligible/Auto Score stored on the row are a "true as of today"
      // DISPLAY snapshot for the Training Management screen only. Actual
      // appraisal scoring NEVER reads these two columns - it always calls
      // calculateTrainingEligibility_()/calculateTrainingScore_() fresh,
      // anchored to that specific appraisal's own "To" date (see
      // getTrainingAppraisalScore_()), so the same record can correctly be
      // eligible for an Aug-2026 appraisal and later show expired here once
      // its validity lapses, without that changing the already-scored past.
      const eligible = calculateTrainingEligibility_(record, new Date());
      const autoScore = eligible ? calculateTrainingScore_(assessmentScore) : 0;
      const ids = readAll(SHEETS.TRAINING_REGISTER).map(r => r['Training Record ID']);
      const id = nextId('TRN', ids);
      appendRow(SHEETS.TRAINING_REGISTER, Object.assign({ 'Training Record ID': id }, record, {
        'Eligible': eligible ? 'Yes' : 'No', 'Auto Score': autoScore,
        'Created By': p.actorUserId, 'Created On': new Date(), 'Updated On': new Date()
      }));
      logAudit(p.actorUserId, 'ADD_TRAINING_RECORD', id + ' | ' + p.staffId + ' | ' + p.trainingId);
      return { trainingRecordId: id, eligible, autoScore, expiryDate };
    } finally {
      lock.releaseLock();
    }
  },

  // p: {actorRole, actorUserId, trainingRecordId, trainer?, mode?, attendance?,
  //     assessmentScore?, completionStatus?, certificateUrl?} - Training Date/
  // Staff ID/Training ID are NOT editable here (would bypass the duplicate
  // key and expiry basis); delete+re-add for a genuine date/training/staff
  // correction. HR-only, same as addTrainingRecord.
  updateTrainingRecord: function (p) {
    requireHR(p.actorRole);
    if (!p.trainingRecordId) throw new Error('trainingRecordId is required');
    const row = readAll(SHEETS.TRAINING_REGISTER).find(r => r['Training Record ID'] === p.trainingRecordId);
    if (!row) throw new Error('Training record not found: ' + p.trainingRecordId);
    const master = readAll(SHEETS.TRAINING_MASTER).find(t => t['Training ID'] === row['Training ID']);
    if (!master) throw new Error('Training Master entry missing for ' + row['Training ID']);

    const updates = {};
    if (p.trainer !== undefined) updates['Trainer'] = p.trainer;
    if (p.mode !== undefined) updates['Mode'] = p.mode;
    if (p.attendance !== undefined) {
      if (['Present', 'Absent'].indexOf(p.attendance) === -1) throw new Error('Attendance must be Present or Absent');
      updates['Attendance'] = p.attendance;
    }
    if (p.completionStatus !== undefined) {
      if (['Completed', 'Not Completed', 'Failed'].indexOf(p.completionStatus) === -1) throw new Error('Invalid Completion Status');
      updates['Completion Status'] = p.completionStatus;
    }
    if (p.assessmentScore !== undefined) updates['Assessment Score'] = clamp(Number(p.assessmentScore) || 0, 0, 100);
    if (p.certificateUrl !== undefined) updates['Certificate / Evidence'] = p.certificateUrl;

    const merged = Object.assign({}, row, updates);
    const expiryDate = calculateTrainingExpiry_(merged['Training Date'], master['Validity Months']);
    const eligible = calculateTrainingEligibility_(merged, new Date());
    const autoScore = eligible ? calculateTrainingScore_(merged['Assessment Score']) : 0;
    updates['Expiry Date'] = expiryDate;
    updates['Eligible'] = eligible ? 'Yes' : 'No';
    updates['Auto Score'] = autoScore;
    updates['Updated On'] = new Date();
    updateRow(SHEETS.TRAINING_REGISTER, row._row, updates);
    logAudit(p.actorUserId, 'UPDATE_TRAINING_RECORD', p.trainingRecordId);
    return { trainingRecordId: p.trainingRecordId, eligible, autoScore, expiryDate };
  },

  submitManualScores: function (p) {
    // p: {actorRole, actorUserId, staffId, periodLabel, workingDays, presentDays,
    //     leaveDays, permissionHours, lateEntries, earlyLeavingCount, shiftCompliancePct,
    //     overtimeHours, discipline, policyCompliance, behaviour, training, hrRemarksRating, hrRemarks}
    requireHR(p.actorRole);
    _requireAppraisalPeriodUnlocked(p.periodLabel);
    const staff = readAll(SHEETS.STAFF).find(s => s['Staff ID'] === p.staffId);
    if (!staff) throw new Error('Staff not found: ' + p.staffId);
    const rm = HR_APPRAISAL_POLICY.RATING_MAX;
    // VALIDATION (audit item 18) - server-side, cross-field checks that a
    // simple clamp/min-0 can't express on its own.
    const wd = Math.max(0, Number(p.workingDays) || 0);
    const pd = Math.max(0, Number(p.presentDays) || 0);
    const ld = Math.max(0, Number(p.leaveDays) || 0);
    if (pd > wd) throw new Error('Present Days (' + pd + ') cannot exceed Working Days (' + wd + ')');
    if (ld > wd) throw new Error('Leave Days (' + ld + ') cannot exceed Working Days (' + wd + ')');
    if (pd + ld > wd) throw new Error('Present Days + Leave Days (' + (pd + ld) + ') cannot exceed Working Days (' + wd + ')');
    const entry = {
      'Working Days': wd,
      'Present Days': pd,
      'Leave Days': ld,
      'Permission Hours': Math.max(0, Number(p.permissionHours) || 0),
      'Late Entries': Math.max(0, Number(p.lateEntries) || 0),
      'Early Leaving Count': Math.max(0, Number(p.earlyLeavingCount) || 0),
      'Shift Compliance %': clamp(Number(p.shiftCompliancePct) || 0, 0, 100),
      'Overtime Hours': Math.max(0, Number(p.overtimeHours) || 0), // info-only, not scored
      'Discipline Rating': clamp(Number(p.discipline) || 0, 0, rm.discipline),
      'Policy Compliance Rating': clamp(Number(p.policyCompliance) || 0, 0, rm.policyCompliance),
      'Behaviour Rating': clamp(Number(p.behaviour) || 0, 0, rm.behaviour),
      // AUTOMATIC TRAINING (Aug 2026): 'Training Rating' is deliberately
      // NOT written here anymore - p.training (if a client still sends it)
      // is ignored. getFullAppraisal()/getAutoHRAppraisal() always compute
      // this field themselves from Training_Register, never from what HR
      // types on this screen. Omitting the key (rather than writing 0)
      // means an update-in-place leaves any pre-migration historical value
      // in the sheet untouched - see migrateTrainingAutomation().
      'HR Remarks Rating': clamp(Number(p.hrRemarksRating) || 0, 0, rm.hrRemarks)
    };
    // CONCURRENCY + DUPLICATE FIX (audit items 19 & 20): the existing-row
    // check (which is what enforces "one active record per Employee +
    // Period", requirement 19) and the update-or-append that follows it are
    // now both done under one script lock, same pattern as
    // addWorkingEntries()/addRecognition()/addDisciplinaryAction() above.
    // Without this, two HR users saving the SAME staff/period at the same
    // instant could both pass the "no existing row" check and both insert a
    // row (a duplicate Employee+Period record), or both compute the same
    // "next" Score ID.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      // PERIOD LABEL FIX: normalize, same reason as getRecognitions() above -
      // without this, a mismatch here would cause a NEW duplicate row to be
      // inserted every time HR "updates" the same staff/period, instead of
      // updating the existing one.
      const existing = readAll(SHEETS.MANUAL_SCORES).find(r => r['Staff ID'] === p.staffId && normalizePeriodLabel(r['Period']) === normalizePeriodLabel(p.periodLabel));
      if (existing) {
        updateRow(SHEETS.MANUAL_SCORES, existing._row, Object.assign({}, entry, {
          'HR Remarks': p.hrRemarks || '', 'Entered By': p.actorUserId, 'Entered On': new Date()
        }));
        logAudit(p.actorUserId, 'UPDATE_HR_APPRAISAL_ENTRY', p.staffId + ' | ' + p.periodLabel);
        return { scoreId: existing['Score ID'], updated: true };
      }
      const ids = readAll(SHEETS.MANUAL_SCORES).map(r => r['Score ID']);
      const id = nextId('MS', ids);
      appendRow(SHEETS.MANUAL_SCORES, Object.assign({
        'Score ID': id, 'Staff ID': p.staffId, 'Staff Name': staff['Staff Name'], 'Period': p.periodLabel
      }, entry, {
        'HR Remarks': p.hrRemarks || '', 'Entered By': p.actorUserId, 'Entered On': new Date()
      }));
      logAudit(p.actorUserId, 'ADD_HR_APPRAISAL_ENTRY', p.staffId + ' | ' + p.periodLabel);
      return { scoreId: id, updated: false };
    } finally {
      lock.releaseLock();
    }
  },

  // ---------- Duplicate Employee+Period report (audit item 19) ----------
  // Read-only, Manager/HR only. Scans HR Manual Scores for any Staff ID +
  // Period combination that has MORE THAN ONE row - this can only happen
  // from data that existed before the submitManualScores() lock+existing-
  // row-check fix above, or a direct sheet edit/paste. Nothing is deleted
  // automatically (requirement 19 explicitly forbids that) - this just
  // reports it so HR can review and, if appropriate, remove the stale row
  // themselves from the sheet. getFullAppraisal() already picks the most
  // recent row deterministically, so a duplicate here does not corrupt the
  // Combined Final Score in the meantime.
  getDuplicateAppraisalReport: function (p) {
    if (p.actorRole !== 'Manager' && p.actorRole !== 'HR') throw new Error('Not authorized: Manager/HR only');
    const rows = readAll(SHEETS.MANUAL_SCORES);
    const groups = {};
    rows.forEach(r => {
      const key = r['Staff ID'] + ' | ' + r['Period'];
      (groups[key] = groups[key] || []).push(r);
    });
    const duplicates = Object.keys(groups).filter(k => groups[k].length > 1).map(key => {
      const rows2 = groups[key];
      return {
        staffId: rows2[0]['Staff ID'], staffName: rows2[0]['Staff Name'], period: rows2[0]['Period'],
        count: rows2.length,
        records: rows2.map(r => ({ scoreId: r['Score ID'], row: r._row, enteredOn: r['Entered On'], enteredBy: r['Entered By'] }))
      };
    });
    return { duplicateCount: duplicates.length, duplicates };
  },

  getManualScores: function (p) {
    // p: {staffId?, periodLabel?}
    // PERIOD LABEL FIX: normalize, same reason as getRecognitions() above.
    let rows = readAll(SHEETS.MANUAL_SCORES);
    if (p && p.staffId) rows = rows.filter(r => r['Staff ID'] === p.staffId);
    if (p && p.periodLabel) {
      const wanted = normalizePeriodLabel(p.periodLabel);
      rows = rows.filter(r => normalizePeriodLabel(r['Period']) === wanted);
    }
    return rows;
  },

  // ---------- Full appraisal: 11 pure-HR KPIs + Bonus - Deduction, PLUS the
  // combined Final Employee Score (Stores KPI % x 70% + HR Score x 30%) ----------
  getFullAppraisal: function (p) {
    // p: {periodLabel, from, to, staffId?}
    if (!p || !p.periodLabel || !p.from || !p.to) throw new Error('periodLabel, from and to are required');
    const from = startOfDay(p.from), to = endOfDay(p.to);
    let staffList = readAll(SHEETS.STAFF).filter(s => s['Status'] === 'Active');
    if (p.staffId) staffList = staffList.filter(s => s['Staff ID'] === p.staffId);
    // PERIOD LABEL FIX: normalize, same reason as getRecognitions() above -
    // otherwise HR-entered Manual Scores can silently fail to be picked up
    // here and getFullAppraisal() falls back to Auto HR data instead.
    const manualScores = readAll(SHEETS.MANUAL_SCORES).filter(r => normalizePeriodLabel(r['Period']) === normalizePeriodLabel(p.periodLabel));
    // PERIOD LABEL FIX: normalize, same reason as getRecognitions() above -
    // otherwise a Recognition can silently fail to reach the HR Final Score's
    // bonusMarks total, not just the display table.
    const recognitions = readAll(SHEETS.RECOGNITION).filter(r => normalizePeriodLabel(r['Period']) === normalizePeriodLabel(p.periodLabel));
    // PERIOD LABEL FIX: normalize, same reason as getRecognitions() above -
    // otherwise a Memo/Disciplinary deduction can silently fail to reach the
    // HR Final Score, even though it shows fine in Memo Management (which
    // already uses normalizePeriodLabel() via getDisciplinaryActions()).
    const memos = readAll(SHEETS.DISCIPLINARY).filter(r => normalizePeriodLabel(r['Period']) === normalizePeriodLabel(p.periodLabel));

    return staffList.map(staff => {
      // DUPLICATE SAFETY (audit item 19): submitManualScores() now blocks new
      // duplicates going forward (see the lock+existing-row check there), but
      // if a Staff ID + Period ever ended up with more than one row in HR
      // Manual Scores (e.g. from before that fix, or a manual sheet edit),
      // picking rows[0] would silently use whichever happened to be first in
      // sheet order - possibly a STALE row, not the corrected one. Use the
      // deterministic rule "most recent 'Entered On', tie-broken by the
      // highest sheet row" instead. Nothing is deleted or hidden here - see
      // getDuplicateAppraisalReport() for surfacing these to HR/Manager.
      const myManualRows = manualScores.filter(r => r['Staff ID'] === staff['Staff ID']);
      const manual = myManualRows.length <= 1 ? myManualRows[0] : myManualRows.slice().sort((a, b) => {
        const ta = a['Entered On'] ? new Date(a['Entered On']).getTime() : 0;
        const tb = b['Entered On'] ? new Date(b['Entered On']).getTime() : 0;
        return tb - ta || (b._row - a._row);
      })[0];
      const auto = !manual ? api.getAutoHRAppraisal({actorRole:'HR', periodLabel:p.periodLabel, from:p.from, to:p.to, staffId:staff['Staff ID']})[0] : null;
      let effectiveManual = manual || (auto ? auto.manualEquivalent : null);
      // AUTOMATIC TRAINING (Aug 2026): Training Rating is no longer an
      // HR-typed manual value - it always comes from Training_Register via
      // getTrainingAppraisalScore_(), for BOTH the manual-row and the
      // auto-fallback branch above. This is the single point where it's
      // applied, so a Training_Register change can never silently affect an
      // already-approved (locked) snapshot - saveAppraisalSnapshot() blocks
      // re-saving an Approved period, so this recompute only ever reaches a
      // still-open period or an explicitly reopened one (see
      // reopenAppraisalSnapshot()), matching the required snapshot/lock
      // behaviour.
      if (effectiveManual) {
        effectiveManual = Object.assign({}, effectiveManual, {
          'Training Rating': getTrainingAppraisalScore_(staff['Staff ID'], from, to)
        });
      }
      const hr = computeHRScores(effectiveManual);
      const storesKpiPct = computeStoresKPIPct(staff['Staff ID'], from, to);

      const myRecognitions = recognitions.filter(r => normalizeStaffId(r['Staff ID']) === normalizeStaffId(staff['Staff ID']));
      const myMemos = memos.filter(r => normalizeStaffId(r['Staff ID']) === normalizeStaffId(staff['Staff ID']));
      const bonusMarks = myRecognitions.reduce((a, r) => a + (Number(r['Bonus Marks']) || 0), 0);
      // Repeated-memo escalation: 3+ (configurable) memos in the SAME Category
      // within this period adds an extra one-time deduction per repeated category.
      // A memo whose Final Decision resolves to "no penalty" (see
      // HR_POLICY.MEMO_DECISION_NO_PENALTY) must not count toward the
      // repeat-offender threshold either, or a withdrawn/no-action/accepted
      // memo would still end up contributing to (and can single-handedly
      // trigger) an extra penalty for that category, contradicting its own
      // "no penalty" outcome. Memos not yet decided (Issued/Replied/
      // Reviewed) still count, since that determination hasn't been made yet.
      const categoryCounts = {};
      myMemos.filter(r => HR_POLICY.MEMO_DECISION_NO_PENALTY.indexOf(r['Final Decision']) === -1)
        .forEach(r => {
          const cat = r['Category'] || r['Type'];
          if (!cat) return;
          categoryCounts[cat] = (categoryCounts[cat] || 0) + 1;
        });
      const repeatCategories = Object.keys(categoryCounts).filter(c => categoryCounts[c] >= HR_POLICY.MEMO_REPEAT_RULE.thresholdCount);
      const repeatEscalation = repeatCategories.length * HR_POLICY.MEMO_REPEAT_RULE.extraDeduction;
      // MEMO DEDUCTION POLICY: resolveMemoEffectiveDeduction() is the SINGLE
      // authoritative resolver (see its definition) - memo creation, employee
      // reply, and HR review never touch Effective Deduction; only a
      // confirmed Final Decision does. Using it here (and nowhere else)
      // guarantees a memo is deducted from the score exactly once.
      const memoDeduction = myMemos.reduce((a, r) => a + resolveMemoEffectiveDeduction(r), 0) + repeatEscalation;

      // SINGLE MEMO-DEDUCTION PATH: memoDeduction is applied here and only
      // here. Auto HR criteria deliberately keep Discipline / Policy /
      // Behaviour neutral so the same memo can never be deducted twice.

      // HR module's own 0-100 score, after bonus/memo adjustment
      const hrFinalScore = Math.round(clamp(hr.hrScore + bonusMarks - memoDeduction, 0, 100) * 100) / 100;

      // Combined Final Employee Score = Stores KPI % x 70% + HR Score x 30%
      const w = HR_APPRAISAL_POLICY.FINAL_SCORE_WEIGHTS;
      const combinedFinalScore = Math.round(
        (storesKpiPct * w.storesKpi / 100 + hrFinalScore * w.hrAppraisal / 100) * 100) / 100;
      const grade = gradeForCombined(combinedFinalScore);

      return {
        staffId: staff['Staff ID'], staffName: staff['Staff Name'],
        designation: staff['Designation'], department: staff['Section'],
        supervisor: staff['Reporting Manager'], appraisalPeriod: p.periodLabel,
        criteria: hr.criteria,
        overtimeHours: manual ? Number(manual['Overtime Hours']) || 0 : 0, // info-only
        manualScoresEntered: !!manual, autoGenerated: !manual && !!auto, autoReviewRequired: !!(auto && auto.autoReviewRequired), missingCriteria: auto ? auto.missingCriteria : [],
        hrScore: hr.hrScore, bonusMarks, memoDeduction, repeatEscalation, repeatCategories, hrFinalScore,
        hrRating: ratingForHR(hrFinalScore),
        storesKpiPct, combinedFinalScore, grade, finalRating: ratingForHR(combinedFinalScore),
        recognitions: myRecognitions.map(r => ({ recognitionId: r['Recognition ID'], type: r['Type'], marks: r['Bonus Marks'], reason: r['Reason'] })),
        memos: myMemos.map(r => {
          const effectiveMarks = resolveMemoEffectiveDeduction(r);
          return {
            memoId: r['Memo ID'], type: r['Type'], category: r['Category'],
            marks: r['Deduction Marks'], effectiveMarks,
            reason: r['Reason'], status: r['Status'],
            // ---- Board/dashboard display fields (section 6) ----
            memoStatus: memoStatusLabel(r['Status']),
            replyStatus: r['Employee Reply'] ? 'Submitted' : 'Pending',
            hrDecision: r['Final Decision'] || 'Pending',
            originalDeduction: Number(r['Deduction Marks']) || 0,
            effectiveDeduction: effectiveMarks,
            appraisalImpact: -effectiveMarks
          };
        })
      };
    });
  },

  // Persists getFullAppraisal() results for this period into HR Appraisal
  // History, so getHRDashboard() can chart Monthly/Yearly trends later.
  // Re-running for the same period REPLACES that period's old snapshot rows -
  // but only while the period is still unlocked. Once HR runs
  // lockAppraisalSnapshot() for a period ('Approval Status' = 'Approved' on
  // every row of that period), this function refuses to silently overwrite
  // it (Approved Result must be immutable / PHASE 17 & PHASE 16 "do not
  // allow later daily data to silently rewrite an already approved
  // appraisal"). An HR user who genuinely needs to correct an approved
  // period must first call reopenAppraisalSnapshot() (an explicit, audited
  // revision action), then save again.
  saveAppraisalSnapshot: function (p) {
    requireHR(p.actorRole);
    // PERIOD LABEL FIX: normalize, same reason as getRecognitions() above -
    // otherwise a locked/approved period could fail to be recognised as such
    // here (letting a re-save slip through) purely because of a
    // whitespace/case/auto-formatted-Date mismatch in the stored Period text.
    const existingBefore = readAll(SHEETS.APPRAISAL_HISTORY).filter(r => normalizePeriodLabel(r['Period']) === normalizePeriodLabel(p.periodLabel));
    if (existingBefore.some(r => r['Approval Status'] === 'Approved')) {
      throw new Error('This period is already Approved and locked. Use "Re-open for Revision" first if it genuinely needs to change.');
    }
    const results = api.getFullAppraisal(p);
    // CONCURRENCY FIX: the delete-existing-rows + insert-new-rows sequence
    // below must not run unlocked - two HR users saving a snapshot for the
    // same Period at (almost) the same instant could otherwise both read/
    // delete the same existing rows and both compute the same "next"
    // Snapshot ID, corrupting or duplicating the snapshot. Same
    // LockService pattern as addWorkingEntries()/addRecognition() etc.
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      // Re-check the lock INSIDE the script lock too - otherwise two HR
      // users, one locking and one saving, could race between the
      // unlocked check above and this write.
      const existing = readAll(SHEETS.APPRAISAL_HISTORY).filter(r => normalizePeriodLabel(r['Period']) === normalizePeriodLabel(p.periodLabel));
      if (existing.some(r => r['Approval Status'] === 'Approved')) {
        throw new Error('This period is already Approved and locked. Use "Re-open for Revision" first if it genuinely needs to change.');
      }
      existing.sort((a, b) => b._row - a._row).forEach(r => deleteRow(SHEETS.APPRAISAL_HISTORY, r._row));
      let nextNum = 1;
      readAll(SHEETS.APPRAISAL_HISTORY).forEach(r => {
        const n = parseInt(String(r['Snapshot ID']).replace('SNAP', ''), 10);
        if (!isNaN(n) && n + 1 > nextNum) nextNum = n + 1;
      });
      results.forEach(r => {
        appendRow(SHEETS.APPRAISAL_HISTORY, {
          'Snapshot ID': 'SNAP' + String(nextNum++).padStart(4, '0'),
          'Period': r.appraisalPeriod, 'Staff ID': r.staffId, 'Staff Name': r.staffName,
          'Department': r.department, 'Designation': r.designation, 'Supervisor': r.supervisor,
          'Attendance Score': r.criteria.attendance, 'Leave Score': r.criteria.leaveManagement,
          'Permission Score': r.criteria.permissionHours, 'Late Coming Score': r.criteria.lateComing,
          'Early Leaving Score': r.criteria.earlyLeaving, 'Shift Compliance Score': r.criteria.shiftCompliance,
          'Discipline Score': r.criteria.discipline, 'Policy Compliance Score': r.criteria.policyCompliance,
          'Behaviour Score': r.criteria.behaviour, 'Training Score': r.criteria.training,
          'HR Remarks Score': r.criteria.hrRemarks,
          'HR Score': r.hrScore, 'Bonus Marks': r.bonusMarks, 'Memo Deduction': r.memoDeduction,
          'HR Final Score': r.hrFinalScore, 'HR Rating': r.hrRating,
          'Stores KPI %': r.storesKpiPct, 'Combined Final Score': r.combinedFinalScore, 'Grade': r.grade,
          'Generated On': new Date(), 'Approval Status': 'Calculated', 'Locked By': '', 'Locked On': ''
        });
      });
    } finally {
      lock.releaseLock();
    }
    logAudit(p.actorUserId, 'SAVE_APPRAISAL_SNAPSHOT', p.periodLabel + ' | ' + results.length + ' staff');
    return { periodLabel: p.periodLabel, staffCount: results.length };
  },

  // Marks every snapshot row of a period as 'Approved' - the final,
  // immutable step of the Draft/Calculated -> Approved appraisal state
  // machine (PHASE 9 / PHASE 17). Once locked, saveAppraisalSnapshot()
  // refuses to overwrite the period until it is explicitly re-opened.
  // Requires the period to already have a saved (Calculated) snapshot -
  // an appraisal cannot be approved before it has actually been
  // calculated/saved (PHASE 9: "must NOT be approved unless it is
  // actually in the correct submitted state").
  lockAppraisalSnapshot: function (p) {
    requireHR(p.actorRole);
    if (!p.periodLabel) throw new Error('Period Label is required');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const rows = readAll(SHEETS.APPRAISAL_HISTORY).filter(r => normalizePeriodLabel(r['Period']) === normalizePeriodLabel(p.periodLabel));
      if (!rows.length) throw new Error('No saved snapshot found for this period. Save the snapshot first.');
      if (rows.some(r => r['Approval Status'] === 'Approved')) throw new Error('This period is already Approved.');
      const now = new Date();
      rows.forEach(r => updateRow(SHEETS.APPRAISAL_HISTORY, r._row, {
        'Approval Status': 'Approved', 'Locked By': p.actorUserId, 'Locked On': now
      }));
      logAudit(p.actorUserId, 'APPROVE_APPRAISAL_SNAPSHOT', p.periodLabel + ' | ' + rows.length + ' staff locked');
      return { periodLabel: p.periodLabel, staffCount: rows.length, status: 'Approved' };
    } finally {
      lock.releaseLock();
    }
  },

  // Explicit, audited revision action: clears the 'Approved' lock on a
  // period so HR can knowingly recalculate/re-save it. This is the
  // "controlled re-open/revision process" PHASE 16 requires instead of
  // ever letting later data silently rewrite an approved result - the
  // prior approved figures remain in the Audit Log even after this runs.
  reopenAppraisalSnapshot: function (p) {
    requireHR(p.actorRole);
    if (!p.periodLabel) throw new Error('Period Label is required');
    if (!p.reason) throw new Error('A reason is required to re-open an approved appraisal period');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const rows = readAll(SHEETS.APPRAISAL_HISTORY).filter(r => normalizePeriodLabel(r['Period']) === normalizePeriodLabel(p.periodLabel));
      if (!rows.length) throw new Error('No saved snapshot found for this period.');
      if (!rows.some(r => r['Approval Status'] === 'Approved')) throw new Error('This period is not currently Approved.');
      rows.forEach(r => updateRow(SHEETS.APPRAISAL_HISTORY, r._row, {
        'Approval Status': 'Reopened', 'Locked By': '', 'Locked On': ''
      }));
      logAudit(p.actorUserId, 'REOPEN_APPRAISAL_SNAPSHOT', p.periodLabel + ' | reason: ' + p.reason);
      return { periodLabel: p.periodLabel, staffCount: rows.length, status: 'Reopened' };
    } finally {
      lock.releaseLock();
    }
  },

  // ---------- HR Management Summary (action-oriented buckets) ----------
  getHRManagementSummary: function (p) {
    // Department-wide summary - Manager/HR only (audit item 4: a Staff user
    // must not be able to pull every employee's HR standing).
    if (p.actorRole !== 'Manager' && p.actorRole !== 'HR') throw new Error('Not authorized: Manager/HR only');
    const results = api.getFullAppraisal(p);
    const byScoreDesc = results.slice().sort((a, b) => b.combinedFinalScore - a.combinedFinalScore);
    return {
      topPerformers: byScoreDesc.filter(r => r.combinedFinalScore >= HR_POLICY.TOP_PERFORMER_MIN_SCORE)
        .map(r => ({ staffId: r.staffId, staffName: r.staffName, finalScore: r.combinedFinalScore, rating: r.finalRating, grade: r.grade })),
      needingImprovement: byScoreDesc.filter(r => r.combinedFinalScore < HR_POLICY.IMPROVEMENT_BELOW_SCORE)
        .map(r => ({ staffId: r.staffId, staffName: r.staffName, finalScore: r.combinedFinalScore, rating: r.finalRating, grade: r.grade })),
      eligibleForPromotion: byScoreDesc.filter(r => r.combinedFinalScore >= HR_POLICY.PROMOTION_MIN_SCORE)
        .map(r => ({ staffId: r.staffId, staffName: r.staffName, finalScore: r.combinedFinalScore })),
      recommendedForIncrement: byScoreDesc.filter(r => r.combinedFinalScore >= HR_POLICY.INCREMENT_MIN_SCORE)
        .map(r => ({ staffId: r.staffId, staffName: r.staffName, finalScore: r.combinedFinalScore })),
      recommendedForTraining: byScoreDesc.filter(r => r.combinedFinalScore < HR_POLICY.TRAINING_BELOW_SCORE
        || r.criteria.training < 6 || r.criteria.discipline < 6)
        .map(r => {
          const maxOf = HR_APPRAISAL_POLICY.WEIGHTS;
          const weakAreas = Object.entries(r.criteria)
            .filter(([k, v]) => v < 0.6 * (maxOf[k] || 10))
            .map(([k]) => k);
          return { staffId: r.staffId, staffName: r.staffName, finalScore: r.combinedFinalScore, weakAreas };
        }),
      disciplinaryConcerns: byScoreDesc.filter(r => r.memos.length > 0)
        .map(r => ({ staffId: r.staffId, staffName: r.staffName, finalScore: r.combinedFinalScore, memos: r.memos }))
    };
  },

  // ---------- Dashboard (department averages, ranking, bonus/memo impact, trend) ----------
  getHRDashboard: function (p) {
    // Department-wide dashboard - Manager/HR only (see getHRManagementSummary).
    if (p.actorRole !== 'Manager' && p.actorRole !== 'HR') throw new Error('Not authorized: Manager/HR only');
    const results = api.getFullAppraisal(p);
    const byDepartment = {};
    results.forEach(r => {
      if (!byDepartment[r.department]) byDepartment[r.department] = { total: 0, count: 0 };
      byDepartment[r.department].total += r.combinedFinalScore;
      byDepartment[r.department].count += 1;
    });
    const departmentWise = Object.keys(byDepartment).map(dep => ({
      department: dep,
      averageScore: Math.round((byDepartment[dep].total / byDepartment[dep].count) * 100) / 100,
      staffCount: byDepartment[dep].count
    }));

    const staffRanking = results.slice().sort((a, b) => b.combinedFinalScore - a.combinedFinalScore)
      .map((r, i) => ({ rank: i + 1, staffId: r.staffId, staffName: r.staffName, department: r.department, finalScore: r.combinedFinalScore, rating: r.finalRating, grade: r.grade }));

    const bonusMemoImpact = results.map(r => ({
      staffId: r.staffId, staffName: r.staffName, hrScore: r.hrScore,
      bonusMarks: r.bonusMarks, memoDeduction: r.memoDeduction, finalScore: r.combinedFinalScore
    }));

    const overallAverage = results.length > 0
      ? Math.round((results.reduce((a, r) => a + r.combinedFinalScore, 0) / results.length) * 100) / 100 : 0;

    // Trend uses saved snapshots (HR Appraisal History), since Final Score for
    // past periods requires that period's own manual scores/bonuses/memos,
    // which only saveAppraisalSnapshot() preserves period-by-period.
    const history = readAll(SHEETS.APPRAISAL_HISTORY);
    const byPeriod = {};
    history.forEach(r => {
      const period = r['Period'];
      if (!byPeriod[period]) byPeriod[period] = { total: 0, count: 0 };
      byPeriod[period].total += Number(r['Combined Final Score']) || 0;
      byPeriod[period].count += 1;
    });
    // Sort period labels chronologically (e.g. "Jul-2026") so the trend chart
    // reflects true time order rather than the order rows were saved into
    // HR Appraisal History. Labels that don't parse as a month-year fall back
    // to a stable alphabetical position after all parsed ones. Uses the
    // shared periodSortKey_()/periodSortCompare_() helpers (see
    // normalizePeriodLabel() above) so this stays single-source with
    // getStaffAppraisalTrend()'s per-staff trend below.
    const monthlyTrend = Object.keys(byPeriod).map(period => ({
      period, averageScore: Math.round((byPeriod[period].total / byPeriod[period].count) * 100) / 100
    })).sort((a, b) => periodSortCompare_(a.period, b.period));
    const byYear = {};
    history.forEach(r => {
      const yearMatch = String(r['Period']).match(/\d{4}/);
      const year = yearMatch ? yearMatch[0] : 'Unknown';
      if (!byYear[year]) byYear[year] = { total: 0, count: 0 };
      byYear[year].total += Number(r['Combined Final Score']) || 0;
      byYear[year].count += 1;
    });
    const yearlyTrend = Object.keys(byYear).map(year => ({
      year, averageScore: Math.round((byYear[year].total / byYear[year].count) * 100) / 100
    })).sort((a, b) => String(a.year).localeCompare(String(b.year)));

    return { departmentWise, staffRanking, bonusMemoImpact, overallAverage, monthlyTrend, yearlyTrend };
  },

  // ---------- Per-Staff Appraisal Trend (Aug 2026) ----------
  // Same source data and same chronological sort as getHRDashboard()'s
  // monthlyTrend/yearlyTrend above (HR Appraisal History snapshots,
  // periodSortKey_()/periodSortCompare_()), just filtered to ONE staff
  // instead of department-averaged - lets a Manager/HR (or the Staff
  // themself) see whether a specific employee's Combined Final Score is
  // trending up or down period over period, not just the department
  // average. A Staff user can only ever pull their own trend - staffId is
  // force-set to the trusted session staffId by handle() for Staff
  // (STAFF_SCOPED_READ_ACTIONS below), with the same explicit ownership
  // check as getMyScore() as defense in depth.
  getStaffAppraisalTrend: function (p) {
    let staffId = p.staffId;
    if (p.actorRole === 'Staff') {
      if (staffId && p.actorStaffId && staffId !== p.actorStaffId) {
        throw new Error('Staff can only view their own Appraisal Trend');
      }
      staffId = p.actorStaffId || staffId;
    }
    if (!staffId) throw new Error('staffId is required');
    const staff = readAll(SHEETS.STAFF).find(s => s['Staff ID'] === staffId);
    if (!staff) throw new Error('Staff not found: ' + staffId);

    // Only this staff's saved snapshot rows - one row per Period normally,
    // but averaged defensively in case a period was ever re-saved without
    // the old row being cleared (saveAppraisalSnapshot() always clears
    // existing rows for a period first, so this is a safety net, not the
    // expected path).
    const rows = readAll(SHEETS.APPRAISAL_HISTORY).filter(r => r['Staff ID'] === staffId);

    const byPeriod = {};
    rows.forEach(r => {
      const period = r['Period'];
      if (!byPeriod[period]) byPeriod[period] = { total: 0, count: 0 };
      byPeriod[period].total += Number(r['Combined Final Score']) || 0;
      byPeriod[period].count += 1;
    });
    const monthlyTrend = Object.keys(byPeriod).map(period => ({
      period, finalScore: Math.round((byPeriod[period].total / byPeriod[period].count) * 100) / 100
    })).sort((a, b) => periodSortCompare_(a.period, b.period));

    const byYear = {};
    rows.forEach(r => {
      const yearMatch = String(r['Period']).match(/\d{4}/);
      const year = yearMatch ? yearMatch[0] : 'Unknown';
      if (!byYear[year]) byYear[year] = { total: 0, count: 0 };
      byYear[year].total += Number(r['Combined Final Score']) || 0;
      byYear[year].count += 1;
    });
    const yearlyTrend = Object.keys(byYear).map(year => ({
      year, averageScore: Math.round((byYear[year].total / byYear[year].count) * 100) / 100
    })).sort((a, b) => String(a.year).localeCompare(String(b.year)));

    // Simple up/down/flat/insufficient-data read on the two most recent
    // chronological snapshots - display/audit label only, same spirit as
    // WORKFLOW_KPI_RATING_BANDS (see HR_APPRAISAL_POLICY above): never
    // feeds back into any score, purely descriptive for the UI arrow.
    let direction = 'insufficient', change = 0;
    if (monthlyTrend.length >= 2) {
      const last = monthlyTrend[monthlyTrend.length - 1].finalScore;
      const prev = monthlyTrend[monthlyTrend.length - 2].finalScore;
      change = Math.round((last - prev) * 100) / 100;
      direction = change > 0 ? 'up' : (change < 0 ? 'down' : 'flat');
    }

    return {
      staffId, staffName: staff['Staff Name'], designation: staff['Designation'], section: staff['Section'],
      monthlyTrend, yearlyTrend, snapshotCount: monthlyTrend.length, direction, change
    };
  },

  // ---------- HR LIVE (TODAY'S) DASHBOARD ----------
  // Separate from getHRDashboard() above (that one is a PERIOD-wise Appraisal
  // Dashboard driven by Manual Scores/Recognition/Memo for a Period Label).
  // This one is a real-time snapshot for a single calendar day, built off the
  // existing Attendance Register + Leave Register + Staff Master - matching
  // the "Total Employees / Present Today / On Leave / Late Comers /
  // Permission Requests / Low Attendance" cards from the HR Dashboard spec.
  // p.date (optional, 'yyyy-MM-dd') - defaults to today in the script's
  // timezone. p.lowAttendanceThresholdPct (optional, default 75) and
  // p.month (optional, 'yyyy-MM', defaults to the month containing p.date)
  // control the Low Attendance calculation below.
  getHRLiveDashboard: function (p) {
    // P1 FIX #1 (audit): exposes department-wide HR data (Total Employees,
    // Present Today, On Leave, Late Comers, Permission Requests, Low
    // Attendance Staff) - Staff users must never be able to call this.
    // p.actorRole here is the server-verified role set by handle() from the
    // trusted session, never a client-supplied value, so this cannot be
    // bypassed by a forged request.
    if (p.actorRole !== 'Manager' && p.actorRole !== 'HR') {
      throw new Error('Not authorized: Manager/HR only');
    }
    const tz = Session.getScriptTimeZone();
    const dateObj = (p && p.date) ? new Date(p.date) : new Date();
    const dayLabel = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    const monthLabel = (p && p.month) || Utilities.formatDate(dateObj, tz, 'yyyy-MM');
    const threshold = (p && p.lowAttendanceThresholdPct) || 75;

    const activeStaff = readAll(SHEETS.STAFF).filter(s => s['Status'] !== 'Inactive');
    const totalEmployees = activeStaff.length;

    // Today's Attendance Register rows (one per staff per day, Status one of
    // Present / Absent / Half Day / Permission / Late - see attStatus select
    // in index.html). Keep only the latest row per staff for the day in case
    // of a re-mark.
    const todaysAttendance = {};
    readAll(SHEETS.ATTENDANCE).forEach(r => {
      if (!r['Date']) return;
      if (Utilities.formatDate(new Date(r['Date']), tz, 'yyyy-MM-dd') !== dayLabel) return;
      todaysAttendance[r['Staff ID']] = r; // last one wins if marked twice
    });
    const todaysRows = Object.values(todaysAttendance);
    const presentToday = todaysRows.filter(r => r['Status'] === 'Present').length;
    const lateComersToday = todaysRows.filter(r => r['Status'] === 'Late').length;
    const permissionRequestsToday = todaysRows.filter(r => r['Status'] === 'Permission' || Number(r['Permission Hours']) > 0).length;

    // On Leave Today: Approved Leave Register rows whose From/To span today.
    const onLeaveRows = readAll(SHEETS.LEAVE).filter(l => {
      if (l['Status'] !== 'Approved') return false;
      if (!l['From Date'] || !l['To Date']) return false;
      const from = Utilities.formatDate(new Date(l['From Date']), tz, 'yyyy-MM-dd');
      const to = Utilities.formatDate(new Date(l['To Date']), tz, 'yyyy-MM-dd');
      return dayLabel >= from && dayLabel <= to;
    });
    const onLeaveToday = onLeaveRows.length;

    // Low Attendance (current month so far): staff whose Present-day % of
    // marked days this month falls below threshold. Staff with no Attendance
    // rows yet this month are skipped (nothing to judge yet), not flagged.
    const monthAttendance = readAll(SHEETS.ATTENDANCE).filter(r =>
      r['Date'] && Utilities.formatDate(new Date(r['Date']), tz, 'yyyy-MM') === monthLabel);
    const byStaffMonth = {};
    monthAttendance.forEach(r => {
      const id = r['Staff ID'];
      if (!byStaffMonth[id]) byStaffMonth[id] = { marked: 0, present: 0, name: r['Staff Name'] };
      byStaffMonth[id].marked += 1;
      if (r['Status'] === 'Present') byStaffMonth[id].present += 1;
    });
    const lowAttendanceStaff = Object.keys(byStaffMonth)
      .map(id => {
        const s = byStaffMonth[id];
        const pct = s.marked > 0 ? Math.round((s.present / s.marked) * 10000) / 100 : 0;
        return { staffId: id, staffName: s.name, attendancePct: pct };
      })
      .filter(x => x.attendancePct < threshold)
      .sort((a, b) => a.attendancePct - b.attendancePct);

    return {
      date: dayLabel, month: monthLabel, totalEmployees, presentToday, onLeaveToday,
      lateComersToday, permissionRequestsToday,
      lowAttendanceCount: lowAttendanceStaff.length, lowAttendanceStaff
    };
  },

  // ---------- FORMULA VERIFICATION PANEL (audit item 12) ----------
  // Manager/HR only. Two parts:
  //  (a) runs the 8 fixed formula test cases from the spec against the
  //      REAL functions (computeHRScores/clamp/gradeForCombined/etc, not a
  //      re-implementation of them) so this can never silently drift from
  //      what the app actually computes;
  //  (b) live data checks (weightage total, duplicate Employee+Period
  //      count, bonus/memo presence) against whatever period the caller
  //      passes in (p.periodLabel/from/to), or against static policy
  //      config for the checks that don't depend on a period.
  // Every row gets its own PASS/WARNING/FAIL - failures are never hidden or
  // silently downgraded to a warning.
  getFormulaVerification: function (p) {
    if (p.actorRole !== 'Manager' && p.actorRole !== 'HR') throw new Error('Not authorized: Manager/HR only');
    const W = HR_APPRAISAL_POLICY.WEIGHTS, FW = HR_APPRAISAL_POLICY.FINAL_SCORE_WEIGHTS;
    const rows = [];
    const add = (check, expected, actual, status) => rows.push({ check, expected, actual, status });
    const eq = (a, b, tol) => Math.abs(Number(a) - Number(b)) <= (tol || 0.01);

    // ---- (a) static policy checks ----
    const kpiCount = Object.keys(W).length;
    add('HR KPI Count', 11, kpiCount, kpiCount === 11 ? 'PASS' : 'FAIL');
    const weightTotal = Math.round(Object.values(W).reduce((a, v) => a + v, 0) * 100) / 100;
    add('HR Weightage Total', '100%', weightTotal + '%', weightTotal === 100 ? 'PASS' : 'FAIL');
    add('Stores Contribution', '70%', FW.storesKpi + '%', FW.storesKpi === 70 ? 'PASS' : 'WARNING');
    add('HR Contribution', '30%', FW.hrAppraisal + '%', FW.hrAppraisal === 30 ? 'PASS' : 'WARNING');
    const combinedContribution = FW.storesKpi + FW.hrAppraisal;
    add('Combined Contribution', '100%', combinedContribution + '%', combinedContribution === 100 ? 'PASS' : 'FAIL');

    // ---- (b) formula test cases, run against the real functions ----
    const combinedOf = (storesPct, hrFinal) => Math.round((storesPct * FW.storesKpi / 100 + hrFinal * FW.hrAppraisal / 100) * 100) / 100;
    const tc = [];
    tc.push(['TC1: Stores 85 + HR 90 -> Combined', 86.50, combinedOf(85, 90)]);
    tc.push(['TC2: Stores 100 + HR 100 -> Combined', 100, combinedOf(100, 100)]);
    tc.push(['TC3: Stores 0 + HR 0 -> Combined', 0, combinedOf(0, 0)]);
    tc.push(['TC4: Base 95 + Bonus 10 -> HR Final (capped)', 100, clamp(95 + 10 - 0, 0, 100)]);
    tc.push(['TC5: Base 10 - Memo 20 -> HR Final (floored)', 0, clamp(10 + 0 - 20, 0, 100)]);
    tc.push(['TC6: Target 100, Approved 40 -> Achievement %', 40, computeAchievementPct(40, 100)]);
    tc.push(['TC7: Target 100, Approved 120 -> Effective Achievement % (capped)', 100, Math.min(computeAchievementPct(120, 100), 100)]);
    tc.forEach(([name, expected, actual]) => {
      const ok = isFinite(actual) && !isNaN(actual) && eq(expected, actual);
      add(name, expected, actual, ok ? 'PASS' : 'FAIL');
    });
    // TC8 checked separately (not through the numeric eq() harness above,
    // since a target of 0 now deliberately returns null/"N/A" - see
    // computeAchievementPct P1 fix - not a number to compare with tolerance).
    add('TC8: Target 0 -> Achievement % is N/A (no division error)', 'N/A (null)',
      computeAchievementPct(50, 0) === null ? 'N/A (null)' : computeAchievementPct(50, 0),
      computeAchievementPct(50, 0) === null ? 'PASS' : 'FAIL');

    // ---- (c) live data checks (only if a period was supplied) ----
    let bonusApplied = false, memoApplied = false, baseRangeOk = true, finalRangeOk = true, combinedRangeOk = true;
    const liveChecked = !!(p.periodLabel && p.from && p.to);
    if (liveChecked) {
      const results = api.getFullAppraisal(p);
      results.forEach(r => {
        if (r.bonusMarks > 0) bonusApplied = true;
        if (r.memoDeduction > 0) memoApplied = true;
        if (r.hrScore < 0 || r.hrScore > 100) baseRangeOk = false;
        if (r.hrFinalScore < 0 || r.hrFinalScore > 100) finalRangeOk = false;
        if (r.combinedFinalScore < 0 || r.combinedFinalScore > 100) combinedRangeOk = false;
      });
    }
    const rangeLabel = ok => liveChecked ? (ok ? '0-100 (live-verified)' : 'OUT OF RANGE') : '0-100 (by design - pass a period to live-verify)';
    add('HR Base Score Range', '0-100', rangeLabel(baseRangeOk), baseRangeOk ? 'PASS' : 'FAIL');
    add('HR Final Score Range', '0-100', rangeLabel(finalRangeOk), finalRangeOk ? 'PASS' : 'FAIL');
    add('Combined Score Range', '0-100', rangeLabel(combinedRangeOk), combinedRangeOk ? 'PASS' : 'FAIL');
    add('Bonus Applied', 'Yes/No', liveChecked ? (bonusApplied ? 'Yes' : 'No') : 'No period given', 'PASS');
    add('Memo Deduction Applied', 'Yes/No', liveChecked ? (memoApplied ? 'Yes' : 'No') : 'No period given', 'PASS');

    // ---- (d) duplicate + zero-target checks ----
    const dupReport = api.getDuplicateAppraisalReport(p);
    add('Duplicate Employee + Period', 0, dupReport.duplicateCount, dupReport.duplicateCount === 0 ? 'PASS' : 'WARNING');

    // "Zero Target Errors" = any Working Register/Monthly Target row with
    // Target <= 0 that produced a non-finite (NaN/Infinity) Achievement % or
    // KPI Score. computeAchievementPct/computeKpiScore already guard against
    // this everywhere in the app (see their comments), so this is a live
    // sweep to CONFIRM that guarantee against real data, not a re-check of
    // the formula in isolation.
    let zeroTargetErrors = 0;
    readAll(SHEETS.REGISTER).forEach(r => {
      const t = Number(r['Target']) || 0;
      const pctVal = Number(r['Achievement %']);
      const scoreVal = Number(r['KPI Score']);
      if (t <= 0 && (!isFinite(pctVal) && r['Achievement %'] !== undefined && r['Achievement %'] !== '')) zeroTargetErrors++;
      if (!isFinite(scoreVal) && r['KPI Score'] !== undefined && r['KPI Score'] !== '') zeroTargetErrors++;
    });
    add('Zero Target Errors', 0, zeroTargetErrors, zeroTargetErrors === 0 ? 'PASS' : 'FAIL');

    // FROZEN TARGET INTEGRITY (audit item - Monthly Target retroactive
    // change, spec section 4/21): every Approved Working Register row must
    // carry its own frozen 'Target' (written at submission - see
    // _addWorkingEntriesLocked) so computeStaffWorkflowKpiGroups_() can use
    // it instead of re-resolving a possibly-since-changed Monthly Target
    // live. Rows missing it are LEGACY / SNAPSHOT NOT AVAILABLE (spec
    // section 24) - flagged here as a WARNING (not FAIL, since this is
    // expected for data created before the freeze fix) so HR/Manager knows
    // which historical periods still rely on live re-resolution.
    let legacyTargetRows = 0;
    readAll(SHEETS.REGISTER).forEach(r => {
      if (r['Approval Status'] === 'Approved' &&
        (r['Target'] === '' || r['Target'] === undefined || r['Target'] === null)) legacyTargetRows++;
    });
    add('Legacy Target Snapshot Rows (Approved, pre-freeze data)', 0, legacyTargetRows,
      legacyTargetRows === 0 ? 'PASS' : 'WARNING');

    // WEIGHTAGE FREEZE CONFIRMATION (spec section 5/22 - Workflow
    // Weightage retroactive change): an Approved row's own stored
    // 'Weightage %' (frozen at submission - see _addWorkingEntriesLocked)
    // is what withTeamSplit()/computeStaffWorkflowKpiGroups_() actually use
    // (row['Weightage %'] -> 'Effective Weightage %'), never a live re-read
    // of the current Workflow Master. This count is INFORMATIONAL, not a
    // failure signal: it simply reports how many Approved rows currently
    // differ from their Workflow's live Weightage % (i.e. how many rows
    // are actively relying on the freeze because someone changed the
    // Workflow Master afterwards), confirming the freeze is real and in
    // effect rather than coincidentally always matching.
    const liveWeightageOf = {};
    readAll(SHEETS.WORKFLOW).forEach(w => { liveWeightageOf[w['Workflow ID']] = Number(w['Weightage %']) || 0; });
    let frozenWeightageRows = 0;
    readAll(SHEETS.REGISTER).forEach(r => {
      if (r['Approval Status'] !== 'Approved') return;
      const live = liveWeightageOf[r['Workflow ID']];
      const stored = Number(r['Weightage %']) || 0;
      if (live !== undefined && stored !== live) frozenWeightageRows++;
    });
    add('Historical Weightage Snapshot Rows (frozen, differ from current Master)', '>= 0',
      frozenWeightageRows, 'PASS');

    // HR LOCK GUARD LIVE CHECK (spec section 12 / INTEGRATION-011): confirms
    // _isAppraisalPeriodLocked() correctly reports lock state for the
    // period being verified here (if one was passed in) - this proves the
    // guard wired into submitManualScores()/addRecognition()/
    // deleteRecognition()/addDisciplinaryAction()/decideMemo() is reading
    // real HR Appraisal History data, not a stub.
    if (p.periodLabel) {
      const locked = _isAppraisalPeriodLocked(p.periodLabel);
      add('HR Appraisal Lock State (' + p.periodLabel + ')', 'Approved or not',
        locked ? 'Locked (Approved)' : 'Not locked', 'PASS');
    }

    // STAFF PERFORMANCE REPORT SHAPE CHECK (regression guard for the
    // "stuck on Loading..." bug fixed Aug 2026): getStaffPerformanceReport()
    // MUST return { rows: [...], summary: {...} } - the frontend's
    // REPORT_DEFS.staffPerf (isSummaryPlusRows:true) has always required
    // this shape, and silently reverting to a bare array here would
    // reintroduce the exact same uncaught-exception/stuck-Loading bug. Run
    // with a tiny 1-day range so this is cheap even on a large Register.
    try {
      const tz = Session.getScriptTimeZone();
      const todayLabel = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
      const probe = api.getStaffPerformanceReport({ from: p.from || todayLabel, to: p.to || todayLabel });
      const shapeOk = probe && Array.isArray(probe.rows) && probe.summary && typeof probe.summary === 'object';
      add('Staff Performance Report Shape ({rows, summary})', 'rows[] + summary{}',
        shapeOk ? 'rows[] + summary{}' : 'INVALID SHAPE', shapeOk ? 'PASS' : 'FAIL');
      const probe2 = api.getSectionPerformanceReport({ from: p.from || todayLabel, to: p.to || todayLabel });
      const shapeOk2 = probe2 && Array.isArray(probe2.rows) && probe2.summary && typeof probe2.summary === 'object';
      add('Section Performance Report Shape ({rows, summary})', 'rows[] + summary{}',
        shapeOk2 ? 'rows[] + summary{}' : 'INVALID SHAPE', shapeOk2 ? 'PASS' : 'FAIL');
    } catch (e) {
      add('Performance Report Shape Check', 'rows[] + summary{}', 'ERROR: ' + e.message, 'FAIL');
    }

    const anyFail = rows.some(r => r.status === 'FAIL');
    const anyWarn = rows.some(r => r.status === 'WARNING');
    const overall = anyFail ? 'FAIL' : (anyWarn ? 'WARNING' : 'PASS');
    add('Formula Validation', 'PASS', overall, overall);

    return { rows, overall, duplicates: dupReport.duplicates };
  },

  // ---------- Notifications (bell icon) - audit fix P0 ----------
  // p.actorUserId / p.actorRole are the server-verified session identity set
  // by handle() - never a client-supplied value - so a user can only ever
  // read/mark notifications addressed to their own User ID or their own Role.
  getNotifications: function (p) {
    const rows = readAll(SHEETS.NOTIFICATIONS)
      .filter(r => notificationBelongsToActor_(r, p.actorUserId, p.actorRole))
      .sort((a, b) => new Date(b['Created On']) - new Date(a['Created On']))
      .slice(0, 50) // most recent 50 - the bell is a live glance, not a full history screen
      .map(r => ({
        notificationId: r['Notification ID'],
        type: r['Type'],
        message: r['Message'],
        refId: r['Ref ID'],
        read: !!r['Read'],
        createdOn: r['Created On']
      }));
    return rows;
  },

  markNotificationRead: function (p) {
    if (!p.notificationId) throw new Error('notificationId is required');
    const row = readAll(SHEETS.NOTIFICATIONS).find(r => r['Notification ID'] === p.notificationId);
    if (!row) throw new Error('Notification not found');
    // OWNERSHIP CHECK - same pattern as replyToMemo: a client could try to
    // mark/read another user's notification by guessing its ID, so this is
    // checked against the server-verified actor, not trusted client input.
    if (!notificationBelongsToActor_(row, p.actorUserId, p.actorRole)) {
      throw new Error('You can only mark your own notifications as read');
    }
    if (!row['Read']) {
      updateRow(SHEETS.NOTIFICATIONS, row._row, { 'Read': true, 'Read On': new Date() });
    }
    return { success: true };
  },

  markAllNotificationsRead: function (p) {
    const now = new Date();
    let count = 0;
    readAll(SHEETS.NOTIFICATIONS).forEach(r => {
      if (!r['Read'] && notificationBelongsToActor_(r, p.actorUserId, p.actorRole)) {
        updateRow(SHEETS.NOTIFICATIONS, r._row, { 'Read': true, 'Read On': now });
        count++;
      }
    });
    return { success: true, count };
  }
};

// SIDEBAR AUTH FIX - same reasoning as resolveUiActor_() above; this one
// was missed in the first pass since it lives far below the other
// wrappers, but it is exactly as reachable via google.script.run and just
// as unauthenticated without this fix.
function ui_getFormulaVerification(params) {
  return api.getFormulaVerification(resolveUiActor_(params));
}

function ratingFor(pct) {
  if (pct >= 90) return 'Outstanding';
  if (pct >= 80) return 'Excellent';
  if (pct >= 70) return 'Good';
  if (pct >= 60) return 'Satisfactory';
  return 'Needs Improvement';
}

// ============================================================
// HR KPI APPRAISAL - HELPERS
// ============================================================
function clamp(n, lo, hi) {
  return Math.max(lo, Math.min(hi, n));
}

// Rating scale exactly as specified by HR (kept separate from the older
// ratingFor() above, which the existing workflow-based Appraisal tab/CSV
// export already depends on with different labels - untouched). Used for
// both the HR module's own 0-100 score AND the combined Final Score.
function ratingForHR(score) {
  const band = HR_APPRAISAL_POLICY.RATING_BANDS.find(b => score >= b.min);
  return band ? band.rating : 'Unsatisfactory';
}

// Grade (A+/A/B/C/D) for the combined Final Employee Score.
function gradeForCombined(score) {
  const band = HR_APPRAISAL_POLICY.GRADE_BANDS.find(b => score >= b.min);
  return band ? band.grade : 'D';
}

// Automated per-KPI Rating (Excellent/Very Good/Good/Average/Needs
// Improvement) + star Score (out of WORKFLOW_KPI_MAX_RATING_SCORE) for a
// given Achievement % - DISPLAY/AUDIT LABEL ONLY (see
// HR_APPRAISAL_POLICY.WORKFLOW_KPI_RATING_BANDS for why this never changes
// the actual KPI Contribution arithmetic). achievementPct === null (target
// Not Set/zero) returns null fields throughout - never a false "Needs
// Improvement" for a KPI that was never even measurable.
function resolveWorkflowKpiRating(achievementPct) {
  if (achievementPct === null || achievementPct === undefined) {
    return { rating: null, ratingScore: null, maxRatingScore: HR_APPRAISAL_POLICY.WORKFLOW_KPI_MAX_RATING_SCORE };
  }
  const pct = Number(achievementPct) || 0;
  const band = HR_APPRAISAL_POLICY.WORKFLOW_KPI_RATING_BANDS.find(b => pct >= b.min);
  return {
    rating: band ? band.rating : 'Needs Improvement',
    ratingScore: band ? band.ratingScore : 1,
    maxRatingScore: HR_APPRAISAL_POLICY.WORKFLOW_KPI_MAX_RATING_SCORE
  };
}

// Computes the 11 pure-HR KPI scores (0-100 total) from one HR Manual
// Scores row. If no entry exists yet for this staff/period, every KPI
// scores 0 (manualScoresEntered will be false on the caller's result).
// Each KPI's Achievement % is capped at 100 before applying its Weightage,
// exactly like the Stores KPI Module's own Achievement %-based scoring.
// ============================================================
// AUTOMATIC TRAINING SCORE - calculation functions
// ============================================================

// Server-side authoritative expiry: Training Date + Training_Master's
// Validity Months for that Training ID. Never trust a client-supplied
// Expiry Date - Training_Register.addTrainingRecord/updateTrainingRecord
// always recompute this themselves.
function calculateTrainingExpiry_(trainingDate, validityMonths) {
  if (!trainingDate) return null;
  const d = new Date(trainingDate);
  const months = Number(validityMonths) || 0;
  if (months <= 0) return null; // 0/blank Validity Months = never expires
  const exp = new Date(d.getFullYear(), d.getMonth() + months, d.getDate());
  exp.setDate(exp.getDate() - 1); // "12 months from 01-Aug" -> 31-Jul next year, not 01-Aug
  return exp;
}

// A record is eligible for a given appraisal period only when Present +
// Completed + trained on/before appraisalTo + not expired as of
// appraisalTo. Deliberately period-anchored, never TODAY() - so a past,
// closed appraisal period always recomputes the same eligibility result
// no matter when it's re-run (test 11 in the UAT suite).
function calculateTrainingEligibility_(record, appraisalTo) {
  const to = endOfDay(appraisalTo);
  if (TRAINING_POLICY.ATTENDANCE_OK.indexOf(record['Attendance']) === -1) return false;
  if (TRAINING_POLICY.COMPLETION_OK.indexOf(record['Completion Status']) === -1) return false;
  if (!record['Training Date']) return false;
  const trainingDate = new Date(record['Training Date']);
  if (trainingDate > to) return false; // future training, relative to the appraisal period, is excluded
  const expiry = record['Expiry Date'] ? new Date(record['Expiry Date']) : null;
  if (expiry && expiry < to) return false; // expired as of appraisal end date
  return true;
}

// Per-record score, from Assessment Score (0-100) via TRAINING_POLICY.SCORE_BANDS.
// Ineligible records are never passed a score above 0 by the caller (see
// getTrainingAppraisalScore_) - this function only handles the banding math.
function calculateTrainingScore_(assessmentScore) {
  const s = clamp(Number(assessmentScore) || 0, 0, 100);
  const band = TRAINING_POLICY.SCORE_BANDS.find(b => s >= b.min);
  return band ? band.score : 0;
}

// THE single authoritative Training Score for a staff member over an
// appraisal period: sums the per-record score of every ELIGIBLE
// Training_Register record whose Training Date falls in [from, to],
// capped at TRAINING_POLICY.MAX_TRAINING_SCORE (test 9 - multiple valid
// trainings never exceed the max). Used by getAutoHRAppraisal() and
// getFullAppraisal() - the ONLY two callers, so Training is always
// computed exactly one way.
function getTrainingAppraisalScore_(staffId, from, to) {
  const fromD = startOfDay(from), toD = endOfDay(to);
  const records = readAll(SHEETS.TRAINING_REGISTER).filter(r =>
    r['Staff ID'] === staffId && r['Training Date'] &&
    new Date(r['Training Date']) >= fromD && new Date(r['Training Date']) <= toD
  );
  const total = records.reduce((sum, r) => {
    const eligible = calculateTrainingEligibility_(r, toD);
    if (!eligible) return sum;
    return sum + calculateTrainingScore_(r['Assessment Score']);
  }, 0);
  return Math.min(total, TRAINING_POLICY.MAX_TRAINING_SCORE);
}

function computeHRScores(row) {
  const P = HR_APPRAISAL_POLICY, W = P.WEIGHTS;
  const scoreFor = (achievementPct, weight) => Math.round((clamp(achievementPct, 0, 100) * weight / 100) * 100) / 100;

  if (!row) {
    const criteria = {};
    Object.keys(W).forEach(k => criteria[k] = 0);
    return { hrScore: 0, criteria };
  }

  const workingDays = Number(row['Working Days']) || 0;
  const presentDays = Number(row['Present Days']) || 0;
  const leaveDays = Number(row['Leave Days']) || 0;
  const permissionHours = Number(row['Permission Hours']) || 0;
  const lateEntries = Number(row['Late Entries']) || 0;
  const earlyLeavingCount = Number(row['Early Leaving Count']) || 0;
  const shiftCompliancePct = Number(row['Shift Compliance %']) || 0;

  const attendancePct = workingDays > 0 ? (presentDays / workingDays) * 100 : 0;
  const leaveExcess = Math.max(0, leaveDays - P.SANCTIONED_LEAVE_DAYS);
  const leavePct = 100 - leaveExcess * P.LEAVE_PENALTY_PER_EXCESS_DAY;
  const permissionExcess = Math.max(0, permissionHours - P.ALLOWED_PERMISSION_HOURS);
  const permissionPct = 100 - permissionExcess * P.PERMISSION_PENALTY_PER_EXCESS_HOUR;
  const latePct = 100 - lateEntries * P.LATE_COMING_PENALTY_PER_ENTRY;
  const earlyPct = 100 - earlyLeavingCount * P.EARLY_LEAVING_PENALTY_PER_ENTRY;
  const shiftPct = shiftCompliancePct;

  const ratingPct = (field, max) => (Number(row[field]) || 0) / max * 100;

  const criteria = {
    attendance: scoreFor(attendancePct, W.attendance),
    leaveManagement: scoreFor(leavePct, W.leaveManagement),
    permissionHours: scoreFor(permissionPct, W.permissionHours),
    lateComing: scoreFor(latePct, W.lateComing),
    earlyLeaving: scoreFor(earlyPct, W.earlyLeaving),
    shiftCompliance: scoreFor(shiftPct, W.shiftCompliance),
    discipline: scoreFor(ratingPct('Discipline Rating', P.RATING_MAX.discipline), W.discipline),
    policyCompliance: scoreFor(ratingPct('Policy Compliance Rating', P.RATING_MAX.policyCompliance), W.policyCompliance),
    behaviour: scoreFor(ratingPct('Behaviour Rating', P.RATING_MAX.behaviour), W.behaviour),
    training: scoreFor(ratingPct('Training Rating', P.RATING_MAX.training), W.training),
    hrRemarks: scoreFor(ratingPct('HR Remarks Rating', P.RATING_MAX.hrRemarks), W.hrRemarks)
  };
  const hrScore = Math.round(Object.values(criteria).reduce((a, v) => a + v, 0) * 100) / 100;
  return { hrScore, criteria };
}

// Stores KPI % for a staff member in a date range - used ONLY to feed the
// combined Final Employee Score (Stores 70% + HR 30%); no longer folded
// into the HR module's own 100-mark score (that was the old "Productivity"
// criterion, removed to avoid double-counting Stores performance).
// Reuses the same Approved + Team-Split logic as the Dashboard/Appraisal
// tab, so this figure always matches "Total KPI %" shown elsewhere.
// DOUBLE-COUNTING FIX (audit item 16): ACT009 "Attendance" is a legacy
// Working Register activity (WF0048 Present / WF0049 Leave / WF0050
// Permission) that is kept for backward compatibility only - it is EXCLUDED
// from the Stores KPI % used in the combined Final Score, because the HR
// Appraisal module already owns Attendance (15) / Leave Management (8) /
// Permission (7) as 3 of its own 11 pure-HR KPIs (see HR_APPRAISAL_POLICY).
// Counting both would double-credit/double-penalize the same underlying
// fact in the blended Final Score = Stores 70% + HR 30%. The Working
// Register entry screen still accepts ACT009 entries (so nothing already
// logged is lost and any process still relying on it keeps working) and it
// still appears in raw Dashboard/Report totals - only the Final Score's
// Stores KPI % input excludes it.
const STORES_KPI_EXCLUDED_ACTIVITY_IDS = ['ACT009'];

// ============================================================
// SINGLE SOURCE OF TRUTH - FIXED MAX/POSSIBLE KPI SCORE (Option B, Aug 2026)
// ============================================================
// REPLACES the old "grows only as a Workflow gets its first Approved
// entry" denominator (previously: sum of each computeStaffWorkflowKpiGroups_()
// group's weightagePct, which only produces a group for a Workflow+Month
// combination the staff actually has an Approved entry for). That model
// made KPI % unexpectedly DROP the moment a staff's Approved work
// legitimately expanded into a brand-new Workflow, because the new
// Workflow's Weightage % entered the denominator immediately while its
// Score started at 0 - a bigger hit to the bottom than the top gained.
//
// NEW MODEL: Max/Possible Score (per Staff) = SUM of Weightage % of every
// ACTIVE + Applicable Workflow for that Staff's Section/Role, regardless
// of whether the Staff has an Approved entry for it yet. This is a FIXED
// total that changes ONLY when Workflow Master itself changes (a workflow
// is added/removed/(de)activated/reweighted) or the Staff's Section/Role
// mapping changes - it never grows or shrinks from day-to-day approval
// activity. ATTENDANCE (ACT009) is always excluded (see
// STORES_KPI_EXCLUDED_ACTIVITY_IDS) because it is scored separately inside
// HR Appraisal (Attendance/Leave/Permission criteria) - including it again
// here would double-count the same fact in the blended Final Score
// (Stores KPI x 70% + HR Appraisal x 30%).
//
// SINGLE SOURCE OF TRUTH (spec item 13): every KPI %/headline-score
// consumer that is meant to represent "Stores KPI" - computeStoresKPIPct(),
// getStoresKpiDebug(), getDashboard()'s Total KPI % card, getMyScore()'s
// Overall % - calls this helper for its denominator, so they can never
// drift apart from each other for the same Staff/period.
//
// staffId is accepted (not currently used to narrow the set) because
// Workflow Master has NO Section/Role applicability column today - audit
// (Aug 2026) confirmed every Staff is eligible for every Active Workflow,
// so the fixed total is one company-wide number applied identically to
// every Staff. If a Section/Role -> Workflow applicability mapping is
// introduced later, filter `workflows` here to that staffId's applicable
// subset BEFORE summing - callers do not need to change.
function computeStaffFixedActiveWorkflowMaxScore_(staffId, workflows) {
  const wfs = workflows || readAll(SHEETS.WORKFLOW);
  const total = wfs
    .filter(w => w['Status'] !== 'Inactive' &&
      STORES_KPI_EXCLUDED_ACTIVITY_IDS.indexOf(w['Activity ID']) === -1)
    .reduce((sum, w) => sum + (Number(w['Weightage %']) || 0), 0);
  return Math.round(total * 100) / 100;
}

function computeStoresKPIPct(staffId, from, to) {
  const approvedInRange = readAll(SHEETS.REGISTER).filter(r => {
    const d = new Date(r['Date']);
    return d >= from && d <= to && r['Approval Status'] === 'Approved' &&
      STORES_KPI_EXCLUDED_ACTIVITY_IDS.indexOf(r['Activity ID']) === -1;
  });
  const regRows = approvedInRange.filter(r => r['Staff ID'] === staffId);
  // FIX (Stores KPI Team-Split Dilution Fix, Aug 2026): Actual is aggregated
  // per Workflow+Month BEFORE Achievement %/KPI Score are computed - see
  // computeStaffWorkflowKpiGroups_() for why this was necessary. Previously
  // this summed each row's already-diluted 'Effective KPI Score'/'Effective
  // Weightage %' directly, which silently re-applied the full Monthly
  // Target to every row a staff had for the same Workflow+Month instead of
  // once to their combined Actual.
  const workflows = readAll(SHEETS.WORKFLOW);
  const monthlyTargets = readAll(SHEETS.TARGETS);
  const groups = computeStaffWorkflowKpiGroups_(regRows, approvedInRange, workflows, monthlyTargets);
  let earned = 0;
  groups.forEach(g => { earned += g.kpiScore; });
  // FIXED DENOMINATOR (Option B, Aug 2026) - see
  // computeStaffFixedActiveWorkflowMaxScore_() above. Replaces the old
  // "possible += g.weightagePct" accumulation, which only counted
  // Workflows the staff already had an Approved entry for.
  const possible = computeStaffFixedActiveWorkflowMaxScore_(staffId, workflows);
  const pctRaw = possible > 0 ? (earned / possible) * 100 : 0;
  // CAP AT 100 (Aug 2026 fix): 'earned' is built from each Register row's
  // FROZEN historical Weightage % (Historical Freeze policy), while
  // 'possible' is the CURRENT/live sum of Weightage % across today's
  // Active workflows (computeStaffFixedActiveWorkflowMaxScore_). If a
  // Workflow's Weightage % was reduced after entries were Approved, or a
  // Workflow was made Inactive after a staff already earned Approved score
  // under it, 'earned' can legitimately exceed 'possible' and pctRaw can
  // exceed 100 - which then pushed Combined Final Score (Stores KPI% x 70%
  // + HR Score x 30%) out of its documented 0-100 range (see Formula
  // Validation test harness: "Combined Score Range"). HR_APPRAISAL_POLICY
  // already states no-overachievement-credit as company policy
  // (MAX_ACHIEVEMENT_PCT_FOR_SCORING = 100, applied per-KPI in
  // computeKpiScore) - this mirrors the same 0-100 clamp already applied to
  // hrFinalScore, so Stores KPI % (and therefore Combined Final Score) can
  // never leave 0-100 regardless of Master Data edits made after the fact.
  return Math.round(clamp(pctRaw, 0, 100) * 100) / 100;
}

// ============================================================
// REPORTS MODULE - PRIVATE HELPERS (Aug 2026 rebuild)
// ============================================================
// Names start with '_' so handle()'s dispatcher can never reach them
// directly (see the '_' guard at the top of handle()) - exactly the same
// convention already used by _addWorkingEntriesLocked.

function _validateReportRange(p) {
  if (!p || !p.from || !p.to) throw new Error('From Date and To Date are required');
  if (new Date(p.from) > new Date(p.to)) throw new Error('From Date must be before or equal to To Date');
}

// List of "MMM-yyyy" month labels covering every calendar month that
// From/To touches (even partially) - e.g. 28 Jul to 3 Aug -> ['Jul-2026','Aug-2026'].
function _monthLabelsInRange(from, to) {
  const labels = [];
  const start = new Date(from), end = new Date(to);
  let cur = new Date(start.getFullYear(), start.getMonth(), 1);
  const endMonth = new Date(end.getFullYear(), end.getMonth(), 1);
  while (cur <= endMonth) {
    labels.push(MONTH_NAMES[cur.getMonth()] + '-' + cur.getFullYear());
    cur.setMonth(cur.getMonth() + 1);
  }
  return labels;
}

// Sums Actual across Approved Working Register rows, counting a team job
// ONCE via its shared Team Ref No (same dedup pattern as getTargetStatus()/
// the Reports "Total Physical Qty" column) - never once-per-participant.
function _dedupApprovedQty(rows) {
  const seenTeamRefs = {};
  let sum = 0;
  rows.forEach(r => {
    const ref = r['Team Ref No'];
    if (ref) {
      if (seenTeamRefs[ref]) return;
      seenTeamRefs[ref] = true;
    }
    sum += Number(r['Actual']) || 0;
  });
  return sum;
}

// SINGLE AUTHORITATIVE aggregation path for Activity / KPI-Workflow /
// Section / Date-Range / Monthly Target & Achievement reports. Returns one
// row per Active Workflow (optionally narrowed by p.activityId/workflowId)
// with Monthly Target resolved via getEffectiveTarget() for every calendar
// month the range touches (summed - never a sum of individual Working
// Register "Target" cells) and Approved Actual = Team-Ref-deduped physical
// quantity from Approved rows in [p.from, p.to], optionally further scoped
// to p.staffId's own rows and/or p.section's staff.
// p: { from, to, activityId?, workflowId?, section?, staffId? }
function _buildWorkflowPerfRows(p) {
  const from = startOfDay(p.from), to = endOfDay(p.to);
  const monthLabels = _monthLabelsInRange(p.from, p.to);

  const staffSectionOf = {};
  readAll(SHEETS.STAFF).forEach(s => staffSectionOf[s['Staff ID']] = s['Section']);

  let workflows = readAll(SHEETS.WORKFLOW).filter(w => w['Status'] !== 'Inactive' &&
    STORES_KPI_EXCLUDED_ACTIVITY_IDS.indexOf(w['Activity ID']) === -1);
  // ATTENDANCE (ACT009) EXCLUDED (Aug 2026, single-source parity - spec
  // item 13): this is the shared row-builder behind the Activity, KPI-
  // Workflow, Section, and Date-Range Performance Reports, so excluding it
  // here makes all of them match Stores KPI / Dashboard / My Score for the
  // same Staff/period, rather than only the Stores KPI functions excluding
  // it. Attendance itself is still fully visible in the dedicated
  // Attendance/Leave module and HR Appraisal - this only keeps it from
  // being double-counted into these general performance reports too.
  if (p.activityId) workflows = workflows.filter(w => w['Activity ID'] === p.activityId);
  if (p.workflowId) workflows = workflows.filter(w => w['Workflow ID'] === p.workflowId);

  const monthlyTargets = readAll(SHEETS.TARGETS);

  // Full date/section-scoped Approved set BEFORE any staffId narrowing -
  // kept separate from the (possibly staffId-filtered) working set below
  // purely so withTeamSplit() below can see every teammate's row and size
  // each Team Ref No group correctly, exactly like every other
  // withTeamSplit()/computeStaffWorkflowKpiGroups_ call in this file.
  let approvedRowsScoped = readAll(SHEETS.REGISTER).filter(r => {
    const d = new Date(r['Date']);
    return r['Approval Status'] === 'Approved' && d >= from && d <= to;
  });
  if (p.section) approvedRowsScoped = approvedRowsScoped.filter(r => staffSectionOf[r['Staff ID']] === p.section);
  let approvedRows = p.staffId
    ? approvedRowsScoped.filter(r => r['Staff ID'] === p.staffId)
    : approvedRowsScoped;

  const rowsByWorkflow = {};
  approvedRows.forEach(r => {
    (rowsByWorkflow[r['Workflow ID']] = rowsByWorkflow[r['Workflow ID']] || []).push(r);
  });

  // TEAM-SPLIT WEIGHTAGE FIX (Option B audit, Aug 2026): when this report is
  // scoped to one staff (p.staffId), a Team-Split workflow used to score
  // against the FULL company-wide Weightage % here - crediting that one
  // staff as if they alone owned the whole team job's weightage, while the
  // same staff's Staff Performance Report / Dashboard / Stores KPI / My
  // Score all correctly divide it via 'Effective Weightage %'
  // (computeStaffWorkflowKpiGroups_ -> withTeamSplit()). That let a team job
  // silently outscore a solo job of identical size in this report even
  // though every other KPI % consumer treated them consistently. Only
  // applies when p.staffId is set - the company/section-wide (no staffId)
  // shape of this report intentionally stays scored on the full Weightage %,
  // since there it is reporting the workflow's own total, not one person's
  // earned share.
  const splitRowsByWorkflow = p.staffId ? withTeamSplit(approvedRows, approvedRowsScoped) : null;

  return workflows.map(wf => {
    const wfRows = rowsByWorkflow[wf['Workflow ID']] || [];
    const approvedActual = Math.round(_dedupApprovedQty(wfRows) * 100) / 100;

    let target = 0, configured = false;
    monthLabels.forEach(m => {
      const resolved = getEffectiveTarget(wf['Workflow ID'], m, { monthlyTargets, workflows: [wf] });
      target += resolved.value;
      if (resolved.configured) configured = true;
    });
    target = Math.round(target * 100) / 100;

    // weightagePct = this staff's own Effective (Team-Split) share when
    // p.staffId is set - summed across their Approved rows for this
    // Workflow in the period, same 'Effective Weightage %' figure
    // computeStaffWorkflowKpiGroups_ uses. Solo (non-team) rows are
    // unaffected: withTeamSplit() only divides TEAM_WORKFLOW_IDS rows that
    // carry a Team Ref No, so weightagePct still equals the full
    // wf['Weightage %'] for a staff who did the job alone. Company/section-
    // wide calls (no p.staffId) keep the original full Weightage %.
    const weightagePct = p.staffId
      ? Math.round((splitRowsByWorkflow || [])
          .filter(r => r['Workflow ID'] === wf['Workflow ID'])
          .reduce((sum, r) => sum + (Number(r['Effective Weightage %']) || 0), 0) * 100) / 100
      : (Number(wf['Weightage %']) || 0);
    // Uncapped Achievement % (so a genuine 120% is visible/auditable as
    // "Over Achieved", never silently clipped) - only the KPI Score
    // contribution is capped, via computeKpiScore()'s existing over-
    // achievement control (HR_APPRAISAL_POLICY.MAX_ACHIEVEMENT_PCT_FOR_SCORING).
    const achievementPct = computeAchievementPct(approvedActual, target);
    const cappedForScoring = achievementPct === null ? null : Math.min(achievementPct, 100);
    const kpiScore = computeKpiScore(cappedForScoring, weightagePct);
    const balance = Math.round((target - approvedActual) * 100) / 100;

    return {
      activityId: wf['Activity ID'], activityName: wf['Activity Name'],
      workflowId: wf['Workflow ID'], workflowName: wf['Workflow Name'], kpiName: wf['KPI Name'],
      unit: wf['Unit'], weightagePct,
      monthlyTarget: target, targetConfigured: configured,
      approvedActual, balance, achievementPct, kpiScore, maxScore: weightagePct,
      entryCount: wfRows.length
    };
  });
}

// Status label for the Monthly Target & Achievement report / KPI-Workflow
// report (spec item 9). achievementPct is the UNCAPPED % (so >100% reads as
// Over Achieved, not silently clamped to Completed).
function _statusForAchievement(achievementPct, targetConfigured) {
  if (!targetConfigured || achievementPct === null || achievementPct === undefined) return 'Not Set';
  if (achievementPct <= 0) return 'Not Started';
  if (achievementPct < 100) return 'In Progress';
  if (achievementPct === 100) return 'Completed';
  return 'Over Achieved';
}

function requireManager(role) {
  if (role !== 'Manager') throw new Error('Manager role required for this action');
}

// HR Appraisal module is edited ONLY by the 'HR' role - a Stores/Department
// Manager (role 'Manager') can VIEW the Scorecard/HR Summary/Department
// Dashboard tabs (those API calls stay read-only, no requireHR needed) but
// is blocked here from writing/editing any HR score, Recognition, Memo, or
// Snapshot. Keeps "HR Appraisal must remain independent from Stores KPI,
// only HR users can edit HR Appraisal, Stores Managers cannot modify HR
// scores" enforced server-side (index.html also hides these controls from
// non-HR users, but that's UI convenience only - this is the real gate).
function requireHR(role) {
  if (role !== 'HR') throw new Error('HR role required for this action - Stores Managers cannot edit HR Appraisal data');
}

// ============================================================
// HR APPRAISAL LOCK GUARD (spec section 12 - "When HR Appraisal is
// locked: HR score must remain frozen... Do NOT bypass the existing HR
// lock" / test INTEGRATION-011/012)
// ============================================================
// saveAppraisalSnapshot() already refuses to overwrite an Approved period's
// snapshot rows in HR Appraisal History, but that alone does not stop the
// INPUTS that feed a live recompute (Manual Scores, Recognition, Memo
// decisions) from being edited after lock - submitManualScores(),
// addRecognition(), deleteRecognition(), addDisciplinaryAction(), and
// decideMemo() previously had no awareness of the lock at all, so editing
// any of them for an already-Approved period would silently change what
// getFullAppraisal()/getHRDashboard()/getHRManagementSummary() compute
// live for that period, even though the frozen APPRAISAL_HISTORY snapshot
// rows themselves stayed untouched - a real "locked appraisal silently
// changes" gap. This single helper is the one place that decides "is this
// Period currently locked", so every write path below stays consistent
// with lockAppraisalSnapshot()/reopenAppraisalSnapshot() and can never
// drift out of sync with each other.
function _isAppraisalPeriodLocked(periodLabel) {
  if (!periodLabel) return false;
  const wanted = normalizePeriodLabel(periodLabel);
  return readAll(SHEETS.APPRAISAL_HISTORY).some(r =>
    normalizePeriodLabel(r['Period']) === wanted && r['Approval Status'] === 'Approved');
}
function _requireAppraisalPeriodUnlocked(periodLabel) {
  if (_isAppraisalPeriodLocked(periodLabel)) {
    throw new Error('This period (' + periodLabel + ') is Approved and locked in HR Appraisal. ' +
      'Use "Re-open for Revision" on the Formula/Snapshot screen first before making this change.');
  }
}

// ============================================================
// SETUP / SEED (run this once manually from the Apps Script editor)
// ============================================================
function setup() {
  // P0 SAFETY GUARD (audit fix - data-loss risk): setup() is a DESTRUCTIVE
  // seed script - every buildSheet() call below deletes-and-recreates a
  // sheet by that name (see buildSheet()'s own guard above for why). It is
  // only meant to be run ONCE, manually, on a brand-new/empty spreadsheet.
  // This pre-flight check looks at every sheet setup() is about to build; if
  // ANY of them already exists with more than just a header row, setup()
  // aborts immediately - before touching a single sheet - instead of wiping
  // production data partway through a run. To intentionally reset a sheet
  // that setup() builds, delete or rename it yourself first, then re-run
  // setup(); to reset just one sheet on purpose, call
  // buildSheet(name, headers, rows, textColumnNames, true) for that sheet
  // directly instead of running the whole setup().
  const sheetsSetupBuilds = [
    SHEETS.COMPANY, SHEETS.SECTION, SHEETS.STAFF, SHEETS.ACTIVITY,
    SHEETS.WORKFLOW, SHEETS.REGISTER, SHEETS.ATTENDANCE, SHEETS.LEAVE,
    SHEETS.USERS, SHEETS.AUDIT, SHEETS.TARGETS, SHEETS.RECOGNITION,
    SHEETS.DISCIPLINARY, SHEETS.MANUAL_SCORES, SHEETS.APPRAISAL_HISTORY
  ];
  const existingWithData = sheetsSetupBuilds
    .map(name => SS.getSheetByName(name))
    .filter(s => s && s.getLastRow() > 1)
    .map(s => s.getName() + ' (' + (s.getLastRow() - 1) + ' row(s))');
  if (existingWithData.length > 0) {
    throw new Error(
      'setup() ABORTED before making any changes. The following sheet(s) ' +
      'already have data and would be PERMANENTLY DELETED if setup() ran: ' +
      existingWithData.join(', ') + '. setup() is only safe to run ONCE, ' +
      'on a brand-new empty spreadsheet. If this really is a fresh setup ' +
      'and that data is disposable test data, delete those sheets manually ' +
      'first, then re-run setup().'
    );
  }

  buildSheet(SHEETS.COMPANY, ['Company Name', 'Department'], [
    ["Sri Narasu's Coffee Company Private Limited", 'Stores Department']
  ]);

  buildSheet(SHEETS.SECTION, ['Section ID', 'Section Name'], [
    ['SEC001', 'Stores']
  ]);

  buildSheet(SHEETS.STAFF, ['Staff ID', 'Staff Name', 'Designation', 'Section', 'Reporting Manager', 'Date of Joining', 'Status'], [
    ['EMP001', 'G. Raja', 'Manager', 'Stores', 'Managing Director', '01-06-2013', 'Active'],
    ['EMP002', 'M. Sakthivel', 'Staff', 'Stores', 'G. Raja', '', 'Active'],
    ['EMP003', 'S. Gowtham', 'Staff', 'Stores', 'G. Raja', '', 'Active'],
    ['EMP004', 'S. Sabarinathan', 'Staff', 'Stores', 'G. Raja', '', 'Active'],
    ['EMP005', 'Anand Kumar', 'Staff', 'Stores', 'G. Raja', '', 'Active'],
    ['EMP006', 'C. Shankar', 'Staff', 'Stores', 'G. Raja', '', 'Active']
  ]);

  const activityRows = [
    ['ACT001', 'REC-PACK', 'Receiving - Packing Material'],
    ['ACT002', 'REC-NONPACK', 'Receiving - Non-Packing Material'],
    ['ACT003', 'BR-DESPATCH', 'Branch Despatch'],
    ['ACT004', 'HO-DESPATCH', 'Head Office Despatch'],
    ['ACT005', 'DESPATCH-PACKMAT', 'Despatch - Packing Material'],
    ['ACT006', 'DESPATCH-JOBWORK', 'Despatch - Job Work'],
    ['ACT007', 'HOUSEKEEPING', 'Housekeeping'],
    ['ACT008', 'MIS-REPORTS', 'MIS Reports'],
    ['ACT009', 'ATTENDANCE', 'Attendance']
  ];
  buildSheet(SHEETS.ACTIVITY, ['Activity ID', 'Activity Code', 'Activity Name'], activityRows);

  // Workflow master seeded with the REAL per-workflow Target / Weightage % / Max Score /
  // Expected Output / Status table supplied by the user (replaces the earlier equal-split
  // placeholder). Columns: WorkflowID, ActivityID, WorkflowName, KPIName, Unit, Target,
  // Weightage %, Max Score, Expected Output, Status.
  // NOTE: source row "WF0059 / Leave" is renumbered WF0049 here to keep IDs sequential
  // (it was almost certainly a typo for WF0049 in the original table).
  const actNameById = {};
  activityRows.forEach(a => actNameById[a[0]] = a[1]);
  // Weightage % / Max Score redesigned per PHYSICAL : SYSTEM effort-based policy
  // (Aug 2026 recalibration, confirmed by Manager): every one of the 47 scoring
  // workflows (WF0001-WF0047, ATTENDANCE/ACT009 excluded exactly as before - see
  // STORES_KPI_EXCLUDED_ACTIVITY_IDS) is first classified as either PHYSICAL
  // (hands-on/manual labour: unloading, picking, packing, material delivery,
  // arrangement, despatch, handover, truck loading, cleaning) or SYSTEM-BASED
  // (desk/paperwork: GRN entry, flow reports, signatures, approvals, DC
  // preparation, gate pass, stock verification, indents, MIS reports). Each
  // group's OLD total weightage was scaled onto a 60 (Physical) : 40 (System)
  // split of the full 100-mark pool, i.e. Physical workflows now collectively
  // outweigh System-based ones ~1.5x per the "physical work = more marks,
  // system work = less marks" instruction. Grand total still = 100.00 exactly.
  // Borderline calls (Checking, Scrap Disposal, Branch/Distributor/SS Selection)
  // were classified System/Physical as marked below - flag to Manager to
  // revisit if the real workflow doesn't match.
  const wfRaw = [
    ['WF0001', 'ACT001', 'Unloading & Bill Confirmation', 'Bills Received', 'Bills', 1, 2.878, 2.878, '100% bills received without shortage'], // Physical
    ['WF0002', 'ACT001', 'GRN Entry', 'GRNs Created', 'Bills', 1, 1.044, 1.044, '100% GRN created within target time'], // System
    ['WF0003', 'ACT001', 'Flow Report Preparation', 'Reports Prepared', 'Bills', 1, 1.044, 1.044, 'Daily flow report prepared accurately'], // System
    ['WF0004', 'ACT001', 'In-Charge Signature', 'Bills Approved', 'Bills', 1, 1.044, 1.044, '100% bills approved on time'], // System
    ['WF0005', 'ACT001', 'Weight / QC / Packing Sample Report', 'Reports Completed', 'Bills', 1, 2.158, 2.158, '100% QC & sample reports completed'], // Physical
    ['WF0006', 'ACT001', 'Bin & Lot Marking', 'Items Marked', 'Items', 1, 2.158, 2.158, '100% materials identified with Bin & Lot'], // Physical
    ['WF0007', 'ACT001', 'Bin Card Update', 'Bin Cards Updated', 'Items', 1, 1.044, 1.044, '100% Bin Cards updated on time'], // System
    ['WF0008', 'ACT001', 'Purchase Copy Submission', 'Bills Submitted', 'Bills', 1, 1.044, 1.044, 'Purchase copies submitted without delay'], // System
    ['WF0009', 'ACT002', 'Unloading', 'Bills Received', 'Bills', 1, 2.590, 2.590, '100% materials unloaded safely'], // Physical
    ['WF0010', 'ACT002', 'Bill Confirmation', 'Bills Verified', 'Bills', 1, 1.253, 1.253, '100% bills verified accurately'], // System
    ['WF0011', 'ACT002', 'GRN Entry', 'GRNs Created', 'Bills', 1, 0.940, 0.940, '100% GRNs created within target time'], // System
    ['WF0012', 'ACT002', 'Flow Report', 'Reports Prepared', 'Bills', 1, 0.940, 0.940, 'Daily reports completed accurately'], // System
    ['WF0013', 'ACT002', 'In-Charge Signature', 'Bills Approved', 'Bills', 1, 0.940, 0.940, '100% approvals completed on time'], // System
    ['WF0014', 'ACT002', 'Section Approval', 'Approvals Completed', 'Bills', 1, 0.836, 0.836, 'Section approvals completed without delay'], // System
    ['WF0015', 'ACT002', 'Material Delivery', 'Deliveries Completed', 'Bills', 1, 2.302, 2.302, 'Materials delivered to users on time'], // Physical
    ['WF0016', 'ACT002', 'Purchase Copy Submission', 'Bills Submitted', 'Bills', 1, 0.940, 0.940, 'Purchase copies submitted on time'], // System
    ['WF0017', 'ACT002', 'Bin / Lot Movement', 'Items Moved', 'Items', 1, 1.439, 1.439, '100% items moved and recorded correctly'], // Physical
    ['WF0018', 'ACT003', 'Branch / Distributor / Super Stockist', 'Destinations Covered', 'Towns', 1, 1.044, 1.044, 'All planned destinations served on schedule'], // System
    ['WF0019', 'ACT003', 'Picking', 'Items Picked', 'Items', 1, 2.302, 2.302, '100% items picked accurately'], // Physical
    ['WF0020', 'ACT003', 'Checking', 'Bills Verified', 'Bills', 1, 1.044, 1.044, 'Bills verified without errors'], // System
    ['WF0021', 'ACT003', 'Packing', 'Bundles Packed', 'Bundles/Kg', 1, 2.158, 2.158, 'Bundles packed correctly without damage'], // Physical
    ['WF0022', 'ACT003', 'DC Preparation', 'DC Created', 'DC', 1, 1.253, 1.253, 'Dispatch Challans prepared accurately'], // System
    ['WF0023', 'ACT003', 'Handover to Despatch', 'Bundles Handed Over', 'Bundles', 1, 2.014, 2.014, 'Bundles handed over without shortage'], // Physical
    ['WF0024', 'ACT003', 'Stock Verification', 'Items Checked', 'Items', 1, 1.358, 1.358, '100% stock verified before dispatch'], // System
    ['WF0025', 'ACT003', 'Gate Pass', 'Gate Pass Issued', 'Gate Pass', 1, 1.044, 1.044, 'Gate passes issued correctly'], // System
    ['WF0026', 'ACT004', 'Section Selection', 'Sections Served', 'Section', 1, 1.462, 1.462, 'Requested sections served on time'], // System
    ['WF0027', 'ACT004', 'Picking', 'Lines Picked', 'Lines', 1, 3.453, 3.453, '100% lines picked accurately'], // Physical
    ['WF0028', 'ACT004', 'DC Preparation', 'DC Created', 'DC', 1, 1.462, 1.462, 'DC prepared without errors'], // System
    ['WF0029', 'ACT004', 'Material Delivery', 'Deliveries Completed', 'Count', 1, 3.165, 3.165, 'Materials delivered as scheduled'], // Physical
    ['WF0030', 'ACT004', 'Gate Pass', 'Gate Pass Issued', 'Count', 1, 1.044, 1.044, 'Gate passes issued correctly'], // System
    ['WF0031', 'ACT004', 'Scrap Disposal', 'Scrap Transactions', 'Count', 1, 2.302, 2.302, 'Scrap disposed as per procedure'], // Physical
    ['WF0032', 'ACT005', 'Sales Indent Preparation', 'Indents Prepared', 'Sections', 1, 1.253, 1.253, 'Sales indents prepared accurately'], // System
    ['WF0033', 'ACT005', 'Picking', 'Lines Picked', 'Lines', 1, 2.590, 2.590, 'Packing materials picked correctly'], // Physical
    ['WF0034', 'ACT005', 'Arrangement', 'Material Arranged', 'Kg', 1, 1.871, 1.871, 'Materials arranged as per FIFO'], // Physical
    ['WF0035', 'ACT005', 'Checking', 'Lines Checked', 'Lines', 1, 1.253, 1.253, 'Dispatch items checked without errors'], // System
    ['WF0036', 'ACT005', 'Despatch', 'Pallets Despatched', 'Pallets', 1, 2.446, 2.446, 'Pallets dispatched on schedule'], // Physical
    ['WF0037', 'ACT005', 'DC Preparation', 'DC Created', 'DC', 1, 1.358, 1.358, 'DC prepared accurately'], // System
    ['WF0038', 'ACT005', 'Stock Verification', 'Items Checked', 'Items', 1, 1.567, 1.567, 'Stock verified before dispatch'], // System
    ['WF0039', 'ACT006', 'Picking', 'Lines Picked', 'Lines', 1, 4.029, 4.029, 'Job work materials picked accurately'], // Physical
    ['WF0040', 'ACT006', 'Checking', 'Lines Checked', 'Lines', 1, 1.880, 1.880, 'Materials checked before loading'], // System
    ['WF0041', 'ACT006', 'Truck Loading', 'Loads Completed', 'Vehicle/Load', 1, 5.755, 5.755, 'Vehicle loaded safely without damage'], // Physical - highest in group
    ['WF0042', 'ACT006', 'DC Preparation', 'DC Created', 'DC', 1, 1.462, 1.462, 'DC prepared accurately'], // System
    ['WF0043', 'ACT007', 'Cleaning Work', 'Area Cleaned', 'Area/Sq.ft', 1, 14.388, 14.388, 'Work area maintained as per 5S & GMP'], // Physical - only workflow in group, highest overall
    ['WF0044', 'ACT008', 'Minimum & Nil Stock Report', 'Reports Submitted', 'Report', 1, 2.611, 2.611, 'Report submitted within schedule'], // System
    ['WF0045', 'ACT008', 'Slow Moving Report', 'Reports Submitted', 'Report', 1, 2.611, 2.611, 'Report submitted accurately'], // System
    ['WF0046', 'ACT008', 'Non-Moving Report', 'Reports Submitted', 'Report', 1, 2.611, 2.611, 'Report submitted accurately'], // System
    ['WF0047', 'ACT008', 'Advertisement Material Report', 'Reports Submitted', 'Report', 1, 2.611, 2.611, 'Report submitted accurately'], // System
    ['WF0048', 'ACT009', 'Present', 'Attendance', 'Days', 1, 50, 50, '100% attendance as per working schedule'], // Attendance - excluded from Stores KPI scoring (STORES_KPI_EXCLUDED_ACTIVITY_IDS), left as-is
    ['WF0049', 'ACT009', 'Leave', 'Leave Taken', 'Days', 1, 25, 25, 'Leave maintained within company norms'], // Attendance - excluded from Stores KPI scoring
    ['WF0050', 'ACT009', 'Permission', 'Permission Hours', 'Hours', 1, 25, 25, 'Permission hours within approved limit'] // Attendance - excluded from Stores KPI scoring
  ];
  const wfFull = wfRaw.map(w => [
    w[0], w[1], actNameById[w[1]], w[2], w[3], w[4], w[5], w[6], w[7], w[8], 'Active'
  ]);
  buildSheet(SHEETS.WORKFLOW, ['Workflow ID', 'Activity ID', 'Activity Name', 'Workflow Name', 'KPI Name', 'Unit',
    'Target', 'Weightage %', 'Max Score', 'Expected Output', 'Status'], wfFull);

  buildSheet(SHEETS.REGISTER, ['Line ID', 'Entry Group ID', 'Date', 'Truck No', 'Staff ID', 'Staff Name', 'Logged By ID', 'Logged By Name',
    'Activity ID', 'Activity Name', 'Team Ref No', 'Workflow ID', 'Workflow Name', 'KPI Name', 'Unit', 'Target', 'Target Source', 'Actual',
    'Achievement %', 'Weightage %', 'Max Score', 'KPI Score', 'Remarks', 'Approval Status', 'Submitted On', 'Approved By', 'Approved On',
    'Rejected By', 'Rejected On', 'Rejection Reason'], []);

  buildSheet(SHEETS.ATTENDANCE, ['Date', 'Staff ID', 'Staff Name', 'Status', 'Permission Hours', 'Remarks'], []);
  applyAttendanceStatusValidation(SS.getSheetByName(SHEETS.ATTENDANCE));

  buildSheet(SHEETS.LEAVE, ['Leave ID', 'Staff ID', 'Staff Name', 'From Date', 'To Date', 'Leave Type',
    'Reason', 'Status', 'Applied On'], []);

  // 'HR' is a third, separate login role (alongside Manager/Staff) - ONLY
  // users with Role='HR' can write to the HR Appraisal module (Manual
  // Scores, Recognition, Memo issue/review/decide, Snapshot - see
  // requireHR() below). A Manager (Stores) can still view the HR Scorecard/
  // Summary/Department Dashboard tabs read-only, but every write endpoint
  // rejects them. Seeded 'hr1' user has no Staff ID (not Stores staff) -
  // change its password before going live, same as the other seeded users.
  // PASSWORD SECURITY (audit item 15): the default 'ChangeMe@123' is never
  // written to the sheet in plaintext, even for a brand-new install - it is
  // hashed with a fresh random salt right here, same as any real
  // changePassword() call, and 'Must Change Password' = Yes forces every
  // seeded account to set its own real password before it can be used for
  // anything else.
  (function seedUsers() {
    const DEFAULT_PW = 'ChangeMe@123';
    const seedRow = (userId, staffId, staffName, role) => {
      const salt = makeSalt();
      return [userId, hashPassword(DEFAULT_PW, salt), salt, staffId, staffName, role, 'Active', 'Yes'];
    };
    buildSheet(SHEETS.USERS, ['User ID', 'Password Hash', 'Salt', 'Staff ID', 'Staff Name', 'Role', 'Status', 'Must Change Password'], [
      seedRow('raja', 'EMP001', 'G. Raja', 'Manager'),
      seedRow('sakthivel', 'EMP002', 'M. Sakthivel', 'Staff'),
      seedRow('gowtham', 'EMP003', 'S. Gowtham', 'Staff'),
      seedRow('sabarinathan', 'EMP004', 'S. Sabarinathan', 'Staff'),
      seedRow('anand', 'EMP005', 'Anand Kumar', 'Staff'),
      seedRow('shankar', 'EMP006', 'C. Shankar', 'Staff'),
      seedRow('hr1', '', 'HR Admin', 'HR')
    ]);
    Logger.log('Seeded users with default password "' + DEFAULT_PW + '" (hashed, never stored in plaintext). ' +
      'Every seeded account has "Must Change Password" = Yes, so each person is forced to set their own real password on first login.');
  })();

  buildSheet(SHEETS.AUDIT, ['Timestamp', 'User ID', 'Action', 'Details'], []);

  buildSheet(SHEETS.TARGETS, ['Target ID', 'Month', 'Activity ID', 'Activity Name', 'Unit',
    'Workflow ID', 'Workflow Name', 'KPI Name', 'Monthly Target', 'Created By', 'Created On'], [], ['Month']);

  buildSheet(SHEETS.RECOGNITION,
    ['Recognition ID', 'Staff ID', 'Staff Name', 'Period', 'Type', 'Bonus Marks', 'Reason', 'Awarded By', 'Awarded On'], [], ['Period']);

  buildSheet(SHEETS.DISCIPLINARY,
    ['Memo ID', 'Memo Date', 'Staff ID', 'Staff Name', 'Department', 'Period', 'Type', 'Category',
      'Deduction Marks', 'Effective Deduction', 'Reason', 'Description', 'Status', 'Issued By', 'Issued On',
      'Employee Reply', 'Employee Replied On', 'Manager Review', 'Manager Reviewed By', 'Manager Reviewed On',
      'HR Review', 'Final Decision', 'Decided On', 'Attachment URL'], [], ['Period']);

  buildSheet(SHEETS.MANUAL_SCORES,
    ['Score ID', 'Staff ID', 'Staff Name', 'Period', 'Working Days', 'Present Days', 'Leave Days',
      'Permission Hours', 'Late Entries', 'Early Leaving Count', 'Shift Compliance %', 'Overtime Hours',
      'Discipline Rating', 'Policy Compliance Rating', 'Behaviour Rating', 'Training Rating',
      'HR Remarks Rating', 'HR Remarks', 'Entered By', 'Entered On'], [], ['Period']);

  buildSheet(SHEETS.APPRAISAL_HISTORY,
    ['Snapshot ID', 'Period', 'Staff ID', 'Staff Name', 'Department', 'Designation', 'Supervisor',
      'Attendance Score', 'Leave Score', 'Permission Score', 'Late Coming Score', 'Early Leaving Score',
      'Shift Compliance Score', 'Discipline Score', 'Policy Compliance Score', 'Behaviour Score',
      'Training Score', 'HR Remarks Score', 'HR Score', 'Bonus Marks', 'Memo Deduction', 'HR Final Score',
      'HR Rating', 'Stores KPI %', 'Combined Final Score', 'Grade', 'Generated On',
      'Approval Status', 'Locked By', 'Locked On'], [], ['Period']);

  // ---- Automatic Training Score module (Aug 2026) ----
  buildSheet(SHEETS.TRAINING_MASTER,
    ['Training ID', 'Training Name', 'Category', 'Mandatory', 'Validity Months', 'Max Score', 'Status'], [
      ['TR001', 'Safety Training', 'Safety', 'Yes', 12, 10, 'Active'],
      ['TR002', 'SOP Training', 'Operations', 'Yes', 12, 10, 'Active'],
      ['TR003', 'Fire Safety', 'Safety', 'Yes', 12, 10, 'Active']
    ], ['Training ID']);

  buildSheet(SHEETS.TRAINING_REGISTER,
    ['Training Record ID', 'Staff ID', 'Training ID', 'Training Name', 'Training Date', 'Trainer', 'Mode',
      'Attendance', 'Assessment Score', 'Completion Status', 'Certificate / Evidence', 'Expiry Date',
      'Eligible', 'Auto Score', 'Created By', 'Created On', 'Updated On'], [],
    ['Training Record ID', 'Staff ID', 'Training ID']);

  // ---- Notification module (Aug 2026 audit fix - P0) ----
  buildSheet(SHEETS.NOTIFICATIONS,
    ['Notification ID', 'To User ID', 'To Role', 'Type', 'Message', 'Ref ID', 'Read', 'Created On', 'Read On'], [],
    ['Notification ID', 'To User ID', 'To Role', 'Type', 'Ref ID']);

  SpreadsheetApp.flush();
  Logger.log('Setup complete. Sheets created/seeded: ' + Object.values(SHEETS).join(', '));
}

// Run this once (from the Apps Script editor) on an EXISTING spreadsheet that
// already has data. Unlike setup(), it only adds the Monthly Targets sheet and
// leaves every other sheet untouched.
function setupMonthlyTargetsOnly() {
  if (SS.getSheetByName(SHEETS.TARGETS)) {
    Logger.log('Monthly Targets sheet already exists - nothing to do.');
    return;
  }
  const s = SS.insertSheet(SHEETS.TARGETS);
  const headers = ['Target ID', 'Month', 'Activity ID', 'Activity Name', 'Unit',
    'Workflow ID', 'Workflow Name', 'KPI Name', 'Monthly Target', 'Created By', 'Created On'];
  s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
  s.getRange(1, headers.indexOf('Month') + 1, 1000, 1).setNumberFormat('@');
  s.setFrozenRows(1);
  s.autoResizeColumns(1, headers.length);
  SpreadsheetApp.flush();
  Logger.log('Monthly Targets sheet created.');
}

// Run this ONCE (from the Apps Script editor) if your Monthly Targets sheet
// already has rows in it and Target Status / Monthly Target Management is
// showing 0 / not finding rows you know you saved. This happens when Sheets
// silently auto-converted a "Jul-2026" style Month label into a real Date.
// This locks the Month column to plain text and rewrites every row's Month
// back to the correct "MMM-yyyy" label text, without touching anything else
// (Working Register, Appraisal, etc. are completely untouched).
function fixMonthlyTargetsMonthFormat() {
  const s = sh(SHEETS.TARGETS);
  const values = s.getDataRange().getValues();
  if (values.length < 2) { Logger.log('No Monthly Target rows to fix.'); return; }
  const headers = values[0];
  const monthCol = headers.indexOf('Month') + 1;
  if (monthCol === 0) { Logger.log('Month column not found.'); return; }
  s.getRange(1, monthCol, 1000, 1).setNumberFormat('@');
  let fixed = 0;
  for (let i = 1; i < values.length; i++) {
    if (values[i].join('') === '') continue;
    const raw = values[i][monthCol - 1];
    const label = normalizeMonthLabel(raw);
    if (raw !== label) {
      s.getRange(i + 1, monthCol).setValue(label);
      fixed++;
    }
  }
  SpreadsheetApp.flush();
  Logger.log('Fixed Month label on ' + fixed + ' row(s). Re-check Target Status now.');
}

// Run this ONCE (from the Apps Script editor) on an EXISTING spreadsheet that
// already has Working Register data, AFTER pasting the updated code.gs that
// makes KPI Score use Weightage % instead of Max Score. It is safe to re-run:
// - Adds a "Weightage %" column to Working Register if it isn't already there.
// - Backfills that column for every existing row from Workflow Master (matched
//   by Workflow ID).
// - Recomputes Achievement % / KPI Score for every row using the new formula
//   (Approved rows included - since Weightage % equals Max Score for every
//   originally-seeded workflow, this will NOT change historical numbers unless
//   you've since edited Weightage % or Max Score independently for a workflow).
// It does NOT touch any other sheet.
function migrateAddWeightageToRegister() {
  const s = sh(SHEETS.REGISTER);
  const lastCol = s.getLastColumn();
  const lastRow = s.getLastRow();
  let headers = s.getRange(1, 1, 1, lastCol).getValues()[0];
  let weightageCol = headers.indexOf('Weightage %') + 1; // 1-based, 0 if absent

  if (weightageCol === 0) {
    const maxScoreCol = headers.indexOf('Max Score') + 1;
    const insertAt = maxScoreCol > 0 ? maxScoreCol : lastCol + 1; // insert just before Max Score
    s.insertColumnBefore(insertAt);
    s.getRange(1, insertAt).setValue('Weightage %').setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
    weightageCol = insertAt;
    Logger.log('Inserted "Weightage %" column at position ' + insertAt + '.');
  } else {
    Logger.log('"Weightage %" column already exists at position ' + weightageCol + '.');
  }

  if (lastRow < 2) {
    Logger.log('No data rows to backfill.');
    return;
  }

  const workflows = readAll(SHEETS.WORKFLOW);
  headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; // refresh after possible insert
  const wfIdCol = headers.indexOf('Workflow ID') + 1;
  const targetCol = headers.indexOf('Target') + 1;
  const actualCol = headers.indexOf('Actual') + 1;
  const achPctCol = headers.indexOf('Achievement %') + 1;
  const kpiScoreCol = headers.indexOf('KPI Score') + 1;

  const data = s.getRange(2, 1, lastRow - 1, s.getLastColumn()).getValues();
  let updated = 0;
  data.forEach((row, i) => {
    if (row.join('') === '') return;
    const wfId = row[wfIdCol - 1];
    const wf = workflows.find(w => w['Workflow ID'] === wfId);
    const weightagePct = Number(wf ? wf['Weightage %'] : 0) || 0;
    const target = Number(row[targetCol - 1]) || 0;
    const actual = Number(row[actualCol - 1]) || 0;
    const achievementPct = computeAchievementPct(actual, target);
    const kpiScore = computeKpiScore(achievementPct, weightagePct);

    s.getRange(i + 2, weightageCol).setValue(weightagePct);
    s.getRange(i + 2, achPctCol).setValue(achievementPct === null ? 'N/A' : achievementPct);
    s.getRange(i + 2, kpiScoreCol).setValue(kpiScore === null ? 0 : kpiScore);
    updated++;
  });
  SpreadsheetApp.flush();
  Logger.log('Backfilled Weightage % and recomputed KPI Score for ' + updated + ' row(s).');
}

// Run this ONCE (from the Apps Script editor) on an EXISTING spreadsheet,
// AFTER pasting the updated code.gs that adds team-based score splitting.
// Adds a "Team Ref No" column to Working Register if it isn't already there
// (inserted just before "Workflow ID"). Existing rows are left blank in that
// column - they simply won't be split (same behaviour as before this fix),
// since there is no way to know which past entries belonged to the same
// job. New entries going forward will require and use it. Safe to
// re-run.
function migrateAddTeamRefNo() {
  const s = sh(SHEETS.REGISTER);
  const lastCol = s.getLastColumn();
  let headers = s.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('Team Ref No') !== -1) {
    Logger.log('"Team Ref No" column already exists.');
    return;
  }
  const wfIdCol = headers.indexOf('Workflow ID') + 1;
  const insertAt = wfIdCol > 0 ? wfIdCol : lastCol + 1; // insert just before Workflow ID
  s.insertColumnBefore(insertAt);
  s.getRange(1, insertAt).setValue('Team Ref No').setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
  SpreadsheetApp.flush();
  Logger.log('Inserted "Team Ref No" column at position ' + insertAt + '. Existing rows left blank (unsplit).');
}

// Run this ONCE (from the Apps Script editor) on an EXISTING spreadsheet,
// AFTER pasting the updated code.gs that adds the Truck Master feature.
// Adds a "Truck No" column to Working Register if it isn't already there
// (inserted just after "Date"). Existing rows are left blank - no way to
// retroactively know which truck an old entry belonged to. New entries
// going forward get one auto-generated by generateTruckNo(). Safe to re-run.
function migrateAddTruckNo() {
  const s = sh(SHEETS.REGISTER);
  const lastCol = s.getLastColumn();
  const headers = s.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.indexOf('Truck No') !== -1) {
    Logger.log('"Truck No" column already exists.');
    return;
  }
  const dateCol = headers.indexOf('Date') + 1;
  const insertAt = dateCol > 0 ? dateCol + 1 : lastCol + 1; // insert just after Date
  s.insertColumnBefore(insertAt);
  s.getRange(1, insertAt).setValue('Truck No').setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
  SpreadsheetApp.flush();
  Logger.log('Inserted "Truck No" column at position ' + insertAt + '. Existing rows left blank.');
}

// Run this ONCE (from the Apps Script editor) on an EXISTING spreadsheet,
// AFTER pasting the updated code.gs that replaces manual "Team Ref No" typing
// with ticking Co-Staff on the submission form. Adds:
//   - "Line ID"        : a unique key per row, needed so a Manager can Approve/
//                        Reject one participant's line without touching others.
//   - "Logged By ID" / "Logged By Name" : who actually filled the form, kept
//                        separate from "Staff ID"/"Staff Name" (who is credited),
//                        since a team submission credits several staff but only
//                        one of them logs it.
// Existing rows are backfilled: Line ID gets a generated unique value; Logged By
// ID/Name are set equal to the row's own Staff ID/Staff Name (the best available
// guess for old data, since who originally typed it in isn't recorded). Safe to
// re-run - rows that already have a Line ID are left untouched.
function migrateAddLineIdAndLoggedBy() {
  const s = sh(SHEETS.REGISTER);
  let headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const lastRow = s.getLastRow();

  function ensureColumn(name, insertBeforeName) {
    headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    if (headers.indexOf(name) !== -1) return;
    const beforeCol = insertBeforeName ? headers.indexOf(insertBeforeName) + 1 : 0;
    const insertAt = beforeCol > 0 ? beforeCol : s.getLastColumn() + 1;
    s.insertColumnBefore(insertAt);
    s.getRange(1, insertAt).setValue(name).setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
    Logger.log('Inserted "' + name + '" column at position ' + insertAt + '.');
  }

  ensureColumn('Line ID', 'Entry Group ID');
  ensureColumn('Logged By ID', 'Activity ID');
  ensureColumn('Logged By Name', 'Activity ID');

  if (lastRow < 2) { SpreadsheetApp.flush(); Logger.log('No data rows to backfill.'); return; }

  headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const lineIdCol = headers.indexOf('Line ID') + 1;
  const staffIdCol = headers.indexOf('Staff ID') + 1;
  const staffNameCol = headers.indexOf('Staff Name') + 1;
  const loggedByIdCol = headers.indexOf('Logged By ID') + 1;
  const loggedByNameCol = headers.indexOf('Logged By Name') + 1;

  const data = s.getRange(2, 1, lastRow - 1, s.getLastColumn()).getValues();
  let updated = 0;
  data.forEach((row, i) => {
    if (row.join('') === '') return;
    const r = i + 2;
    if (!row[lineIdCol - 1]) {
      s.getRange(r, lineIdCol).setValue('WL' + new Date().getTime() + '_' + i);
    }
    if (!row[loggedByIdCol - 1]) s.getRange(r, loggedByIdCol).setValue(row[staffIdCol - 1]);
    if (!row[loggedByNameCol - 1]) s.getRange(r, loggedByNameCol).setValue(row[staffNameCol - 1]);
    updated++;
  });
  SpreadsheetApp.flush();
  Logger.log('Backfilled Line ID / Logged By for ' + updated + ' row(s).');
}

// Run this ONCE (from the Apps Script editor) on an EXISTING spreadsheet,
// AFTER pasting the updated code.gs that resolves KPI Target as Monthly
// Target -> Default KPI Target ("Expected Output"). Safe to re-run.
//   - Adds a "Target Source" column to Working Register if it isn't already
//     there (inserted just after "Target").
//   - Recomputes Target / Target Source / Achievement % / KPI Score for every
//     row NOT already Approved or Rejected (so historical, already-decided
//     entries are left exactly as they were - only Draft/Submitted rows,
//     which can still be edited, are refreshed).
// It does NOT touch any other sheet.
// Run this ONCE (from the Apps Script editor) on an EXISTING spreadsheet,
// AFTER pasting the updated code.gs that gives Approve/Reject their own
// audit columns (P1 fix - see buildApprovalFields). Adds 'Rejected By',
// 'Rejected On', 'Rejection Reason' to Working Register if not already
// there (inserted right after 'Approved On'). Existing Rejected rows keep
// whatever is in 'Approved By'/'Approved On' from before this fix (that was
// a real, if mislabeled, record of who acted and when) - they are copied
// forward into the new Rejected columns so no audit history is lost, and
// 'Rejection Reason' is left blank for them since no reason was captured
// pre-fix. Safe to re-run.
function migrateAddApprovalAuditFields() {
  const s = sh(SHEETS.REGISTER);
  let headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const approvedOnCol = headers.indexOf('Approved On') + 1;
  const newCols = ['Rejected By', 'Rejected On', 'Rejection Reason'];
  let inserted = 0;
  newCols.forEach((name, idx) => {
    headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; // refresh each loop
    if (headers.indexOf(name) !== -1) { Logger.log('"' + name + '" already exists - skipped.'); return; }
    const insertAt = approvedOnCol + inserted + 1;
    s.insertColumnAfter(approvedOnCol + inserted);
    s.getRange(1, insertAt).setValue(name).setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
    inserted++;
    Logger.log('Inserted "' + name + '" column at position ' + insertAt + '.');
  });

  const lastRow = s.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.flush(); Logger.log('No data rows to backfill.'); return; }
  headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Approval Status') + 1;
  const approvedByCol = headers.indexOf('Approved By') + 1;
  const approvedOnCol2 = headers.indexOf('Approved On') + 1;
  const rejectedByCol = headers.indexOf('Rejected By') + 1;
  const rejectedOnCol = headers.indexOf('Rejected On') + 1;
  const data = s.getRange(2, 1, lastRow - 1, s.getLastColumn()).getValues();
  let backfilled = 0;
  data.forEach((row, i) => {
    if (row.join('') === '') return;
    if (row[statusCol - 1] !== 'Rejected') return;
    // Copy the old (mislabeled) Approved By/On into the new Rejected By/On,
    // then clear the old columns so a Rejected row no longer shows an
    // "Approved By" name.
    s.getRange(i + 2, rejectedByCol).setValue(row[approvedByCol - 1] || '');
    s.getRange(i + 2, rejectedOnCol).setValue(row[approvedOnCol2 - 1] || '');
    s.getRange(i + 2, approvedByCol).setValue('');
    s.getRange(i + 2, approvedOnCol2).setValue('');
    backfilled++;
  });
  SpreadsheetApp.flush();
  Logger.log('Migrated approval audit fields. Backfilled ' + backfilled + ' previously-Rejected row(s).');
}

function migrateAddTargetSourceToRegister() {
  const s = sh(SHEETS.REGISTER);
  let headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  let targetSourceCol = headers.indexOf('Target Source') + 1;

  if (targetSourceCol === 0) {
    const targetCol = headers.indexOf('Target') + 1;
    const insertAt = targetCol > 0 ? targetCol + 1 : s.getLastColumn() + 1;
    s.insertColumnAfter(targetCol > 0 ? targetCol : s.getLastColumn());
    s.getRange(1, insertAt).setValue('Target Source').setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
    targetSourceCol = insertAt;
    Logger.log('Inserted "Target Source" column at position ' + insertAt + '.');
  } else {
    Logger.log('"Target Source" column already exists at position ' + targetSourceCol + '.');
  }

  const lastRow = s.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.flush(); Logger.log('No data rows to backfill.'); return; }

  headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0]; // refresh after possible insert
  const dateCol = headers.indexOf('Date') + 1;
  const wfIdCol = headers.indexOf('Workflow ID') + 1;
  const targetCol = headers.indexOf('Target') + 1;
  const actualCol = headers.indexOf('Actual') + 1;
  const achPctCol = headers.indexOf('Achievement %') + 1;
  const weightagePctCol = headers.indexOf('Weightage %') + 1;
  const kpiScoreCol = headers.indexOf('KPI Score') + 1;
  const statusCol = headers.indexOf('Approval Status') + 1;

  const workflows = readAll(SHEETS.WORKFLOW);
  const monthlyTargets = readAll(SHEETS.TARGETS);
  const data = s.getRange(2, 1, lastRow - 1, s.getLastColumn()).getValues();
  let updated = 0;
  data.forEach((row, i) => {
    if (row.join('') === '') return;
    const status = row[statusCol - 1];
    if (status === 'Approved' || status === 'Rejected') return; // leave decided history untouched
    const wfId = row[wfIdCol - 1];
    const wf = workflows.find(w => w['Workflow ID'] === wfId);
    const month = monthLabelOf(row[dateCol - 1]);
    const resolved = getEffectiveTarget(wfId, month, { monthlyTargets, workflows });
    const target = resolved.configured ? resolved.value : (Number(row[targetCol - 1]) || 0);
    const actual = Number(row[actualCol - 1]) || 0;
    const achievementPct = computeAchievementPct(actual, target);
    const weightagePct = Number(wf ? wf['Weightage %'] : row[weightagePctCol - 1]) || 0;
    const kpiScore = computeKpiScore(achievementPct, weightagePct);

    s.getRange(i + 2, targetCol).setValue(target);
    s.getRange(i + 2, targetSourceCol).setValue(resolved.source);
    s.getRange(i + 2, achPctCol).setValue(achievementPct === null ? 'N/A' : achievementPct);
    s.getRange(i + 2, kpiScoreCol).setValue(kpiScore === null ? 0 : kpiScore);
    updated++;
  });
  SpreadsheetApp.flush();
  Logger.log('Backfilled Target Source and recomputed Target/Achievement %/KPI Score for ' + updated + ' non-decided row(s).');
}

// Run this ONCE (from the Apps Script editor) on an EXISTING spreadsheet to
// add the 4 new HR KPI Appraisal sheets (Recognition Register, Disciplinary
// Register, HR Manual Scores, HR Appraisal History). Non-destructive: skips
// any sheet that already exists, so it's safe to re-run and does not touch
// Staff Master, Working Register, Attendance, or anything else.
function setupHRAppraisalModule() {
  function ensureSheet(name, headers, textColumnNames) {
    if (SS.getSheetByName(name)) {
      Logger.log('"' + name + '" already exists - skipped.');
      return;
    }
    buildSheet(name, headers, [], textColumnNames);
    Logger.log('Created sheet "' + name + '".');
  }

  ensureSheet(SHEETS.RECOGNITION,
    ['Recognition ID', 'Staff ID', 'Staff Name', 'Period', 'Type', 'Bonus Marks', 'Reason', 'Awarded By', 'Awarded On'],
    ['Period']);

  ensureSheet(SHEETS.DISCIPLINARY,
    ['Memo ID', 'Memo Date', 'Staff ID', 'Staff Name', 'Department', 'Period', 'Type', 'Category',
      'Deduction Marks', 'Effective Deduction', 'Reason', 'Description', 'Status', 'Issued By', 'Issued On',
      'Employee Reply', 'Employee Replied On', 'Manager Review', 'Manager Reviewed By', 'Manager Reviewed On',
      'HR Review', 'Final Decision', 'Decided On', 'Attachment URL'],
    ['Period']);

  ensureSheet(SHEETS.MANUAL_SCORES,
    ['Score ID', 'Staff ID', 'Staff Name', 'Period', 'Working Days', 'Present Days', 'Leave Days',
      'Permission Hours', 'Late Entries', 'Early Leaving Count', 'Shift Compliance %', 'Overtime Hours',
      'Discipline Rating', 'Policy Compliance Rating', 'Behaviour Rating', 'Training Rating',
      'HR Remarks Rating', 'HR Remarks', 'Entered By', 'Entered On'],
    ['Period']);

  ensureSheet(SHEETS.APPRAISAL_HISTORY,
    ['Snapshot ID', 'Period', 'Staff ID', 'Staff Name', 'Department', 'Designation', 'Supervisor',
      'Attendance Score', 'Leave Score', 'Permission Score', 'Late Coming Score', 'Early Leaving Score',
      'Shift Compliance Score', 'Discipline Score', 'Policy Compliance Score', 'Behaviour Score',
      'Training Score', 'HR Remarks Score', 'HR Score', 'Bonus Marks', 'Memo Deduction', 'HR Final Score',
      'HR Rating', 'Stores KPI %', 'Combined Final Score', 'Grade', 'Generated On'],
    ['Period']);

  SpreadsheetApp.flush();
  Logger.log('setupHRAppraisalModule() complete.');
}

// Run this ONCE (from the Apps Script editor) on a spreadsheet that already
// has the OLD 8-criteria HR Manual Scores / HR Appraisal History sheets
// (Quality/Teamwork/Initiative/Service columns). Renames those two sheets to
// "...(Old)" as a backup and rebuilds them fresh with the new pure-HR v2
// schema (Attendance/Leave/Permission/Late/Early/Shift/Discipline/Policy/
// Behaviour/Training/HR Remarks). Old data is preserved in the renamed
// sheets but NOT auto-converted, since the old criteria (Quality, Teamwork,
// Initiative, Service, Productivity) have no equivalent in the new schema -
// review the "...(Old)" sheets manually before deleting them.
function migrateHRAppraisalV2() {
  [SHEETS.MANUAL_SCORES, SHEETS.APPRAISAL_HISTORY].forEach(name => {
    const sh = SS.getSheetByName(name);
    if (sh) {
      const backupName = name + ' (Old v1)';
      if (!SS.getSheetByName(backupName)) {
        sh.setName(backupName);
        Logger.log('Renamed "' + name + '" -> "' + backupName + '" (backup).');
      } else {
        Logger.log('Backup "' + backupName + '" already exists - leaving "' + name + '" untouched. Rename/delete manually if you want a re-migration.');
        return;
      }
    }
  });
  buildSheet(SHEETS.MANUAL_SCORES,
    ['Score ID', 'Staff ID', 'Staff Name', 'Period', 'Working Days', 'Present Days', 'Leave Days',
      'Permission Hours', 'Late Entries', 'Early Leaving Count', 'Shift Compliance %', 'Overtime Hours',
      'Discipline Rating', 'Policy Compliance Rating', 'Behaviour Rating', 'Training Rating',
      'HR Remarks Rating', 'HR Remarks', 'Entered By', 'Entered On'], [], ['Period']);
  buildSheet(SHEETS.APPRAISAL_HISTORY,
    ['Snapshot ID', 'Period', 'Staff ID', 'Staff Name', 'Department', 'Designation', 'Supervisor',
      'Attendance Score', 'Leave Score', 'Permission Score', 'Late Coming Score', 'Early Leaving Score',
      'Shift Compliance Score', 'Discipline Score', 'Policy Compliance Score', 'Behaviour Score',
      'Training Score', 'HR Remarks Score', 'HR Score', 'Bonus Marks', 'Memo Deduction', 'HR Final Score',
      'HR Rating', 'Stores KPI %', 'Combined Final Score', 'Grade', 'Generated On',
      'Approval Status', 'Locked By', 'Locked On'], [], ['Period']);
  SpreadsheetApp.flush();
  Logger.log('migrateHRAppraisalV2() complete. Old data backed up in "(Old v1)" sheets - review then delete when ready.');
}

// Run ONCE (from the Apps Script editor) on an EXISTING spreadsheet to add
// the separate 'HR' login role, now required by requireHR() to write to the
// HR Appraisal module (Manual Scores, Recognition, Memo issue/review/decide,
// Snapshot). Before this migration, those endpoints accepted 'Manager' -
// after re-deploying the updated code.gs/index.html, any existing 'Manager'
// user will get "HR role required" on those actions until you either (a)
// run this once to seed a starter HR login, or (b) manually add a Role='HR'
// row to the Users sheet / change an existing user's Role to 'HR'. Safe to
// re-run - skips if a Role='HR' user already exists.
function migrateAddHRRole() {
  const users = readAll(SHEETS.USERS);
  if (users.some(u => u['Role'] === 'HR')) {
    Logger.log('An HR-role user already exists in Users sheet - nothing to do.');
    return;
  }
  // SECURITY FIX (live-readiness audit, Aug 2026): this used to write the
  // seed password in plaintext to the 'Password' column and Logger.log it
  // in cleartext, unlike the hashed seedUsers() path used on a fresh
  // install - inconsistent with the rest of the hardening. Now hashes+salts
  // it the same way, blanks the legacy 'Password' column, and sets 'Must
  // Change Password' = Yes so hr1 is forced to set a real password on
  // first login instead of being able to stay on the default indefinitely.
  const DEFAULT_PW = 'ChangeMe@123';
  const salt = makeSalt();
  appendRow(SHEETS.USERS, {
    'User ID': 'hr1', 'Password': '', 'Password Hash': hashPassword(DEFAULT_PW, salt), 'Salt': salt,
    'Staff ID': '', 'Staff Name': 'HR Admin', 'Role': 'HR', 'Status': 'Active', 'Must Change Password': 'Yes'
  });
  Logger.log('Added seed HR user (User ID: hr1). Password is hashed, not stored/logged in plaintext - ' +
    '"Must Change Password" is set to Yes, so hr1 will be forced to set a real password on first login. ' +
    'Add/convert real HR users the same way (a Users sheet row with Role = "HR").');
}

// P0 SAFETY GUARD (audit fix - data-loss risk): buildSheet() used to
// unconditionally delete-and-recreate any existing sheet with the same
// name, no questions asked. That's fine for a brand-new spreadsheet, but if
// setup() is ever re-run (by mistake, by a new team member who doesn't know
// better, or via a stale bookmark in the Apps Script editor) on a
// spreadsheet that already has real production data, every one of those
// sheets - Staff Master, Working Register, Users, Audit Log, everything -
// would be silently wiped with no prompt and no way to undo it.
// NOW: buildSheet() refuses to delete a sheet that already has data (more
// than just a header row) unless the caller explicitly passes force=true.
// setup() itself (below) also does its own pre-flight check across every
// sheet it's about to touch, so it aborts BEFORE deleting anything at all,
// not partway through. Non-destructive callers (migrateTrainingAutomation()
// etc.) were never affected - they only call buildSheet() for a sheet that
// doesn't exist yet.
function buildSheet(name, headers, rows, textColumnNames, force) {
  let s = SS.getSheetByName(name);
  if (s) {
    const hasData = s.getLastRow() > 1; // more than just the header row
    if (hasData && !force) {
      throw new Error(
        'buildSheet() refused to delete sheet "' + name + '" - it already ' +
        'has ' + (s.getLastRow() - 1) + ' data row(s). Recreating it would ' +
        'PERMANENTLY DELETE that data. If this sheet genuinely needs to be ' +
        'reset, call buildSheet(..., true) to confirm, or clear/rename the ' +
        'sheet yourself first.'
      );
    }
    SS.deleteSheet(s);
  }
  s = SS.insertSheet(name);
  s.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
  if (rows.length > 0) {
    s.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  // Force specific columns to stay plain text (e.g. "Jul-2026" style labels),
  // otherwise Sheets can silently auto-convert them into real Date values,
  // which breaks any later exact-string match against them.
  (textColumnNames || []).forEach(colName => {
    const colIdx = headers.indexOf(colName) + 1;
    if (colIdx > 0) s.getRange(1, colIdx, 1000, 1).setNumberFormat('@');
  });
  s.setFrozenRows(1);
  s.autoResizeColumns(1, headers.length);
}

// ============================================================
// ATTENDANCE STATUS DROPDOWN FIX (Aug 2026)
// ============================================================
// Applies a dropdown to the Status column (col 4) covering every value in
// ATTENDANCE_STATUS_VALUES, using "show a warning" instead of "reject
// input" - so even if the list ever needs to change in future, a script
// write is never silently blocked by the sheet itself. The real
// enforcement lives server-side in markAttendance(); this dropdown is
// just so the column looks/behaves sensibly if anyone opens the sheet
// directly. Called automatically by setup() on a brand-new spreadsheet.
function applyAttendanceStatusValidation(sheet) {
  if (!sheet) return;
  const statusCol = 4; // Date, Staff ID, Staff Name, Status, ...
  const numRows = Math.max(sheet.getMaxRows() - 1, 999);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ATTENDANCE_STATUS_VALUES, true)
    .setAllowInvalid(true) // "show warning", never blocks a script write
    .build();
  sheet.getRange(2, statusCol, numRows, 1).setDataValidation(rule);
}

// Run this ONCE (from the Apps Script editor) on your EXISTING spreadsheet
// to fix an Attendance Register sheet that predates this change - e.g. if
// its Status dropdown was only ever set up with some of the 5 values
// (commonly missing Absent/Permission/Late). Non-destructive: only
// touches the data validation rule on the Status column, never the data
// itself. Safe to re-run any number of times.
function migrateAttendanceStatusDropdown() {
  const sheet = SS.getSheetByName(SHEETS.ATTENDANCE);
  if (!sheet) throw new Error('Sheet "' + SHEETS.ATTENDANCE + '" not found');
  applyAttendanceStatusValidation(sheet);
  Logger.log('Attendance Status dropdown fixed: ' + ATTENDANCE_STATUS_VALUES.join(', '));
}

// ============================================================
// AUTOMATIC TRAINING SCORE - safe migration (Aug 2026)
// ============================================================
// Run this ONCE (from the Apps Script editor) on an EXISTING spreadsheet
// that predates the Automatic Training Score module. Idempotent and
// non-destructive: only CREATES Training_Master/Training_Register if they
// don't already exist (never uses buildSheet()'s delete-and-recreate on an
// existing sheet - unlike setup(), which is only ever run on a brand-new,
// empty spreadsheet). Safe to re-run any number of times.
//
// Does NOT touch 'HR Manual Scores' (the 'Training Rating' column and any
// historical values already in it are left exactly as they are - see the
// comment on submitManualScores() for why) and does NOT recalculate or
// rewrite any existing 'HR Appraisal History' snapshot row, Approved or
// not - getFullAppraisal() only starts using the automatic Training score
// the next time it's actually called (a new period, or an explicitly
// reopened one), never retroactively.
function migrateTrainingAutomation() {
  let created = [];
  if (!SS.getSheetByName(SHEETS.TRAINING_MASTER)) {
    buildSheet(SHEETS.TRAINING_MASTER,
      ['Training ID', 'Training Name', 'Category', 'Mandatory', 'Validity Months', 'Max Score', 'Status'], [
        ['TR001', 'Safety Training', 'Safety', 'Yes', 12, 10, 'Active'],
        ['TR002', 'SOP Training', 'Operations', 'Yes', 12, 10, 'Active'],
        ['TR003', 'Fire Safety', 'Safety', 'Yes', 12, 10, 'Active']
      ], ['Training ID']);
    created.push(SHEETS.TRAINING_MASTER);
  }
  if (!SS.getSheetByName(SHEETS.TRAINING_REGISTER)) {
    buildSheet(SHEETS.TRAINING_REGISTER,
      ['Training Record ID', 'Staff ID', 'Training ID', 'Training Name', 'Training Date', 'Trainer', 'Mode',
        'Attendance', 'Assessment Score', 'Completion Status', 'Certificate / Evidence', 'Expiry Date',
        'Eligible', 'Auto Score', 'Created By', 'Created On', 'Updated On'], [],
      ['Training Record ID', 'Staff ID', 'Training ID']);
    created.push(SHEETS.TRAINING_REGISTER);
  }
  SpreadsheetApp.flush();
  Logger.log(created.length
    ? 'migrateTrainingAutomation() complete. Created: ' + created.join(', ') + '. ' +
      'Existing HR Manual Scores / HR Appraisal History left untouched. ' +
      'Add real Training_Master rows before recording Training_Register entries.'
    : 'migrateTrainingAutomation(): Training_Master and Training_Register already exist - nothing to do.');
}

// Run this ONCE (from the Apps Script editor) on an EXISTING, already-live
// spreadsheet to add the Notification Register sheet (P0 audit fix - the
// frontend bell already calls getNotifications/markNotificationRead/
// markAllNotificationsRead; this creates the backend storage those actions
// need). Safe to re-run - does nothing if the sheet already exists.
function migrateAddNotificationRegister() {
  if (SS.getSheetByName(SHEETS.NOTIFICATIONS)) {
    Logger.log('migrateAddNotificationRegister(): Notification Register already exists - nothing to do.');
    return;
  }
  buildSheet(SHEETS.NOTIFICATIONS,
    ['Notification ID', 'To User ID', 'To Role', 'Type', 'Message', 'Ref ID', 'Read', 'Created On', 'Read On'], [],
    ['Notification ID', 'To User ID', 'To Role', 'Type', 'Ref ID']);
  SpreadsheetApp.flush();
  Logger.log('migrateAddNotificationRegister() complete. Notification Register sheet created.');
}

// Run this ONCE (from the Apps Script editor) on an EXISTING spreadsheet that
// already has a "Disciplinary Register" sheet from before the Memo
// Management upgrade (Phase 2). Adds the new lifecycle columns (Memo Date,
// Department, Category, Description, Status, Employee Reply, Manager Review,
// HR Review, Final Decision, Attachment URL etc.) without touching any
// existing data. Existing memo rows get Status = "Issued" so they still show
// up correctly in the new Memo Management UI. Safe to re-run.
function migrateAddMemoManagementFields() {
  const s = sh(SHEETS.DISCIPLINARY);
  let headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];

  function ensureColumn(name, insertBeforeName) {
    headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    if (headers.indexOf(name) !== -1) return;
    const beforeCol = insertBeforeName ? headers.indexOf(insertBeforeName) + 1 : 0;
    const insertAt = beforeCol > 0 ? beforeCol : s.getLastColumn() + 1;
    s.insertColumnBefore(insertAt);
    s.getRange(1, insertAt).setValue(name).setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
    Logger.log('Inserted "' + name + '" column at position ' + insertAt + '.');
  }

  ensureColumn('Memo Date', 'Staff ID');
  ensureColumn('Department', 'Period');
  ensureColumn('Category', 'Deduction Marks');
  ensureColumn('Status', 'Issued By');
  ensureColumn('Description', 'Status');
  ensureColumn('Employee Reply', null);
  ensureColumn('Employee Replied On', null);
  ensureColumn('Manager Review', null);
  ensureColumn('Manager Reviewed By', null);
  ensureColumn('Manager Reviewed On', null);
  ensureColumn('HR Review', null);
  ensureColumn('Final Decision', null);
  ensureColumn('Decided On', null);
  ensureColumn('Attachment URL', null);

  const lastRow = s.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.flush(); Logger.log('No existing memo rows to backfill.'); return; }

  headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const staffIdCol = headers.indexOf('Staff ID') + 1;
  const deptCol = headers.indexOf('Department') + 1;
  const statusCol = headers.indexOf('Status') + 1;
  const memoDateCol = headers.indexOf('Memo Date') + 1;
  const issuedOnCol = headers.indexOf('Issued On') + 1;
  const staffRows = readAll(SHEETS.STAFF);

  const data = s.getRange(2, 1, lastRow - 1, s.getLastColumn()).getValues();
  let updated = 0;
  data.forEach((row, i) => {
    if (row.join('') === '') return;
    const r = i + 2;
    if (!row[statusCol - 1]) s.getRange(r, statusCol).setValue('Issued');
    if (!row[deptCol - 1]) {
      const staff = staffRows.find(x => x['Staff ID'] === row[staffIdCol - 1]);
      if (staff) s.getRange(r, deptCol).setValue(staff['Section']);
    }
    if (!row[memoDateCol - 1] && row[issuedOnCol - 1]) s.getRange(r, memoDateCol).setValue(row[issuedOnCol - 1]);
    updated++;
  });
  SpreadsheetApp.flush();
  Logger.log('migrateAddMemoManagementFields() complete. Backfilled ' + updated + ' existing memo row(s) with Status/Department/Memo Date.');
}

// Run this ONCE, manually from the Apps Script editor, on any spreadsheet
// that was set up BEFORE the 'Effective Deduction' column existed (memo
// withdrawn/adjusted deduction policy fix). Safe to re-run - does nothing
// to rows that already have a value in this column.
//
// Deduction Marks stays untouched (the original, for audit history).
// Effective Deduction is backfilled per existing Final Decision:
//   no Final Decision yet (still open)         -> 0 (BUSINESS RULE: never
//                                                 auto-deduct a memo with no
//                                                 confirmed penalty decision
//                                                 - see item 10, "if unclear,
//                                                 do not deduct")
//   any no-penalty decision (see
//     HR_POLICY.MEMO_DECISION_NO_PENALTY,
//     new + legacy labels)                      -> 0
//   Deduction Adjusted                          -> Deduction Marks (this
//                                                  sheet's Deduction Marks
//                                                  was already overwritten
//                                                  with the adjusted value by
//                                                  the old decideMemo() code,
//                                                  before this fix - the true
//                                                  original can't be
//                                                  recovered for these older
//                                                  rows)
//   Penalty Confirmed / Warning Upheld / Escalated -> Deduction Marks
function migrateAddEffectiveDeductionField() {
  const s = sh(SHEETS.DISCIPLINARY);
  let headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  if (headers.indexOf('Effective Deduction') === -1) {
    const insertAt = headers.indexOf('Deduction Marks') + 2; // right after Deduction Marks
    s.insertColumnBefore(insertAt);
    s.getRange(1, insertAt).setValue('Effective Deduction').setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
    Logger.log('Inserted "Effective Deduction" column at position ' + insertAt + '.');
  }
  headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const lastRow = s.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.flush(); Logger.log('No existing memo rows to backfill.'); return; }

  const deductionCol = headers.indexOf('Deduction Marks') + 1;
  const effectiveCol = headers.indexOf('Effective Deduction') + 1;
  const finalDecisionCol = headers.indexOf('Final Decision') + 1;
  const data = s.getRange(2, 1, lastRow - 1, s.getLastColumn()).getValues();
  let updated = 0;
  data.forEach((row, i) => {
    if (row.join('') === '') return;
    if (row[effectiveCol - 1] !== '' && row[effectiveCol - 1] !== null) return; // already backfilled
    const r = i + 2;
    const original = Number(row[deductionCol - 1]) || 0;
    const finalDecision = row[finalDecisionCol - 1];
    let effective;
    if (!finalDecision || HR_POLICY.MEMO_DECISION_NO_PENALTY.indexOf(finalDecision) !== -1) effective = 0;
    else effective = original; // Penalty Confirmed / Deduction Adjusted / legacy Warning Upheld / Escalated
    s.getRange(r, effectiveCol).setValue(effective);
    updated++;
  });
  SpreadsheetApp.flush();
  Logger.log('migrateAddEffectiveDeductionField() complete. Backfilled ' + updated + ' existing memo row(s).');
}

// Run this ONCE, manually from the Apps Script editor, to correct memo rows
// created by the PRIOR version of addDisciplinaryAction(), which incorrectly
// set 'Effective Deduction' = the full deduction at issue time (before any
// reply/review/Final Decision) - see the memo-reply/appraisal-score fix.
// This corrects the bug going forward WITHOUT silently changing any memo
// that already has a confirmed Final Decision (a Closed memo's Effective
// Deduction is exactly what decideMemo() intentionally set and must be left
// alone) and WITHOUT touching Deduction Marks (the original, for history).
//
// Only touches rows where:
//   - there is NO Final Decision yet (Status is not 'Closed'), AND
//   - Effective Deduction currently equals Deduction Marks (i.e. it still
//     holds the old pre-decision placeholder, not a value HR set)
// For those rows, Effective Deduction is reset to 0 (no confirmed penalty
// yet) and logged for HR review, per item 10 ("if the final decision is
// unavailable or ambiguous, do NOT automatically deduct - mark it for HR
// review instead"). Safe to re-run - already-zeroed/already-decided rows
// are left untouched every time.
function migrateZeroPendingMemoEffectiveDeduction() {
  const s = sh(SHEETS.DISCIPLINARY);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const lastRow = s.getLastRow();
  if (lastRow < 2) { Logger.log('No existing memo rows to check.'); return; }

  const memoIdCol = headers.indexOf('Memo ID') + 1;
  const statusCol = headers.indexOf('Status') + 1;
  const deductionCol = headers.indexOf('Deduction Marks') + 1;
  const effectiveCol = headers.indexOf('Effective Deduction') + 1;
  const finalDecisionCol = headers.indexOf('Final Decision') + 1;
  if (!statusCol || !deductionCol || !effectiveCol) {
    Logger.log('Required columns not found - run migrateAddMemoManagementFields()/migrateAddEffectiveDeductionField() first.');
    return;
  }
  const data = s.getRange(2, 1, lastRow - 1, s.getLastColumn()).getValues();
  let fixed = 0;
  const fixedMemoIds = [];
  data.forEach((row, i) => {
    if (row.join('') === '') return;
    const status = row[statusCol - 1];
    const finalDecision = finalDecisionCol ? row[finalDecisionCol - 1] : '';
    if (status === 'Closed' || finalDecision) return; // already decided - never touch
    const original = Number(row[deductionCol - 1]) || 0;
    const effective = row[effectiveCol - 1];
    // Only reset rows still holding the old "full deduction at issue" bug
    // value; a 0 (already correct) or any other HR-entered value is left as-is.
    if (original > 0 && Number(effective) === original) {
      s.getRange(i + 2, effectiveCol).setValue(0);
      fixed++;
      fixedMemoIds.push(row[memoIdCol - 1]);
    }
  });
  SpreadsheetApp.flush();
  Logger.log('migrateZeroPendingMemoEffectiveDeduction() complete. Reset ' + fixed +
    ' pending (not-yet-decided) memo row(s) to Effective Deduction = 0: ' + fixedMemoIds.join(', '));
}

// Run this ONCE, manually from the Apps Script editor, on any spreadsheet
// that was set up BEFORE the 'Must Change Password' column existed (audit
// item 31). Safe to re-run - does nothing if the column is already there.
// Existing accounts are left as 'No' (not forced to change) since they are
// already in real use; only brand-new accounts added afterwards should be
// created with 'Yes' if you want to force a change on first login.
// Run ONCE (from the Apps Script editor) on an EXISTING spreadsheet to add
// the Approved-snapshot lock fields to 'HR Appraisal History'. Purely
// additive (ensureColumn-style, never deletes/recreates the sheet - PHASE
// 19 sheet safety) and safe to re-run. Existing snapshot rows are backfilled
// with a blank Approval Status ('Not Set', not 'Approved') so nothing that
// was already saved under the old code is retroactively treated as locked -
// HR must explicitly run lockAppraisalSnapshot() per period going forward.
function migrateAddAppraisalApprovalFields() {
  const s = sh(SHEETS.APPRAISAL_HISTORY);
  function ensureColumn(name) {
    const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    if (headers.indexOf(name) === -1) {
      const col = s.getLastColumn() + 1;
      s.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
      Logger.log('Inserted "' + name + '" column at position ' + col + '.');
    }
  }
  ensureColumn('Approval Status');
  ensureColumn('Locked By');
  ensureColumn('Locked On');
  SpreadsheetApp.flush();
  Logger.log('migrateAddAppraisalApprovalFields() complete. Existing snapshot rows left unlocked - run lockAppraisalSnapshot() per period to approve them.');
}

function migrateAddMustChangePasswordField() {
  const s = sh(SHEETS.USERS);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  if (headers.indexOf('Must Change Password') !== -1) {
    Logger.log('Must Change Password column already exists - nothing to do.');
    return;
  }
  const col = s.getLastColumn() + 1;
  s.getRange(1, col).setValue('Must Change Password').setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
  const lastRow = s.getLastRow();
  if (lastRow >= 2) s.getRange(2, col, lastRow - 1, 1).setValue('No');
  SpreadsheetApp.flush();
  Logger.log('migrateAddMustChangePasswordField() complete.');
}

// Run this ONCE, manually from the Apps Script editor, on any spreadsheet
// that was set up BEFORE the password-hashing fix (audit item 15). Safe to
// re-run - does nothing if the columns already exist. Only ADDS the two new
// empty columns; it deliberately does NOT touch the existing 'Password'
// column or hash anything itself - login() does the actual per-user
// hash+blank the very next time each person logs in (see the "SAFE
// MIGRATION path" comment on api.login above), so no password is
// hashed/migrated without that user proving they still know it first.
// Run this ONCE, manually from the Apps Script editor, on any spreadsheet
// that was set up BEFORE the login-lockout fix (audit fix - brute-force
// protection). Safe to re-run - only adds columns that don't already exist.
// Purely additive (ensureColumn-style, never deletes/recreates the sheet),
// so it's safe to run on a Users sheet that already has real accounts in
// it. New columns are left blank for every existing row (0 failed
// attempts, not locked) - nobody gets retroactively locked out by running
// this. See LOGIN_LOCKOUT_POLICY and api.login() for how these are used.
function migrateAddLoginLockoutFields() {
  const s = sh(SHEETS.USERS);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  let added = 0;
  ['Failed Login Count', 'Locked Until', 'Last Failed Login'].forEach(name => {
    if (headers.indexOf(name) !== -1) return;
    const col = s.getLastColumn() + 1;
    s.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
    added++;
  });
  SpreadsheetApp.flush();
  Logger.log(added > 0
    ? 'migrateAddLoginLockoutFields() complete - added ' + added + ' column(s). Existing accounts start with 0 failed attempts and are not locked.'
    : 'Failed Login Count / Locked Until / Last Failed Login columns already exist - nothing to do.');
}

function migrateAddPasswordHashFields() {
  const s = sh(SHEETS.USERS);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  let added = 0;
  ['Password Hash', 'Salt'].forEach(name => {
    if (headers.indexOf(name) !== -1) return;
    const col = s.getLastColumn() + 1;
    s.getRange(1, col).setValue(name).setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
    added++;
  });
  SpreadsheetApp.flush();
  Logger.log(added > 0
    ? 'migrateAddPasswordHashFields() complete - added ' + added + ' column(s). Existing plaintext passwords are unchanged and will be hashed automatically the next time each user logs in.'
    : 'Password Hash / Salt columns already exist - nothing to do.');
}

// Run this ONCE, manually from the Apps Script editor, on the LIVE
// spreadsheet to add the 'Email' column to Staff Master - required for
// sendEmailNotification_() (see addNotification_ above) to have anywhere
// to send to. Safe/idempotent (no-op if the column already exists). Adds
// the column ONLY - does not and cannot populate real addresses; a
// Manager/HR user must fill in each staff member's Email afterwards
// (via the sheet directly, or the Edit Staff screen once a frontend field
// is added there - out of scope of this backend migration). Until an
// address is filled in for a given staff member, sendEmailNotification_()
// simply skips them (see its own comment) - the existing in-app bell
// notification is completely unaffected either way.
function migrateAddStaffEmailField() {
  const s = sh(SHEETS.STAFF);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  if (headers.indexOf('Email') !== -1) {
    Logger.log('Email column already exists on Staff Master - nothing to do.');
    return;
  }
  const col = s.getLastColumn() + 1;
  s.getRange(1, col).setValue('Email').setFontWeight('bold').setBackground('#6F4E37').setFontColor('#FFFFFF');
  SpreadsheetApp.flush();
  Logger.log('migrateAddStaffEmailField() complete - added \'Email\' column to Staff Master (blank for all existing staff). ' +
    'Fill in each staff member\'s email address to start receiving email notifications alongside the existing in-app bell notifications.');
}

// ============================================================
// WORKFLOW WEIGHTAGE RECALIBRATION - PHYSICAL:SYSTEM 60:40 (Aug 2026)
// ============================================================
// Run this ONCE, manually from the Apps Script editor, on the LIVE
// spreadsheet to push the new "physical work = more marks, system work =
// less marks" weightages (see the wfRaw table inside setup() above, which
// this function's data is copied from) onto an EXISTING Workflow Master
// sheet that already has real Working Register history against it.
//
// Deliberately does NOT call buildSheet()/setup() - it edits the 'Weightage
// %' and 'Max Score' cells of the 47 already-existing WF0001-WF0047 rows
// IN PLACE, by Workflow ID, leaving every other column (Workflow Name, KPI
// Name, Unit, Target, Expected Output, Status) and every other row
// (WF0048-WF0050 Attendance, and anything a Manager has added since) fully
// untouched. Safe to re-run - re-applies the same numbers every time.
//
// NOTE ON HISTORY: this only changes the Workflow Master going forward.
// Past Working Register rows keep whatever 'Weightage %'/'Max Score' was
// captured on them at submission time (by design - see
// migrateAddWeightageToRegister() above), so already-Approved KPI scores
// for earlier periods are NOT retroactively recalculated. Only NEW entries
// logged after this migration will use the new Physical:System weightages.
function migrateWorkflowWeightagePhysicalSystem60_40() {
  const NEW_WEIGHTAGE = {
    'WF0001': 2.878, 'WF0002': 1.044, 'WF0003': 1.044, 'WF0004': 1.044,
    'WF0005': 2.158, 'WF0006': 2.158, 'WF0007': 1.044, 'WF0008': 1.044,
    'WF0009': 2.590, 'WF0010': 1.253, 'WF0011': 0.940, 'WF0012': 0.940,
    'WF0013': 0.940, 'WF0014': 0.836, 'WF0015': 2.302, 'WF0016': 0.940,
    'WF0017': 1.439, 'WF0018': 1.044, 'WF0019': 2.302, 'WF0020': 1.044,
    'WF0021': 2.158, 'WF0022': 1.253, 'WF0023': 2.014, 'WF0024': 1.358,
    'WF0025': 1.044, 'WF0026': 1.462, 'WF0027': 3.453, 'WF0028': 1.462,
    'WF0029': 3.165, 'WF0030': 1.044, 'WF0031': 2.302, 'WF0032': 1.253,
    'WF0033': 2.590, 'WF0034': 1.871, 'WF0035': 1.253, 'WF0036': 2.446,
    'WF0037': 1.358, 'WF0038': 1.567, 'WF0039': 4.029, 'WF0040': 1.880,
    'WF0041': 5.755, 'WF0042': 1.462, 'WF0043': 14.388, 'WF0044': 2.611,
    'WF0045': 2.611, 'WF0046': 2.611, 'WF0047': 2.611
    // WF0048-WF0050 (Attendance) deliberately excluded - see function header.
  };
  const s = sh(SHEETS.WORKFLOW);
  const headers = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  const idCol = headers.indexOf('Workflow ID') + 1;
  const weightageCol = headers.indexOf('Weightage %') + 1;
  const maxScoreCol = headers.indexOf('Max Score') + 1;
  if (!idCol || !weightageCol || !maxScoreCol) {
    Logger.log('Required columns (Workflow ID / Weightage % / Max Score) not found on ' + SHEETS.WORKFLOW + ' - aborting.');
    return;
  }
  const lastRow = s.getLastRow();
  if (lastRow < 2) { Logger.log('No workflow rows found - nothing to migrate.'); return; }

  const ids = s.getRange(2, idCol, lastRow - 1, 1).getValues();
  let updated = 0;
  const missing = [];
  let newTotal = 0;
  ids.forEach((row, i) => {
    const wfId = row[0];
    if (!wfId) return;
    if (Object.prototype.hasOwnProperty.call(NEW_WEIGHTAGE, wfId)) {
      const val = NEW_WEIGHTAGE[wfId];
      s.getRange(i + 2, weightageCol).setValue(val);
      s.getRange(i + 2, maxScoreCol).setValue(val);
      newTotal += val;
      updated++;
    } else if (wfId !== 'WF0048' && wfId !== 'WF0049' && wfId !== 'WF0050') {
      missing.push(wfId); // unexpected extra workflow row not covered by this recalibration
    }
  });
  SpreadsheetApp.flush();
  Logger.log('migrateWorkflowWeightagePhysicalSystem60_40() complete. Updated ' + updated +
    ' of 47 expected workflow row(s). New Physical:System-scaled total = ' +
    (Math.round(newTotal * 100) / 100) + ' (should be ~100.00).' +
    (missing.length ? ' Rows found on the sheet but NOT in the recalibration table (left untouched - review manually): ' + missing.join(', ') : ''));
}
