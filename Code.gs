/**
 * DepEd Region XI Classroom Thermal Conditions Survey
 * Public survey/dashboard; authorized administrators can review, export, and delete responses.
 *
 * SETUP:
 * 1. Create a NEW Google Spreadsheet.
 * 2. Copy its ID from the spreadsheet URL.
 * 3. Replace PASTE_NEW_SPREADSHEET_ID_HERE below with that ID.
 */
const SPREADSHEET_ID = '1rZndi93gwRzYC53rkgvZf4cdr9ZWNjR1p7gqtGoBwh0';
const RESPONSE_SHEET = 'ThermalSurveyResponses';
const AUDIT_SHEET = 'ThermalSurveyAuditLog';
const ADMIN_EMAILS = [
  'shellalu.delacerna@deped.gov.ph',
  'pprd.region11@deped.gov.ph'
];

const HEADERS = [
  'School ID', 'School Name', 'Sector', 'Division', 'Municipality',
  'Classification', 'Location', 'Total Learners', 'Total Classrooms',
  'Excessive Heat (Q1)', 'Hot Classrooms (Q2)', 'Thermal Comfort (Q10)',
  'Total Est. Cost (PHP)', 'Last Updated', 'Full JSON Payload'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('DepEd Region XI | Classroom Thermal Conditions Survey')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDb_() {
  if (!SPREADSHEET_ID) {
    throw new Error('Please add your new Google Spreadsheet ID in Code.gs first.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateSheet_() {
  const ss = getDb_();
  let sh = ss.getSheetByName(RESPONSE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RESPONSE_SHEET);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#fff');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < HEADERS.length) {
    sh.getRange(1, sh.getLastColumn() + 1, 1, HEADERS.length - sh.getLastColumn())
      .setValues([HEADERS.slice(sh.getLastColumn())]);
  }
  return sh;
}

function getAudit_() {
  const ss = getDb_();
  let sh = ss.getSheetByName(AUDIT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(AUDIT_SHEET);
    sh.appendRow(['Timestamp', 'Admin Email', 'Action', 'School ID', 'Details']);
    sh.getRange(1, 1, 1, 5).setFontWeight('bold')
      .setBackground('#7f1d1d').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function currentAdmin_() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  if (!ADMIN_EMAILS.includes(email)) {
    throw new Error(
      'Administrator access denied. The signed-in email was not available or is not authorized. ' +
      'Detected account: ' + (email || 'not available')
    );
  }
  return email;
}

function logAdmin_(email, action, schoolId, details) {
  getAudit_().appendRow([new Date(), email, action, schoolId || '', details || '']);
}

/** Called by the survey form. Saves or updates one record per School ID. */
function saveSurveyResponse(payload) {
  if (!payload) throw new Error('No survey data received.');
  const sid = String(payload.schoolId || payload.profile?.schoolId || '').trim();
  if (!sid) throw new Error('School ID is required.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getOrCreateSheet_();
    const values = sh.getDataRange().getValues();
    let row = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === sid) {
        row = i + 1;
        break;
      }
    }

    const p = payload.profile || {};
    const p3 = payload.part3 || {};
    const record = [
      sid, p.schoolName || '', p.sector || '', p.division || '',
      p.municipality || '', p.classification || '', p.location || '',
      p.totalLearners || 0, p.totalClassrooms || 0,
      p3.q1 || '', p3.q2 || 0, p3.q10 || '',
      payload.totalEstimatedCost || 0, new Date(), JSON.stringify(payload)
    ];

    if (row > 0) {
      sh.getRange(row, 1, 1, record.length).setValues([record]);
    } else {
      sh.appendRow(record);
    }
    return { ok: true, schoolId: sid, savedAt: new Date().toISOString() };
  } finally {
    lock.releaseLock();
  }
}

/** Public data only: status directory and high-level part completion counts. */
function getPublicDashboard() {
  const sh = getOrCreateSheet_();
  const vals = sh.getDataRange().getValues();
  const rows = [];
  const partNames = [
    'Part II • School Profile', 'Part III • Thermal Conditions',
    'Part IV • Infrastructure', 'Part V • Effects',
    'Part VI • Incidents', 'Part VII • Preparedness',
    'Part VIII • Resource Interventions', 'Part IX • Classroom Assessment',
    'Part X • Open-ended'
  ];
  const partCounts = partNames.map(name => ({ name, count: 0 }));

  for (let i = 1; i < vals.length; i++) {
    const r = vals[i];
    if (!r[0]) continue;

    let obj = {};
    try { obj = JSON.parse(r[14] || '{}'); } catch (e) {}

    const updated = r[13] instanceof Date ? r[13].toISOString() : String(r[13] || '');
    rows.push({
      schoolId: String(r[0]),
      schoolName: String(r[1] || ''),
      division: String(r[3] || ''),
      municipality: String(r[4] || ''),
      status: 'Submitted',
      lastUpdated: updated
    });

    const groups = [
      obj.profile, obj.part3, obj.part4, obj.part5, obj.part6,
      obj.part7, obj.part8, obj.part9, obj.part10
    ];
    groups.forEach((group, index) => {
      if (group && (Array.isArray(group) ? group.length : Object.keys(group).length)) {
        partCounts[index].count++;
      }
    });
  }

  return {
    ok: true,
    total: 2706,
    submitted: rows.length,
    pending: Math.max(0, 2706 - rows.length),
    rows: rows,
    parts: partCounts
  };
}

/** Used only to show or hide administrator controls in the page. */
function getAdminSession() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  return { email: email, isAdmin: ADMIN_EMAILS.includes(email) };
}

/** Full response details are returned only after server-side admin authorization. */
function adminGetResponse(schoolId) {
  currentAdmin_();
  const sh = getOrCreateSheet_();
  const vals = sh.getDataRange().getValues();

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(schoolId).trim()) {
      let record = {};
      try { record = JSON.parse(vals[i][14] || '{}'); } catch (e) {}
      return { ok: true, record: record };
    }
  }
  throw new Error('Response not found.');
}

/** Permanently deletes a response; the action is retained in the audit sheet. */
function adminDeleteResponse(schoolId) {
  const email = currentAdmin_();
  const sh = getOrCreateSheet_();
  const vals = sh.getDataRange().getValues();
  let row = -1;

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(schoolId).trim()) {
      row = i + 1;
      break;
    }
  }
  if (row < 0) throw new Error('School response not found.');

  sh.deleteRow(row);
  logAdmin_(email, 'DELETE', schoolId, 'Response permanently deleted');
  return { ok: true, schoolId: String(schoolId), admin: email };
}

/** Detailed export is restricted to authorized administrators. */
function adminExportRows() {
  currentAdmin_();
  return getOrCreateSheet_().getDataRange().getDisplayValues();
}

/** Optional diagnostic: run this from the Apps Script editor after setting the new ID. */
function testSpreadsheetConnection() {
  const ss = getDb_();
  Logger.log('Connected to: ' + ss.getName());
  Logger.log('Spreadsheet URL: ' + ss.getUrl());
  const sh = getOrCreateSheet_();
  Logger.log('Response sheet: ' + sh.getName());
  Logger.log('Last row: ' + sh.getLastRow());
}
