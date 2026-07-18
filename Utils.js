/**
 * Sheet Names
 */
const SHEETS = {
  RECRUITERS: "Recruiters",
  TEMPLATES: "Templates",
  SETTINGS: "Settings",
  DASHBOARD: "Dashboard",
  DASHBOARDHISTORY : "Dashboard_History",
  LOGS: "Logs",
  ACTIVITY_LOG: "ActivityLog"
};

/**
 * Recruiters Sheet Columns
 */
const COL = {
  RECRUITER: 1,
  EMAIL: 2,
  COMPANY: 3,
  POSITION: 4,
  SUBJECT_OVERRIDE: 5,
  STATUS: 6,
  ATTEMPT: 7,
  SENT_DATE: 8,
  FOLLOWUP_DATE: 9,
  REPLY: 10,
  RETRY_COUNT: 11,
  ERROR: 12,
  THREAD_ID: 13
};

/**
 * Status Constants
 */
const STATUS = {
  PENDING: "PENDING",
  PROCESSING: "PROCESSING",
  SENT: "SENT",
  FAILED: "FAILED",
  COMPLETED: "COMPLETED"
};

/**
 * Returns Sheet
 */
function getSheet(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

/**
 * Reads all settings from Settings sheet
 */
function getSettings() {

  const sheet = getSheet(SHEETS.SETTINGS);

  const values = sheet.getDataRange().getValues();

  let settings = {};

  for (let i = 1; i < values.length; i++) {

    const key = values[i][0];
    const value = values[i][1];

    settings[key] = value;

  }

  return settings;

}

function getGreeting(name) {

  if (!name || name.trim() === "") {
    return "Hello,";
  }

  return `Hi ${name},`;

}
/**
 * Current Date
 */
function now() {
  return new Date();
}

function logActivity(email, action, status, error) {

  const sheet = getSheet(SHEETS.LOGS);

  sheet.appendRow([
    new Date(),
    email,
    action,
    status,
    error
  ]);

}

  function getDriveFileId(url) {
  const match = url.match(/[-\w]{25,}/);
  if (!match) {
    throw new Error("Invalid Google Drive URL");
  }
  return match[0];
}