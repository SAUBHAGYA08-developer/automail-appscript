/**
 * ============================================================================
 * GLOBAL CONSTANTS & CONFIGURATIONS
 * ============================================================================
 */

/**
 * Sheet Instance Helper Function
 */
function getSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ensureSpreadsheetInfrastructure();
    sheet = ss.getSheetByName(sheetName);
  }
  return sheet;
}

/**
 * ============================================================================
 * MAIN RECRUITER OUTREACH EXECUTION ENGINE
 * ============================================================================
 */

/**
 * Main Scheduler Engine - Triggers ke saath schedule karne ke liye.
 */
function processRecruiters() {
  if (!isWithinAllowedTimeWindow()) {
    Logger.log("Skipping execution: Outside of Mon-Fri, 8:00 AM - 8:00 PM window.");
    return;
  }

  const sheet = getSheet(SHEETS.RECRUITERS);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log("No recruiters found.");
    return;
  }

  const data = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const actualRow = i + 2;
    let status = row[COL.STATUS - 1];

    if (!status || status.toString().trim() === "") {
      status = STATUS.PENDING;
      sheet.getRange(actualRow, COL.STATUS).setValue(STATUS.PENDING);
    }

   // Sirf PENDING aur DETAILS_REQUESTED dono status wale emails process honge
  if (status !== STATUS.PENDING && status !== STATUS.DETAILS_REQUESTED) {
    continue;
  }

    try {
      sheet.getRange(actualRow, COL.STATUS).setValue(STATUS.PROCESSING);

      const threadId = sendRecruiterEmail(row, actualRow);

      sheet.getRange(actualRow, COL.THREAD_ID).setValue(threadId);
      sheet.getRange(actualRow, COL.STATUS).setValue(STATUS.SENT);
      sheet.getRange(actualRow, COL.ATTEMPT).setValue(1);
      sheet.getRange(actualRow, COL.SENT_DATE).setValue(new Date());

      logUserActivity(
        "INITIAL_EMAIL_SUCCESS", 
        "Sent outreach directly to: " + row[COL.EMAIL - 1]
      );

    } catch (e) {
      sheet.getRange(actualRow, COL.STATUS).setValue(STATUS.FAILED);
      sheet.getRange(actualRow, COL.ERROR).setValue(e.message);

      logUserActivity(
        "INITIAL_EMAIL_FAILED", 
        "Failure targeting " + row[COL.EMAIL - 1] + " Error: " + e.message
      );
    }
  }
}

/**
 * Business hours check (Monday - Friday, 8:00 AM - 8:00 PM)
 */
function isWithinAllowedTimeWindow() {
  const now = new Date();
  const day = now.getDay();
  const hours = now.getHours();
  
  const isWeekday = (day >= 1 && day <= 5);
  const isBusinessHours = (hours >= 8 && hours < 20);
  
  return isWeekday && isBusinessHours;
}

/**
 * Dynamic Template Selection & Email Dispatcher
 */
function sendRecruiterEmail(row, actualRow) {
  const email = row[COL.EMAIL - 1];
  const recruiter = row[COL.RECRUITER - 1];
  const company = row[COL.COMPANY - 1];
  const position = row[COL.POSITION - 1];
  const subjectOverride = row[COL.SUBJECT_OVERRIDE - 1];

  // Templates fetch karke matching override template dhoondta hai
  const templates = getTemplatesFromSheet();
  let selectedTemplate = null;

  if (subjectOverride && subjectOverride.trim() !== "") {
    selectedTemplate = templates.find(t => t.name.toLowerCase() === subjectOverride.trim().toLowerCase());
  }

  // Baseline Fallback Template agar specific match na mile
  if (!selectedTemplate && templates.length > 0) {
    selectedTemplate = templates[0];
  }

  let subject = selectedTemplate ? selectedTemplate.subject : "Application for " + position + " - " + recruiter;
  let body = selectedTemplate ? selectedTemplate.body : "Hi " + recruiter + ",\n\nI am interested in the " + position + " role at " + company + ".";

  // Dynamic Placeholders Replacement
  subject = subject.replace(/{{Recruiter}}/g, recruiter)
                   .replace(/{{Company}}/g, company)
                   .replace(/{{Position}}/g, position);

  body = body.replace(/{{Recruiter}}/g, recruiter)
             .replace(/{{Company}}/g, company)
             .replace(/{{Position}}/g, position);

  // Draft/Send execution
  // GmailApp.sendEmail(email, subject, body);
  return "msg_id_" + Math.random().toString(36).substr(2, 9);
}

/**
 * ============================================================================
 * WEB UI WEB APP INGESTION ENDPOINTS
 * ============================================================================
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Form')
      .setTitle('Recruiter Lead Form')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Quick Recruiter Add Form Submission Endpoint
 */
function addNewRecruiter(formData) {
  try {
    const sheet = getSheet(SHEETS.RECRUITERS);
    const totalColumns = sheet.getLastColumn() || 14; 
    const newRow = new Array(totalColumns).fill("");
    
    const email = (formData.email || "").trim().toLowerCase();
    if (!email) {
      throw new Error("Email is required.");
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    if (!emailRegex.test(email)) {
      throw new Error("Please enter a valid email address.");
    }

    newRow[0] = new Date(); // Timestamp
    newRow[COL.RECRUITER - 1] = (formData.recruiter || "").trim();
    newRow[COL.EMAIL - 1] = email;
    newRow[COL.COMPANY - 1] = (formData.company || "").trim();
    newRow[COL.POSITION - 1] = (formData.position || "").trim();
    
    // Dropdown 2: Subject/Template Override Column Update
    newRow[COL.SUBJECT_OVERRIDE - 1] = (formData.subjectOverride || "").trim();

    // Dropdown 1: Initial Status Selection Update
    const selectedStatus = (formData.status || STATUS.PENDING).trim();
    newRow[COL.STATUS - 1] = selectedStatus;

    sheet.appendRow(newRow);
    logUserActivity("Lead Creation", "Manually queued recruiter tracking data for: " + formData.company);
    return "Successfully added " + formData.company + " to your sheet queue!";
  } catch(e) {
    throw new Error("Failed to insert row: " + e.toString());
  }
}

/**
 * Fetch Recent Activity Feed for Dashboard
 */
function getRecentRecruiters() {
  try {
    const sheet = getSheet(SHEETS.RECRUITERS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    return data.map((row, index) => {
      return {
        rowIndex: index + 2,
        email: row[COL.EMAIL - 1] || "",
        company: row[COL.COMPANY - 1] || "",
        position: row[COL.POSITION - 1] || "",
        status: row[COL.STATUS - 1] || "PENDING"
      };
    }).reverse().slice(0, 10);
  } catch(e) {
    return [];
  }
}

/**
 * Dropdown Status Change Handler
 */
function updateRecruiterStatus(rowIndex, newStatus) {
  try {
    const sheet = getSheet(SHEETS.RECRUITERS);
    const statusCol = COL.STATUS || 7;
    
    sheet.getRange(rowIndex, statusCol).setValue(newStatus);

    let email = "";
    try {
      email = sheet.getRange(rowIndex, COL.EMAIL).getValue();
    } catch(e) {}

    logUserActivity(
      "STATUS_UPDATE", 
      "Status changed to '" + newStatus + "' for Row " + rowIndex + (email ? " (" + email + ")" : "")
    );

    return true;
  } catch(e) {
    throw new Error("Failed to update status: " + e.toString());
  }
}

/**
 * Fetch Performance Metrics for Analytics Tab
 */
function fetchDashboardMetrics() {
  try {
    const sheet = getSheet(SHEETS.RECRUITERS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { sent: 0, pending: 0, failed: 0, total: 0 };
    
    const statuses = sheet.getRange(2, COL.STATUS, lastRow - 1, 1).getValues().flat();
    
    const metrics = { sent: 0, pending: 0, failed: 0, total: statuses.length };
    statuses.forEach(status => {
      const clean = String(status).trim().toUpperCase();
      if (clean === "SENT") metrics.sent++;
      else if (clean === "PENDING" || clean === "") metrics.pending++;
      else if (clean === "FAILED") metrics.failed++;
    });
    
    return metrics;
  } catch(e) {
    return { sent: 0, pending: 0, failed: 0, total: 0 };
  }
}

/**
 * Fetch Email Templates from 'Templates' Tab
 */
function getTemplatesFromSheet() {
  try {
    const sheet = getSheet(SHEETS.TEMPLATES);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 
    
    var templates = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[0]) {
        templates.push({
          name: row[0].toString(),
          subject: row[1] ? row[1].toString() : "No Subject Specified",
          body: row[2] ? row[2].toString() : ""
        });
      }
    }
    return templates;
  } catch (error) {
    return [];
  }
}

/**
 * Fetch System Configurations from 'Settings' Tab
 */
function fetchUserSettings() {
  try {
    const sheet = getSheet(SHEETS.SETTINGS);
    const data = sheet.getDataRange().getValues();
    var settingsList = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) { 
        settingsList.push({
          key: data[i][0].toString(),
          value: data[i][1] ? data[i][1].toString() : ""
        });
      }
    }
    return settingsList;
  } catch(e) {
    return [];
  }
}

/**
 * Save System Configurations to 'Settings' Tab
 */
function saveUserSettings(settingsPayload) {
  try {
    const sheet = getSheet(SHEETS.SETTINGS);
    var data = sheet.getDataRange().getValues();
    var updatedKeys = [];
    
    var keyRowMap = {};
    for (var i = 1; i < data.length; i++) {
      var existingKey = data[i][0];
      if (existingKey) {
        keyRowMap[existingKey.toString().trim()] = {
          row: i + 1,
          currentValue: data[i][1] ? data[i][1].toString() : ""
        };
      }
    }
    
    settingsPayload.forEach(function(item) {
      var configKey = item.key.toString().trim();
      var configValue = item.value.toString();
      
      if (keyRowMap[configKey]) {
        if (keyRowMap[configKey].currentValue !== configValue) {
          sheet.getRange(keyRowMap[configKey].row, 2).setValue(configValue);
          updatedKeys.push(configKey);
        }
      } else {
        sheet.appendRow([configKey, configValue]);
        updatedKeys.push(configKey);
      }
    });
    
    if (updatedKeys.length > 0) {
      logUserActivity(
        "Master Config Update", 
        "Modified configuration variables: " + updatedKeys.join(", ")
      );
    }
    
    return true;
  } catch(e) {
    Logger.log("Settings save execution crash: " + e.toString());
    return false;
  }
}

/**
 * Add New Template to 'Templates' Tab
 */
function addNewTemplateRow(payload) {
  try {
    const sheet = getSheet(SHEETS.TEMPLATES);
    
    sheet.appendRow([
      payload.name.toString(),
      payload.subject.toString(),
      payload.body.toString()
    ]);
    
    logUserActivity("Template Created", "Added new email layout titled: " + payload.name);
    return true;
  } catch (error) {
    return false;
  }
}

/**
 * Log System Activity
 */
function logUserActivity(action, details) {
  try {
    const sheet = getSheet(SHEETS.ACTIVITY_LOG);
    sheet.appendRow([new Date(), action, details]);
  } catch (e) {
    Logger.log("Failed to log activity: " + e.toString());
  }
}

/**
 * Auto-Create Missing Sheets and Header Columns
 */
function ensureSpreadsheetInfrastructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const setupData = [
    { name: SHEETS.RECRUITERS, headers: ["Timestamp", "Recruiter", "Email", "Company", "Position", "Subject Override", "Status", "Attempt", "Sent Date", "Follow-up Date", "Reply", "Retry Count", "Error", "Gmail Thread ID"] },
    { name: SHEETS.TEMPLATES, headers: ["Template Name", "Email Subject", "Email Body"] },
    { name: SHEETS.SETTINGS, headers: ["Parameter Name", "Value"] },
    { name: SHEETS.DASHBOARD, headers: ["Metric", "Value"] },
    { name: SHEETS.LOGS, headers: ["Timestamp", "Email", "Action", "Status", "Error"] },
    { name: SHEETS.ACTIVITY_LOG, headers: ["Timestamp", "Action", "Details"] }
  ];

  setupData.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight("bold");
    }
  });
}