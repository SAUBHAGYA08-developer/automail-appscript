/**
 * ============================================================================
 * MAIN RECRUITER OUTREACH EXECUTION ENGINE
 * ============================================================================
 */

/**
 * Main Scheduler Engine - Inko triggers ke saath set kiya ja sakta hai.
 */
function processRecruiters() {
  if (!isWithinAllowedTimeWindow()) {
    Logger.log("Skipping execution: Outside of Mon-Fri, 8:00 AM - 8:00 PM window.");
    return;
  }

  // Utils file mein defined global variables (SHEETS, COL, STATUS) auto-resolve honge
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

    if (status !== STATUS.PENDING) {
      continue;
    }

    try {
      sheet.getRange(actualRow, COL.STATUS).setValue(STATUS.PROCESSING);

      // Note: sendRecruiterEmail function definition niche di gayi hai
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
 * Checks if the current time is between Monday and Friday, 8:00 AM and 8:00 PM.
 * @return {boolean}
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
 * Mock Email Dispatcher Stub. Yahaan apna core GmailApp logic connect karein.
 */
function sendRecruiterEmail(row, actualRow) {
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

function addNewRecruiter(formData) {
  try {
    const sheet = getSheet(SHEETS.RECRUITERS);
    const totalColumns = sheet.getLastColumn() || 10; 
    const newRow = new Array(totalColumns).fill("");
    
    newRow[0] = new Date(); // Column A: Timestamp
    newRow[COL.RECRUITER - 1] = formData.recruiter;
    newRow[COL.EMAIL - 1] = formData.email;
    newRow[COL.COMPANY - 1] = formData.company;
    newRow[COL.POSITION - 1] = formData.position;
    newRow[COL.STATUS - 1] = STATUS.PENDING; 

    sheet.appendRow(newRow);
    logUserActivity("Lead Creation", "Manually queued recruiter tracking data for: " + formData.company);
    return "Successfully added " + formData.company + " to your sheet queue!";
  } catch(e) {
    throw new Error("Failed to insert row: " + e.toString());
  }
}

function getRecentRecruiters() {
  try {
    const sheet = getSheet(SHEETS.RECRUITERS);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    return data.map(row => {
      return {
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

function getTemplatesFromSheet() {
  try {
    // Check template key name dynamically, safely falls back to "Templates" sheet
    var templateSheetName = (SHEETS && SHEETS.TEMPLATES) ? SHEETS.TEMPLATES : "Templates";
    const sheet = getSheet(templateSheetName);
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
 * UI Payload ko process karke Settings tab me updates single-write transaction se push karta hai.
 */
function saveUserSettings(settingsPayload) {
  try {
    const sheet = getSheet(SHEETS.SETTINGS);
    var data = sheet.getDataRange().getValues();
    var updatedKeys = [];
    
    // Quick index reference map generation
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

function addNewTemplateRow(payload) {
  try {
    var templateSheetName = (SHEETS && SHEETS.TEMPLATES) ? SHEETS.TEMPLATES : "Templates";
    const sheet = getSheet(templateSheetName);
    
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
function logUserActivity(action, details) {
  try {
    const sheet = getSheet(SHEETS.ACTIVITY_LOG);
    sheet.appendRow([new Date(), action, details]);
  } catch (e) {
    Logger.log("Failed to log activity: " + e.toString());
  }
}

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
