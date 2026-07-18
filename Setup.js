const ss = SpreadsheetApp.getActiveSpreadsheet();
// Inside your Setup.gs file:
let templatesSheet = ss.getSheetByName(SHEETS.TEMPLATES);
if (!templatesSheet) {
  templatesSheet = ss.insertSheet(SHEETS.TEMPLATES);
  templatesSheet.appendRow(["Template Name", "Email Subject", "Email Body"]);
  templatesSheet.getRange("A1:C1").setFontWeight("bold");
  
  // SEED THE CORE LOGIC TEMPLATES IMMEDIATELY
  templatesSheet.appendRow([
    "INITIAL_DEFAULT",
    "Introduction / Open Opportunities",
    "Hi {{greeting}},\n\nI wanted to reach out regarding open tracks at your organization..."
  ]);
  templatesSheet.appendRow([
    "INITIAL_COMPANY",
    "Excited about roles at {{company}}",
    "Hi {{greeting}},\n\nI've been following the recent work done by the teams at {{company}}..."
  ]);
  templatesSheet.appendRow([
    "INITIAL_POSITION",
    "Application for {{position}} role",
    "Hi {{greeting}},\n\nI am writing to express my interest in the {{position}} opening..."
  ]);
  templatesSheet.appendRow([
    "FOLLOWUP1",
    "Re: Quick Follow-up",
    "Hi {{greeting}},\n\nJust dropping a quick note to bump this in your inbox..."
  ]);
}

// Inside your Setup.gs file -> ensureSpreadsheetInfrastructure() function:
let settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
if (!settingsSheet) {
  settingsSheet = ss.insertSheet(SHEETS.SETTINGS);
  settingsSheet.appendRow(["Parameter Name", "Value"]);
  settingsSheet.getRange("A1:B1").setFontWeight("bold");
  
  // Base configuration options
  settingsSheet.appendRow(["Resume URL", "https://link-to-your-resume.com"]);
  settingsSheet.appendRow(["Sender Name", ""]);
  settingsSheet.appendRow(["Signature", "Best regards,\nYour Name"]);
  
  // SEED THE TIMING CONFIGURATIONS
  settingsSheet.appendRow(["Followup 1 Days", "3"]); // Sent 3 days after initial email
  settingsSheet.appendRow(["Followup 2 Days", "7"]); // Sent 7 days after initial email (4 days after Followup 1)
  settingsSheet.appendRow(["Followup 3 Days", "14"]); // Sent 14 days after initial email
}

// Setup.gs mein ye add kar lo:
let logSheet = ss.getSheetByName(SHEETS.ACTIVITY_LOG);
if (!logSheet) {
  logSheet = ss.insertSheet(SHEETS.ACTIVITY_LOG);
  logSheet.appendRow(["Timestamp", "Action", "Details"]);
  logSheet.getRange("A1:C1").setFontWeight("bold");
}