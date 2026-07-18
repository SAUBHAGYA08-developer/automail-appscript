function processFollowups() {
  console.log("========== PROCESS FOLLOWUPS STARTED ==========");

  const sheet = getSheet(SHEETS.RECRUITERS);
  const settings = getSettings();
  
  console.log(JSON.stringify(settings));

  // 1. Dynamic Staging Lookup: Map specific step intervals with fallback configurations
  const timingConfig = {
    1: Number(settings["Followup 1 Days"] ?? settings["Follow-up Days"] ?? 3),
    2: Number(settings["Followup 2 Days"] ?? settings["Follow-up Days"] ?? 7),
    3: Number(settings["Followup 3 Days"] ?? settings["Follow-up Days"] ?? 14)
  };
  
  const maxFollowups = Number(settings["Max Follow-ups"] ?? 3);

  console.log("Dynamic Sequence Timing Map: ", JSON.stringify(timingConfig));
  console.log("Max Follow-ups Threshold: " + maxFollowups);

  const lastRow = sheet.getLastRow();
  console.log("Last Row Index: " + lastRow);

  if (lastRow <= 1) {
    console.log("No recruiter data found.");
    return;
  }

  const data = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();

  // Normalize current timestamp date boundaries to ensure clean, consistent comparisons
  const today = new Date();
  const normalizedToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const actualRow = i + 2;

    console.log("--------------------------------");
    console.log("Checking Row : " + actualRow);

    const status = row[COL.STATUS - 1];
    const reply = row[COL.REPLY - 1];
    const attempt = Number(row[COL.ATTEMPT - 1] || 0);
    const sentDateVal = row[COL.SENT_DATE - 1];
    const threadId = row[COL.THREAD_ID - 1];

    if (status !== STATUS.SENT) {
      console.log("Skipped -> Status is not SENT");
      continue;
    }

    if (reply === "YES") {
      console.log("Skipped -> Recruiter already replied");
      continue;
    }

    if (attempt >= maxFollowups) {
      console.log("Skipped -> Max followups already reached or exceeded");
      continue;
    }

    if (!sentDateVal) {
      console.log("Skipped -> Sent Date structural marker is missing");
      continue;
    }

    const nextAttempt = attempt + 1;
    
    // 2. Fetch the dynamic day delay interval required for THIS specific upcoming sequence phase
    const requiredDaysDelay = timingConfig[nextAttempt] ?? 5; 

    // Calculate days elapsed relative to the base initial Sent Date
    const sentDate = new Date(sentDateVal);
    const normalizedSentDate = new Date(sentDate.getFullYear(), sentDate.getMonth(), sentDate.getDate());
    
    const nextFollowupDate = new Date(normalizedSentDate);
    nextFollowupDate.setDate(nextFollowupDate.getDate() + requiredDaysDelay);

    console.log("Attempt Step Context : Phase " + nextAttempt);
    console.log("Required Delay Target: " + requiredDaysDelay + " days from Initial Sent Date");
    console.log("Target Execution Date: " + nextFollowupDate);

    if (normalizedToday < nextFollowupDate) {
      console.log("Skipped -> Follow-up phase not due yet");
      continue;
    }

    try {
      console.log("Executing Send Action: Deploying Follow-up " + nextAttempt + "...");

      // Triggers thread bouncer via raw Gmail message injection
      sendFollowupEmail(
        threadId,
        nextAttempt,
        row
      );

      console.log("Follow-up sent successfully.");

      // Update structural grid metrics
      sheet.getRange(actualRow, COL.ATTEMPT).setValue(nextAttempt);
      sheet.getRange(actualRow, COL.FOLLOWUP_DATE).setValue(today);

      // If this was the final configured message block, close out the target lifecycle state
      if (nextAttempt === maxFollowups) {
        sheet.getRange(actualRow, COL.STATUS).setValue("COMPLETED");
        console.log("Pipeline Row Status Promoted: COMPLETED");
      }

      logActivity(
        row[COL.EMAIL - 1],
        "FOLLOWUP_" + nextAttempt,
        "SUCCESS",
        "Dispatched successfully at day mark index."
      );

    } catch (e) {
      console.log("ERROR EXECUTING ATTEMPT STEP: " + e.message);

      sheet.getRange(actualRow, COL.ERROR).setValue(e.message);

      logActivity(
        row[COL.EMAIL - 1],
        "FOLLOWUP_" + nextAttempt,
        "FAILED",
        e.message
      );
    }
  }

  console.log("========== PROCESS FOLLOWUPS COMPLETED ==========");
}