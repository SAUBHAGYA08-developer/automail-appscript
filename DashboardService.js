function updateDashboard() {

  const dashboard = getSheet(SHEETS.DASHBOARD);
  const recruiters = getSheet(SHEETS.RECRUITERS);

  const settings = getSettings();

  const lastRow = recruiters.getLastRow();

  if (lastRow <= 1) {
    dashboard.getRange("A1:B12").clearContent();
    return;
  }

  const data = recruiters
    .getRange(2, 1, lastRow - 1, recruiters.getLastColumn())
    .getValues();


  let totalRecruiters = 0;
  let pending = 0;
  let sent = 0;
  let failed = 0;
  let replies = 0;

  let followup1 = 0;
  let followup2 = 0;
  let followup3 = 0;

  let todaysEmails = 0;
  let totalEmailsSent = 0;


  const today = new Date();


  data.forEach(row => {

    totalRecruiters++;

    const status = row[COL.STATUS - 1];
    const sentDate = row[COL.SENT_DATE - 1];
    const attempt = Number(row[COL.ATTEMPT - 1] || 0);
    const reply = row[COL.REPLY - 1];


    // Status count

    if (status === STATUS.PENDING) {
      pending++;
    }

    if (status === STATUS.SENT) {
      sent++;
    }

    if (status === STATUS.FAILED) {
      failed++;
    }


    // Reply count

    if (reply === "YES") {
      replies++;
    }


    // Email attempts

    if (attempt > 0) {
      totalEmailsSent++;
    }


    if (attempt === 2) {
      followup1++;
    }

    if (attempt === 3) {
      followup2++;
    }

    if (attempt >= 4) {
      followup3++;
    }


    // Today's email count

    if (sentDate instanceof Date) {

      if (
        sentDate.getDate() === today.getDate() &&
        sentDate.getMonth() === today.getMonth() &&
        sentDate.getFullYear() === today.getFullYear()
      ) {
        todaysEmails++;
      }

    }


  });


  const dailyLimit = Number(
    settings["Daily Email Limit"] || 40
  );


  const remainingToday = Math.max(
    0,
    dailyLimit - todaysEmails
  );


  const dashboardData = [

    ["Metric", "Value"],

    ["Today's Emails", todaysEmails],

    ["Daily Limit", dailyLimit],

    ["Remaining Today", remainingToday],

    ["Total Recruiters", totalRecruiters],

    ["Pending", pending],

    ["Sent", sent],

    ["Failed", failed],

    ["Replies Received", replies],

    ["Follow-up 1 Sent", followup1],

    ["Follow-up 2 Sent", followup2],

    ["Follow-up 3 Sent", followup3],

    ["Total Emails Sent", totalEmailsSent]

  ];


  dashboard
    .getRange(
      1,
      1,
      dashboardData.length,
      2
    )
    .setValues(dashboardData);


}