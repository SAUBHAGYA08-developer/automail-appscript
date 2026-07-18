function saveDashboardHistory() {

  const dashboard = getSheet(SHEETS.DASHBOARD);
  const history = getSheet("Dashboard_History");

  // Create header if missing
  const headers = [
    "Date",
    "Today's Emails",
    "Total Recruiters",
    "Sent",
    "Failed",
    "Replies",
    "Total Emails Sent"
  ];

  if (
    history.getLastRow() === 0 ||
    history.getRange(1, 1).getValue() !== "Date"
  ) {
    history.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const dashboardData = dashboard
    .getRange(1, 1, dashboard.getLastRow(), 2)
    .getValues();

  const metrics = {};

  dashboardData.forEach(row => {
    metrics[row[0]] = row[1];
  });

  history.appendRow([
    new Date(),
    metrics["Today's Emails"] || 0,
    metrics["Total Recruiters"] || 0,
    metrics["Sent"] || 0,
    metrics["Failed"] || 0,
    metrics["Replies Received"] || 0,
    metrics["Total Emails Sent"] || 0
  ]);

}