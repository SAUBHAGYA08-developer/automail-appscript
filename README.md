# Recruiter Outreach Engine Architecture

Welcome to the 6-Sheet Recruiter Outreach Platform. This engine leverages Google Apps Script (GAS) to sync dynamic browser dashboards directly with a secure Google Spreadsheet data pipeline.

## 🗂️ 6-Sheet Architecture Ecosystem
The core execution relies on six system-managed tables mapped in your project configurations:
1. **Recruiters**: The main transaction queue where candidates/recruiter leads are tracked.
2. **Dashboard**: The direct UI interface landing for high-level state analytics.
3. **Logs**: A delivery diagnostic panel to audit transactional mail exceptions.
4. **Settings**: The system parameters vault holding metadata variables and custom scheduling constraints.
5. **Dashboard_History**: Archive structures managing periodic trend snapshots.
6. **ActivityLog**: The security audit trail handling actions performed by dashboard managers.

## 🔄 Dynamic Lifecycle & System Data Flow

1. **Infrastructure Provisioning (`Setup.gs`)**
   Every time the deployment script wakes up, `ensureSpreadsheetInfrastructure()` acts as a structural gateway. If a user has manually corrupted or deleted database sheets (like `Settings` or `Templates`), the system immediately self-heals by provisioning fresh tables and seeding fallback parameters (`Followup 1 Days`, baseline templates, structural indices).

2. **UI Hydration & Core Mapping (`Code.gs`)**
   When a user loads the Web UI form, backend engines execute `fetchUserSettings()` and `getTemplatesFromSheet()`. Instead of hardcoding keys inside the HTML interface, client code parses live spreadsheet data vectors directly, generating matching dynamic layouts on the fly.

3. **Outreach & Scheduling Engine (`Code.gs`)**
   The time-gated scheduler `processRecruiters()` evaluates pending target batches. It intercepts data arrays, verifies business operational constraints (strictly Mon-Fri, 8 AM - 8 PM), dispatches outbound templates via transactional threads, and stamps performance markers seamlessly back to the source rows.

4. **Transactional Updates & Logs Loop (`Code.gs`)**
   Saving form configuration values transforms frontend UI object inputs into safe transaction payloads. The controller invokes `saveUserSettings(payload)`, verifies current properties against delta changes to mitigate performance bottleneck overrides, and explicitly commits states to the `ActivityLog` trail.

To make your repository "plug-and-play" for others, your `README.md` needs to clearly explain the **Apps Script binding** process. Since Google Apps Script lives *inside* a specific Google Sheet, users cannot simply clone the repo and run it; they must "link" the code to their own spreadsheet.

Here is a structured template you can copy and paste into your `README.md`:

---

# Automail Outreach System

This is an automated recruiter outreach platform built on Google Apps Script and Google Sheets.

## 🚀 Setup Instructions

### 1. Prepare your Google Sheet

1. Create a new [Google Sheet](https://sheets.new).
2. Create the following four tabs exactly as named:
* `Recruiters` (Columns: `Name`, `Email`, `Company`, `Position`, `Status`, `Date`)
* `Templates` (Columns: `Name`, `Subject`, `Body`)
* `Settings` (Columns: `Key`, `Value`)
* `ACTIVITY_LOG` (Columns: `Timestamp`, `Action`, `Email`, `Status`)



### 2. Connect the Code

1. In your Google Sheet, go to **Extensions > Apps Script**.
2. Delete any existing code (`function myFunction...`).
3. Copy the contents of `Code.gs`, `Setup.gs`, and `Form.html` from this repository into your project editor.
* *Note: In the editor, use the `+` icon to create a new "HTML" file for `Form.html`.*


4. Save the project (Ctrl + S).

### 3. Initialize the System

1. In the Apps Script editor, select the `Setup.gs` file.
2. Select the `bootstrap` function from the dropdown menu and click **Run**.
3. Grant the required Google permissions when prompted (if you see "Unsafe," click **Advanced > Go to [Project Name] (unsafe)**).

### 4. Configure Triggers

1. On the left sidebar, click the **Triggers (alarm clock icon)**.
2. Click **+ Add Trigger**.
3. Select `runOutreachLoop` as the function to run.
4. Set event source to **Time-driven**.
5. Select **Hour timer** (e.g., every 1 hour).
6. Click **Save**.

### 5. Deploy as Web App

1. Click the **Deploy** button (top right) > **New Deployment**.
2. Select **Web App**.
3. Set "Execute as" to **Me** and "Who has access" to **Anyone** (or "Only Myself" for private use).
4. Click **Deploy** and copy the **Web App URL**.
5. Open this URL in your browser to access your Outreach Dashboard!

---

## 🛠 Troubleshooting

* **Permissions:** Always run the `bootstrap` function first to ensure the script has access to your Sheets.
* **Dashboard Errors:** Ensure the names of your Sheet tabs exactly match the code requirements.
* **GitHub Integration:** If you are using the GitHub Assistant extension, ensure you have generated a **Classic Personal Access Token** with `repo` scope permissions.

---

### A few tips for your README:

* **Note on API:** Remind them that if they use `MailApp` or `GmailApp`, the script will automatically ask for email-sending permissions during the first execution.
* **Keep it private:** Remind them not to share their sheet URL publicly, as it contains sensitive recruiter contact information.
