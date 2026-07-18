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
