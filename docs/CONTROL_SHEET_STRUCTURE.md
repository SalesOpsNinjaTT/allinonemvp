# Control Sheet Structure

The Control Sheet is the central hub for managing all salesperson dashboards.

## 📋 Control Sheet Tabs

### Tab 1: 👥 Salespeople Config

This is the master list of all salespeople. Add/remove people here.

| Column A | Column B              | Column C      | Column D                              |
|----------|-----------------------|---------------|---------------------------------------|
| Name     | Email                 | Sheet ID      | Sheet URL                             |
| John Doe | john.doe@company.com  | (auto-filled) | (auto-filled)                         |
| Jane Smith | jane.smith@company.com | (auto-filled) | (auto-filled)                        |

**How it works:**
- **Add new person**: Add name + email, leave Sheet ID blank
- **Script runs**: Auto-creates individual sheet, fills in Sheet ID and URL
- **Remove person**: Delete row (their sheet remains but won't update)
- **Self-provisioning**: New sheets created automatically on first run

---

### Tab 2: 🔧 Tech Access

Automation account emails for sharing and permissions.

| Column A | Column B                          |
|----------|-----------------------------------|
| Purpose  | Email                             |
| Script   | automation@company.com            |
| Backup   | backup-automation@company.com     |

**Purpose:**
- Lists all automation accounts that need access
- Used for auto-sharing individual sheets
- Ensures scripts can access created sheets

---

### Tab 3: 🎯 Goals & Quotas

Monthly goals and targets per salesperson.

| Column A | Column B              | Column C      | Column D           |
|----------|-----------------------|---------------|--------------------|
| Email    | Monthly Goal          | Q1 Target     | YTD Target         |
| john.doe@company.com | 15 | 45 | 180 |
| jane.smith@company.com | 20 | 60 | 240 |

**Purpose:**
- Store enrollment goals and targets
- Used by Enrollment Tracker component
- Used by Summary Dashboard for aggregations

---

### Tab 4: 📊 Summary Dashboard (Optional)

Overall metrics across all salespeople.

**Sections:**
1. **Current Month Summary**
   - Total enrollments across all AEs
   - Average deal close rate
   - Top performers

2. **Pipeline Overview**
   - Total active deals
   - Deals by stage
   - Average deal value

3. **Call Quality Summary**
   - Average scores across all AEs
   - Lowest/highest performers by metric

4. **Goals Progress**
   - % to goal by AE
   - On track / Behind / Ahead status

**Purpose:**
- Management reporting
- Quick overview of entire team
- Identify trends and outliers

---

## 🔄 How Self-Provisioning Works

### Initial State
```
Salespeople Config:
| Name     | Email                 | Sheet ID | Sheet URL |
|----------|-----------------------|----------|-----------|
| John Doe | john.doe@company.com  |          |           |
```

### After First Run
```
Salespeople Config:
| Name     | Email                 | Sheet ID           | Sheet URL                    |
|----------|-----------------------|--------------------|------------------------------|
| John Doe | john.doe@company.com  | 1aB2cD3eF4gH5iJ... | https://docs.google.com/...  |
```

**What happened:**
1. Script reads Salespeople Config
2. Finds John Doe has no Sheet ID
3. Creates new Google Sheet: "John Doe - Dashboard"
4. Adds 4 tabs: Pipeline Review, Bonus Calculation, Enrollment Tracker, Operational Metrics
5. Shares with John Doe's email
6. Writes Sheet ID and URL back to Control Sheet
7. Next run: Uses existing sheet, doesn't create duplicate

---

## 📂 Individual Sheet Structure

Each salesperson gets their own Google Sheet with 4 tabs:

### Sheet: "[Name] - Dashboard"

**Tab 1: 📊 Pipeline Review**
- Deal list with call quality scores
- Note1, Note2 columns for manual notes
- Hyperlinked to HubSpot
- Auto-refreshed from HubSpot

**Tab 2: 💰 Bonus Calculation**
- Commission dashboard
- Monthly bonus breakdown
- Business month override
- Goal tracking

**Tab 3: 📚 Enrollment Tracker**
- Month-to-month enrollment tracking
- Achievement metrics
- Goal progress
- Historical comparisons

**Tab 4: 📞 Operational Metrics**
- Call metrics
- Schedule adherence
- Performance analytics
- Trend analysis

---

## 🔐 Permissions & Sharing

### Control Sheet
- **Owner**: Your account (konstantin.gevorkov@tripleten.com)
- **Editors**: Automation accounts (from Tech Access tab)
- **Viewers**: Management/leadership (optional)

### Individual Salesperson Sheets
- **Owner**: Your account (or automation account)
- **Editor**: The salesperson (their email from Salespeople Config)
- **Editors**: Automation accounts (from Tech Access tab)
- **Viewers**: Their manager (optional)

---

## 🎯 Main Function Structure

```javascript
function generateAllDashboards() {
  // 1. Read Control Sheet config
  const salespeople = readSalespeopleConfig();
  const goals = readGoals();
  
  // 2. For each salesperson
  salespeople.forEach(person => {
    // 3. Get or create their individual sheet
    const sheet = getOrCreatePersonSheet(person);
    
    // 4. Update each tab
    updatePipelineReview(sheet, person);
    updateBonusCalculation(sheet, person);
    updateEnrollmentTracker(sheet, person, goals);
    updateOperationalMetrics(sheet, person);
  });
  
  // 5. Update summary dashboard
  updateSummaryDashboard(salespeople);
}
```

---

## 🚀 Adding a New Salesperson

**Steps:**
1. Open Control Sheet
2. Go to "👥 Salespeople Config" tab
3. Add new row:
   - Name: "New Person"
   - Email: "newperson@company.com"
   - Leave Sheet ID blank
   - Leave Sheet URL blank
4. Go to "🎯 Goals & Quotas" tab
5. Add their goals
6. Run script (or wait for scheduled run)
7. ✅ Done! Their sheet is created and they receive access

---

## 🗑️ Removing a Salesperson

**Steps:**
1. Open Control Sheet
2. Go to "👥 Salespeople Config" tab
3. Delete their row (or move to "Inactive" section)
4. Their individual sheet remains but won't update
5. Optionally: Manually delete their individual sheet
6. Optionally: Remove from Goals & Quotas tab

---

## 📊 Reporting Opportunities

### From Control Sheet
- **Summary Dashboard**: Real-time view of all salespeople
- **Export to BigQuery**: For advanced analytics (future)
- **Scheduled Reports**: Email summary to management
- **Trend Analysis**: Track performance over time

### From Individual Sheets
- Each salesperson can export their own data
- Managers can view their team's sheets
- Leadership can view summary only

---

## 🔧 Configuration Management

### Script Properties (Set in Project Settings)
- `HUBSPOT_ACCESS_TOKEN` - HubSpot API key
- `CONTROL_SHEET_ID` - Control Sheet ID (auto-detected)

### In Control Sheet Tabs
- `Salespeople Config` - Who to process
- `Tech Access` - Automation accounts
- `Goals & Quotas` - Targets and goals

---

**Last Updated**: November 19, 2025

