# Control Sheet Structure

Central hub for managing all salesperson dashboards.

## Tabs

### 👥 Salespeople Config
| Name | Email | Sheet ID | Sheet URL |
|------|-------|----------|-----------|
| John Doe | john@company.com | (auto) | (auto) |

**Add new person**: Add name + email, script auto-creates their sheet

### 🔧 Tech Access
| Purpose | Email |
|---------|-------|
| Script | automation@company.com |

**Purpose**: Automation accounts for sharing

### 🎯 Goals & Quotas
| Email | Monthly Goal |
|-------|--------------|
| john@company.com | 15 |

**Purpose**: Enrollment targets for Enrollment Tracker component

### 📊 Summary Dashboard
- Total enrollments across all AEs
- Pipeline overview (total deals, by stage)
- Call quality averages
- Goals progress tracking

---

## Individual Sheets (Auto-Created)

**"[Name] - Dashboard"** with 4 tabs:
1. 📊 Pipeline Review - Deals + call quality + notes
2. 💰 Bonus Calculation - Commission dashboard
3. 📚 Enrollment Tracker - Monthly enrollments
4. 📞 Operational Metrics - Call metrics

---

## Self-Provisioning

1. Add name + email to Salespeople Config
2. Leave Sheet ID blank
3. Run script
4. Sheet auto-creates, ID auto-fills

---

## Main Function

```javascript
function generateAllDashboards() {
  const salespeople = readSalespeopleConfig();
  salespeople.forEach(person => {
    const sheet = getOrCreatePersonSheet(person); // Self-provision
    updatePipelineReview(sheet, person);
    updateBonusCalculation(sheet, person);
    updateEnrollmentTracker(sheet, person);
    updateOperationalMetrics(sheet, person);
  });
  updateSummaryDashboard(salespeople);
}
```
