# Enrollment Tracker - Sheet Structure

## Overview

The Enrollment Tracker displays enrollments recorded in HubSpot with rolling monthly history. It fetches the current month and last month from HubSpot, and preserves older historical data in the sheet.

## Data Fetching Strategy

**HubSpot Filters:**
- `closed_status` = "Closed won"
- `dealstage` = "90284262" (Partnership Confirmed)
- `closedate` >= first day of last month
- Owner ID matches the AE

**Refresh Logic:**
- ✅ **Current Month**: Always fetched and rewritten
- ✅ **Last Month**: Always fetched and rewritten (catches backdated deals)
- ✅ **Older Months**: Preserved from existing sheet (historical record)

This creates a rolling history where data naturally "scrolls down" each month.

## Sheet Layout

```
┌──────────────────────────────────────────┐
│  NOVEMBER 2025 - ENROLLMENTS (Total: 12) │
├──────────────────────────────────────────┤
│  Deal Name | Platform Email | Program... │
│  Student 1 | email@... | MBA...         │
│  Student 2 | email@... | DBA...         │
└──────────────────────────────────────────┘

┌──────────────────────────────────────────┐
│  NOVEMBER 2025 - EASY STARTS (Total: 3)  │
├──────────────────────────────────────────┤
│  Deal Name | Easy Start | Platform...    │
│  Student 3 | Started | email@...         │
└──────────────────────────────────────────┘

┌──────────────────────────────────────────┐
│  OCTOBER 2025 - ENROLLMENTS              │
├──────────────────────────────────────────┤
│  [Same structure as above]               │
└──────────────────────────────────────────┘

┌──────────────────────────────────────────┐
│  OCTOBER 2025 - EASY STARTS              │
├──────────────────────────────────────────┤
│  [Same structure as above]               │
└──────────────────────────────────────────┘

[Older months preserved below...]
```

## Field Configuration

### Regular Enrollments Columns

| Column | HubSpot Property | Type | Notes |
|--------|------------------|------|-------|
| Deal Name | `dealname` | Text | Hyperlinked to HubSpot |
| Platform Email | `platform_email` | Text | Student's platform email |
| Program | `program` | Text | Program name (MBA, DBA, etc.) |
| Cohort Start Date | `start_date` | Date | When cohort starts |
| Close Date | `closedate` | Date | Deal close date |
| Warm Handoff | `warm_handoff_scoring` | Number | 🟥 0-2 🟨 2-3 🟩 3-5 |
| Ask for Referral | `s_closing_the_deal__a_ask_for_referral` | Number | 🟥 0-2 🟨 2-3 🟩 3-5 |

### Easy Starts Columns

Same as Regular Enrollments, plus:
- **Easy Start Option** (after Deal Name): Shows the value ("Waiting for the start", "Started", "Didn't pay")

## Enrollment vs Easy Start Logic

**Easy Start Criteria:**
- `easy_start_option` IS one of:
  - "Waiting for the start"
  - "Started"
  - "Didn't pay"

**Regular Enrollment:**
- `easy_start_option` is NOT one of the above values (or is empty)

## Formatting

### Section Headers
- **Background**: Dark blue (#1155CC)
- **Font**: Bold, white, size 12
- **Behavior**: Merged across all columns

### Column Headers
- **Background**: Medium blue (#4285F4)
- **Font**: Bold, white, centered
- **Position**: Immediately below each section header

### Data Rows
- **Normal styling** (no special formatting)
- **Deal Name**: Hyperlinked to HubSpot deal page

### Color Coding
- **Warm Handoff** and **Ask for Referral** columns use gradient:
  - 🟥 0 = Light red (#F4C7C3)
  - 🟨 2.5 = Light yellow (#FCE8B2)
  - 🟩 5 = Light green (#B7E1CD)

## Historical Preservation Logic

1. **On refresh**, the sheet reads all existing data
2. **Identifies** the first two months (current + last)
3. **Preserves** everything after the second month's Easy Starts section
4. **Writes** new current + last month data at the top
5. **Appends** preserved historical data at the bottom

**Result**: Each month, the historical data "rolls down" and becomes permanent.

## Example Evolution Over Time

### November 2025 Refresh
```
November 2025 (fetched)
  - Enrollments
  - Easy Starts
October 2025 (fetched)
  - Enrollments
  - Easy Starts
September 2025 (preserved)
  - Enrollments
  - Easy Starts
```

### December 2025 Refresh
```
December 2025 (fetched)
  - Enrollments
  - Easy Starts
November 2025 (fetched)
  - Enrollments
  - Easy Starts
October 2025 (preserved - was "last month")
  - Enrollments
  - Easy Starts
September 2025 (preserved)
  - Enrollments
  - Easy Starts
```

## Integration

### Files Modified/Created

1. **HubSpotClient.js** - Added `fetchEnrollmentDeals()` and `fetchEnrollmentDealsPage()`
2. **EnrollmentTracker.js** - New component (full implementation)
3. **Main.js** - Added call to `updateEnrollmentTracker()`

### How to Run

Run `generateAllDashboards()` from the Control Sheet. This will:
1. Fetch enrollment deals for each AE
2. Group by month and split by type
3. Update the 📚 Enrollment Tracker tab
4. Preserve historical data

### Manual Refresh

To refresh a single AE's enrollment tracker:
```javascript
function refreshSingleAEEnrollments() {
  const config = loadConfiguration();
  const person = config.salespeople[0]; // Change index as needed
  
  const sheet = SpreadsheetApp.openById(person.sheetId);
  const result = updateEnrollmentTracker(sheet, person);
  
  Logger.log(result);
}
```

## Benefits of This Approach

✅ **Efficient**: Only 2 months of API calls per refresh  
✅ **Historical**: Never loses old data  
✅ **Accurate**: Catches backdated enrollments in last month  
✅ **Scalable**: Sheet grows naturally over time  
✅ **Simple**: No complex storage or caching needed  

## Future Enhancements

Possible additions (not currently implemented):
- Monthly goal tracking (enrollments vs goal)
- Year-to-date totals
- Comparison to previous months
- Easy Start conversion tracking

