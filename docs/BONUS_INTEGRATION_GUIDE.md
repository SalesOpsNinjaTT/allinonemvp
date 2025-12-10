# Bonus Calculation Integration Guide

## Overview

This guide explains how the new All-in-One Dashboard integrates with the existing Clean 2.2 bonus calculation system.

**Strategy:** Instead of re-implementing the bonus calculation logic, we **copy the already-calculated "📈 Monthly Bonus Report" tab** from each person's legacy bonus sheet into their new All-in-One Dashboard.

---

## How It Works

### 1. **Data Source**
- **Legacy Control Sheet:** `1OTyjtQ27iCXhqgscNomEktNQg2qlS6Uk0_j6EeeyEFI`
- **Tab:** `👥 Salespeople Config`
- **Columns Used:**
  - Column A: Salesperson Name
  - Column B: auth_email (Email)
  - Column C: Spreadsheet ID (Legacy bonus sheet)

### 2. **Process Flow**

```
1. Read NEW All-in-One Control Sheet
   └─ Get list of salespeople

2. For each person:
   ├─ Look up their legacy bonus sheet ID from legacy Control Sheet
   ├─ Open their legacy bonus sheet
   ├─ Find the "📈 Monthly Bonus Report" tab
   ├─ Copy ALL data + formatting (values, colors, fonts, number formats)
   └─ Paste into new "💰 Bonus Calculation" tab

3. Log success/failure for each person
```

### 3. **What Gets Copied**

✅ All cell values  
✅ All formatting (backgrounds, font colors, font weights, font sizes)  
✅ Number formats (currency, percentages, dates)  
✅ Column widths (auto-resized)  
✅ Frozen rows (if any)  

❌ Does NOT copy formulas (only values)  
❌ Does NOT copy conditional formatting rules (Google Sheets limitation)  

---

## Configuration

### Legacy Sheet Column Mapping

If the legacy Control Sheet has a different structure, update these constants in `BonusCalculation.js`:

```javascript
const LEGACY_COLUMNS = {
  NAME: 1,        // Column A: Salesperson Name
  EMAIL: 2,       // Column B: auth_email
  SHEET_ID: 3,    // Column C: Spreadsheet ID
  SHEET_URL: 4    // Column D: Spreadsheet URL (optional)
};
```

### Legacy Control Sheet ID

To change which legacy Control Sheet to read from:

```javascript
const LEGACY_CONTROL_SHEET_ID = '1OTyjtQ27iCXhqgscNomEktNQg2qlS6Uk0_j6EeeyEFI';
```

---

## Testing

### Test with Single Person

**Before running on everyone, test with one person:**

1. Open Apps Script Editor
2. Select function: `testBonusCopy`
3. Click Run ▶️
4. Check logs for success/failure

**Expected Output:**
```
=== Testing Bonus Calculation Copy ===
Testing bonus copy for: John Doe (john@company.com)
[Bonus Calculation] Updating for John Doe...
  Step 1: Looking up legacy bonus sheet ID...
  Found legacy sheet ID: 1ABC123...
  Step 2: Opening legacy bonus sheet...
  Step 3: Copying data from legacy sheet...
  Step 4: Writing to new dashboard...
  Step 5: Auto-resizing columns...
[Bonus Calculation] Complete for John Doe (2.3s)
✅ Bonus copy successful
   Rows copied: 45
   Columns copied: 8
   Duration: 2.3s
```

### Run for All Salespeople

**Once testing succeeds, run full generation:**

```javascript
generateAllDashboards()
```

This will:
1. Update Director Hub
2. Update each AE's:
   - Pipeline Review
   - Enrollment Tracker
   - **Bonus Calculation** (copied from legacy)
3. Sync director flags

---

## Error Handling

### Scenario 1: No Legacy Sheet ID Found

**Message:** `No legacy sheet ID found`

**Cause:** Person exists in new system but not in legacy system

**Solution:** 
- Add the person to the legacy Control Sheet, OR
- Manually enter their legacy sheet ID in the legacy Control Sheet, OR
- They'll see "No legacy bonus data found" message in their tab

### Scenario 2: Legacy Sheet Has No "Monthly Bonus Report" Tab

**Message:** `No Monthly Bonus Report tab in legacy sheet`

**Cause:** Legacy sheet exists but is missing the bonus report tab

**Solution:**
- Run the legacy `generateSalespersonDashboards()` function first
- Or manually create/rename the tab to "📈 Monthly Bonus Report"

### Scenario 3: Legacy Bonus Tab is Empty

**Message:** `Legacy bonus tab is empty`

**Cause:** The tab exists but has no data

**Solution:**
- Run the legacy bonus calculation to populate data
- Check if the person has any bonus data in their `💎 Bonuses 3.0` source tab

---

## Matching Logic

The system tries to match people between the new and legacy systems using:

1. **Email Match (Primary):** Case-insensitive email comparison
2. **Name Match (Fallback):** Case-insensitive name comparison

**Example:**
- New system: Email = `john.doe@company.com`
- Legacy system: Email = `John.Doe@Company.com`
- ✅ **MATCH** (case-insensitive)

---

## Maintenance

### When to Re-Run

**Run `generateAllDashboards()` whenever:**
- Bonus calculations are updated in the legacy system
- You want to refresh all bonus data
- New salespeople are added

**Frequency:** 
- Daily after bonus calculations run, OR
- On-demand when directors need updated data

### Business Month Override

The legacy system respects the **Business Month Override** (Cell L2 in legacy Control Sheet).

- If `L2 = 2025-11`, bonuses for November will show
- If `L2 is empty`, current month bonuses show

**This is handled by the legacy system** - you just copy the results!

---

## Troubleshooting

### Issue: "Cannot find legacy Control Sheet"

**Error:** `Exception: Requested entity was not found.`

**Solution:** Check that `LEGACY_CONTROL_SHEET_ID` is correct

### Issue: "Permission denied when opening legacy sheet"

**Error:** `Exception: You do not have permission to access...`

**Solution:** 
- Ensure the Apps Script has permission to access legacy sheets
- Share legacy sheets with the script's service account, OR
- Run script as a user who has access to all legacy sheets

### Issue: Formatting looks different

**Cause:** Some complex formatting may not copy perfectly

**Solution:**
- Check if conditional formatting rules are being used (these don't copy)
- Manually reapply conditional formatting in new sheets if needed

---

## Future Enhancements

### Optional: Copy Dashboard Summary Tab

Currently we only copy the "📈 Monthly Bonus Report" tab.

To also copy the "📊 Dashboard" summary:

```javascript
// In Main.js, add after updateBonusCalculation():
const dashboardResult = copyLegacyDashboard(sheet, person);
```

This will create a "📊 Legacy Dashboard" tab with the summary metrics.

### Optional: Copy Multiple Tabs

To copy additional tabs (e.g., "📚 Enrollments"):

1. Add tab name constant to `BonusCalculation.js`
2. Create new copy function similar to `updateBonusCalculation()`
3. Call it from `Main.js`

---

## Summary

✅ **Pros of this approach:**
- No need to re-implement complex bonus calculation logic
- Guaranteed data consistency with legacy system
- Simple copy operation (fast and reliable)
- Easy to maintain and debug

⚠️ **Cons:**
- Dependent on legacy system structure
- Requires legacy system to be up-to-date
- Formulas are converted to values (static data)

**Bottom Line:** This is a pragmatic solution that gets bonus data into the new dashboard quickly while the legacy system continues to work.

---

**Last Updated:** December 10, 2025  
**Component:** `BonusCalculation.js`  
**Status:** ✅ Ready for Testing

