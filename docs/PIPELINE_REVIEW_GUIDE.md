# Pipeline Review System - Complete Guide

## 🎯 Overview

The Pipeline Review system provides **bi-directional synchronization** between individual Account Executive (AE) sheets and Director consolidated views.

### System Architecture

```
┌─────────────────────────────────────────────────────────┐
│  Individual AE Sheets (One per AE)                      │
│  - John Doe's Pipeline Review                           │
│  - Jane Smith's Pipeline Review                         │
│                                                         │
│  AE Can:                                                │
│  - See their own deals only                             │
│  - Add/edit NOTES                                       │
│  - Receive HIGHLIGHTING from Director                   │
└─────────────────┬───────────────────────────────────────┘
                  │
                  │ ┌─────────────────────────┐
                  ├─┤ NOTES flow UP          │
                  │ └─────────────────────────┘
                  │
                  │ ┌─────────────────────────┐
                  └─┤ HIGHLIGHTING flows DOWN │
                    └─────────────────────────┘
                  │
┌─────────────────┴───────────────────────────────────────┐
│  Control Sheet - Director Tabs (Multiple)               │
│                                                         │
│  Tab: "🎯 Director - Kraken Team"                       │
│  Tab: "🎯 Director - Victorious Team"                   │
│  Tab: "🎯 Asst Director - DeeStroyers Team"            │
│                                                         │
│  Each Director Can:                                     │
│  - See ALL deals from THEIR TEAM                        │
│  - See ALL notes from their AEs                         │
│  - HIGHLIGHT deals (rows or cells)                      │
│  - Highlighting syncs back to AE sheets                 │
└─────────────────────────────────────────────────────────┘
```

---

## ✅ What's Already Implemented

### 1. **Individual AE Pipeline Review** (`PipelineReview.js`)
- ✅ Fetches deals from HubSpot for each AE
- ✅ Displays with call quality scores (color-coded)
- ✅ Manual "Notes" column (editable)
- ✅ Preserves notes across refreshes (by Deal ID)
- ✅ Applies highlighting from Director

### 2. **Director Consolidated Views** (`DirectorHub.js`)
- ✅ One tab per Director/Asst Director
- ✅ Shows all deals from their team
- ✅ Collects notes from all AE sheets
- ✅ Director can highlight rows/cells
- ✅ Highlighting syncs to AE sheets

### 3. **Bi-Directional Sync**
- ✅ **AE → Director**: Notes flow up automatically
- ✅ **Director → AE**: Highlighting flows down automatically
- ✅ Both preserved by Deal ID across refreshes

### 4. **Call Quality Scoring**
- ✅ 7 call quality dimensions from HubSpot
- ✅ Color-coded: Red (0-2), Yellow (3), Green (4-5)
- ✅ Applied to both AE and Director views

---

## 📋 Pipeline Review Columns

### **Individual AE Sheet: "📊 Pipeline Review"**

| Column | Field | Source | Editable | Description |
|--------|-------|--------|----------|-------------|
| A | Deal ID | HubSpot | ❌ No (Hidden) | Primary key for sync |
| B | Deal Name | HubSpot | ❌ No | Hyperlinked to HubSpot |
| C | Stage | HubSpot | ❌ No | Current pipeline stage |
| D | Last Activity | HubSpot | ❌ No | Last engagement date |
| E | Next Activity | HubSpot | ❌ No | Scheduled next action |
| F | Why Not Purchase Today | HubSpot | ❌ No | AI-generated blocker |
| G | Call Quality Score | HubSpot | ❌ No | Overall score (0-5) 🟥🟨🟩 |
| H | Questioning | HubSpot | ❌ No | Discovery score 🟥🟨🟩 |
| I | Building Value | HubSpot | ❌ No | Value prop score 🟥🟨🟩 |
| J | Funding Options | HubSpot | ❌ No | Funding discussion score 🟥🟨🟩 |
| K | Addressing Objections | HubSpot | ❌ No | Objection handling score 🟥🟨🟩 |
| L | Closing the Deal | HubSpot | ❌ No | Closing technique score 🟥🟨🟩 |
| M | Ask for Referral | HubSpot | ❌ No | Referral request score 🟥🟨🟩 |
| **N** | **Notes** | **Manual** | **✅ YES** | **AE fills this, syncs to Director** |

### **Director's Consolidated View**

Same columns as AE sheet, PLUS:
- **Column C**: Owner Name (AE Name) - added after Deal Name
- **Column O**: Notes from AE (read-only for director)
- **Highlighting**: Director can highlight any row/cell

---

## 🚀 Setup Instructions

### Step 1: Configure Director Config Tab

In the Control Sheet, create/update the **"👥 Director Config"** tab:

```
| Director Name | Director Email        | Team Name    | Type      | Tab Name                        |
|---------------|-----------------------|--------------|-----------|----------------------------------|
| Sarah Connor  | sarah@company.com     | Kraken       | Director  | 🎯 Director - Kraken Team        |
| John Matrix   | john@company.com      | Victorious   | Director  | 🎯 Director - Victorious Team    |
| Kyle Reese    | kyle@company.com      | DeeStroyers  | Asst Dir  | 🎯 Asst Director - DeeStroyers   |
```

**Columns:**
- **A**: Director Name
- **B**: Director Email
- **C**: Team Name (must match Salespeople Config)
- **D**: Type (Director or Asst Dir)
- **E**: Tab Name (auto-generated if blank)

### Step 2: Update Salespeople Config

In **"👥 Salespeople Config"**, ensure **Team** column is filled:

```
| Name       | Email            | Sheet ID | Sheet URL | HubSpot User ID | Team       | Role |
|------------|------------------|----------|-----------|-----------------|------------|------|
| John Doe   | john@company.com | ...      | ...       | 12345678        | Kraken     | AE   |
| Jane Smith | jane@company.com | ...      | ...       | 87654321        | Victorious | AE   |
```

**Important:** Team names must match exactly between Salespeople Config and Director Config.

### Step 3: Verify HubSpot Properties

Ensure these properties exist in your HubSpot account:
- `dealname`
- `dealstage`
- `hubspot_owner_id`
- `notes_last_updated`
- `notes_next_activity_date`
- `why_not_purchase_today_`
- `call_quality_score`
- `s_discovery_a_questioning_technique__details`
- `s_building_value_a_tailoring_features_and_benefits__details`
- `s_funding_options__a_identifying_funding_needs__details`
- `s_addressing_objections_a_identifying_and_addressing_objections_and_obstacles__details`
- `s_closing_the_deal__a_assuming_the_sale__details`
- `s_closing_the_deal__a_ask_for_referral__details`

### Step 4: Run the Script

1. Open Control Sheet
2. Go to: Extensions → Apps Script
3. Select function: `generateAllDashboards`
4. Click **Run** ▶️

**What happens:**
1. ✅ Creates/updates Pipeline Review tab for each AE
2. ✅ Creates/updates Director tabs in Control Sheet
3. ✅ Syncs notes from AEs → Directors
4. ✅ Syncs highlighting from Directors → AEs

---

## 🔄 How Bi-Directional Sync Works

### **Flow 1: AE Notes → Director** (Automatic)

1. AE adds notes in Column N of their Pipeline Review
2. Next script run: `collectNotesFromTeamAEs()` reads all AE notes by Deal ID
3. `buildConsolidatedPipelineDataArray()` merges notes into Director view
4. Director sees all notes from their team in Column O

### **Flow 2: Director Highlighting → AE** (Automatic)

1. Director highlights deals in their consolidated tab (any rows/cells)
2. Next script run: `captureDirectorHighlighting()` captures by Deal ID
3. `syncDirectorFlagsToAESheets()` applies highlighting to AE sheets
4. AE sees highlighting on their Pipeline Review

---

## 🎨 Highlighting Rules

**Directors can highlight:**
- ✅ Entire rows (all columns)
- ✅ Individual cells
- ✅ Background colors
- ✅ Font colors

**What gets preserved:**
- ✅ Background colors for all cells in a row
- ✅ Font colors for all cells in a row
- ✅ Preserved by Deal ID (not row position)

**Example Use Cases:**
- 🟢 Green highlight = Hot deal, close soon
- 🔴 Red highlight = Needs urgent attention
- 🟡 Yellow highlight = Follow up required
- 🔵 Blue highlight = On hold, check next week

---

## 🧪 Testing

### Test with One AE

```javascript
function testSingleSalesperson() {
  // Defined in Main.js
  // Tests with first AE in Salespeople Config
}
```

### Test Director Pipeline

```javascript
function testDirectorPipeline() {
  const config = loadConfiguration();
  const directors = getDirectorConfig();
  const controlSheet = SpreadsheetApp.openById(CONTROL_SHEET_ID);
  
  if (directors.length > 0) {
    updateDirectorConsolidatedPipeline(directors[0], config, controlSheet);
    Logger.log('✅ Test complete');
  }
}
```

---

## 🐛 Troubleshooting

### Issue: "No deals showing in AE sheet"
**Cause:** HubSpot User ID might be wrong or missing  
**Fix:** Check Column E in Salespeople Config, verify HubSpot owner ID

### Issue: "Notes not syncing to Director"
**Cause:** Deal ID might have changed or AE sheet not accessible  
**Fix:** 
- Verify AE Sheet ID in Salespeople Config
- Check that script has permission to access AE sheets
- Run script again (notes sync on next run)

### Issue: "Highlighting not syncing to AE"
**Cause:** Deal ID mismatch or AE sheet not found  
**Fix:**
- Verify highlighting is applied in Director tab BEFORE running script
- Check that Deal IDs match between Director and AE sheets
- Run script again (highlighting syncs on next run)

### Issue: "Call quality scores not showing"
**Cause:** HubSpot properties missing or not accessible  
**Fix:**
- Verify call quality properties exist in HubSpot
- Check HubSpot API token has permission to read these fields
- Properties must be numeric (0-5 scale)

### Issue: "Director tab not created"
**Cause:** Director Config not set up  
**Fix:**
- Create "👥 Director Config" tab in Control Sheet
- Add at least one director with Team Name
- Team Name must match Salespeople Config exactly

---

## 📊 Performance

**Expected execution time:**
- **Per AE**: 3-5 seconds (fetch + write + format)
- **Per Director**: 2-3 seconds (collect notes + write)
- **Total for 10 AEs + 2 Directors**: ~40-60 seconds

**Optimization tips:**
- Run hourly (not more frequently)
- Use time-based trigger for automatic updates
- Monitor Apps Script execution logs

---

## 🔐 Security & Permissions

### Who Can See What

**AEs:**
- ✅ See only their own deals
- ✅ Can edit Notes column
- ✅ See highlighting from their Director
- ❌ Cannot see other AEs' deals
- ❌ Cannot see Director tabs

**Directors:**
- ✅ See all deals from their team
- ✅ See all notes from their AEs
- ✅ Can highlight any deal
- ✅ Can view their consolidated tab
- ❌ Cannot see other teams' deals

**Script Access:**
- ✅ Read access to all AE sheets
- ✅ Write access to all AE sheets (for highlighting sync)
- ✅ Read/write access to Control Sheet

---

## 📝 Next Steps

1. ✅ **Set up Director Config** (Step 1 above)
2. ✅ **Update Salespeople Config** with Team names (Step 2 above)
3. ✅ **Run `generateAllDashboards()`** (Step 4 above)
4. ✅ **Test with 1-2 AEs** to verify notes sync
5. ✅ **Test highlighting** to verify Director → AE flow
6. ✅ **Roll out to full team**

---

## 🎯 Key Files

- `src/components/PipelineReview.js` - Individual AE pipeline logic
- `src/components/DirectorHub.js` - Director consolidated views
- `src/components/DirectorMenu.js` - Director → AE sync logic
- `src/services/ConfigManager.js` - Config loading
- `src/services/HubSpotClient.js` - HubSpot API integration
- `src/Main.js` - Main orchestration

---

## ✅ Success Criteria

- ✅ Each AE sees their own Pipeline Review with call quality scores
- ✅ AEs can add notes in Notes column
- ✅ Directors see consolidated view with all team deals
- ✅ Directors see all AE notes in their view
- ✅ Director highlighting appears in AE sheets on next run
- ✅ AE notes appear in Director view on next run
- ✅ Data preserved across refreshes (by Deal ID)

---

**Last Updated:** December 10, 2025  
**Status:** ✅ Fully Implemented & Ready for Use

