# Pipeline Review - Sheet Structure

## 📊 Complete Column Layout

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                          PIPELINE REVIEW SHEET                                   │
├─────────────────────────────────────────────────────────────────────────────────┤
│ [HubSpot Data - Overwritten on Refresh] │ [Manual - Preserved by Deal ID]      │
├──────────────────────────────────────────┼──────────────────────────────────────┤
│ A-G: Core Fields                         │ O-P: Manual Notes                     │
│ H-N: Call Quality Scores (🟥🟨🟩)       │                                      │
└──────────────────────────────────────────┴──────────────────────────────────────┘
```

## Columns A-G: Core Deal Fields (HubSpot)

| Col | Field                 | Source   | Type      | Display                    |
|-----|-----------------------|----------|-----------|----------------------------|
| A   | Deal ID               | HubSpot  | Text      | Hidden (for reference)     |
| B   | Deal Name             | HubSpot  | Text      | Hyperlinked to HubSpot     |
| C   | Stage                 | HubSpot  | Text      | Current pipeline stage     |
| D   | Last Activity         | HubSpot  | Date      | EST format: yyyy-MM-dd     |
| E   | Next Activity         | HubSpot  | Date      | EST format: yyyy-MM-dd     |
| F   | Next Task Name        | HubSpot  | Text      | ⏳ BLANK (future feature) |
| G   | Why Not Purchase      | HubSpot  | Text      | Key blocker information    |

## Columns H-N: Call Quality Scores (HubSpot)

| Col | Field                 | Score Range | Color Coding         |
|-----|-----------------------|-------------|----------------------|
| H   | Call Quality Score    | 1-5         | 🟥 1-2, 🟨 3, 🟩 4-5 |
| I   | Questioning Technique | 1-5         | 🟥 1-2, 🟨 3, 🟩 4-5 |
| J   | Building Value        | 1-5         | 🟥 1-2, 🟨 3, 🟩 4-5 |
| K   | Funding Options       | 1-5         | 🟥 1-2, 🟨 3, 🟩 4-5 |
| L   | Addressing Objections | 1-5         | 🟥 1-2, 🟨 3, 🟩 4-5 |
| M   | Closing the Deal      | 1-5         | 🟥 1-2, 🟨 3, 🟩 4-5 |
| N   | Ask for Referral      | 1-5         | 🟥 1-2, 🟨 3, 🟩 4-5 |

## Columns O-P: Manual Notes (Editable)

| Col | Field   | Editable | Preserved | Format Preserved |
|-----|---------|----------|-----------|------------------|
| O   | Note 1  | ✅ Yes   | ✅ Yes    | ✅ Yes           |
| P   | Note 2  | ✅ Yes   | ✅ Yes    | ✅ Yes           |

---

## 🔄 Refresh Behavior

### What Gets Overwritten (Columns A-N)
- All HubSpot data refreshes completely
- Deal Name, Stage, Dates, Call Quality Scores all update
- Automatic color coding reapplies

### What Gets Preserved (Columns O-P)
- **Manual Notes**: Note1 and Note2 values preserved by Deal ID
- **Formatting**: Background colors, font colors, bold, italics preserved for ALL columns
- **New Deals**: Get empty Note1/Note2 columns
- **Deleted Deals**: Manual data lost (deal no longer in HubSpot)

---

## 🎨 Formatting Rules

### Automatic Formatting (Applied on Every Refresh)
- **Call Quality columns (H-N)**: Red-Yellow-Green gradient based on score
  - Red (1-2): Poor performance
  - Yellow (3): Medium performance
  - Green (4-5): Excellent performance

### Custom Formatting (Preserved by Deal ID)
- Any manual formatting AEs apply (highlighting rows, changing colors, etc.)
- Preserved across refreshes as long as Deal ID still exists

---

## 🔒 Column Protection

- **Columns A-N**: Protected (read-only) - HubSpot source data
- **Columns O-P**: Editable - Manual notes for AE use

---

## 💡 Use Cases for Manual Notes

**Note 1 Examples:**
- "Follow up after budget meeting"
- "Waiting on VP approval"
- "Hot lead - close by EOW"

**Note 2 Examples:**
- "Competitor mentioned: Acme Corp"
- "Priority: HIGH"
- "Demo scheduled"

---

## 🚀 Future Enhancements

### Next Task Name (Column F)
- Currently blank
- Requires HubSpot API key with task read permissions
- When enabled, will show task title from HubSpot
- Column position reserved now for future use

### Additional Manual Columns
- Easy to add more manual columns (Note 3, Flag, Priority, etc.)
- Just add to `MANUAL_FIELDS` config array
- Preservation logic automatically extends

---

**Last Updated**: November 19, 2025

