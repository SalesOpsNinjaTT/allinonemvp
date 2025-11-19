# Pipeline Review - Sheet Structure

## Column Layout (16 columns)

```
[A-G: Core Fields] → [H-N: Call Quality 🟥🟨🟩] → [O-P: Manual Notes]
```

### A-G: Core Fields (HubSpot)
| Col | Field | Notes |
|-----|-------|-------|
| A | Deal ID | Hidden |
| B | Deal Name | Hyperlinked to HubSpot |
| C | Stage | Current pipeline stage |
| D | Last Activity | Date (EST) |
| E | Next Activity | Date (EST) |
| F | Next Task Name | ⏳ Blank (future - needs API permission) |
| G | Why Not Purchase Today | Key blocker |

### H-N: Call Quality Scores (HubSpot)
| Col | Field | Color |
|-----|-------|-------|
| H | Call Quality Score | 🟥 1-2, 🟨 3, 🟩 4-5 |
| I | Questioning Technique | 🟥 1-2, 🟨 3, 🟩 4-5 |
| J | Building Value | 🟥 1-2, 🟨 3, 🟩 4-5 |
| K | Funding Options | 🟥 1-2, 🟨 3, 🟩 4-5 |
| L | Addressing Objections | 🟥 1-2, 🟨 3, 🟩 4-5 |
| M | Closing the Deal | 🟥 1-2, 🟨 3, 🟩 4-5 |
| N | Ask for Referral | 🟥 1-2, 🟨 3, 🟩 4-5 |

### O-P: Manual Notes (Editable)
| Col | Field | Preserved |
|-----|-------|-----------|
| O | Note 1 | ✅ Yes, by Deal ID |
| P | Note 2 | ✅ Yes, by Deal ID |

---

## Refresh Behavior

**Overwritten** (A-N): All HubSpot data, auto color-coding reapplied

**Preserved** (O-P): Manual notes + all formatting (backgrounds, colors) by Deal ID

**New deals**: Empty notes

**Deleted deals**: Notes lost (deal gone from HubSpot)

---

## Protection

- **A-N**: Read-only (HubSpot data)
- **O-P**: Editable (manual notes)

---

## Field Configuration

```javascript
const CORE_FIELDS = [
  { property: 'dealId', header: 'Deal ID', hidden: true },
  { property: 'dealname', header: 'Deal Name', hyperlink: true },
  { property: 'dealstage', header: 'Stage' },
  { property: 'notes_last_updated', header: 'Last Activity', type: 'date' },
  { property: 'notes_next_activity_date', header: 'Next Activity', type: 'date' },
  { property: 'next_task_name', header: 'Next Task Name', enabled: false },
  { property: 'why_not_purchase_today_', header: 'Why Not Purchase Today' }
];

const CALL_QUALITY_FIELDS = [
  { property: 'call_quality_score', header: 'Call Quality Score', colorCode: true },
  // ... more fields
];

const MANUAL_FIELDS = [
  { header: 'Note 1', editable: true, preserve: true },
  { header: 'Note 2', editable: true, preserve: true }
];
```

**To customize**: Edit config arrays, no code changes needed
