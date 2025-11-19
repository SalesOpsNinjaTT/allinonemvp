# Architecture

## System Overview

```
┌──────────────────┐
│ CONTROL SHEET    │ ← Apps Script runs here
│ (Central Hub)    │
├──────────────────┤
│ Config           │ → Salespeople list
│ Goals            │ → Targets
│ Summary          │ → Team metrics
└────────┬─────────┘
         │
         ├─── Auto-creates & updates ───┐
         ▼                               ▼
   [John - Dashboard]            [Jane - Dashboard]
   ├─ Pipeline Review            ├─ Pipeline Review
   ├─ Bonus Calculation          ├─ Bonus Calculation
   ├─ Enrollment Tracker         ├─ Enrollment Tracker
   └─ Operational Metrics        └─ Operational Metrics
```

## Data Flow

```
HubSpot API → Apps Script → Individual Sheets (per salesperson)
                    ↓
             Control Sheet (summary)
```

## Key Components

**Control Sheet** - Single source for configuration
- Salespeople list (self-provisioning)
- Goals/targets
- Summary dashboard

**Individual Sheets** - One per salesperson
- 4 tabs (Pipeline, Bonuses, Enrollment, Ops)
- Auto-created on first run
- Personal data workspace

**Apps Script** - Backend orchestration
- Runs from Control Sheet
- Fetches from HubSpot
- Updates all sheets
- Preserves manual notes/formatting

## Self-Provisioning

Add to Salespeople Config → Script auto-creates sheet → ID auto-fills

## Code Structure

```
src/
├── Main.js                     # Orchestration
├── services/
│   ├── HubSpotClient.js        # API integration
│   └── ConfigManager.js        # Config loading
└── components/
    ├── PipelineReview.js       # Component 1
    ├── BonusCalculation.js     # Component 2
    ├── EnrollmentTracker.js    # Component 3
    └── OperationalMetrics.js   # Component 4
```

## Principles

- **EST timestamps only**
- **Idempotent operations** (can rerun safely)
- **Preserve manual notes** by Deal ID
- **Self-provisioning** (auto-create sheets)
- **Config-driven** (easy customization)

