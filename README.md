# All in One Dashboard

MVP dashboard consolidating Pipeline Review, Bonus Calculation, Enrollment Tracker, and Operational Metrics into a single Google Sheet per salesperson.

## Project Structure

```
all-in-one-dashboard/
├── src/
│   ├── services/       # Shared services (HubSpot API, Sheets operations, etc.)
│   ├── components/     # Main components (Pipeline, Bonuses, Enrollment, Ops)
│   ├── models/         # Data models
│   └── utils/          # Helper functions
├── config/             # Configuration files
├── docs/               # Documentation
├── .clasp.json         # Clasp configuration
├── appsscript.json     # Apps Script manifest
└── README.md
```

## Technology Stack

- **Platform**: Google Apps Script
- **Data Storage**: Google Sheets
- **APIs**: HubSpot CRM
- **Development**: clasp (local development)
- **Version Control**: Git

## Setup

### Prerequisites

- Node.js and npm installed
- Google account with Apps Script access
- clasp CLI installed: `npm install -g @google/clasp`

### Initial Setup

1. **Login to clasp**:
   ```bash
   clasp login
   ```

2. **Clone this repository**:
   ```bash
   git clone [repository-url]
   cd all-in-one-dashboard
   ```

3. **Create Apps Script project**:
   ```bash
   clasp create --type standalone --title "All in One Dashboard"
   ```

4. **Push code to Apps Script**:
   ```bash
   clasp push
   ```

## Development Workflow

See [PRINCIPLES.md](../PRINCIPLES.md) for detailed development principles.

### Basic Workflow

1. **Pull latest from Apps Script** (if needed):
   ```bash
   clasp pull
   ```

2. **Make changes locally** in `src/` directory

3. **Push to Apps Script**:
   ```bash
   clasp push
   ```

4. **View logs**:
   ```bash
   clasp logs
   ```

5. **Open in browser**:
   ```bash
   clasp open
   ```

## Components

### 1. Pipeline Review
- HubSpot deal synchronization
- Notes preservation
- Individual AE sheets
- Status: Extracting from VERSION 2.9

### 2. Bonus Calculation
- Commission dashboards
- Business month override
- Self-provisioning
- Status: Extracting from Clean 2.2

### 3. Enrollment Tracker
- HubSpot-based enrollment tracking
- Achievement metrics
- Status: Build from scratch

### 4. Operational Metrics
- Call metrics
- Performance analytics
- Status: Lowest priority

## Timeline

- **Phase 1-2**: Foundation & Setup (Weeks 1-2)
- **Phase 3**: Pipeline Review Migration (Weeks 3-4)
- **Phase 4**: Bonus Calculation Migration (Weeks 5-6)
- **Phase 5**: Enrollment Tracker (Weeks 7-9)
- **Phase 6**: Operational Metrics (Weeks 10-12) - Optional
- **Phase 7**: Polish & Rollout (Weeks 13-14)

**MVP Target**: 6 weeks (Pipeline + Bonuses)

## Key Principles

- **EST timestamps only** - All time handling in Eastern Time
- **Idempotent operations** - Can rerun safely
- **Keep raw data untouched** - Never modify source tabs
- **Fail fast on bad data** - Don't guess, surface errors
- **One component at a time** - Finish before starting next
- **Small sample first** - Test with 1-2 people before full rollout

## Configuration

Configuration will be stored in Google Sheets and Script Properties:

- `HUBSPOT_ACCESS_TOKEN` - HubSpot API token (Script Properties)
- Salesperson list and mapping (Google Sheets)
- Goal tracking data (Google Sheets)

## Testing

- Start with 1-2 pilot salespeople
- Validate data accuracy before full rollout
- Keep legacy systems running in parallel during transition

## Documentation

- [Roadmap](../ALL_IN_ONE_DASHBOARD_ROADMAP.md) - Full project roadmap
- [Principles](../PRINCIPLES.md) - Development principles
- [Legacy Pipeline Review](../Pipeline%20Review%20Legacy%20MIGRATION_ANALYSIS.md) - Pipeline Review migration analysis
- [Legacy Bonuses](../Bonuses%20Legacy%20Clean%202.2%20-%20DOCUMENTATION.md) - Bonus calculation documentation

## Status

**Current Phase**: Phase 1 - Setup & Foundation

**Last Updated**: November 19, 2025

