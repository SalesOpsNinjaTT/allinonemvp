# All in One Dashboard

Single dashboard per salesperson consolidating: Pipeline Review, Bonus Calculation, Enrollment Tracker, Operational Metrics.

## Architecture

**Control Sheet** (central hub) → **Individual Sheets** (per salesperson, 4 tabs each)

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)

## Technology

- Google Apps Script + Sheets
- HubSpot API
- clasp (local dev)

## Quick Start

```bash
# Login
clasp login

# Push code
clasp push

# Open in browser
clasp open
```

See [docs/SETUP_GUIDE.md](docs/SETUP_GUIDE.md) for full setup

## Components

1. **Pipeline Review** - HubSpot deals + call quality scores + manual notes
2. **Bonus Calculation** - Commission dashboard + business month override
3. **Enrollment Tracker** - Monthly enrollment tracking (HubSpot source)
4. **Operational Metrics** - Call metrics + performance analytics

## Timeline

~14 weeks total, **MVP in 6 weeks** (Pipeline + Bonuses)

## Documentation

- [ARCHITECTURE.md](docs/ARCHITECTURE.md) - System overview
- [CONTROL_SHEET_STRUCTURE.md](docs/CONTROL_SHEET_STRUCTURE.md) - Control Sheet tabs
- [PIPELINE_REVIEW_STRUCTURE.md](docs/PIPELINE_REVIEW_STRUCTURE.md) - Pipeline Review columns
- [SETUP_GUIDE.md](docs/SETUP_GUIDE.md) - Development setup
- [Roadmap](../ALL_IN_ONE_DASHBOARD_ROADMAP.md) - Full plan
- [Principles](../PRINCIPLES.md) - Development standards

## Status

**Phase**: Phase 1 - Setup & Foundation  
**Updated**: November 19, 2025

