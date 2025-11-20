# Project Information

## Repository
- **GitHub**: https://github.com/SalesOpsNinjaTT/allinonemvp.git

## Apps Script Project
- **Type**: Container-bound (attached to Control Sheet) ✅
- **Account**: konstantin.gevorkov@tripleten.com (Work Account) ✅
- **Script ID**: `1AlRJpl-uvs14yfoIpk_Y7rQ4XcKjm4KOSu7sHtDAKOFwyxjc-7uDqm4z`
- **Editor URL**: https://script.google.com/u/0/home/projects/1AlRJpl-uvs14yfoIpk_Y7rQ4XcKjm4KOSu7sHtDAKOFwyxjc-7uDqm4z/edit
- **Access via**: Control Sheet → Extensions → Apps Script

## Control Sheet
- **Name**: All-In-One 2.0 MVP ✅
- **Sheet ID**: `1-zipx1vWfjYaMjgl7BbqfCVjl8NZch9DMk5T-DRfnnQ`
- **Sheet URL**: https://docs.google.com/spreadsheets/d/1-zipx1vWfjYaMjgl7BbqfCVjl8NZch9DMk5T-DRfnnQ/edit

## Setup Status

✅ **Completed**:
- Project structure created
- Git repository initialized
- GitHub repository created and code pushed
- Apps Script project created
- Code pushed to Apps Script (10 files)
- Control Sheet created with 4 tabs
- ConfigManager & SheetProvisioner implemented

## Quick Commands

### Open Apps Script in browser
```bash
clasp open
```

### Push local changes to Apps Script
```bash
clasp push
```

### View execution logs
```bash
clasp logs
```

### Pull changes from Apps Script to local
```bash
clasp pull
```

## Next Steps

1. **Add Salespeople to Control Sheet**:
   - Open Control Sheet (URL above)
   - Go to "👥 Salespeople Config" tab
   - Add Name and Email for each salesperson

2. **Configure HubSpot API Token**:
   - In Apps Script: Project Settings > Script Properties
   - Add: `HUBSPOT_ACCESS_TOKEN` = `your-token`

3. **Test Self-Provisioning**:
   - In Apps Script, run function: `testSingleSalesperson()`
   - Check logs for new individual sheet URL
   - Verify sheet was created with 4 tabs

4. **Continue Phase 1**: Extract HubSpot API client
   - See: `docs/SETUP_GUIDE.md` and `ALL_IN_ONE_DASHBOARD_ROADMAP.md`

---

**Last Updated**: November 20, 2025 (Migrated to container-bound script)

