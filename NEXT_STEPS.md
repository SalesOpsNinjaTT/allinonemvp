# Next Steps - Phase 1

## Week 1: Control Sheet Setup (This Week)

### 1. Create Control Sheet
- [ ] Create new Google Sheet: "All in One - Control Sheet"
- [ ] Add tab: `👥 Salespeople Config` with columns: Name | Email | Sheet ID | Sheet URL
- [ ] Add tab: `🔧 Tech Access` with columns: Purpose | Email
- [ ] Add tab: `🎯 Goals & Quotas` with columns: Email | Monthly Goal
- [ ] Add tab: `📊 Summary Dashboard` (blank for now)
- [ ] Add 2-3 test salespeople to Salespeople Config

### 2. Configure Apps Script
- [ ] Open Control Sheet → Extensions → Apps Script
- [ ] Link to existing project (Script ID: `1ndG3sIH-I_4Sz8OoHKRCU8Mr1ly2myU25Iw2tlJaEYf70hz3fv8wgD8k`)
- [ ] OR create new project and update our code
- [ ] Set Script Property: `HUBSPOT_ACCESS_TOKEN` = your token

---

## Week 2: Extract & Test Core Code

### 3. Extract HubSpot API Client
- [ ] Clone Call-Quality-Review repo locally
- [ ] Copy `HubSpotAPI.gs` → `src/services/HubSpotClient.js`
- [ ] Copy `DataFetcher.gs` logic
- [ ] Test: Fetch deals from HubSpot

### 4. Extract Self-Provisioning Logic
- [ ] Open Bonuses Clean 2.2 code
- [ ] Copy config loader (reads Salespeople Config)
- [ ] Copy self-provisioning logic (creates new sheets)
- [ ] Test: Add test person, verify sheet auto-creates

### 5. Push & Test
- [ ] `clasp push`
- [ ] Run main function
- [ ] Verify: Individual sheet created for test person
- [ ] Verify: Sheet ID filled in Control Sheet

---

## Success Criteria - End of Week 2

✅ Control Sheet exists with config tabs  
✅ HubSpot API can fetch deals  
✅ Self-provisioning works (add person → sheet auto-creates)  
✅ Ready to build Pipeline Review component

---

**Start with #1** - Create the Control Sheet

