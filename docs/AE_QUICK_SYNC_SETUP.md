# 📤 Quick Sync Setup for Account Executives

## One-Time Setup (2 minutes)

### Step 1: Open Apps Script Editor

1. Open your individual dashboard sheet
2. Click: **Extensions → Apps Script**
3. You'll see a code editor

### Step 2: Copy and Paste This Code

Delete any existing code and paste this:

```javascript
// ============================================
// Quick Sync for Account Executives
// Auto-generated code - do not edit
// ============================================

// Control Sheet ID (where directors see consolidated data)
const CONTROL_SHEET_ID = '1OTyjtQ27iCXhqgscNomEktNQg2qlS6Uk0_j6EeeyEFI';

// Tab names
const TAB_PIPELINE = '📊 Pipeline Review';

// ============================================
// MENU: Creates "⚡ Quick Sync" menu on sheet open
// ============================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Quick Sync')
    .addItem('📤 Push My Notes to Director', 'pushMyNotesToDirector')
    .addToUi();
}

// ============================================
// FUNCTION: Push notes to director's sheet
// ============================================
function pushMyNotesToDirector() {
  const startTime = new Date();
  const mySheet = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    mySheet.toast('📤 Pushing your notes to director...', 'Syncing', 3);
    
    // Step 1: Read my notes from Pipeline Review
    const pipelineSheet = mySheet.getSheetByName(TAB_PIPELINE);
    if (!pipelineSheet) {
      throw new Error('Pipeline Review tab not found');
    }
    
    const lastRow = pipelineSheet.getLastRow();
    if (lastRow < 2) {
      mySheet.toast('ℹ️ No deals found in your Pipeline Review', 'No Data', 3);
      return;
    }
    
    // Read Deal IDs (column A) and Notes (column G)
    const NOTES_COLUMN_INDEX = 6; // Column G (0-based)
    const data = pipelineSheet.getRange(2, 1, lastRow - 1, Math.max(7, pipelineSheet.getLastColumn())).getValues();
    
    const myNotes = new Map();
    data.forEach(row => {
      const dealId = row[0]?.toString();
      const notes = row[NOTES_COLUMN_INDEX] || '';
      if (dealId) {
        myNotes.set(dealId, notes);
      }
    });
    
    if (myNotes.size === 0) {
      mySheet.toast('ℹ️ No deals found to sync', 'No Data', 3);
      return;
    }
    
    // Step 2: Find my team and director from Control Sheet
    const controlSheet = SpreadsheetApp.openById(CONTROL_SHEET_ID);
    const configSheet = controlSheet.getSheetByName('👥 Salespeople Config');
    const directorConfigSheet = controlSheet.getSheetByName('👥 Director Config');
    
    if (!configSheet) {
      throw new Error('Salespeople Config not found in Control Sheet');
    }
    
    // Find myself in the config
    const configData = configSheet.getDataRange().getValues();
    const mySheetId = mySheet.getId();
    let myName = '';
    let myTeam = '';
    
    for (let i = 1; i < configData.length; i++) { // Skip header
      const sheetId = configData[i][2]?.toString(); // Column C
      if (sheetId === mySheetId) {
        myName = configData[i][0]; // Column A
        myTeam = configData[i][5]; // Column F
        break;
      }
    }
    
    if (!myTeam) {
      throw new Error('Could not find your team assignment. Please contact admin.');
    }
    
    // Find my director's tab
    const directorData = directorConfigSheet ? directorConfigSheet.getDataRange().getValues() : [];
    let directorTabName = '';
    
    for (let i = 1; i < directorData.length; i++) { // Skip header
      const team = directorData[i][1]; // Column B
      if (team.toLowerCase() === myTeam.toLowerCase()) {
        directorTabName = directorData[i][2]; // Column C
        break;
      }
    }
    
    if (!directorTabName) {
      throw new Error(`No director configured for ${myTeam} team. Please contact admin.`);
    }
    
    // Step 3: Update notes in director's sheet
    const directorSheet = controlSheet.getSheetByName(directorTabName);
    if (!directorSheet) {
      throw new Error(`Director sheet "${directorTabName}" not found. Please run full dashboard generation first.`);
    }
    
    const directorLastRow = directorSheet.getLastRow();
    if (directorLastRow < 2) {
      mySheet.toast('ℹ️ Director sheet is empty. Please run full dashboard generation first.', 'No Data', 5);
      return;
    }
    
    // Read director sheet (Deal ID in column A, Notes in column H)
    const DIRECTOR_NOTES_COLUMN = 8; // Column H (1-based)
    const directorData = directorSheet.getRange(2, 1, directorLastRow - 1, 1).getValues();
    
    let updatedCount = 0;
    directorData.forEach((row, index) => {
      const dealId = row[0]?.toString();
      if (dealId && myNotes.has(dealId)) {
        const rowIndex = index + 2;
        directorSheet.getRange(rowIndex, DIRECTOR_NOTES_COLUMN).setValue(myNotes.get(dealId));
        updatedCount++;
      }
    });
    
    const duration = (new Date() - startTime) / 1000;
    mySheet.toast(`✅ Pushed notes for ${updatedCount} deal(s) to ${directorTabName} (${duration.toFixed(1)}s)`, 'Success', 5);
    Logger.log(`[Quick Sync] ${myName} pushed ${updatedCount} notes to ${directorTabName} (${duration}s)`);
    
  } catch (error) {
    Logger.log(`[Quick Sync] ERROR: ${error.message}`);
    mySheet.toast(`❌ Failed: ${error.message}`, 'Error', 10);
    
    // Send error email to admin
    try {
      MailApp.sendEmail({
        to: 'konstantin.gevorkov@tripleten.com',
        subject: '[All-in-One Dashboard] Quick Sync Error',
        body: `Error occurred during Quick Sync.\n\nUser: ${Session.getActiveUser().getEmail()}\nTimestamp: ${new Date().toISOString()}\n\nError: ${error.message}\n\nStack: ${error.stack || 'No stack trace'}`
      });
    } catch (e) {
      // Silent fail on email
    }
  }
}
```

### Step 3: Save

1. Click the **Save icon** (💾) or press `Ctrl+S` (Windows) / `Cmd+S` (Mac)
2. You may be prompted to name the project - call it "Quick Sync"

### Step 4: Close and Reopen

1. Close the Apps Script editor tab
2. Close your dashboard sheet
3. Reopen your dashboard sheet
4. You should now see a new menu: **⚡ Quick Sync**

---

## ✅ Testing

1. Add some notes in your **📊 Pipeline Review** tab (Column G)
2. Click: **⚡ Quick Sync → 📤 Push My Notes to Director**
3. You should see a toast: "✅ Pushed notes for X deal(s)..."
4. Check your director's consolidated sheet to verify notes appeared

---

## 🐛 Troubleshooting

### "Pipeline Review tab not found"
- Make sure your sheet has a tab named exactly: `📊 Pipeline Review`

### "Could not find your team assignment"
- Contact admin - your team needs to be configured in the Control Sheet

### "No director configured for X team"
- Contact admin - a director needs to be assigned to your team

### Menu doesn't appear after reopening
1. Open **Extensions → Apps Script**
2. Verify the code is there
3. Try clicking **Run → onOpen** manually
4. Close and reopen the sheet again

---

## 📧 Support

If you encounter any issues, email: **konstantin.gevorkov@tripleten.com**

The system automatically sends error reports to the admin, but feel free to reach out directly.

