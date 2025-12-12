/**
 * Sheet Provisioner
 * Handles self-provisioning of individual salesperson sheets
 * 
 * Note: TAB constants are defined in Constants.js
 */

/**
 * Creates custom menu for AE sheets
 * Called by onOpen trigger in individual AE sheets
 */
function onOpenAESheet() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('⚡ Quick Sync')
      .addItem('📤 Push My Notes to Director', 'pushMyNotesToDirector')
      .addToUi();
  } catch (error) {
    Logger.log(`[AE onOpen] Error: ${error.message}`);
  }
}

// Removed: installAEMenuTrigger() - triggers don't work across separate spreadsheets
// See installQuickSyncMenusOnAllAESheets() for new approach

/**
 * UTILITY: Install Quick Sync menus on all existing AE sheets
 * Run this once to add menus to sheets created before this feature
 * 
 * NOTE: This creates script files in each AE sheet and installs triggers
 */
function installQuickSyncMenusOnAllAESheets() {
  try {
    Logger.log('[Quick Sync Setup] Installing menus on all AE sheets...');
    Logger.log('This will create script files in each AE sheet (one-time setup)');
    
    const config = loadConfiguration();
    const { salespeople } = config;
    let successCount = 0;
    let errorCount = 0;
    
    salespeople.forEach(person => {
      if (!person.sheetId) {
        Logger.log(`  Skipping ${person.name} (no sheet ID)`);
        return;
      }
      
      try {
        Logger.log(`  Setting up ${person.name}...`);
        const sheet = SpreadsheetApp.openById(person.sheetId);
        
        // Create the script project with menu code
        copyQuickSyncCodeToAESheet(sheet, person);
        
        successCount++;
      } catch (error) {
        Logger.log(`  ❌ Error for ${person.name}: ${error.message}`);
        errorCount++;
      }
    });
    
    Logger.log(`\n[Quick Sync Setup] Complete: ${successCount} installed, ${errorCount} errors`);
    Logger.log('\n⚠️ IMPORTANT: AEs must close and reopen their sheets to see the menu.');
    Logger.log('📝 If menu still doesn\'t appear, AEs should:');
    Logger.log('   1. Open their sheet');
    Logger.log('   2. Extensions → Apps Script');
    Logger.log('   3. Verify "QuickSync.gs" file exists');
    Logger.log('   4. Close and reopen the sheet');
    
    return `Success! Installed on ${successCount} sheets. AEs should close and reopen.`;
    
  } catch (error) {
    Logger.log(`[Quick Sync Setup] Error: ${error.message}`);
    throw error;
  }
}

/**
 * Copies the Quick Sync code to an individual AE sheet's script project
 * This is required because AE sheets are separate spreadsheets with their own script projects
 * 
 * @param {Spreadsheet} spreadsheet - The AE's spreadsheet
 * @param {Object} person - Person config object
 */
function copyQuickSyncCodeToAESheet(spreadsheet, person) {
  try {
    // Unfortunately, we cannot programmatically create script files in other spreadsheets
    // via SpreadsheetApp. This requires the Apps Script API which has complex setup.
    
    // WORKAROUND: Create a helper sheet with instructions for manual setup
    Logger.log(`    ⚠️ Cannot programmatically add script to ${person.name}'s sheet`);
    Logger.log(`    📝 Manual setup required (see documentation)`);
    
    // Create or update an "Instructions" sheet with setup code
    let instructionsSheet = spreadsheet.getSheetByName('⚡ Quick Sync Setup');
    if (!instructionsSheet) {
      instructionsSheet = spreadsheet.insertSheet('⚡ Quick Sync Setup');
    }
    
    // Write instructions
    const instructions = [
      ['⚡ QUICK SYNC SETUP INSTRUCTIONS'],
      [''],
      ['To enable the "📤 Push My Notes to Director" button:'],
      [''],
      ['1. Click: Extensions → Apps Script'],
      ['2. Delete any existing code in the editor'],
      ['3. Copy and paste the code from the box below'],
      ['4. Click the Save icon (💾)'],
      ['5. Close the Apps Script tab'],
      ['6. Close and reopen this spreadsheet'],
      ['7. You should see a new menu: "⚡ Quick Sync"'],
      [''],
      ['============ COPY CODE BELOW ============'],
      [''],
      [`// Quick Sync Menu for ${person.name}`],
      ['// Auto-generated code - do not edit'],
      [''],
      ['// Control Sheet ID'],
      [`const CONTROL_SHEET_ID = '${CONTROL_SHEET_ID}';`],
      [''],
      ['// Tab names'],
      [`const TAB_PIPELINE = '${TAB_PIPELINE}';`],
      [''],
      ['// Create menu on open'],
      ['function onOpen() {'],
      ['  SpreadsheetApp.getUi()'],
      ['    .createMenu("⚡ Quick Sync")'],
      ['    .addItem("📤 Push My Notes to Director", "pushMyNotesToDirector")'],
      ['    .addToUi();'],
      ['}'],
      [''],
      ['// Push notes function'],
      ['function pushMyNotesToDirector() {'],
      ['  // Call the main Control Sheet script via URL'],
      ['  const url = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";'],
      ['  // TODO: This requires Web App deployment'],
      ['  SpreadsheetApp.getActiveSpreadsheet().toast("Feature requires admin setup", "Not Ready", 5);'],
      ['}']
    ];
    
    instructionsSheet.clear();
    instructionsSheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
    instructionsSheet.getRange('A1').setFontWeight('bold').setFontSize(14);
    instructionsSheet.getRange('A3:A9').setWrap(true);
    instructionsSheet.setColumnWidth(1, 800);
    
    Logger.log(`    Created setup instructions in "⚡ Quick Sync Setup" tab`);
    
  } catch (error) {
    Logger.log(`    Error creating instructions: ${error.message}`);
    throw error;
  }
}

/**
 * Get or create individual sheet for a salesperson
 * @param {Object} person - Person object {name, email, sheetId}
 * @param {Array} techAccessEmails - Automation account emails
 * @param {Sheet} configSheet - Config sheet to update with new Sheet ID
 * @param {number} rowIndex - Row index in config sheet (for updating)
 * @returns {Spreadsheet} Individual spreadsheet
 */
function getOrCreatePersonSheet(person, techAccessEmails, configSheet, rowIndex) {
  let sheetId = person.sheetId;
  
  // Self-provisioning: Create new sheet if ID is missing
  if (!sheetId || sheetId === '') {
    Logger.log(`[SELF-PROVISION] Creating new sheet for ${person.name}...`);
    
    const newSpreadsheet = SpreadsheetApp.create(`${person.name} - Dashboard`);
    sheetId = newSpreadsheet.getId();
    const sheetUrl = newSpreadsheet.getUrl();
    
    // Initialize with 4 tabs
    initializeIndividualSheet(newSpreadsheet);
    
    // Share with salesperson
    if (person.email) {
      try {
        newSpreadsheet.addEditor(person.email);
        Logger.log(`  Shared with: ${person.email}`);
      } catch (e) {
        Logger.log(`  Warning: Could not share with ${person.email}: ${e.message}`);
      }
    }
    
    // Share with tech access accounts
    techAccessEmails.forEach(email => {
      try {
        newSpreadsheet.addEditor(email);
        Logger.log(`  Shared with tech access: ${email}`);
      } catch (e) {
        Logger.log(`  Warning: Could not share with ${email}: ${e.message}`);
      }
    });
    
    // Update Control Sheet with new ID and URL
    if (configSheet && rowIndex) {
      configSheet.getRange(rowIndex, 3).setValue(sheetId);  // Column C: Sheet ID
      configSheet.getRange(rowIndex, 4).setValue(sheetUrl); // Column D: Sheet URL
      Logger.log(`  Updated Control Sheet with Sheet ID: ${sheetId}`);
    }
    
    Logger.log(`[SELF-PROVISION] ✅ Created sheet for ${person.name}`);
    return newSpreadsheet;
  }
  
  // Open existing sheet
  try {
    return SpreadsheetApp.openById(sheetId);
  } catch (e) {
    throw new Error(`Could not open sheet for ${person.name} (ID: ${sheetId}): ${e.message}`);
  }
}

/**
 * Initialize individual sheet with 4 tabs
 * @param {Spreadsheet} ss - Spreadsheet to initialize
 */
function initializeIndividualSheet(ss) {
  // Delete default Sheet1 if it exists
  const defaultSheet = ss.getSheetByName('Sheet1');
  
  // Create 4 tabs
  const pipelineTab = ss.insertSheet(TAB_PIPELINE);
  const bonusesTab = ss.insertSheet(TAB_BONUS);
  const enrollmentTab = ss.insertSheet(TAB_ENROLLMENT);
  const opsTab = ss.insertSheet(TAB_OPERATIONAL);
  
  // Set up Pipeline Review tab
  pipelineTab.getRange('A1').setValue('Pipeline Review - Will be populated by script');
  pipelineTab.setFrozenRows(1);
  
  // Set up Bonuses tab
  bonusesTab.getRange('A1').setValue('Bonus Calculation - Will be populated by script');
  bonusesTab.setFrozenRows(1);
  
  // Set up Enrollment tab
  enrollmentTab.getRange('A1').setValue('Enrollment Tracker - Will be populated by script');
  enrollmentTab.setFrozenRows(1);
  
  // Set up Ops Metrics tab
  opsTab.getRange('A1').setValue('Operational Metrics - Will be populated by script');
  opsTab.setFrozenRows(1);
  
  // Delete Sheet1 after creating new tabs
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
  
  // Make Pipeline Review the active tab
  ss.setActiveSheet(pipelineTab);
  
  Logger.log('  Initialized 4 tabs: Pipeline Review, Bonus Calculation, Enrollment Tracker, Operational Metrics');
}

