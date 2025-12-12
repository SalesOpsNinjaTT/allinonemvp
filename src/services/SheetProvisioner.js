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

/**
 * Installs onOpen trigger for an individual AE sheet
 * @param {Spreadsheet} spreadsheet - The AE's spreadsheet
 */
function installAEMenuTrigger(spreadsheet) {
  try {
    // Delete any existing onOpen triggers for this spreadsheet
    const existingTriggers = ScriptApp.getUserTriggers(spreadsheet);
    existingTriggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'onOpenAESheet') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create new trigger
    ScriptApp.newTrigger('onOpenAESheet')
      .forSpreadsheet(spreadsheet)
      .onOpen()
      .create();
    
    Logger.log(`  Installed Quick Sync menu trigger for ${spreadsheet.getName()}`);
  } catch (error) {
    Logger.log(`  Warning: Could not install menu trigger: ${error.message}`);
  }
}

/**
 * UTILITY: Install Quick Sync menus on all existing AE sheets
 * Run this once to add menus to sheets created before this feature
 */
function installQuickSyncMenusOnAllAESheets() {
  try {
    Logger.log('[Quick Sync Setup] Installing menus on all AE sheets...');
    
    const salespeople = loadConfiguration();
    let successCount = 0;
    let errorCount = 0;
    
    salespeople.forEach(person => {
      if (!person.sheetId) {
        Logger.log(`  Skipping ${person.name} (no sheet ID)`);
        return;
      }
      
      try {
        const sheet = SpreadsheetApp.openById(person.sheetId);
        installAEMenuTrigger(sheet);
        successCount++;
      } catch (error) {
        Logger.log(`  ❌ Error for ${person.name}: ${error.message}`);
        errorCount++;
      }
    });
    
    Logger.log(`[Quick Sync Setup] Complete: ${successCount} installed, ${errorCount} errors`);
    Logger.log('AEs should close and reopen their sheets to see the new menu.');
    
    return `Success! Installed on ${successCount} sheets. AEs should close and reopen.`;
    
  } catch (error) {
    Logger.log(`[Quick Sync Setup] Error: ${error.message}`);
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
    
    // Install onOpen trigger for Quick Sync menu
    installAEMenuTrigger(newSpreadsheet);
    
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

