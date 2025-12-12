/**
 * Sheet Provisioner
 * Handles self-provisioning of individual salesperson sheets
 * 
 * Note: TAB constants are defined in Constants.js
 */

// Removed: AE Quick Sync functions - no longer needed
// Directors now use "Pull Notes from Team" button in Control Sheet

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

