/**
 * Bonus Calculation Component
 * 
 * Copies the "📈 Monthly Bonus Report" tab from legacy Clean 2.2 bonus sheets
 * into the new All-in-One Dashboard
 * 
 * Strategy: Don't re-implement bonus calculation - just copy the already-calculated data
 * from the legacy system into the new unified dashboard.
 * 
 * Note: TAB_BONUSES is defined in SheetProvisioner.js
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

// Legacy system Control Sheet ID
const LEGACY_CONTROL_SHEET_ID = '1OTyjtQ27iCXhqgscNomEktNQg2qlS6Uk0_j6EeeyEFI';

// Legacy tab names (from Clean 2.2)
const LEGACY_TAB_BONUS_REPORT = '📈 Monthly Bonus Report';
const LEGACY_TAB_DASHBOARD = '📊 Dashboard';

// Legacy Control Sheet column mapping (1-based indices)
// Adjust these if the legacy sheet has a different structure
const LEGACY_COLUMNS = {
  NAME: 1,        // Column A: Salesperson Name
  EMAIL: 2,       // Column B: auth_email
  SHEET_ID: 3,    // Column C: Spreadsheet ID
  SHEET_URL: 4    // Column D: Spreadsheet URL (optional)
};

/**
 * Updates the Bonus Calculation tab for a salesperson by copying from legacy system
 * @param {Spreadsheet} individualSheet - The salesperson's new All-in-One Dashboard
 * @param {Object} person - Person object {name, email}
 * @returns {Object} Update result
 */
function updateBonusCalculation(individualSheet, person) {
  try {
    Logger.log(`[Bonus Calculation] Updating for ${person.name}...`);
    
    const startTime = new Date();
    const sheet = individualSheet.getSheetByName(TAB_BONUSES);
    
    if (!sheet) {
      throw new Error(`Bonus Calculation tab not found for ${person.name}`);
    }
    
    // Step 1: Look up legacy bonus sheet ID
    Logger.log('  Step 1: Looking up legacy bonus sheet ID...');
    const legacySheetId = getLegacyBonusSheetId(person);
    
    if (!legacySheetId) {
      Logger.log(`  No legacy bonus sheet found for ${person.name}, skipping`);
      sheet.clear();
      sheet.getRange('A1').setValue(`No legacy bonus data found for ${person.name}`)
        .setFontWeight('bold')
        .setFontColor('#cc0000');
      return {
        success: false,
        error: 'No legacy sheet ID found'
      };
    }
    
    Logger.log(`  Found legacy sheet ID: ${legacySheetId}`);
    
    // Step 2: Open legacy sheet and find Monthly Bonus Report tab
    Logger.log('  Step 2: Opening legacy bonus sheet...');
    const legacySheet = SpreadsheetApp.openById(legacySheetId);
    const legacyBonusTab = legacySheet.getSheetByName(LEGACY_TAB_BONUS_REPORT) || 
                           legacySheet.getSheetByName('Monthly Bonus Report');
    
    if (!legacyBonusTab) {
      Logger.log(`  No "Monthly Bonus Report" tab found in legacy sheet for ${person.name}`);
      sheet.clear();
      sheet.getRange('A1').setValue(`Legacy bonus sheet found, but no "Monthly Bonus Report" tab`)
        .setFontWeight('bold')
        .setFontColor('#ff9900');
      return {
        success: false,
        error: 'No Monthly Bonus Report tab in legacy sheet'
      };
    }
    
    // Step 3: Copy all data from legacy tab
    Logger.log('  Step 3: Copying data from legacy sheet...');
    const lastRow = legacyBonusTab.getLastRow();
    const lastCol = legacyBonusTab.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      Logger.log(`  Legacy bonus tab is empty for ${person.name}`);
      sheet.clear();
      sheet.getRange('A1').setValue(`Legacy bonus data is empty`)
        .setFontWeight('bold')
        .setFontColor('#ff9900');
      return {
        success: false,
        error: 'Legacy bonus tab is empty'
      };
    }
    
    // Copy values
    const sourceRange = legacyBonusTab.getRange(1, 1, lastRow, lastCol);
    const values = sourceRange.getValues();
    
    // Copy formatting
    const backgrounds = sourceRange.getBackgrounds();
    const fontColors = sourceRange.getFontColors();
    const fontWeights = sourceRange.getFontWeights();
    const fontSizes = sourceRange.getFontSizes();
    const numberFormats = sourceRange.getNumberFormats();
    const horizontalAlignments = sourceRange.getHorizontalAlignments();
    
    // Step 4: Clear and write to new sheet
    Logger.log('  Step 4: Writing to new dashboard...');
    sheet.clear();
    
    const targetRange = sheet.getRange(1, 1, lastRow, lastCol);
    targetRange.setValues(values);
    
    // Apply formatting
    targetRange.setBackgrounds(backgrounds);
    targetRange.setFontColors(fontColors);
    targetRange.setFontWeights(fontWeights);
    targetRange.setFontSizes(fontSizes);
    targetRange.setNumberFormats(numberFormats);
    targetRange.setHorizontalAlignments(horizontalAlignments);
    
    // Step 5: Auto-resize columns
    Logger.log('  Step 5: Auto-resizing columns...');
    for (let col = 1; col <= lastCol; col++) {
      sheet.autoResizeColumn(col);
    }
    
    // Step 6: Check for frozen rows and apply
    const frozenRows = legacyBonusTab.getFrozenRows();
    if (frozenRows > 0) {
      sheet.setFrozenRows(frozenRows);
    }
    
    const duration = (new Date() - startTime) / 1000;
    Logger.log(`[Bonus Calculation] Complete for ${person.name} (${duration}s)`);
    
    return {
      success: true,
      rowCount: lastRow,
      colCount: lastCol,
      duration: duration
    };
    
  } catch (error) {
    Logger.log(`[Bonus Calculation] Error for ${person.name}: ${error.message}`);
    
    // Write error message to sheet
    try {
      const sheet = individualSheet.getSheetByName(TAB_BONUSES);
      if (sheet) {
        sheet.clear();
        sheet.getRange('A1').setValue(`Error loading bonus data: ${error.message}`)
          .setFontWeight('bold')
          .setFontColor('#cc0000');
      }
    } catch (e) {
      // Ignore
    }
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Looks up the legacy bonus sheet ID for a person
 * @param {Object} person - Person object {name, email}
 * @returns {string|null} Legacy sheet ID or null if not found
 */
function getLegacyBonusSheetId(person) {
  try {
    // Open legacy Control Sheet
    const legacyControlSheet = SpreadsheetApp.openById(LEGACY_CONTROL_SHEET_ID);
    const configSheet = legacyControlSheet.getSheetByName('👥 Salespeople Config') ||
                        legacyControlSheet.getSheetByName('Salespeople Config');
    
    if (!configSheet) {
      Logger.log(`  WARNING: Could not find Salespeople Config tab in legacy Control Sheet`);
      return null;
    }
    
    const lastRow = configSheet.getLastRow();
    if (lastRow < 2) {
      return null;
    }
    
    // Read all config data using configured column mapping
    const maxCol = Math.max(LEGACY_COLUMNS.NAME, LEGACY_COLUMNS.EMAIL, LEGACY_COLUMNS.SHEET_ID);
    const data = configSheet.getRange(2, 1, lastRow - 1, maxCol).getValues();
    
    // Try to match by email first (more reliable), then by name
    for (let i = 0; i < data.length; i++) {
      const legacyName = data[i][LEGACY_COLUMNS.NAME - 1];
      const legacyEmail = data[i][LEGACY_COLUMNS.EMAIL - 1];
      const legacySheetId = data[i][LEGACY_COLUMNS.SHEET_ID - 1];
      
      // Match by email (case-insensitive)
      if (legacyEmail && person.email && 
          legacyEmail.toString().toLowerCase().trim() === person.email.toString().toLowerCase().trim()) {
        return legacySheetId ? legacySheetId.toString() : null;
      }
      
      // Fallback: Match by name (case-insensitive)
      if (legacyName && person.name &&
          legacyName.toString().toLowerCase().trim() === person.name.toString().toLowerCase().trim()) {
        return legacySheetId ? legacySheetId.toString() : null;
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log(`  ERROR looking up legacy sheet ID for ${person.name}: ${error.message}`);
    return null;
  }
}

/**
 * Optional: Also copies the Dashboard summary tab from legacy system
 * @param {Spreadsheet} individualSheet - The salesperson's new All-in-One Dashboard
 * @param {Object} person - Person object {name, email}
 * @returns {Object} Update result
 */
function copyLegacyDashboard(individualSheet, person) {
  try {
    Logger.log(`[Legacy Dashboard] Copying for ${person.name}...`);
    
    // Look up legacy sheet ID
    const legacySheetId = getLegacyBonusSheetId(person);
    if (!legacySheetId) {
      return { success: false, error: 'No legacy sheet ID found' };
    }
    
    // Open legacy sheet
    const legacySheet = SpreadsheetApp.openById(legacySheetId);
    const legacyDashboardTab = legacySheet.getSheetByName(LEGACY_TAB_DASHBOARD) ||
                                legacySheet.getSheetByName('Dashboard');
    
    if (!legacyDashboardTab) {
      return { success: false, error: 'No Dashboard tab in legacy sheet' };
    }
    
    // Create new tab for legacy dashboard (optional - you can skip this if not needed)
    const newTabName = '📊 Legacy Dashboard';
    let targetTab = individualSheet.getSheetByName(newTabName);
    
    if (!targetTab) {
      targetTab = individualSheet.insertSheet(newTabName);
    } else {
      targetTab.clear();
    }
    
    // Copy data and formatting
    const lastRow = legacyDashboardTab.getLastRow();
    const lastCol = legacyDashboardTab.getLastColumn();
    
    if (lastRow > 0 && lastCol > 0) {
      const sourceRange = legacyDashboardTab.getRange(1, 1, lastRow, lastCol);
      const targetRange = targetTab.getRange(1, 1, lastRow, lastCol);
      
      targetRange.setValues(sourceRange.getValues());
      targetRange.setBackgrounds(sourceRange.getBackgrounds());
      targetRange.setFontColors(sourceRange.getFontColors());
      targetRange.setFontWeights(sourceRange.getFontWeights());
      targetRange.setFontSizes(sourceRange.getFontSizes());
      targetRange.setNumberFormats(sourceRange.getNumberFormats());
      
      // Auto-resize
      for (let col = 1; col <= lastCol; col++) {
        targetTab.autoResizeColumn(col);
      }
    }
    
    Logger.log(`[Legacy Dashboard] Complete for ${person.name}`);
    return { success: true };
    
  } catch (error) {
    Logger.log(`[Legacy Dashboard] Error for ${person.name}: ${error.message}`);
    return { success: false, error: error.message };
  }
}

