/**
 * All in One Dashboard - Main Entry Point
 * 
 * Phase 1: Setup & Foundation
 * - Control Sheet configuration loading
 * - Self-provisioning of individual sheets
 * - Ready for component integration
 * 
 * Note: CONTROL_SHEET_ID is defined in ConfigManager.js
 */

/**
 * Main function to generate/update all salesperson dashboards
 * Run this from Control Sheet
 */
function generateAllDashboards() {
  const startTime = new Date();
  Logger.log('=== All in One Dashboard Generation ===');
  Logger.log(`Start time: ${startTime.toISOString()}`);
  
  try {
    // 1. Load configuration from Control Sheet
    const config = loadConfiguration();
    const { salespeople, goalsMap, techAccessEmails, configSheet } = config;
    
    Logger.log(`\nProcessing ${salespeople.length} salespeople...`);
    
    // 1.5. Update Director Hub (team-wide view)
    Logger.log(`\n=== Updating Director Hub ===`);
    const controlSheet = SpreadsheetApp.openById(CONTROL_SHEET_ID);
    const hubResult = updateDirectorHub(controlSheet, salespeople);
    if (hubResult.success) {
      Logger.log(`✅ Director Hub: ${hubResult.dealCount} deals (${hubResult.duration}s)`);
    } else {
      Logger.log(`❌ Director Hub failed: ${hubResult.error}`);
    }
    
    let processedCount = 0;
    let errorCount = 0;
    
    // 2. For each salesperson
    salespeople.forEach((person, index) => {
      Logger.log(`\n--- Processing ${person.name} (${person.email}) ---`);
      
      try {
        // 3. Get or create their individual sheet (self-provisioning)
        const rowIndex = index + 2; // +2 because row 1 is header, index starts at 0
        const sheet = getOrCreatePersonSheet(person, techAccessEmails, configSheet, rowIndex);
        
        Logger.log(`  Sheet ID: ${sheet.getId()}`);
        Logger.log(`  Sheet URL: ${sheet.getUrl()}`);
        
        // Update Pipeline Review tab
        const pipelineResult = updatePipelineReview(sheet, person);
        if (pipelineResult.success) {
          Logger.log(`  ✅ Pipeline Review: ${pipelineResult.dealCount} deals`);
        } else {
          Logger.log(`  ❌ Pipeline Review failed: ${pipelineResult.error}`);
        }
        
        // Update Enrollment Tracker tab
        const enrollmentResult = updateEnrollmentTracker(sheet, person);
        if (enrollmentResult.success) {
          Logger.log(`  ✅ Enrollment Tracker: ${enrollmentResult.enrollmentCount} enrollments`);
        } else {
          Logger.log(`  ❌ Enrollment Tracker failed: ${enrollmentResult.error}`);
        }
        
        // Update Bonus Calculation tab (copies from legacy Clean 2.2 system)
        const bonusResult = updateBonusCalculation(sheet, person);
        if (bonusResult.success) {
          Logger.log(`  ✅ Bonus Calculation: ${bonusResult.rowCount} rows copied from legacy system`);
        } else {
          Logger.log(`  ❌ Bonus Calculation failed: ${bonusResult.error}`);
        }
        
        // TODO: Phase 4 - Update Operational Metrics tab
        
        Logger.log(`✅ ${person.name} - SUCCESS`);
        processedCount++;
        
      } catch (e) {
        errorCount++;
        Logger.log(`❌ ${person.name} - ERROR: ${e.message}`);
        Logger.log(`   Stack: ${e.stack}`);
      }
    });
    
    // 3. Update all Director Consolidated Pipelines (bi-directional sync)
    Logger.log(`\n=== Updating Director Consolidated Pipelines ===`);
    const directorResult = updateAllDirectorConsolidatedPipelines();
    if (directorResult.success) {
      Logger.log(`✅ Director Pipelines: ${directorResult.directorCount} directors updated`);
    } else {
      Logger.log(`❌ Director Pipelines failed: ${directorResult.error}`);
    }
    
    // 4. Sync director flags from Director Hub to AE sheets
    Logger.log(`\n=== Syncing Director Flags to AE Sheets ===`);
    syncDirectorFlagsToAESheets(controlSheet, salespeople);
    
    // 5. Summary
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    const message = `Complete: ${processedCount} successful, ${errorCount} errors (${duration}s)`;
    
    Logger.log(`\n=== ${message} ===`);
    
    // Show toast if running from a spreadsheet context
    try {
      const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
      if (activeSheet) {
        activeSheet.toast(message, 'Dashboard Generation', 10);
      }
    } catch (e) {
      // No active spreadsheet, that's okay
    }
    
  } catch (error) {
    Logger.log(`FATAL ERROR: ${error.message}`);
    Logger.log(error.stack);
    
    // Show toast if running from a spreadsheet context
    try {
      const activeSheet = SpreadsheetApp.getActiveSpreadsheet();
      if (activeSheet) {
        activeSheet.toast(`Error: ${error.message}`, 'Error', 10);
      }
    } catch (e) {
      // No active spreadsheet, that's okay
    }
    
    throw error;
  }
}

/**
 * Test function for development
 * Run this to test with a single salesperson
 */
function testSingleSalesperson() {
  Logger.log('=== Testing with single salesperson ===');
  
  try {
    const config = loadConfiguration();
    const firstPerson = config.salespeople[0];
    
    if (!firstPerson) {
      throw new Error('No salespeople found in Salespeople Config');
    }
    
    Logger.log(`Testing with: ${firstPerson.name} (${firstPerson.email})`);
    
    const sheet = getOrCreatePersonSheet(
      firstPerson, 
      config.techAccessEmails, 
      config.configSheet, 
      2 // Row 2 in config sheet
    );
    
    Logger.log(`✅ Test successful`);
    Logger.log(`   Sheet ID: ${sheet.getId()}`);
    Logger.log(`   Sheet URL: ${sheet.getUrl()}`);
    
  } catch (e) {
    Logger.log(`❌ Test failed: ${e.message}`);
    Logger.log(e.stack);
  }
}

/**
 * Test function for bonus calculation copy
 * Run this to test copying bonus data for a single person
 */
function testBonusCopy() {
  Logger.log('=== Testing Bonus Calculation Copy ===');
  
  try {
    const config = loadConfiguration();
    const firstPerson = config.salespeople[0];
    
    if (!firstPerson) {
      throw new Error('No salespeople found in Salespeople Config');
    }
    
    Logger.log(`Testing bonus copy for: ${firstPerson.name} (${firstPerson.email})`);
    
    // Get their sheet
    if (!firstPerson.sheetId) {
      throw new Error(`No sheet ID found for ${firstPerson.name}. Run testSingleSalesperson() first.`);
    }
    
    const sheet = SpreadsheetApp.openById(firstPerson.sheetId);
    
    // Test bonus calculation copy
    const result = updateBonusCalculation(sheet, firstPerson);
    
    if (result.success) {
      Logger.log(`✅ Bonus copy successful`);
      Logger.log(`   Rows copied: ${result.rowCount}`);
      Logger.log(`   Columns copied: ${result.colCount}`);
      Logger.log(`   Duration: ${result.duration}s`);
    } else {
      Logger.log(`❌ Bonus copy failed: ${result.error}`);
    }
    
  } catch (e) {
    Logger.log(`❌ Test failed: ${e.message}`);
    Logger.log(e.stack);
  }
}

/**
 * Setup function - run once to initialize Control Sheet
 * Creates the necessary tabs if they don't exist
 */
function setupControlSheet() {
  Logger.log('=== Control Sheet Setup ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check if tabs exist
  const configSheet = findTab(ss, TAB_CONFIG, 'Salespeople Config');
  const goalsSheet = findTab(ss, TAB_GOALS, 'Goals & Quotas');
  const techSheet = findTab(ss, TAB_TECH, 'tech access');
  const summarySheet = findTab(ss, TAB_SUMMARY, 'Summary Dashboard');
  
  // Create missing tabs
  if (!configSheet) {
    Logger.log('Creating Salespeople Config tab...');
    const newConfig = ss.insertSheet(TAB_CONFIG);
    newConfig.getRange('A1:D1').setValues([['Name', 'Email', 'Sheet ID', 'Sheet URL']]);
    newConfig.getRange('A1:D1').setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    newConfig.setFrozenRows(1);
    Logger.log('✅ Created Salespeople Config tab');
  }
  
  if (!goalsSheet) {
    Logger.log('Creating Goals & Quotas tab...');
    const newGoals = ss.insertSheet(TAB_GOALS);
    newGoals.getRange('A1:B1').setValues([['Email', 'Monthly Goal']]);
    newGoals.getRange('A1:B1').setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    newGoals.setFrozenRows(1);
    Logger.log('✅ Created Goals & Quotas tab');
  }
  
  if (!techSheet) {
    Logger.log('Creating Tech Access tab...');
    const newTech = ss.insertSheet(TAB_TECH);
    newTech.getRange('A1:B1').setValues([['Purpose', 'Email']]);
    newTech.getRange('A1:B1').setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    newTech.setFrozenRows(1);
    Logger.log('✅ Created Tech Access tab');
  }
  
  if (!summarySheet) {
    Logger.log('Creating Summary Dashboard tab...');
    const newSummary = ss.insertSheet(TAB_SUMMARY);
    newSummary.getRange('A1').setValue('Summary Dashboard - Coming soon');
    Logger.log('✅ Created Summary Dashboard tab');
  }
  
  Logger.log('\n=== Setup Complete ===');
  Logger.log('Next steps:');
  Logger.log('1. Add salespeople to Salespeople Config tab');
  Logger.log('2. Add goals to Goals & Quotas tab');
  Logger.log('3. Add tech access emails to Tech Access tab');
  Logger.log('4. Set HUBSPOT_ACCESS_TOKEN in Script Properties');
  Logger.log('5. Run generateAllDashboards() or testSingleSalesperson()');
}
