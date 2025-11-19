/**
 * Setup.js
 * One-time setup script to create the Control Sheet
 */

function createControlSheet() {
  const SHEET_NAME = 'All-In-One 2.0 MVP';
  
  Logger.log('🚀 Creating Control Sheet...');
  
  // Create new spreadsheet
  const ss = SpreadsheetApp.create(SHEET_NAME);
  const sheetId = ss.getId();
  const sheetUrl = ss.getUrl();
  
  Logger.log(`✅ Created: ${SHEET_NAME}`);
  Logger.log(`   ID: ${sheetId}`);
  Logger.log(`   URL: ${sheetUrl}`);
  
  // Create all tabs FIRST
  setupSalespeopleConfig(ss);
  setupGoalsAndQuotas(ss);
  setupTechAccess(ss);
  setupInstructions(ss);
  
  // Then delete default Sheet1
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    ss.deleteSheet(defaultSheet);
  }
  
  Logger.log('');
  Logger.log('🎉 SETUP COMPLETE!');
  Logger.log('');
  Logger.log('📋 NEXT STEPS:');
  Logger.log('1. Open your Control Sheet:');
  Logger.log(`   ${sheetUrl}`);
  Logger.log('');
  Logger.log('2. Add your salespeople to "👥 Salespeople Config" tab');
  Logger.log('');
  Logger.log('3. Set HubSpot token in Script Properties:');
  Logger.log('   - Go to Project Settings (gear icon)');
  Logger.log('   - Add Script Property: HUBSPOT_ACCESS_TOKEN');
  Logger.log('');
  Logger.log('4. Run testSingleSalesperson() to test');
  Logger.log('');
  
  return sheetUrl;
}

function setupSalespeopleConfig(ss) {
  const sheet = ss.insertSheet('👥 Salespeople Config');
  
  // Headers
  const headers = [['Name', 'Email', 'Sheet ID', 'Sheet URL']];
  sheet.getRange('A1:D1').setValues(headers)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // Sample data (commented out row)
  const sampleData = [
    ['Example: John Doe', 'john.doe@example.com', '', '']
  ];
  sheet.getRange('A2:D2').setValues(sampleData)
    .setFontColor('#999999')
    .setFontStyle('italic');
  
  // Format columns
  sheet.setColumnWidth(1, 150); // Name
  sheet.setColumnWidth(2, 200); // Email
  sheet.setColumnWidth(3, 300); // Sheet ID
  sheet.setColumnWidth(4, 400); // Sheet URL
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  Logger.log('✅ Created: 👥 Salespeople Config');
}

function setupGoalsAndQuotas(ss) {
  const sheet = ss.insertSheet('🎯 Goals & Quotas');
  
  // Headers
  const headers = [['Email', 'Monthly Goal']];
  sheet.getRange('A1:B1').setValues(headers)
    .setFontWeight('bold')
    .setBackground('#34a853')
    .setFontColor('#ffffff');
  
  // Sample data
  const sampleData = [
    ['john.doe@example.com', 15]
  ];
  sheet.getRange('A2:B2').setValues(sampleData)
    .setFontColor('#999999')
    .setFontStyle('italic');
  
  // Format columns
  sheet.setColumnWidth(1, 200); // Email
  sheet.setColumnWidth(2, 120); // Monthly Goal
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  Logger.log('✅ Created: 🎯 Goals & Quotas');
}

function setupTechAccess(ss) {
  const sheet = ss.insertSheet('🔧 Tech Access');
  
  // Headers
  const headers = [['Email']];
  sheet.getRange('A1').setValues(headers)
    .setFontWeight('bold')
    .setBackground('#ea4335')
    .setFontColor('#ffffff');
  
  // Sample data
  const sampleData = [
    ['tech.admin@example.com']
  ];
  sheet.getRange('A2').setValues(sampleData)
    .setFontColor('#999999')
    .setFontStyle('italic');
  
  // Format columns
  sheet.setColumnWidth(1, 200);
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  Logger.log('✅ Created: 🔧 Tech Access');
}

function setupInstructions(ss) {
  const sheet = ss.insertSheet('📖 Instructions');
  
  const instructions = [
    ['All-In-One Dashboard 2.0 MVP - Control Sheet'],
    [''],
    ['PURPOSE:'],
    ['This Control Sheet manages all individual salesperson dashboards.'],
    [''],
    ['HOW TO USE:'],
    [''],
    ['1️⃣ Add Salespeople'],
    ['   Go to "👥 Salespeople Config" tab'],
    ['   Add Name and Email for each salesperson'],
    ['   Leave Sheet ID and Sheet URL blank (auto-populated)'],
    [''],
    ['2️⃣ Set Goals'],
    ['   Go to "🎯 Goals & Quotas" tab'],
    ['   Match Email to salesperson and set Monthly Goal'],
    [''],
    ['3️⃣ Add Tech Access'],
    ['   Go to "🔧 Tech Access" tab'],
    ['   Add emails of people who should have access to all sheets'],
    [''],
    ['4️⃣ Run Refresh'],
    ['   From menu: All in One Dashboard → 🔄 Refresh All Dashboards'],
    ['   Or run main() function from Apps Script'],
    [''],
    ['WHAT HAPPENS:'],
    ['• Individual sheets are auto-created for each salesperson'],
    ['• Each sheet has 4 tabs: Pipeline Review, Bonus Calculation, Enrollment Tracker, Operational Metrics'],
    ['• Sheets are automatically shared with the salesperson + tech access team'],
    ['• Sheet IDs and URLs are saved back to the Salespeople Config tab'],
    [''],
    ['TROUBLESHOOTING:'],
    ['• Check Execution Log in Apps Script (View → Logs)'],
    ['• Verify HubSpot token is set in Script Properties'],
    ['• Make sure emails are valid and have Google accounts'],
    [''],
    ['GitHub: https://github.com/SalesOpsNinjaTT/allinonemvp']
  ];
  
  sheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
  
  // Format title
  sheet.getRange('A1').setFontSize(14).setFontWeight('bold').setBackground('#fbbc04');
  
  // Format section headers
  sheet.getRange('A3').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A6').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A21').setFontWeight('bold').setFontSize(11);
  sheet.getRange('A29').setFontWeight('bold').setFontSize(11);
  
  // Set column width
  sheet.setColumnWidth(1, 600);
  
  // Set as active sheet
  ss.setActiveSheet(sheet);
  
  Logger.log('✅ Created: 📖 Instructions');
}

/**
 * Test function - creates a test salesperson sheet
 */
function testSingleSalesperson() {
  Logger.log('=== Testing Self-Provisioning ===');
  Logger.log('');
  
  const configManager = new ConfigManager();
  const sheetProvisioner = new SheetProvisioner(configManager);
  
  const salespeople = configManager.getSalespeopleConfig();
  
  if (salespeople.length === 0) {
    Logger.log('❌ No salespeople found in config. Please add at least one person.');
    return;
  }
  
  const testPerson = salespeople[0];
  Logger.log(`Testing with: ${testPerson.name} (${testPerson.email})`);
  Logger.log('');
  
  try {
    const individualSheet = sheetProvisioner.getOrCreateIndividualSheet(testPerson);
    Logger.log('');
    Logger.log('✅ Test successful!');
    Logger.log(`   Sheet ID: ${individualSheet.getId()}`);
    Logger.log(`   Sheet URL: ${individualSheet.getUrl()}`);
    Logger.log('');
    Logger.log('Check your Google Drive for the new sheet!');
  } catch (e) {
    Logger.log('');
    Logger.log('❌ Test failed!');
    Logger.log(`   Error: ${e.message}`);
    Logger.log(`   Stack: ${e.stack}`);
  }
}

