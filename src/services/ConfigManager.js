/**
 * Configuration Manager
 * Reads configuration from Control Sheet
 */

// Control Sheet Tab Names
const TAB_CONFIG = '👥 Salespeople Config';
const TAB_GOALS = '🎯 Goals & Quotas';
const TAB_TECH = '🔧 Tech Access';
const TAB_SUMMARY = '📊 Summary Dashboard';

/**
 * Finds a tab by checking both emoji and non-emoji versions
 * @param {Spreadsheet} ss - Spreadsheet object
 * @param {string} emojiName - Tab name with emoji
 * @param {string} plainName - Tab name without emoji
 * @returns {Sheet|null}
 */
function findTab(ss, emojiName, plainName) {
  return ss.getSheetByName(emojiName) || ss.getSheetByName(plainName);
}

/**
 * Load configuration from Control Sheet
 * @returns {Object} Configuration object
 */
function loadConfiguration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Read Salespeople Config
  const configSheet = findTab(ss, TAB_CONFIG, 'Salespeople Config');
  if (!configSheet) {
    throw new Error(`Config sheet not found. Expected: ${TAB_CONFIG} or 'Salespeople Config'`);
  }
  
  const salespeople = configSheet.getRange('A2:D' + configSheet.getLastRow())
    .getValues()
    .filter(row => row[0] && row[1]) // Name and Email required
    .map(row => ({
      name: row[0],
      email: row[1],
      sheetId: row[2] || '', // May be empty for new people
      sheetUrl: row[3] || ''
    }));
  
  // Read Goals
  const goalsSheet = findTab(ss, TAB_GOALS, 'Goals & Quotas');
  const goalsMap = new Map();
  if (goalsSheet) {
    const goalsData = goalsSheet.getRange('A2:B' + goalsSheet.getLastRow()).getValues();
    goalsData.forEach(row => {
      if (row[0]) goalsMap.set(row[0], row[1] || 0);
    });
  }
  
  // Read Tech Access
  const techAccessSheet = findTab(ss, TAB_TECH, 'tech access');
  const techAccessEmails = [];
  if (techAccessSheet) {
    const emails = techAccessSheet.getRange('B2:B' + techAccessSheet.getLastRow()).getValues();
    emails.forEach(row => {
      if (row[0]) techAccessEmails.push(row[0]);
    });
  }
  
  Logger.log(`Configuration loaded: ${salespeople.length} salespeople, ${goalsMap.size} goals, ${techAccessEmails.length} tech access emails`);
  
  return {
    salespeople,
    goalsMap,
    techAccessEmails,
    configSheet // Return for updating Sheet IDs
  };
}

/**
 * Get HubSpot API token from Script Properties
 * @returns {string} API token
 */
function getHubSpotToken() {
  const token = PropertiesService.getScriptProperties().getProperty('HUBSPOT_ACCESS_TOKEN');
  if (!token) {
    throw new Error('HubSpot API token not configured. Set HUBSPOT_ACCESS_TOKEN in Script Properties.');
  }
  return token;
}
