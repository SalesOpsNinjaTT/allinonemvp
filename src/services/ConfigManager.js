/**
 * Configuration Manager
 * Reads configuration from Control Sheet
 */

// Control Sheet ID - All-In-One 2.0 MVP
const CONTROL_SHEET_ID = '1-zipx1vWfjYaMjgl7BbqfCVjl8NZch9DMk5T-DRfnnQ';

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
  const ss = SpreadsheetApp.openById(CONTROL_SHEET_ID);
  
  // Read Salespeople Config
  const configSheet = findTab(ss, TAB_CONFIG, 'Salespeople Config');
  if (!configSheet) {
    throw new Error(`Config sheet not found. Expected: ${TAB_CONFIG} or 'Salespeople Config'`);
  }
  
  const salespeople = configSheet.getRange('A2:F' + configSheet.getLastRow())
    .getValues()
    .filter(row => row[0] && row[1]) // Name and Email required
    .map(row => ({
      name: row[0],
      email: row[1],
      sheetId: row[2] || '', // May be empty for new people
      sheetUrl: row[3] || '',
      hubspotUserId: row[4] ? row[4].toString() : '', // HubSpot User ID
      role: row[5] || '' // Role (AE, SrAE, CAE, etc.)
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

/**
 * Load tier levels configuration from Control Sheet
 * @returns {Map} Map of role name to tier levels array
 */
function getTierLevels() {
  const ss = SpreadsheetApp.openById(CONTROL_SHEET_ID);
  
  // Try to find Tier Levels sheet
  const tierSheet = ss.getSheetByName('🎯 Tier Levels') || ss.getSheetByName('Tier Levels');
  
  if (!tierSheet) {
    Logger.log('Warning: Tier Levels sheet not found, returning empty map');
    return new Map();
  }
  
  const lastRow = tierSheet.getLastRow();
  if (lastRow < 2) {
    return new Map();
  }
  
  // Read tier data: Tier | Min | Max | Percent | Role | Alt. Role Name
  const data = tierSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  // Group by role
  const tiersByRole = new Map();
  
  data.forEach(row => {
    const tier = row[0];
    const min = row[1];
    const max = row[2];
    const percent = row[3];
    const role = row[4]; // Primary role name
    const altRole = row[5]; // Alternative role name
    
    if (!role) return; // Skip rows without role
    
    const tierInfo = {
      tier: tier,
      min: min,
      max: max === 100500 ? Infinity : max, // Handle placeholder
      percent: percent,
      displayMax: max === 100500 ? '35+' : max // For display purposes
    };
    
    // Add to primary role name
    if (!tiersByRole.has(role)) {
      tiersByRole.set(role, []);
    }
    tiersByRole.get(role).push(tierInfo);
    
    // Also add to alternative role name if it exists
    if (altRole && altRole !== '') {
      if (!tiersByRole.has(altRole)) {
        tiersByRole.set(altRole, []);
      }
      tiersByRole.get(altRole).push(tierInfo);
    }
  });
  
  // Sort tiers by tier number for each role
  tiersByRole.forEach((tiers, role) => {
    tiers.sort((a, b) => a.tier - b.tier);
  });
  
  Logger.log(`Loaded tier levels for ${tiersByRole.size} roles`);
  
  return tiersByRole;
}
