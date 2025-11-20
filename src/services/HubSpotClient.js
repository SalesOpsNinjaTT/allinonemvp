/**
 * HubSpot API Client
 * Handles all communication with HubSpot CRM API v3
 * 
 * Extracted and adapted from Call-Quality-Review project
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

const HUBSPOT_API_CONFIG = {
  BASE_URL: 'https://api.hubapi.com',
  ENDPOINTS: {
    DEALS_SEARCH: '/crm/v3/objects/deals/search',
    DEALS_BY_ID: '/crm/v3/objects/deals',
    PROPERTIES: '/crm/v3/properties/deals',
    OWNERS: '/crm/v3/owners'
  },
  DEFAULT_BATCH_SIZE: 100,
  MAX_RESULTS: 10000,
  MAX_RETRIES: 3
};

// ============================================================================
// API KEY MANAGEMENT
// ============================================================================

/**
 * Gets the HubSpot API access token from Script Properties
 * @returns {string} The HubSpot access token
 * @throws {Error} If token is not configured
 */
function getHubSpotAccessToken() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const token = scriptProperties.getProperty('HUBSPOT_ACCESS_TOKEN');
  
  if (!token) {
    throw new Error(
      'HubSpot API token not configured. ' +
      'Set HUBSPOT_ACCESS_TOKEN in Script Properties'
    );
  }
  
  return token;
}

// ============================================================================
// MAIN FETCH FUNCTIONS
// ============================================================================

/**
 * Fetches deals for a specific owner
 * @param {string} ownerEmailOrId - Email or User ID of the deal owner
 * @param {Array<string>} properties - Array of property names to fetch
 * @param {Object} options - Optional filters (dateRange, hubspotUserId, etc.)
 * @returns {Array<Object>} Array of deal objects
 */
function fetchDealsByOwner(ownerEmailOrId, properties, options = {}) {
  try {
    Logger.log(`Fetching deals for ${ownerEmailOrId}...`);
    
    let ownerId;
    
    // Check if we have a direct User ID in options
    if (options.hubspotUserId && options.hubspotUserId !== '') {
      ownerId = options.hubspotUserId.toString();
      Logger.log(`  Using provided User ID: ${ownerId}`);
    } else {
      // Fall back to email lookup (requires API permission)
      ownerId = getOwnerIdByEmail(ownerEmailOrId);
      if (!ownerId) {
        Logger.log(`  Warning: Owner not found for ${ownerEmailOrId}, returning 0 deals`);
        return [];
      }
      Logger.log(`  Looked up User ID: ${ownerId}`);
    }
    
    const accessToken = getHubSpotAccessToken();
    const allDeals = [];
    let after = null;
    let pageCount = 0;
    
    do {
      pageCount++;
      Logger.log(`  Page ${pageCount}...`);
      
      const response = fetchDealsPage(accessToken, properties, ownerId, after, options);
      
      if (response.results && response.results.length > 0) {
        allDeals.push(...response.results);
      }
      
      after = response.paging && response.paging.next ? response.paging.next.after : null;
      
      // Safety limit
      if (allDeals.length >= HUBSPOT_API_CONFIG.MAX_RESULTS) {
        Logger.log(`  Reached maximum limit of ${HUBSPOT_API_CONFIG.MAX_RESULTS}`);
        break;
      }
      
    } while (after);
    
    Logger.log(`  Total deals fetched: ${allDeals.length}`);
    
    // Apply client-side GTC filter (ability_to_pay >= 3 AND warm_prospects >= 3)
    const gtcDeals = allDeals.filter(deal => {
      const abilityToPay = extractNumericProperty(deal, 'ability_to_pay');
      const warmProspects = extractNumericProperty(deal, 'warm_prospects');
      
      // Both must be >= 3 (GTC = "Good To Close")
      // If value is empty/null, it's not GTC
      return (typeof abilityToPay === 'number' && abilityToPay >= 3) &&
             (typeof warmProspects === 'number' && warmProspects >= 3);
    });
    
    Logger.log(`  GTC-filtered deals: ${gtcDeals.length} (from ${allDeals.length} total)`);
    return gtcDeals;
    
  } catch (error) {
    Logger.log(`Error fetching deals for ${ownerEmailOrId}: ${error.message}`);
    throw error;
  }
}

/**
 * Fetches enrollment deals for a specific owner
 * @param {string} ownerEmailOrId - Email or User ID of the deal owner
 * @param {Array<string>} properties - Array of property names to fetch
 * @param {Object} options - Optional filters (hubspotUserId, months, etc.)
 * @returns {Array<Object>} Array of enrollment deal objects
 */
function fetchEnrollmentDeals(ownerEmailOrId, properties, options = {}) {
  try {
    Logger.log(`Fetching enrollment deals for ${ownerEmailOrId}...`);
    
    let ownerId;
    
    // Check if we have a direct User ID in options
    if (options.hubspotUserId && options.hubspotUserId !== '') {
      ownerId = options.hubspotUserId.toString();
      Logger.log(`  Using provided User ID: ${ownerId}`);
    } else {
      // Fall back to email lookup (requires API permission)
      ownerId = getOwnerIdByEmail(ownerEmailOrId);
      if (!ownerId) {
        Logger.log(`  Warning: Owner not found for ${ownerEmailOrId}, returning 0 deals`);
        return [];
      }
      Logger.log(`  Looked up User ID: ${ownerId}`);
    }
    
    const accessToken = getHubSpotAccessToken();
    const allDeals = [];
    let after = null;
    let pageCount = 0;
    
    do {
      pageCount++;
      Logger.log(`  Page ${pageCount}...`);
      
      const response = fetchEnrollmentDealsPage(accessToken, properties, ownerId, after, options);
      
      if (response.results && response.results.length > 0) {
        allDeals.push(...response.results);
      }
      
      after = response.paging && response.paging.next ? response.paging.next.after : null;
      
      // Safety limit
      if (allDeals.length >= HUBSPOT_API_CONFIG.MAX_RESULTS) {
        Logger.log(`  Reached maximum limit of ${HUBSPOT_API_CONFIG.MAX_RESULTS}`);
        break;
      }
      
    } while (after);
    
    Logger.log(`  Total enrollment deals fetched: ${allDeals.length}`);
    return allDeals;
    
  } catch (error) {
    Logger.log(`Error fetching enrollment deals for ${ownerEmailOrId}: ${error.message}`);
    throw error;
  }
}

/**
 * Fetches a single page of deals from HubSpot
 * @param {string} accessToken - HubSpot access token
 * @param {Array<string>} properties - Properties to fetch
 * @param {string} ownerId - Owner ID filter (numeric)
 * @param {string} after - Pagination cursor
 * @param {Object} options - Additional filter options
 * @returns {Object} API response object
 */
function fetchDealsPage(accessToken, properties, ownerId, after, options = {}) {
  const url = HUBSPOT_API_CONFIG.BASE_URL + HUBSPOT_API_CONFIG.ENDPOINTS.DEALS_SEARCH;
  
  // Build filters (all filters in one group = AND logic)
  const filters = [];
  
  // Filter 1: Owner ID
  if (ownerId) {
    filters.push({
      propertyName: 'hubspot_owner_id',
      operator: 'EQ',
      value: ownerId.toString()
    });
  }
  
  // Filter 2: Deal Stage (Partnership Proposal, Negotiation, Demonstrating Value)
  filters.push({
    propertyName: 'dealstage',
    operator: 'IN',
    values: ['90284260', '90284261', '90284259']
  });
  
  // Filter 3: Create Date (last 120 days)
  const now = new Date();
  const days120Ago = new Date();
  days120Ago.setDate(now.getDate() - 120);
  
  filters.push({
    propertyName: 'createdate',
    operator: 'GTE',
    value: days120Ago.getTime().toString()
  });
  
  // Filter 4: NOT Closed Lost
  filters.push({
    propertyName: 'closed_status',
    operator: 'NEQ',
    value: 'Closed lost (please specify the reason)'
  });
  
  // Filter 5: GTC Filter - Ability to Pay >= 3 (between 3-5)
  filters.push({
    propertyName: 'ability_to_pay',
    operator: 'GTE',
    value: '3'
  });
  
  // Filter 6: GTC Filter - Warm Prospects >= 3 (between 3-5)
  filters.push({
    propertyName: 'warm_prospects',
    operator: 'GTE',
    value: '3'
  });
  
  const filterGroups = [{ filters: filters }];
  
  const payload = {
    properties: properties,
    limit: HUBSPOT_API_CONFIG.DEFAULT_BATCH_SIZE
  };
  
  if (filterGroups.length > 0) {
    payload.filterGroups = filterGroups;
  }
  
  if (after) {
    payload.after = after;
  }
  
  const options_fetch = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${accessToken}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options_fetch);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (statusCode !== 200) {
    Logger.log(`HubSpot API Error (${statusCode}): ${responseText}`);
    throw new Error(`HubSpot API error: ${statusCode}`);
  }
  
  return JSON.parse(responseText);
}

/**
 * Fetches a single page of enrollment deals from HubSpot
 * @param {string} accessToken - HubSpot access token
 * @param {Array<string>} properties - Properties to fetch
 * @param {string} ownerId - Owner ID filter (numeric)
 * @param {string} after - Pagination cursor
 * @param {Object} options - Additional filter options
 * @returns {Object} API response object
 */
function fetchEnrollmentDealsPage(accessToken, properties, ownerId, after, options = {}) {
  const url = HUBSPOT_API_CONFIG.BASE_URL + HUBSPOT_API_CONFIG.ENDPOINTS.DEALS_SEARCH;
  
  // Build filters (all filters in one group = AND logic)
  const filters = [];
  
  // Filter 1: Owner ID
  if (ownerId) {
    filters.push({
      propertyName: 'hubspot_owner_id',
      operator: 'EQ',
      value: ownerId.toString()
    });
  }
  
  // Filter 2: Deal Stage (Partnership Confirmed = 90284262)
  filters.push({
    propertyName: 'dealstage',
    operator: 'EQ',
    value: '90284262'
  });
  
  // Filter 3: Closed Status (Closed won)
  filters.push({
    propertyName: 'closed_status',
    operator: 'EQ',
    value: 'Closed won'
  });
  
  // Filter 4: Close Date - Current month and last month
  // Calculate first day of last month
  const now = new Date();
  const firstDayOfCurrentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const firstDayOfLastMonth = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  
  filters.push({
    propertyName: 'closedate',
    operator: 'GTE',
    value: firstDayOfLastMonth.getTime().toString()
  });
  
  const filterGroups = [{ filters: filters }];
  
  const payload = {
    properties: properties,
    limit: HUBSPOT_API_CONFIG.DEFAULT_BATCH_SIZE
  };
  
  if (filterGroups.length > 0) {
    payload.filterGroups = filterGroups;
  }
  
  if (after) {
    payload.after = after;
  }
  
  const options_fetch = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${accessToken}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options_fetch);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (statusCode !== 200) {
    Logger.log(`HubSpot API Error (${statusCode}): ${responseText}`);
    throw new Error(`HubSpot API error: ${statusCode}`);
  }
  
  return JSON.parse(responseText);
}

// ============================================================================
// OWNER LOOKUP
// ============================================================================

/**
 * Gets owner ID by email address
 * @param {string} email - Owner email address
 * @returns {string|null} Owner ID or null if not found
 */
function getOwnerIdByEmail(email) {
  try {
    const accessToken = getHubSpotAccessToken();
    const url = HUBSPOT_API_CONFIG.BASE_URL + HUBSPOT_API_CONFIG.ENDPOINTS.OWNERS;
    
    const options = {
      method: 'get',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    
    if (statusCode !== 200) {
      Logger.log(`Failed to fetch owners: ${statusCode}`);
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (data.results) {
      // Search for owner by email
      const owner = data.results.find(owner => 
        owner.email && owner.email.toLowerCase() === email.toLowerCase()
      );
      
      if (owner) {
        return owner.id.toString();
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log(`Error getting owner ID for ${email}: ${error.message}`);
    return null;
  }
}

// ============================================================================
// URL BUILDING
// ============================================================================

/**
 * Builds a URL to a HubSpot deal
 * @param {string} dealId - The HubSpot deal ID
 * @returns {string} Full URL to the deal in HubSpot
 */
function buildDealUrl(dealId) {
  // EU region URL (portal ID: 25196166)
  return `https://app-eu1.hubspot.com/contacts/25196166/record/0-3/${dealId}/`;
}

// ============================================================================
// DATA EXTRACTION HELPERS
// ============================================================================

/**
 * Extracts property value from a deal object
 * @param {Object} deal - The deal object from HubSpot
 * @param {string} propertyName - The property name to extract
 * @returns {*} The property value (or empty string if not found)
 */
function extractDealProperty(deal, propertyName) {
  if (!deal || !deal.properties) {
    return '';
  }
  
  const value = deal.properties[propertyName];
  
  // Handle null/undefined
  if (value === null || value === undefined) {
    return '';
  }
  
  return value;
}

/**
 * Extracts and formats a date property from a deal
 * @param {Object} deal - The deal object from HubSpot
 * @param {string} propertyName - The date property name
 * @param {string} timezone - Timezone for formatting (default: 'America/New_York')
 * @returns {Date|string} Formatted date or empty string
 */
function extractDateProperty(deal, propertyName, timezone = 'America/New_York') {
  const value = extractDealProperty(deal, propertyName);
  
  if (!value || value === '' || value === '0' || value === 'null') {
    return '';
  }
  
  try {
    const valueStr = value.toString().trim();
    
    // Check if it's an ISO date string (e.g., 2025-08-18T17:46:59.963Z)
    if (/^\d{4}-\d{2}-\d{2}T/.test(valueStr)) {
      const date = new Date(valueStr);
      if (!isNaN(date.getTime())) {
        return date;
      }
    }
    
    // Check if it's already in YYYY-MM-DD format
    if (/^\d{4}-\d{2}-\d{2}$/.test(valueStr)) {
      return valueStr;
    }
    
    // Try parsing as Unix timestamp in milliseconds
    const timestamp = parseInt(valueStr);
    if (isNaN(timestamp) || timestamp === 0) {
      return '';
    }
    
    // Validate timestamp is reasonable (after year 2000, before year 2100)
    if (timestamp < 946684800000 || timestamp > 4102444800000) {
      return '';
    }
    
    // Convert to Date object
    const date = new Date(timestamp);
    if (isNaN(date.getTime())) {
      return '';
    }
    
    // Return as Date object for Google Sheets (will auto-format)
    return date;
    
  } catch (error) {
    Logger.log(`Error parsing date ${value}: ${error.message}`);
    return '';
  }
}

/**
 * Extracts a numeric property from a deal
 * @param {Object} deal - The deal object from HubSpot
 * @param {string} propertyName - The property name
 * @returns {number|string} Numeric value or empty string
 */
function extractNumericProperty(deal, propertyName) {
  const value = extractDealProperty(deal, propertyName);
  
  if (value === '' || value === null || value === undefined) {
    return '';
  }
  
  const numValue = parseFloat(value);
  return isNaN(numValue) ? '' : numValue;
}

// ============================================================================
// TESTING FUNCTIONS
// ============================================================================

/**
 * Test function to verify HubSpot API connection
 * Can be run from the Apps Script editor
 */
function testHubSpotConnection() {
  try {
    Logger.log('=== Testing HubSpot Connection ===');
    
    const accessToken = getHubSpotAccessToken();
    Logger.log('✅ Access token found');
    
    // Test with a simple properties fetch
    const url = HUBSPOT_API_CONFIG.BASE_URL + HUBSPOT_API_CONFIG.ENDPOINTS.DEALS_SEARCH;
    
    const payload = {
      properties: ['dealname', 'hubspot_owner_id'],
      limit: 5
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    
    if (statusCode !== 200) {
      throw new Error(`API returned status ${statusCode}: ${response.getContentText()}`);
    }
    
    const data = JSON.parse(response.getContentText());
    const dealCount = data.results ? data.results.length : 0;
    
    Logger.log(`✅ Successfully fetched ${dealCount} test deals`);
    Logger.log('=== Test Complete ===');
    
    return true;
    
  } catch (error) {
    Logger.log(`❌ Test failed: ${error.message}`);
    return false;
  }
}
