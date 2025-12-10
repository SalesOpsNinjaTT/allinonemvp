/**
 * Pipeline Review Component
 * 
 * Displays pipeline data from HubSpot for individual Account Executives
 * - Fetches deals from HubSpot (with call quality scores)
 * - Preserves manual notes across refreshes (by Deal ID)
 * - Applies highlighting from Director (bi-directional sync)
 * - Color-codes call quality scores (red-yellow-green)
 * 
 * Tab Name: 📊 Pipeline Review
 */

// ============================================================================
// CONSTANTS
// ============================================================================

// Note: TAB_PIPELINE is defined in SheetProvisioner.js

// Colors for call quality scoring (same as Enrollment Tracker)
const SCORE_COLORS = {
  RED: '#f4cccc',      // 0-2: Red
  YELLOW: '#fff2cc',   // 3: Yellow
  GREEN: '#d9ead3'     // 4-5: Green
};

// HubSpot properties to fetch
const PIPELINE_PROPERTIES = [
  'dealname',
  'dealstage',
  'hubspot_owner_id',
  'notes_last_updated',
  'notes_next_activity_date',
  'why_not_purchase_today_',
  'createdate',
  'closedate',
  'hs_deal_stage_probability',
  // Call Quality properties
  'call_quality_score',
  's_discovery_a_questioning_technique__details',
  's_building_value_a_tailoring_features_and_benefits__details',
  's_funding_options__a_identifying_funding_needs__details',
  's_addressing_objections_a_identifying_and_addressing_objections_and_obstacles__details',
  's_closing_the_deal__a_assuming_the_sale__details',
  's_closing_the_deal__a_ask_for_referral__details'
];

// Field configuration for Pipeline Review
const PIPELINE_FIELDS = [
  { property: 'dealId', header: 'Deal ID', hidden: true, type: 'text' },
  { property: 'dealname', header: 'Deal Name', hyperlink: true, type: 'text' },
  { property: 'dealstage', header: 'Stage', type: 'text' },
  { property: 'notes_last_updated', header: 'Last Activity', type: 'date' },
  { property: 'notes_next_activity_date', header: 'Next Activity', type: 'date' },
  { property: 'why_not_purchase_today_', header: 'Why Not Purchase Today', type: 'text' },
  { property: 'call_quality_score', header: 'Call Quality Score', colorCode: true, type: 'number' },
  { property: 's_discovery_a_questioning_technique__details', header: 'Questioning', colorCode: true, type: 'number' },
  { property: 's_building_value_a_tailoring_features_and_benefits__details', header: 'Building Value', colorCode: true, type: 'number' },
  { property: 's_funding_options__a_identifying_funding_needs__details', header: 'Funding Options', colorCode: true, type: 'number' },
  { property: 's_addressing_objections_a_identifying_and_addressing_objections_and_obstacles__details', header: 'Addressing Objections', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_assuming_the_sale__details', header: 'Closing the Deal', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_ask_for_referral__details', header: 'Ask for Referral', colorCode: true, type: 'number' }
];

// Manual editable columns (appended after HubSpot fields)
const MANUAL_FIELDS = [
  { header: 'Notes', editable: true, preserve: true }
];

// ============================================================================
// MAIN PIPELINE REVIEW FUNCTION
// ============================================================================

/**
 * Updates the Pipeline Review tab for a salesperson
 * @param {Spreadsheet} individualSheet - The salesperson's individual sheet
 * @param {Object} person - Person object {name, email, hubspotUserId, team}
 * @returns {Object} Update result
 */
function updatePipelineReview(individualSheet, person) {
  try {
    Logger.log(`[Pipeline Review] Updating for ${person.name}...`);
    
    const startTime = new Date();
    const sheet = individualSheet.getSheetByName(TAB_PIPELINE);
    
    if (!sheet) {
      throw new Error(`Pipeline Review tab not found for ${person.name}`);
    }
    
    // Step 1: Capture existing notes and highlighting (by Deal ID)
    Logger.log('  Step 1: Capturing notes and highlighting...');
    const preservedData = capturePreservedData(sheet);
    
    // Step 2: Fetch deals from HubSpot
    Logger.log('  Step 2: Fetching deals from HubSpot...');
    const deals = fetchDealsForAE(person);
    Logger.log(`  Found ${deals.length} deals`);
    
    // Step 3: Build sheet data
    Logger.log('  Step 3: Building sheet data...');
    const dataArray = buildPipelineDataArray(deals, preservedData);
    
    // Step 4: Clear and write data
    Logger.log('  Step 4: Writing data to sheet...');
    clearSheetData(sheet);
    writeDataToSheet(sheet, dataArray);
    
    // Step 5: Apply formatting
    Logger.log('  Step 5: Applying formatting...');
    applyPipelineFormatting(sheet, dataArray, preservedData);
    
    const duration = (new Date() - startTime) / 1000;
    Logger.log(`[Pipeline Review] Complete for ${person.name} (${duration}s)`);
    
    return {
      success: true,
      dealCount: deals.length,
      duration: duration
    };
    
  } catch (error) {
    Logger.log(`[Pipeline Review] Error for ${person.name}: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================================================
// DATA CAPTURE (Preserve Notes & Highlighting)
// ============================================================================

/**
 * Captures preserved data (notes and highlighting) by Deal ID
 * @param {Sheet} sheet - Pipeline Review sheet
 * @returns {Map} Map of Deal ID to preserved data
 */
function capturePreservedData(sheet) {
  const preservedMap = new Map();
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return preservedMap; // Empty sheet
  }
  
  const lastCol = sheet.getLastColumn();
  
  // Read all data
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const backgrounds = sheet.getRange(2, 1, lastRow - 1, lastCol).getBackgrounds();
  const fontColors = sheet.getRange(2, 1, lastRow - 1, lastCol).getFontColors();
  
  data.forEach((row, index) => {
    const dealId = row[0]?.toString(); // Column A is Deal ID
    if (!dealId) return;
    
    // Find notes column (last column)
    const notesColumnIndex = PIPELINE_FIELDS.length; // Notes is after all HubSpot fields
    const notes = row[notesColumnIndex] || '';
    
    preservedMap.set(dealId, {
      notes: notes,
      backgrounds: backgrounds[index],
      fontColors: fontColors[index]
    });
  });
  
  Logger.log(`  Captured data for ${preservedMap.size} deals`);
  return preservedMap;
}

// ============================================================================
// HUBSPOT DATA FETCHING
// ============================================================================

/**
 * Fetches deals from HubSpot for a specific AE
 * @param {Object} person - Person object
 * @returns {Array} Array of deals
 */
function fetchDealsForAE(person) {
  const token = getHubSpotToken();
  
  // Build filter for this AE's deals
  const filterGroups = [];
  
  // Filter by email or HubSpot User ID
  if (person.hubspotUserId && person.hubspotUserId !== '') {
    filterGroups.push({
      filters: [{
        propertyName: 'hubspot_owner_id',
        operator: 'EQ',
        value: person.hubspotUserId
      }]
    });
  } else {
    // Fallback: filter by email (if HubSpot User ID not available)
    Logger.log(`  Warning: No HubSpot User ID for ${person.name}, using email filter`);
    filterGroups.push({
      filters: [{
        propertyName: 'hubspot_owner_email',
        operator: 'EQ',
        value: person.email
      }]
    });
  }
  
  // Fetch from HubSpot
  const url = 'https://api.hubapi.com/crm/v3/objects/deals/search';
  
  const payload = {
    filterGroups: filterGroups,
    properties: PIPELINE_PROPERTIES,
    limit: 100
  };
  
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': `Bearer ${token}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error(`HubSpot API error: ${responseCode} - ${response.getContentText()}`);
    }
    
    const result = JSON.parse(response.getContentText());
    return result.results || [];
    
  } catch (error) {
    Logger.log(`  Error fetching deals: ${error.message}`);
    return [];
  }
}

// ============================================================================
// DATA BUILDING
// ============================================================================

/**
 * Builds data array for sheet from deals and preserved data
 * @param {Array} deals - HubSpot deals
 * @param {Map} preservedMap - Preserved notes and highlighting
 * @returns {Array} 2D array for sheet
 */
function buildPipelineDataArray(deals, preservedMap) {
  const dataArray = [];
  
  // Header row
  const headers = PIPELINE_FIELDS.map(f => f.header).concat(MANUAL_FIELDS.map(f => f.header));
  dataArray.push(headers);
  
  // Data rows
  deals.forEach(deal => {
    const dealId = deal.id.toString();
    const properties = deal.properties || {};
    const preserved = preservedMap.get(dealId) || {};
    
    const row = [];
    
    // HubSpot fields
    PIPELINE_FIELDS.forEach(field => {
      let value = properties[field.property] || '';
      
      // Special handling for Deal ID
      if (field.property === 'dealId') {
        value = dealId;
      }
      
      // Format dates
      if (field.type === 'date' && value) {
        value = formatDate(value);
      }
      
      row.push(value);
    });
    
    // Manual fields (notes) - restored from preserved data
    row.push(preserved.notes || '');
    
    dataArray.push(row);
  });
  
  return dataArray;
}

/**
 * Formats a date value to yyyy-MM-dd
 * @param {string|number} dateValue - Date value from HubSpot
 * @returns {string} Formatted date
 */
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  try {
    let date;
    if (typeof dateValue === 'string') {
      date = new Date(dateValue);
    } else if (typeof dateValue === 'number') {
      date = new Date(dateValue);
    } else {
      return '';
    }
    
    if (isNaN(date.getTime())) return '';
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
  } catch (e) {
    return '';
  }
}

// ============================================================================
// SHEET WRITING
// ============================================================================

/**
 * Clears sheet data (keeps formatting)
 * @param {Sheet} sheet - Pipeline Review sheet
 */
function clearSheetData(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clear();
  }
  sheet.clearConditionalFormatRules();
}

/**
 * Writes data array to sheet
 * @param {Sheet} sheet - Pipeline Review sheet
 * @param {Array} dataArray - 2D data array
 */
function writeDataToSheet(sheet, dataArray) {
  if (dataArray.length === 0) return;
  
  // Write all data
  sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
  
  // Add hyperlinks for Deal Name column
  const dealNameColIndex = PIPELINE_FIELDS.findIndex(f => f.property === 'dealname') + 1;
  const dealIdColIndex = 1; // Deal ID is always column A
  
  for (let i = 2; i <= dataArray.length; i++) {
    const dealId = sheet.getRange(i, dealIdColIndex).getValue();
    if (dealId) {
      const dealUrl = `https://app.hubspot.com/contacts/47363978/deal/${dealId}`;
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(sheet.getRange(i, dealNameColIndex).getValue())
        .setLinkUrl(dealUrl)
        .build();
      sheet.getRange(i, dealNameColIndex).setRichTextValue(richText);
    }
  }
}

// ============================================================================
// FORMATTING
// ============================================================================

/**
 * Applies formatting to Pipeline Review sheet
 * @param {Sheet} sheet - Pipeline Review sheet
 * @param {Array} dataArray - 2D data array
 * @param {Map} preservedMap - Preserved highlighting by Deal ID
 */
function applyPipelineFormatting(sheet, dataArray, preservedMap) {
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, dataArray[0].length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  sheet.setFrozenRows(1);
  
  // Hide Deal ID column
  sheet.hideColumns(1);
  
  // Apply call quality color coding
  applyCallQualityFormatting(sheet, dataArray);
  
  // Apply preserved highlighting (from Director)
  applyPreservedHighlighting(sheet, dataArray, preservedMap);
  
  // Protect HubSpot data columns (all except Notes)
  protectDataColumns(sheet, dataArray[0].length);
}

/**
 * Applies red-yellow-green color coding to call quality columns
 * @param {Sheet} sheet - Pipeline Review sheet
 * @param {Array} dataArray - 2D data array
 */
function applyCallQualityFormatting(sheet, dataArray) {
  // Find call quality column indices
  const callQualityFields = PIPELINE_FIELDS.filter(f => f.colorCode);
  
  callQualityFields.forEach(field => {
    const colIndex = PIPELINE_FIELDS.findIndex(f => f.property === field.property) + 1;
    
    // Apply to data rows (skip header)
    for (let row = 2; row <= dataArray.length; row++) {
      const value = sheet.getRange(row, colIndex).getValue();
      
      if (value === '' || value === null) continue;
      
      const numValue = parseFloat(value);
      if (isNaN(numValue)) continue;
      
      let color = SCORE_COLORS.GREEN;
      if (numValue <= 2) {
        color = SCORE_COLORS.RED;
      } else if (numValue === 3) {
        color = SCORE_COLORS.YELLOW;
      }
      
      sheet.getRange(row, colIndex).setBackground(color);
    }
  });
}

/**
 * Applies preserved highlighting from Director
 * @param {Sheet} sheet - Pipeline Review sheet
 * @param {Array} dataArray - 2D data array
 * @param {Map} preservedMap - Preserved highlighting by Deal ID
 */
function applyPreservedHighlighting(sheet, dataArray, preservedMap) {
  const dealIdColIndex = 1;
  
  for (let row = 2; row <= dataArray.length; row++) {
    const dealId = sheet.getRange(row, dealIdColIndex).getValue()?.toString();
    if (!dealId) continue;
    
    const preserved = preservedMap.get(dealId);
    if (!preserved) continue;
    
    // Apply row backgrounds and font colors
    const rowRange = sheet.getRange(row, 1, 1, dataArray[0].length);
    
    if (preserved.backgrounds) {
      rowRange.setBackgrounds([preserved.backgrounds]);
    }
    
    if (preserved.fontColors) {
      rowRange.setFontColors([preserved.fontColors]);
    }
  }
}

/**
 * Protects HubSpot data columns (all except Notes)
 * @param {Sheet} sheet - Pipeline Review sheet
 * @param {number} totalCols - Total number of columns
 */
function protectDataColumns(sheet, totalCols) {
  const notesColIndex = totalCols; // Last column is Notes
  
  // Protect columns A through (Notes - 1)
  const protection = sheet.getRange(2, 1, sheet.getMaxRows() - 1, notesColIndex - 1).protect();
  protection.setDescription('HubSpot data (read-only)');
  protection.setWarningOnly(false);
  
  // Remove all editors except the script
  const me = Session.getEffectiveUser();
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}
