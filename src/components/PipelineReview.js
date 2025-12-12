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

// Note: TAB_PIPELINE and STAGE_MAP are defined in Constants.js

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
  // Call Quality properties (SCORES, not __details text) - ALL 13 fields
  's_discovery_a_questioning_technique',
  's_discovery_a_empathy__rapport_building_and_active_listening',
  's_building_value_a_recap_of_students_needs',
  's_building_value_a_tailoring_features_and_benefits',
  's_gaining_an_affirmation_and_program_requirements__a_gaining_affirmation',
  's_gaining_an_affirmation_and_program_requirements__a_essential_program_requirements',
  's_funding_options__a_identifying_funding_needs',
  's_funding_options__a_presenting_funding_solutions',
  's_funding_options__a_securing_financial_commitment',
  's_addressing_objections_a_identifying_and_addressing_objections_and_obstacles',
  's_closing_the_deal__a_creating_a_sense_of_urgency',
  's_closing_the_deal__a_assuming_the_sale',
  's_closing_the_deal__a_ask_for_referral'
];

// Field configuration for Pipeline Review
// ORDER: Deal Name → Stage → Last Activity → Next Activity → Notes → Why Not Purchase → [Call Quality]
const PIPELINE_FIELDS = [
  { property: 'dealId', header: 'Deal ID', hidden: true, type: 'text' },
  { property: 'dealname', header: 'Deal Name', hyperlink: true, type: 'text' },
  { property: 'dealstage', header: 'Stage', type: 'text' },
  { property: 'notes_last_updated', header: 'Last Activity', type: 'date' },
  { property: 'notes_next_activity_date', header: 'Next Activity', type: 'date' },
  { property: '__notes__', header: 'Notes', editable: true, preserve: true, type: 'manual' }, // Manual field
  { property: 'why_not_purchase_today_', header: 'Why Not Purchase Today', type: 'text' },
  // Call Quality columns at the END (ALL 13 fields)
  { property: 's_discovery_a_questioning_technique', header: 'Questioning', colorCode: true, type: 'number' },
  { property: 's_discovery_a_empathy__rapport_building_and_active_listening', header: 'Trust', colorCode: true, type: 'number' },
  { property: 's_building_value_a_recap_of_students_needs', header: 'Recap Needs', colorCode: true, type: 'number' },
  { property: 's_building_value_a_tailoring_features_and_benefits', header: 'Building Value', colorCode: true, type: 'number' },
  { property: 's_gaining_an_affirmation_and_program_requirements__a_gaining_affirmation', header: 'Program Alignment', colorCode: true, type: 'number' },
  { property: 's_gaining_an_affirmation_and_program_requirements__a_essential_program_requirements', header: 'Requirements', colorCode: true, type: 'number' },
  { property: 's_funding_options__a_identifying_funding_needs', header: 'Funding Needs', colorCode: true, type: 'number' },
  { property: 's_funding_options__a_presenting_funding_solutions', header: 'Funding Solution', colorCode: true, type: 'number' },
  { property: 's_funding_options__a_securing_financial_commitment', header: 'Funding Commitment', colorCode: true, type: 'number' },
  { property: 's_addressing_objections_a_identifying_and_addressing_objections_and_obstacles', header: 'Objections', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_creating_a_sense_of_urgency', header: 'Urgency', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_assuming_the_sale', header: 'Assume Sale', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_ask_for_referral', header: 'Referral', colorCode: true, type: 'number' }
];

// DEPRECATED: Manual fields now integrated into PIPELINE_FIELDS
const MANUAL_FIELDS = [];

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
    
    // 🛡️ SAFEGUARD: Handle empty deal list gracefully
    if (deals.length === 0) {
      Logger.log('  No deals found - clearing sheet and showing message');
      sheet.clear();
      sheet.getRange('A1').setValue('📊 Pipeline Review')
        .setFontWeight('bold')
        .setFontSize(14)
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF');
      sheet.getRange('A3').setValue('ℹ️ No active deals found matching filters:')
        .setFontWeight('bold');
      sheet.getRange('A4').setValue('• Stages: Negotiation, Partnership Proposal')
        .setFontSize(10);
      sheet.getRange('A5').setValue('• Created in last 90 days')
        .setFontSize(10);
      sheet.getRange('A6').setValue('• Not Closed Lost')
        .setFontSize(10);
      
      return {
        success: true,
        dealCount: 0,
        duration: (new Date() - startTime) / 1000
      };
    }
    
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
    
    // Find notes column
    const notesFieldIndex = PIPELINE_FIELDS.findIndex(f => f.property === '__notes__');
    const notes = notesFieldIndex >= 0 ? (row[notesFieldIndex] || '') : '';
    
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
  
  // Calculate 90-day window
  const today = new Date();
  const ninetyDaysAgo = new Date(today);
  ninetyDaysAgo.setDate(today.getDate() - 90);
  
  // Build filter for this AE's deals
  const filterGroups = [];
  
  // Filter by email or HubSpot User ID + Stage + Date + Exclude Closed Lost
  const baseFilters = [];
  
  // 1. Owner filter
  if (person.hubspotUserId && person.hubspotUserId !== '') {
    baseFilters.push({
      propertyName: 'hubspot_owner_id',
      operator: 'EQ',
      value: person.hubspotUserId
    });
  } else {
    // Fallback: filter by email (if HubSpot User ID not available)
    Logger.log(`  Warning: No HubSpot User ID for ${person.name}, using email filter`);
    baseFilters.push({
      propertyName: 'hubspot_owner_email',
      operator: 'EQ',
      value: person.email
    });
  }
  
  // 2. Stage filter - Only Negotiation (90284261) and Partnership Proposal (90284260)
  baseFilters.push({
    propertyName: 'dealstage',
    operator: 'IN',
    values: ['90284261', '90284260']
  });
  
  // 3. Date filter - Last 90 days
  baseFilters.push({
    propertyName: 'createdate',
    operator: 'GTE',
    value: ninetyDaysAgo.getTime()
  });
  
  // 4. Exclude Closed Lost deals
  baseFilters.push({
    propertyName: 'closed_status',
    operator: 'NEQ',
    value: 'Closed lost (please specify the reason)'
  });
  
  filterGroups.push({
    filters: baseFilters
  });
  
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
  const headers = PIPELINE_FIELDS.map(f => f.header);
  dataArray.push(headers);
  
  // Data rows
  deals.forEach(deal => {
    const dealId = deal.id.toString();
    const properties = deal.properties || {};
    const preserved = preservedMap.get(dealId) || {};
    
    const row = [];
    
    // All fields (including manual Notes field)
    PIPELINE_FIELDS.forEach(field => {
      let value = '';
      
      // Special handling for Deal ID
      if (field.property === 'dealId') {
        value = dealId;
      }
      // Special handling for Notes (manual field)
      else if (field.property === '__notes__') {
        value = preserved.notes || '';
      }
      // Special handling for Stage (map ID to name)
      else if (field.property === 'dealstage') {
        const stageId = properties.dealstage || '';
        value = STAGE_MAP[stageId] || stageId; // Fallback to ID if not found
      }
      // HubSpot properties
      else {
        value = properties[field.property] || '';
        
        // Format dates
        if (field.type === 'date' && value) {
          value = formatDate(value);
        }
      }
      
      row.push(value);
    });
    
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
 * Writes data array to sheet with batch hyperlink creation (FAST)
 * @param {Sheet} sheet - Pipeline Review sheet
 * @param {Array} dataArray - 2D data array
 */
function writeDataToSheet(sheet, dataArray) {
  if (dataArray.length === 0) return;
  
  // Write all data
  sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
  
  // Add hyperlinks for Deal Name column (BATCH operation)
  const dealNameColIndex = PIPELINE_FIELDS.findIndex(f => f.property === 'dealname') + 1;
  const dealIdColIndex = 1; // Deal ID is always column A
  
  const rowCount = dataArray.length - 1; // Exclude header
  if (rowCount > 0) {
    const richTextValues = [];
    
    for (let i = 1; i <= rowCount; i++) {
      const dealId = dataArray[i][dealIdColIndex - 1]; // -1 for 0-based array index
      const dealName = dataArray[i][dealNameColIndex - 1];
      
      if (dealId && dealName) {
        const dealUrl = `https://app-eu1.hubspot.com/contacts/25196166/record/0-3/${dealId}`;
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(dealName)
          .setLinkUrl(dealUrl)
          .build();
        richTextValues.push([richText]);
      } else {
        richTextValues.push([dealName || '']);
      }
    }
    
    // Set all hyperlinks in one operation
    sheet.getRange(2, dealNameColIndex, rowCount, 1).setRichTextValues(richTextValues);
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
  
  // Apply preserved highlighting (from Director) - MVP: Skip for now (not critical, slows down)
  // TODO: Re-enable when director highlighting is working
  // applyPreservedHighlighting(sheet, dataArray, preservedMap);
  
  // Protect HubSpot data columns (all except Notes)
  protectDataColumns(sheet, dataArray[0].length);
}

/**
 * Applies red-yellow-green color coding to call quality columns
 * Uses conditional formatting rules (FAST - single operation per column)
 * @param {Sheet} sheet - Pipeline Review sheet
 * @param {Array} dataArray - 2D data array
 */
function applyCallQualityFormatting(sheet, dataArray) {
  const dataRowCount = dataArray.length - 1; // Exclude header
  if (dataRowCount < 1) return;
  
  // Find call quality column indices
  const callQualityFields = PIPELINE_FIELDS.filter(f => f.colorCode);
  const rules = [];
  
  callQualityFields.forEach(field => {
    const colIndex = PIPELINE_FIELDS.findIndex(f => f.property === field.property) + 1;
    
    if (colIndex > 0) {
      const range = sheet.getRange(2, colIndex, dataRowCount, 1);
      
      // Gradient: Red (0) → Yellow (2.5) → Green (5)
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .setGradientMinpointWithValue('#F4C7C3', SpreadsheetApp.InterpolationType.NUMBER, '0')
        .setGradientMidpointWithValue('#FCE8B2', SpreadsheetApp.InterpolationType.NUMBER, '2.5')
        .setGradientMaxpointWithValue('#B7E1CD', SpreadsheetApp.InterpolationType.NUMBER, '5')
        .setRanges([range])
        .build();
      
      rules.push(rule);
    }
  });
  
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
  }
}

/**
 * Applies preserved highlighting from Director (BATCH operation, only for rows with highlighting)
 * @param {Sheet} sheet - Pipeline Review sheet
 * @param {Array} dataArray - 2D data array
 * @param {Map} preservedMap - Preserved highlighting by Deal ID
 */
function applyPreservedHighlighting(sheet, dataArray, preservedMap) {
  if (preservedMap.size === 0) return; // No highlighting to restore
  
  const currentColCount = dataArray[0].length;
  const rowCount = dataArray.length - 1; // Exclude header
  
  // Only apply highlighting to rows that actually have it (SPARSE operation)
  for (let i = 1; i <= rowCount; i++) {
    const dealId = dataArray[i][0]?.toString(); // Deal ID is first column in dataArray
    const preserved = preservedMap.get(dealId);
    
    if (preserved && (preserved.backgrounds || preserved.fontColors)) {
      const rowIndex = i + 1; // +1 for header
      const rowRange = sheet.getRange(rowIndex, 1, 1, currentColCount);
      
      // Truncate or pad to match current column count
      if (preserved.backgrounds) {
        const backgrounds = preserved.backgrounds.slice(0, currentColCount);
        while (backgrounds.length < currentColCount) {
          backgrounds.push('#ffffff');
        }
        rowRange.setBackgrounds([backgrounds]);
      }
      
      if (preserved.fontColors) {
        const fontColors = preserved.fontColors.slice(0, currentColCount);
        while (fontColors.length < currentColCount) {
          fontColors.push('#000000');
        }
        rowRange.setFontColors([fontColors]);
      }
    }
  }
  
  Logger.log(`  Applied highlighting to ${preservedMap.size} deals`);
}

/**
 * Protects HubSpot data columns (all except Notes) - MVP: Just protect actual data rows
 * @param {Sheet} sheet - Pipeline Review sheet
 * @param {number} totalCols - Total number of columns
 */
function protectDataColumns(sheet, totalCols) {
  // MVP: Skip protection for now - too slow and not critical
  // TODO: Re-enable after finding faster approach
  // 
  // const notesColIndex = totalCols; // Last column is Notes
  // const actualDataRows = sheet.getLastRow() - 1; // Only protect actual data, not 1000 empty rows
  // 
  // if (actualDataRows > 0) {
  //   const protection = sheet.getRange(2, 1, actualDataRows, notesColIndex - 1).protect();
  //   protection.setDescription('HubSpot data (read-only)');
  //   protection.setWarningOnly(true); // Warning only, not strict
  // }
}
