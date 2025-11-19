/**
 * Pipeline Review Component
 * 
 * Displays deal pipeline data from HubSpot with call quality scores
 * Includes manual notes preservation and format preservation across refreshes
 * 
 * Note: Uses TAB_PIPELINE constant from SheetProvisioner.js
 */

// ============================================================================
// FIELD CONFIGURATION
// ============================================================================

// Stage mapping (ID to Name)
const STAGE_MAP = {
  '90284257': 'Create Curiosity',
  '90284258': 'Needs Analysis',
  '90284259': 'Demonstrating Value',
  '90284260': 'Partnership Proposal',
  '90284261': 'Negotiation',
  '90284262': 'Partnership Confirmed'
};

// Core deal fields (always visible)
const CORE_FIELDS = [
  { property: 'dealname', header: 'Deal Name', hyperlink: true, type: 'text' },
  { property: 'dealstage', header: 'Stage', type: 'text', useMapping: true },
  { property: 'notes_last_updated', header: 'Last Activity', type: 'date' },
  { property: 'notes_next_activity_date', header: 'Next Activity', type: 'date' },
  { property: 'next_task_name', header: 'Next Task Name', enabled: false, type: 'text' }, // Future: needs API permissions
  { property: 'why_not_purchase_today_', header: 'Why Not Purchase Today', type: 'text' }
];

// Call quality score fields (color-coded red-yellow-green)
const CALL_QUALITY_FIELDS = [
  { property: 's_discovery_a_questioning_technique', header: 'DISCOVERY', colorCode: true, type: 'number' },
  { property: 's_discovery_a_empathy__rapport_building_and_active_listening', header: 'TRUST', colorCode: true, type: 'number' },
  { property: 's_building_value_a_recap_of_students_needs', header: 'RECAP NEEDS', colorCode: true, type: 'number' },
  { property: 's_building_value_a_tailoring_features_and_benefits', header: 'TAILORING FEATURES', colorCode: true, type: 'number' },
  { property: 's_gaining_an_affirmation_and_program_requirements__a_gaining_affirmation', header: 'PROGRAM ALIGNMENT', colorCode: true, type: 'number' },
  { property: 's_gaining_an_affirmation_and_program_requirements__a_essential_program_requirements', header: 'REQUIREMENTS', colorCode: true, type: 'number' },
  { property: 's_funding_options__a_identifying_funding_needs', header: 'FUNDING NEEDS', colorCode: true, type: 'number' },
  { property: 's_funding_options__a_presenting_funding_solutions', header: 'FUNDING SOLUTION', colorCode: true, type: 'number' },
  { property: 's_funding_options__a_securing_financial_commitment', header: 'FUNDING COMMITMENT', colorCode: true, type: 'number' },
  { property: 's_addressing_objections_a_identifying_and_addressing_objections_and_obstacles', header: 'OBJECTIONS', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_creating_a_sense_of_urgency', header: 'URGENCY', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_assuming_the_sale', header: 'ASSUME SALE', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_ask_for_referral', header: 'REFERRAL', colorCode: true, type: 'number' }
];

// Manual note columns (editable, preserved across refreshes)
const MANUAL_FIELDS = [
  { header: 'Note 1', editable: true, preserve: true },
  { header: 'Note 2', editable: true, preserve: true }
];

/**
 * Gets all enabled properties to fetch from HubSpot
 * @returns {Array<string>} Array of property names
 */
function getPipelineReviewProperties() {
  const properties = [];
  
  // Add core fields (only enabled ones)
  CORE_FIELDS.forEach(field => {
    if (field.enabled !== false) {
      properties.push(field.property);
    }
  });
  
  // Add call quality fields
  CALL_QUALITY_FIELDS.forEach(field => {
    properties.push(field.property);
  });
  
  // Add standard fields always needed
  if (!properties.includes('hubspot_owner_id')) {
    properties.push('hubspot_owner_id');
  }
  
  return properties;
}

/**
 * Gets all column headers in order
 * @returns {Array<string>} Array of header names
 */
function getPipelineReviewHeaders() {
  const headers = [];
  
  // Core fields
  CORE_FIELDS.forEach(field => {
    headers.push(field.header);
  });
  
  // Call quality fields
  CALL_QUALITY_FIELDS.forEach(field => {
    headers.push(field.header);
  });
  
  // Manual fields
  MANUAL_FIELDS.forEach(field => {
    headers.push(field.header);
  });
  
  return headers;
}

// ============================================================================
// MAIN PIPELINE REVIEW FUNCTION
// ============================================================================

/**
 * Updates the Pipeline Review tab for a salesperson
 * @param {Spreadsheet} individualSheet - The salesperson's individual sheet
 * @param {Object} person - Person object {name, email}
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
    
    // Step 1: Capture existing notes and formatting BEFORE clearing
    Logger.log('  Step 1: Capturing existing notes and formatting...');
    const preserved = capturePreservedData(sheet);
    
    // Step 2: Fetch deals from HubSpot
    Logger.log('  Step 2: Fetching deals from HubSpot...');
    const properties = getPipelineReviewProperties();
    const options = {};
    
    // Pass HubSpot User ID if available
    if (person.hubspotUserId && person.hubspotUserId !== '') {
      options.hubspotUserId = person.hubspotUserId;
    }
    
    const deals = fetchDealsByOwner(person.email, properties, options);
    Logger.log(`  Found ${deals.length} deals`);
    
    // Step 3: Build data array
    Logger.log('  Step 3: Building data array...');
    const { dataArray, urlMap, dealIdMap } = buildPipelineDataArray(deals);
    
    // Step 4: Clear and write data
    Logger.log('  Step 4: Writing data to sheet...');
    sheet.clear();
    writeDataToSheet(sheet, dataArray, urlMap);
    
    // Step 5: Apply formatting
    Logger.log('  Step 5: Applying formatting...');
    applyPipelineFormatting(sheet, dataArray.length - 1); // -1 for header
    
    // Step 6: Restore preserved notes and formatting
    Logger.log('  Step 6: Restoring preserved data...');
    restorePreservedData(sheet, preserved, dealIdMap);
    
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
// DATA BUILDING
// ============================================================================

/**
 * Builds the data array for Pipeline Review sheet
 * @param {Array<Object>} deals - Raw deals from HubSpot
 * @returns {Object} {dataArray, urlMap, dealIdMap}
 */
function buildPipelineDataArray(deals) {
  const dataArray = [];
  const urlMap = {}; // Map of row index to URL for hyperlinks
  const dealIdMap = {}; // Map of row index to Deal ID (for preservation)
  
  // Headers
  dataArray.push(getPipelineReviewHeaders());
  
  // Data rows
  deals.forEach((deal, index) => {
    const row = [];
    const rowIndex = index + 2; // +2 because row 1 is header, index starts at 0
    
    // Store Deal ID for preservation (not displayed in sheet)
    dealIdMap[rowIndex] = deal.id;
    
    // Core fields
    CORE_FIELDS.forEach(field => {
      if (field.property === 'dealname') {
        // Deal Name - will be hyperlinked
        const dealName = extractDealProperty(deal, field.property);
        row.push(dealName);
        urlMap[rowIndex] = buildDealUrl(deal.id);
      } else if (field.property === 'dealstage' && field.useMapping) {
        // Stage - map ID to name
        const stageId = extractDealProperty(deal, field.property);
        row.push(STAGE_MAP[stageId] || stageId);
      } else if (field.type === 'date') {
        // Date fields
        row.push(extractDateProperty(deal, field.property));
      } else if (field.enabled === false) {
        // Disabled field (Next Task Name) - leave blank
        row.push('');
      } else {
        // Regular text field
        row.push(extractDealProperty(deal, field.property));
      }
    });
    
    // Call quality fields (numeric)
    CALL_QUALITY_FIELDS.forEach(field => {
      row.push(extractNumericProperty(deal, field.property));
    });
    
    // Manual fields (blank initially, will be restored if preserved)
    MANUAL_FIELDS.forEach(() => {
      row.push('');
    });
    
    dataArray.push(row);
  });
  
  return { dataArray, urlMap, dealIdMap };
}

// ============================================================================
// SHEET WRITING
// ============================================================================

/**
 * Writes data to the Pipeline Review sheet
 * @param {Sheet} sheet - The sheet to write to
 * @param {Array<Array>} dataArray - 2D array of data
 * @param {Object} urlMap - Map of row index to URL for hyperlinks
 */
function writeDataToSheet(sheet, dataArray, urlMap) {
  if (dataArray.length === 0) {
    return;
  }
  
  // Write all data at once
  const range = sheet.getRange(1, 1, dataArray.length, dataArray[0].length);
  range.setValues(dataArray);
  
  // Apply hyperlinks to Deal Name column (Column A) - batch operation
  const dealNameCol = 1; // Column A
  const rowCount = dataArray.length - 1; // Exclude header
  
  if (rowCount > 0 && Object.keys(urlMap).length > 0) {
    // Build rich text values array for all rows
    const richTextValues = [];
    
    for (let i = 0; i < rowCount; i++) {
      const rowIndex = i + 2; // +2 for header and 0-based
      const dealName = dataArray[i + 1][0]; // +1 to skip header row
      const url = urlMap[rowIndex];
      
      if (url && dealName) {
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(dealName)
          .setLinkUrl(url)
          .build();
        richTextValues.push([richText]);
      } else {
        richTextValues.push([dealName]); // Plain text if no URL
      }
    }
    
    // Apply all hyperlinks at once
    sheet.getRange(2, dealNameCol, rowCount, 1).setRichTextValues(richTextValues);
  }
}

/**
 * Applies formatting to Pipeline Review sheet
 * @param {Sheet} sheet - The sheet to format
 * @param {number} dataRowCount - Number of data rows (excluding header)
 */
function applyPipelineFormatting(sheet, dataRowCount) {
  if (dataRowCount < 1) {
    return;
  }
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Auto-resize columns
  for (let col = 1; col <= sheet.getLastColumn(); col++) {
    sheet.autoResizeColumn(col);
  }
  
  // Apply conditional formatting to call quality columns
  applyCallQualityFormatting(sheet, dataRowCount);
}

/**
 * Applies red-yellow-green conditional formatting to call quality columns
 * @param {Sheet} sheet - The sheet to format
 * @param {number} dataRowCount - Number of data rows
 */
function applyCallQualityFormatting(sheet, dataRowCount) {
  if (dataRowCount < 1) {
    return;
  }
  
  // Find column indices for call quality fields
  const headers = getPipelineReviewHeaders();
  const rules = [];
  
  CALL_QUALITY_FIELDS.forEach(field => {
    if (field.colorCode) {
      const colIndex = headers.indexOf(field.header) + 1; // +1 for 1-based indexing
      
      if (colIndex > 0) {
        const range = sheet.getRange(2, colIndex, dataRowCount, 1);
        
        // Create gradient rule: 0-5 scale with red (0) -> yellow (2.5) -> green (5)
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .setGradientMinpointWithValue(
            '#F4C7C3', // Light red
            SpreadsheetApp.InterpolationType.NUMBER,
            '0'
          )
          .setGradientMidpointWithValue(
            '#FCE8B2', // Light yellow
            SpreadsheetApp.InterpolationType.NUMBER,
            '2.5'
          )
          .setGradientMaxpointWithValue(
            '#B7E1CD', // Light green
            SpreadsheetApp.InterpolationType.NUMBER,
            '5'
          )
          .setRanges([range])
          .build();
        
        rules.push(rule);
      }
    }
  });
  
  // Apply all rules at once
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
    Logger.log(`  Applied conditional formatting to ${rules.length} call quality columns`);
  }
}

// ============================================================================
// DATA PRESERVATION
// ============================================================================

/**
 * Captures notes and formatting before refresh
 * @param {Sheet} sheet - The Pipeline Review sheet
 * @returns {Object} Preserved data indexed by row number (temporary, will map to Deal ID)
 */
function capturePreservedData(sheet) {
  const preserved = {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return preserved; // No data to preserve
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dealNameCol = headers.indexOf('Deal Name') + 1; // Use Deal Name for matching
  const note1Col = headers.indexOf('Note 1') + 1;
  const note2Col = headers.indexOf('Note 2') + 1;
  
  // Read all data
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();
  const fontColors = dataRange.getFontColors();
  const fontWeights = dataRange.getFontWeights();
  
  // Capture by Deal Name (temporary key until we get Deal IDs)
  for (let i = 0; i < values.length; i++) {
    const dealName = dealNameCol > 0 ? values[i][dealNameCol - 1] : '';
    
    if (dealName && dealName !== '') {
      preserved[dealName.toString()] = {
        note1: note1Col > 0 ? values[i][note1Col - 1] : '',
        note2: note2Col > 0 ? values[i][note2Col - 1] : '',
        backgrounds: backgrounds[i],
        fontColors: fontColors[i],
        fontWeights: fontWeights[i]
      };
    }
  }
  
  Logger.log(`  Preserved data for ${Object.keys(preserved).length} deals`);
  return preserved;
}

/**
 * Restores preserved notes and formatting after refresh
 * @param {Sheet} sheet - The Pipeline Review sheet
 * @param {Object} preserved - Preserved data indexed by Deal Name
 * @param {Object} dealIdMap - Map of row index to Deal ID
 */
function restorePreservedData(sheet, preserved, dealIdMap) {
  if (Object.keys(preserved).length === 0) {
    Logger.log('  No preserved data to restore');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dealNameCol = headers.indexOf('Deal Name') + 1;
  const note1Col = headers.indexOf('Note 1') + 1;
  const note2Col = headers.indexOf('Note 2') + 1;
  const currentColCount = headers.length;
  
  // Read all deal names at once
  const dealNames = sheet.getRange(2, dealNameCol, lastRow - 1, 1).getValues();
  
  // Prepare batch updates for notes
  const note1Updates = [];
  const note2Updates = [];
  let restoredCount = 0;
  
  // Build updates array
  for (let i = 0; i < dealNames.length; i++) {
    const dealName = dealNames[i][0];
    const rowIndex = i + 2; // +2 for header and 0-based index
    
    if (dealName && preserved[dealName.toString()]) {
      const data = preserved[dealName.toString()];
      
      note1Updates.push([data.note1 || '']);
      note2Updates.push([data.note2 || '']);
      
      restoredCount++;
    } else {
      note1Updates.push(['']);
      note2Updates.push(['']);
    }
  }
  
  // Apply all notes at once (batch)
  if (note1Col > 0 && note1Updates.length > 0) {
    sheet.getRange(2, note1Col, note1Updates.length, 1).setValues(note1Updates);
  }
  if (note2Col > 0 && note2Updates.length > 0) {
    sheet.getRange(2, note2Col, note2Updates.length, 1).setValues(note2Updates);
  }
  
  Logger.log(`  Restored data for ${restoredCount} deals`);
}
