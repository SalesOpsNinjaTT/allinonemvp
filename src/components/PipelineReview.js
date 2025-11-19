/**
 * Pipeline Review Component
 * 
 * Displays deal pipeline data from HubSpot with call quality scores
 * Includes manual notes preservation and format preservation across refreshes
 */

// Tab name constant (must match SheetProvisioner)
const TAB_PIPELINE = '📊 Pipeline Review';

// ============================================================================
// FIELD CONFIGURATION
// ============================================================================

// Core deal fields (always visible)
const CORE_FIELDS = [
  { property: 'hs_object_id', header: 'Deal ID', hidden: true, type: 'text' },
  { property: 'dealname', header: 'Deal Name', hyperlink: true, type: 'text' },
  { property: 'dealstage', header: 'Stage', type: 'text' },
  { property: 'notes_last_updated', header: 'Last Activity', type: 'date' },
  { property: 'notes_next_activity_date', header: 'Next Activity', type: 'date' },
  { property: 'next_task_name', header: 'Next Task Name', enabled: false, type: 'text' }, // Future: needs API permissions
  { property: 'why_not_purchase_today_', header: 'Why Not Purchase Today', type: 'text' }
];

// Call quality score fields (color-coded red-yellow-green)
const CALL_QUALITY_FIELDS = [
  { property: 'call_quality_score', header: 'Call Quality Score', colorCode: true, type: 'number' },
  { property: 's_discovery_a_questioning_technique__details', header: 'Questioning', colorCode: true, type: 'number' },
  { property: 's_building_value_a_recap_of_students_needs', header: 'Building Value', colorCode: true, type: 'number' },
  { property: 's_funding_options__a_identifying_funding_needs', header: 'Funding Options', colorCode: true, type: 'number' },
  { property: 's_addressing_objections_a_identifying_and_addressing_objections_and_obstacles', header: 'Addressing Objections', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_ask_for_referral', header: 'Closing the Deal', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_assuming_the_sale', header: 'Ask for Referral', colorCode: true, type: 'number' }
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
    const deals = fetchDealsByOwner(person.email, properties);
    Logger.log(`  Found ${deals.length} deals`);
    
    // Step 3: Build data array
    Logger.log('  Step 3: Building data array...');
    const { dataArray, urlMap } = buildPipelineDataArray(deals);
    
    // Step 4: Clear and write data
    Logger.log('  Step 4: Writing data to sheet...');
    sheet.clear();
    writeDataToSheet(sheet, dataArray, urlMap);
    
    // Step 5: Apply formatting
    Logger.log('  Step 5: Applying formatting...');
    applyPipelineFormatting(sheet, dataArray.length - 1); // -1 for header
    
    // Step 6: Restore preserved notes and formatting
    Logger.log('  Step 6: Restoring preserved data...');
    restorePreservedData(sheet, preserved, dataArray);
    
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
 * @returns {Object} {dataArray, urlMap}
 */
function buildPipelineDataArray(deals) {
  const dataArray = [];
  const urlMap = {}; // Map of row index to URL for hyperlinks
  
  // Headers
  dataArray.push(getPipelineReviewHeaders());
  
  // Data rows
  deals.forEach((deal, index) => {
    const row = [];
    const rowIndex = index + 2; // +2 because row 1 is header, index starts at 0
    
    // Core fields
    CORE_FIELDS.forEach(field => {
      if (field.hidden) {
        // Deal ID - hidden but included
        row.push(deal.id || '');
      } else if (field.property === 'dealname') {
        // Deal Name - will be hyperlinked
        const dealName = extractDealProperty(deal, field.property);
        row.push(dealName);
        urlMap[rowIndex] = buildDealUrl(deal.id);
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
  
  return { dataArray, urlMap };
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
  
  // Apply hyperlinks to Deal Name column (Column B, row 2 onwards)
  Object.keys(urlMap).forEach(rowIndex => {
    const url = urlMap[rowIndex];
    const cell = sheet.getRange(parseInt(rowIndex), 2); // Column B (Deal Name)
    cell.setFormula(`=HYPERLINK("${url}", "${cell.getValue()}")`);
  });
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
  
  // Hide Deal ID column (Column A)
  sheet.hideColumns(1);
  
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
 * @returns {Object} Preserved data indexed by Deal ID
 */
function capturePreservedData(sheet) {
  const preserved = {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return preserved; // No data to preserve
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dealIdCol = headers.indexOf('Deal ID') + 1; // +1 for 1-based
  const note1Col = headers.indexOf('Note 1') + 1;
  const note2Col = headers.indexOf('Note 2') + 1;
  
  if (dealIdCol === 0) {
    Logger.log('  Warning: Deal ID column not found, cannot preserve data');
    return preserved;
  }
  
  // Read all data
  const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const values = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();
  const fontColors = dataRange.getFontColors();
  const fontWeights = dataRange.getFontWeights();
  
  // Capture by Deal ID
  for (let i = 0; i < values.length; i++) {
    const dealId = values[i][dealIdCol - 1]; // -1 for 0-based array index
    
    if (dealId && dealId !== '') {
      preserved[dealId.toString()] = {
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
 * @param {Object} preserved - Preserved data indexed by Deal ID
 * @param {Array<Array>} dataArray - Current data array (to get Deal IDs)
 */
function restorePreservedData(sheet, preserved, dataArray) {
  if (Object.keys(preserved).length === 0) {
    Logger.log('  No preserved data to restore');
    return;
  }
  
  const headers = dataArray[0];
  const dealIdCol = headers.indexOf('Deal ID');
  const note1Col = headers.indexOf('Note 1');
  const note2Col = headers.indexOf('Note 2');
  
  let restoredCount = 0;
  
  // Restore for each row
  for (let i = 1; i < dataArray.length; i++) {
    const dealId = dataArray[i][dealIdCol];
    
    if (dealId && preserved[dealId.toString()]) {
      const data = preserved[dealId.toString()];
      const rowIndex = i + 1; // +1 for 1-based sheet index
      
      // Restore notes
      if (note1Col >= 0 && data.note1) {
        sheet.getRange(rowIndex, note1Col + 1).setValue(data.note1);
      }
      if (note2Col >= 0 && data.note2) {
        sheet.getRange(rowIndex, note2Col + 1).setValue(data.note2);
      }
      
      // Restore formatting for entire row
      const rowRange = sheet.getRange(rowIndex, 1, 1, headers.length);
      if (data.backgrounds) {
        rowRange.setBackgrounds([data.backgrounds]);
      }
      if (data.fontColors) {
        rowRange.setFontColors([data.fontColors]);
      }
      if (data.fontWeights) {
        rowRange.setFontWeights([data.fontWeights]);
      }
      
      restoredCount++;
    }
  }
  
  Logger.log(`  Restored data for ${restoredCount} deals`);
}
