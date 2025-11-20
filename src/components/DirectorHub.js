/**
 * Director Hub Component
 * 
 * Team-wide view for directors to review all deals across their AEs
 * Includes manual flagging system and automatic conditional formatting
 */

// Tab name constant
const TAB_DIRECTOR_HUB = '👔 Director Hub';

// Director priority flags
const PRIORITY_FLAGS = {
  HOT: '🟢',
  COLD: '🔴',
  ATTENTION: '🟡',
  NONE: ''
};

/**
 * Updates the Director Hub with all deals from the team
 * @param {Spreadsheet} controlSheet - The control spreadsheet
 * @param {Array<Object>} salespeople - Array of salesperson configs
 * @returns {Object} Update result
 */
function updateDirectorHub(controlSheet, salespeople) {
  try {
    Logger.log('[Director Hub] Updating...');
    const startTime = new Date();
    
    // Get or create Director Hub sheet
    let sheet = controlSheet.getSheetByName(TAB_DIRECTOR_HUB);
    if (!sheet) {
      sheet = controlSheet.insertSheet(TAB_DIRECTOR_HUB);
      Logger.log('  Created Director Hub sheet');
    }
    
    // Step 1: Capture existing director flags/notes
    Logger.log('  Step 1: Capturing director directives...');
    const directives = captureDirectorDirectives(sheet);
    
    // Step 2: Fetch and aggregate deals from all AEs
    Logger.log('  Step 2: Fetching deals for all AEs...');
    const allDeals = [];
    const properties = getPipelineReviewProperties();
    
    salespeople.forEach(person => {
      Logger.log(`    Fetching for ${person.name}...`);
      const options = {};
      if (person.hubspotUserId && person.hubspotUserId !== '') {
        options.hubspotUserId = person.hubspotUserId;
      }
      
      const deals = fetchDealsByOwner(person.email, properties, options);
      
      // Add owner name to each deal
      deals.forEach(deal => {
        deal.ownerName = person.name;
      });
      
      allDeals.push(...deals);
    });
    
    Logger.log(`  Found ${allDeals.length} total deals across team`);
    
    // Step 3: Build data array with director columns
    Logger.log('  Step 3: Building data array...');
    const { dataArray, urlMap, dealIdMap } = buildDirectorHubDataArray(allDeals);
    
    // Step 4: Clear and write data
    Logger.log('  Step 4: Writing data to sheet...');
    sheet.clear();
    writeDirectorHubData(sheet, dataArray, urlMap);
    
    // Step 5: Apply formatting
    Logger.log('  Step 5: Applying formatting...');
    applyDirectorHubFormatting(sheet, dataArray.length - 1);
    
    // Step 6: Restore director directives
    Logger.log('  Step 6: Restoring director directives...');
    restoreDirectorDirectives(sheet, directives, dealIdMap);
    
    // Step 7: Apply conditional formatting (red for blank Next Activity)
    Logger.log('  Step 7: Applying conditional formatting...');
    applyNextActivityConditionalFormatting(sheet, dataArray.length - 1);
    
    const duration = (new Date() - startTime) / 1000;
    Logger.log(`[Director Hub] Complete (${duration}s)`);
    
    return {
      success: true,
      dealCount: allDeals.length,
      duration: duration
    };
    
  } catch (error) {
    Logger.log(`[Director Hub] Error: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Builds the data array for Director Hub
 * @param {Array<Object>} deals - All deals from all AEs
 * @returns {Object} {dataArray, urlMap, dealIdMap}
 */
function buildDirectorHubDataArray(deals) {
  const dataArray = [];
  const urlMap = {};
  const dealIdMap = {};
  
  // Headers - add Owner, Director Priority, and Director Note
  const headers = [
    'Owner',
    ...getPipelineReviewHeaders(),
    'Director Priority',
    'Director Note'
  ];
  dataArray.push(headers);
  
  // Data rows
  deals.forEach((deal, index) => {
    const row = [];
    const rowIndex = index + 2; // +2 for header and 0-based
    
    // Store Deal ID for preservation
    dealIdMap[rowIndex] = deal.id;
    
    // Owner name (first column)
    row.push(deal.ownerName || '');
    
    // Core fields
    CORE_FIELDS.forEach(field => {
      if (field.property === 'dealname') {
        const dealName = extractDealProperty(deal, field.property);
        row.push(dealName);
        urlMap[rowIndex] = buildDealUrl(deal.id);
      } else if (field.property === 'dealstage' && field.useMapping) {
        const stageId = extractDealProperty(deal, field.property);
        row.push(STAGE_MAP[stageId] || stageId);
      } else if (field.type === 'date') {
        row.push(extractDateProperty(deal, field.property));
      } else if (field.enabled === false) {
        row.push('');
      } else {
        row.push(extractDealProperty(deal, field.property));
      }
    });
    
    // Call quality fields
    CALL_QUALITY_FIELDS.forEach(field => {
      row.push(extractNumericProperty(deal, field.property));
    });
    
    // Manual fields (Note 1, Note 2)
    MANUAL_FIELDS.forEach(() => {
      row.push('');
    });
    
    // Director columns (initially empty)
    row.push(''); // Director Priority
    row.push(''); // Director Note
    
    dataArray.push(row);
  });
  
  return { dataArray, urlMap, dealIdMap };
}

/**
 * Writes data to Director Hub sheet
 * @param {Sheet} sheet - The sheet
 * @param {Array<Array>} dataArray - 2D data array
 * @param {Object} urlMap - Map of row to URL
 */
function writeDirectorHubData(sheet, dataArray, urlMap) {
  if (dataArray.length === 0) return;
  
  // Write all data
  const range = sheet.getRange(1, 1, dataArray.length, dataArray[0].length);
  range.setValues(dataArray);
  
  // Apply hyperlinks to Deal Name column (Column B - after Owner)
  const dealNameCol = 2;
  const rowCount = dataArray.length - 1;
  
  if (rowCount > 0 && Object.keys(urlMap).length > 0) {
    const richTextValues = [];
    
    for (let i = 0; i < rowCount; i++) {
      const rowIndex = i + 2;
      const dealName = dataArray[i + 1][1]; // Column B (index 1)
      const url = urlMap[rowIndex];
      
      if (url && dealName) {
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(dealName)
          .setLinkUrl(url)
          .build();
        richTextValues.push([richText]);
      } else {
        richTextValues.push([dealName]);
      }
    }
    
    sheet.getRange(2, dealNameCol, rowCount, 1).setRichTextValues(richTextValues);
  }
}

/**
 * Applies formatting to Director Hub
 * @param {Sheet} sheet - The sheet
 * @param {number} dataRowCount - Number of data rows
 */
function applyDirectorHubFormatting(sheet, dataRowCount) {
  if (dataRowCount < 1) return;
  
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  
  // Freeze header row and first two columns (Owner + Deal Name)
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  
  // Auto-resize columns
  for (let col = 1; col <= sheet.getLastColumn(); col++) {
    sheet.autoResizeColumn(col);
  }
  
  // Apply call quality conditional formatting
  applyCallQualityFormattingDirectorHub(sheet, dataRowCount);
}

/**
 * Applies conditional formatting to call quality columns in Director Hub
 * @param {Sheet} sheet - The sheet
 * @param {number} dataRowCount - Number of data rows
 */
function applyCallQualityFormattingDirectorHub(sheet, dataRowCount) {
  if (dataRowCount < 1) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rules = [];
  
  CALL_QUALITY_FIELDS.forEach(field => {
    if (field.colorCode) {
      const colIndex = headers.indexOf(field.header) + 1;
      
      if (colIndex > 0) {
        const range = sheet.getRange(2, colIndex, dataRowCount, 1);
        
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .setGradientMinpointWithValue('#F4C7C3', SpreadsheetApp.InterpolationType.NUMBER, '0')
          .setGradientMidpointWithValue('#FCE8B2', SpreadsheetApp.InterpolationType.NUMBER, '2.5')
          .setGradientMaxpointWithValue('#B7E1CD', SpreadsheetApp.InterpolationType.NUMBER, '5')
          .setRanges([range])
          .build();
        
        rules.push(rule);
      }
    }
  });
  
  if (rules.length > 0) {
    sheet.setConditionalFormatRules(rules);
  }
}

/**
 * Applies conditional formatting to Next Activity column (red if blank)
 * @param {Sheet} sheet - The sheet
 * @param {number} dataRowCount - Number of data rows
 */
function applyNextActivityConditionalFormatting(sheet, dataRowCount) {
  if (dataRowCount < 1) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const nextActivityCol = headers.indexOf('Next Activity') + 1;
  
  if (nextActivityCol > 0) {
    const range = sheet.getRange(2, nextActivityCol, dataRowCount, 1);
    
    // Red background if blank
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenCellEmpty()
      .setBackground('#F4C7C3')
      .setRanges([range])
      .build();
    
    const existingRules = sheet.getConditionalFormatRules();
    existingRules.push(rule);
    sheet.setConditionalFormatRules(existingRules);
    
    Logger.log('  Applied blank Next Activity conditional formatting');
  }
}

/**
 * Captures director directives before refresh
 * @param {Sheet} sheet - Director Hub sheet
 * @returns {Object} Directives indexed by Deal ID
 */
function captureDirectorDirectives(sheet) {
  const directives = {};
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return directives;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dealNameCol = headers.indexOf('Deal Name') + 1;
  const priorityCol = headers.indexOf('Director Priority') + 1;
  const noteCol = headers.indexOf('Director Note') + 1;
  
  if (priorityCol === 0 || noteCol === 0) return directives;
  
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const backgrounds = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getBackgrounds();
  
  for (let i = 0; i < data.length; i++) {
    const dealName = data[i][dealNameCol - 1];
    const priority = data[i][priorityCol - 1];
    const note = data[i][noteCol - 1];
    const rowBackground = backgrounds[i];
    
    if (dealName && (priority || note)) {
      directives[dealName] = {
        priority: priority,
        note: note,
        background: rowBackground
      };
    }
  }
  
  Logger.log(`  Captured ${Object.keys(directives).length} directives`);
  return directives;
}

/**
 * Restores director directives after refresh
 * @param {Sheet} sheet - Director Hub sheet
 * @param {Object} directives - Directives indexed by Deal Name
 * @param {Object} dealIdMap - Map of row index to Deal ID
 */
function restoreDirectorDirectives(sheet, directives, dealIdMap) {
  if (Object.keys(directives).length === 0) {
    Logger.log('  No directives to restore');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const dealNameCol = headers.indexOf('Deal Name') + 1;
  const priorityCol = headers.indexOf('Director Priority') + 1;
  const noteCol = headers.indexOf('Director Note') + 1;
  
  const dealNames = sheet.getRange(2, dealNameCol, lastRow - 1, 1).getValues();
  
  let restoredCount = 0;
  
  for (let i = 0; i < dealNames.length; i++) {
    const dealName = dealNames[i][0];
    const rowIndex = i + 2;
    
    if (dealName && directives[dealName]) {
      const directive = directives[dealName];
      
      // Restore priority and note
      if (directive.priority) {
        sheet.getRange(rowIndex, priorityCol).setValue(directive.priority);
      }
      if (directive.note) {
        sheet.getRange(rowIndex, noteCol).setValue(directive.note);
      }
      
      // Restore row background
      if (directive.background) {
        const rowRange = sheet.getRange(rowIndex, 1, 1, headers.length);
        rowRange.setBackgrounds([directive.background]);
      }
      
      restoredCount++;
    }
  }
  
  Logger.log(`  Restored ${restoredCount} directives`);
}

