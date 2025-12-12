/**
 * Director Hub Component
 * 
 * Team-wide view for directors to review all deals across their AEs
 * Includes manual flagging system and automatic conditional formatting
 */

// Tab name constant
const TAB_DIRECTOR_HUB = '👔 Director Hub';

// Stage mapping (ID to Name)
const STAGE_MAP = {
  '90284257': 'Create Curiosity',
  '90284258': 'Needs Analysis',
  '90284259': 'Demonstrating Value',
  '90284260': 'Partnership Proposal',
  '90284261': 'Negotiation',
  '90284262': 'Partnership Confirmed'
};

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
    const properties = PIPELINE_PROPERTIES; // Defined in PipelineReview.js
    
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
  
  // Headers - add Owner at the beginning
  // ORDER: Owner → Deal Name → Stage → [Call Quality] → Activity Dates
  const headers = [
    'Owner',
    'Deal ID',
    'Deal Name',
    'Stage',
    // Call Quality columns
    'Questioning',
    'Building Value',
    'Funding Options',
    'Addressing Objections',
    'Closing the Deal',
    'Ask for Referral',
    // Activity dates and other info
    'Last Activity',
    'Next Activity',
    'Why Not Purchase Today'
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
    
    // First 5 core fields (B-F)
    for (let i = 0; i < 5 && i < CORE_FIELDS.length; i++) {
      const field = CORE_FIELDS[i];
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
    }
    
    // Director fields at positions 7-8 (G-H) - initially empty
    DIRECTOR_FIELDS.forEach(() => {
      row.push('');
    });
    
    // Remaining core fields (from position 6 onwards)
    for (let i = 5; i < CORE_FIELDS.length; i++) {
      const field = CORE_FIELDS[i];
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
    }
    
    // Call quality fields
    CALL_QUALITY_FIELDS.forEach(field => {
      row.push(extractNumericProperty(deal, field.property));
    });
    
    // Manual fields (Note 1, Note 2)
    MANUAL_FIELDS.forEach(() => {
      row.push('');
    });
    
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
  
  // Set width and text wrapping for specific columns
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const whyNotPurchaseCol = headers.indexOf('Why Not Purchase Today') + 1;
  const callsHistoryCol = headers.indexOf('calls history') + 1;
  
  if (whyNotPurchaseCol > 0) {
    sheet.setColumnWidth(whyNotPurchaseCol, 250); // 250 pixels width
    if (dataRowCount > 0) {
      sheet.getRange(2, whyNotPurchaseCol, dataRowCount, 1).setWrap(true);
    }
  }
  
  if (callsHistoryCol > 0) {
    sheet.setColumnWidth(callsHistoryCol, 250); // 250 pixels width
    if (dataRowCount > 0) {
      sheet.getRange(2, callsHistoryCol, dataRowCount, 1).setWrap(true);
    }
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

// ============================================================================
// MULTI-DIRECTOR CONSOLIDATED PIPELINE
// ============================================================================

/**
 * Updates all director consolidated pipeline tabs
 * Bi-directional sync: AE notes → Director, Director highlighting → AE
 */
function updateAllDirectorConsolidatedPipelines() {
  try {
    Logger.log('[Director Pipelines] Updating all director consolidated pipelines...');
    const startTime = new Date();
    
    // Load configuration
    const config = loadConfiguration();
    const directors = getDirectorConfig();
    
    if (directors.length === 0) {
      Logger.log('[Director Pipelines] No directors configured, skipping');
      return { success: true, directorCount: 0 };
    }
    
    const controlSheet = SpreadsheetApp.openById(CONTROL_SHEET_ID);
    
    // Update each director's tab
    let successCount = 0;
    directors.forEach(director => {
      Logger.log(`  Processing ${director.type} - ${director.name} (${director.team} team)...`);
      
      try {
        updateDirectorConsolidatedPipeline(director, config, controlSheet);
        successCount++;
      } catch (e) {
        Logger.log(`    Error updating ${director.name}: ${e.message}`);
      }
    });
    
    const duration = (new Date() - startTime) / 1000;
    Logger.log(`[Director Pipelines] Complete: ${successCount}/${directors.length} directors updated (${duration}s)`);
    
    return {
      success: true,
      directorCount: successCount,
      duration: duration
    };
    
  } catch (error) {
    Logger.log(`[Director Pipelines] Error: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Updates a single director's consolidated pipeline tab
 * @param {Object} director - Director config
 * @param {Object} config - Main configuration
 * @param {Spreadsheet} controlSheet - Control sheet
 */
function updateDirectorConsolidatedPipeline(director, config, controlSheet) {
  // Get or create director's tab
  let sheet = controlSheet.getSheetByName(director.tabName);
  if (!sheet) {
    sheet = controlSheet.insertSheet(director.tabName);
    Logger.log(`    Created tab: ${director.tabName}`);
  }
  
  // Step 1: Capture director's existing highlighting
  Logger.log(`    Capturing director highlighting...`);
  const preservedHighlighting = captureDirectorHighlighting(sheet);
  
  // Step 2: Get all AEs from this director's team
  const teamAEs = config.salespeople.filter(p => p.team === director.team);
  Logger.log(`    Found ${teamAEs.length} AEs in ${director.team} team`);
  
  if (teamAEs.length === 0) {
    Logger.log(`    No AEs in team, skipping`);
    return;
  }
  
  // Step 3: Fetch deals from HubSpot for this team
  const allDeals = [];
  teamAEs.forEach(person => {
    try {
      const deals = fetchDealsForAE(person);
      deals.forEach(deal => {
        deal.ownerName = person.name;
      });
      allDeals.push(...deals);
    } catch (e) {
      Logger.log(`      Error fetching deals for ${person.name}: ${e.message}`);
    }
  });
  
  Logger.log(`    Fetched ${allDeals.length} deals from team`);
  
  // Step 3.5: Sort deals by priority (most recent first) and limit to prevent timeout
  // Priority: Recent Next Activity > Recent Last Activity > Stage
  const MAX_DIRECTOR_DEALS = 200; // Limit to prevent timeout with large teams
  allDeals.sort((a, b) => {
    const aNext = a.properties?.notes_next_activity_date || '';
    const bNext = b.properties?.notes_next_activity_date || '';
    const aLast = a.properties?.notes_last_updated || '';
    const bLast = b.properties?.notes_last_updated || '';
    
    // Prioritize deals with upcoming next activity
    if (aNext && !bNext) return -1;
    if (!aNext && bNext) return 1;
    if (aNext && bNext) return bNext.localeCompare(aNext); // Most recent first
    
    // Then by last activity
    return bLast.localeCompare(aLast); // Most recent first
  });
  
  if (allDeals.length > MAX_DIRECTOR_DEALS) {
    Logger.log(`    ⚠️ Limiting to ${MAX_DIRECTOR_DEALS} most recent deals (from ${allDeals.length})`);
    allDeals.splice(MAX_DIRECTOR_DEALS);
  }
  
  // Step 4: Collect notes from all individual AE sheets
  Logger.log(`    Collecting notes from AE sheets...`);
  const notesMap = collectNotesFromTeamAEs(teamAEs);
  
  // Step 5: Build consolidated data array
  Logger.log(`    Building data array for ${allDeals.length} deals...`);
  const dataArray = buildConsolidatedPipelineDataArray(allDeals, notesMap);
  Logger.log(`    Built array: ${dataArray.length} rows × ${dataArray[0].length} columns`);
  
  // Step 6: Write data to director's tab
  Logger.log(`    Writing data to sheet...`);
  sheet.clear();
  writeConsolidatedPipelineData(sheet, dataArray);
  Logger.log(`    Data written`);
  
  // Step 7: Apply formatting and restore highlighting
  Logger.log(`    Applying formatting...`);
  applyConsolidatedPipelineFormatting(sheet, dataArray, preservedHighlighting);
  Logger.log(`    Formatting applied`);
  
  Logger.log(`    ✓ ${director.tabName} updated`);
}

/**
 * Captures director's highlighting by Deal ID
 * @param {Sheet} sheet - Director's tab
 * @returns {Map} Map of Deal ID to highlighting data
 */
function captureDirectorHighlighting(sheet) {
  const highlightMap = new Map();
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return highlightMap;
  
  const lastCol = sheet.getLastColumn();
  
  // Read Deal IDs (column A, hidden)
  const dealIds = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  // Read backgrounds and font colors for entire rows
  const backgrounds = sheet.getRange(2, 1, lastRow - 1, lastCol).getBackgrounds();
  const fontColors = sheet.getRange(2, 1, lastRow - 1, lastCol).getFontColors();
  
  dealIds.forEach((row, index) => {
    const dealId = row[0]?.toString();
    if (!dealId) return;
    
    highlightMap.set(dealId, {
      backgrounds: backgrounds[index],
      fontColors: fontColors[index]
    });
  });
  
  Logger.log(`    Captured highlighting for ${highlightMap.size} deals`);
  return highlightMap;
}

/**
 * Collects notes from all team AEs' individual sheets
 * @param {Array} teamAEs - Array of AE configs
 * @returns {Map} Map of Deal ID to notes
 */
function collectNotesFromTeamAEs(teamAEs) {
  const notesMap = new Map();
  
  teamAEs.forEach(person => {
    try {
      if (!person.sheetId) {
        Logger.log(`      No sheet ID for ${person.name}, skipping`);
        return;
      }
      
      const aeSheet = SpreadsheetApp.openById(person.sheetId);
      const pipelineSheet = aeSheet.getSheetByName(TAB_PIPELINE);
      
      if (!pipelineSheet) {
        Logger.log(`      No Pipeline Review tab for ${person.name}, skipping`);
        return;
      }
      
      const lastRow = pipelineSheet.getLastRow();
      if (lastRow < 2) return;
      
      // Read Deal IDs and Notes (Notes is in column 6, not last column)
      // Column order: Deal ID | Deal Name | Stage | Last Activity | Next Activity | Notes | Why Not Purchase | [Call Quality 13 fields]
      const NOTES_COLUMN_INDEX = 5; // Column F (0-based index)
      const data = pipelineSheet.getRange(2, 1, lastRow - 1, Math.max(6, pipelineSheet.getLastColumn())).getValues();
      
      data.forEach(row => {
        const dealId = row[0]?.toString();
        const notes = row[NOTES_COLUMN_INDEX] || ''; // Column 6 (index 5)
        
        if (dealId && notes) {
          notesMap.set(dealId, notes);
        }
      });
      
    } catch (e) {
      Logger.log(`      Error collecting notes from ${person.name}: ${e.message}`);
    }
  });
  
  Logger.log(`    Collected notes for ${notesMap.size} deals`);
  return notesMap;
}

/**
 * Builds consolidated pipeline data array
 * @param {Array} allDeals - All deals from team
 * @param {Map} notesMap - Notes collected from AEs
 * @returns {Array} 2D data array
 */
function buildConsolidatedPipelineDataArray(allDeals, notesMap) {
  const dataArray = [];
  
  // Headers - Optimized order per director request
  // ORDER: Deal Name → Owner → Stage → Last Activity → Next Activity → Notes → Why Not Purchase → [Call Quality 13]
  const headers = [
    'Deal ID', // Hidden (A)
    'Deal Name', // (B)
    'Owner', // AE Name (C)
    'Stage', // (D)
    'Last Activity', // (E)
    'Next Activity', // (F)
    'Notes', // From AE (G)
    'Why Not Purchase Today', // (H)
    // Call Quality columns (matching PipelineReview.js order - ALL 13 fields) (I-U)
    'Questioning',
    'Trust',
    'Recap Needs',
    'Building Value',
    'Program Alignment',
    'Requirements',
    'Funding Needs',
    'Funding Solution',
    'Funding Commitment',
    'Objections',
    'Urgency',
    'Assume Sale',
    'Referral'
  ];
  
  dataArray.push(headers);
  
  // Data rows
  allDeals.forEach(deal => {
    const dealId = deal.id.toString();
    const properties = deal.properties || {};
    
    // Map stage ID to stage name
    const stageId = properties.dealstage || '';
    const stageName = STAGE_MAP[stageId] || stageId; // Fallback to ID if not found
    
    const row = [
      dealId, // A - Deal ID
      properties.dealname || '', // B - Deal Name
      deal.ownerName || '', // C - Owner (AE Name)
      stageName, // D - Stage (NAME, not ID)
      formatDate(properties.notes_last_updated), // E - Last Activity
      formatDate(properties.notes_next_activity_date), // F - Next Activity
      notesMap.get(dealId) || '', // G - Notes from AE
      properties.why_not_purchase_today_ || '', // H - Why Not Purchase Today
      // Call Quality scores (ALL 13 fields) I-U
      properties.s_discovery_a_questioning_technique || '',
      properties.s_discovery_a_empathy__rapport_building_and_active_listening || '',
      properties.s_building_value_a_recap_of_students_needs || '',
      properties.s_building_value_a_tailoring_features_and_benefits || '',
      properties.s_gaining_an_affirmation_and_program_requirements__a_gaining_affirmation || '',
      properties.s_gaining_an_affirmation_and_program_requirements__a_essential_program_requirements || '',
      properties.s_funding_options__a_identifying_funding_needs || '',
      properties.s_funding_options__a_presenting_funding_solutions || '',
      properties.s_funding_options__a_securing_financial_commitment || '',
      properties.s_addressing_objections_a_identifying_and_addressing_objections_and_obstacles || '',
      properties.s_closing_the_deal__a_creating_a_sense_of_urgency || '',
      properties.s_closing_the_deal__a_assuming_the_sale || '',
      properties.s_closing_the_deal__a_ask_for_referral || ''
    ];
    
    dataArray.push(row);
  });
  
  return dataArray;
}

/**
 * Writes consolidated pipeline data to director's tab
 * @param {Sheet} sheet - Director's tab
 * @param {Array} dataArray - 2D data array
 */
function writeConsolidatedPipelineData(sheet, dataArray) {
  if (dataArray.length === 0) return;
  
  // Write all data at once (FAST)
  sheet.getRange(1, 1, dataArray.length, dataArray[0].length).setValues(dataArray);
  
  // Add hyperlinks for Deal Name column (Column B) - BATCH OPERATION
  const richTextValues = [];
  for (let i = 1; i < dataArray.length; i++) { // Skip header (index 0)
    const dealId = dataArray[i][0]; // Deal ID (Column A)
    const dealName = dataArray[i][1]; // Deal Name (Column B)
    
    if (dealId && dealName) {
      const dealUrl = `https://app-eu1.hubspot.com/contacts/25196166/record/0-3/${dealId}`;
      const richText = SpreadsheetApp.newRichTextValue()
        .setText(dealName)
        .setLinkUrl(dealUrl)
        .build();
      richTextValues.push([richText]);
    } else {
      richTextValues.push([SpreadsheetApp.newRichTextValue().setText(dealName || '').build()]);
    }
  }
  
  // Apply all hyperlinks at once (FAST - single API call instead of 300+)
  if (richTextValues.length > 0) {
    sheet.getRange(2, 2, richTextValues.length, 1).setRichTextValues(richTextValues);
  }
}

/**
 * Applies formatting to consolidated pipeline and restores highlighting
 * @param {Sheet} sheet - Director's tab
 * @param {Array} dataArray - 2D data array
 * @param {Map} highlightMap - Preserved highlighting
 */
function applyConsolidatedPipelineFormatting(sheet, dataArray, highlightMap) {
  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, dataArray[0].length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#4285F4')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4); // Freeze Deal ID, Deal Name, Owner, Notes
  
  // Hide Deal ID column
  sheet.hideColumns(1);
  
  // Apply call quality color coding
  applyCallQualityFormattingToConsolidated(sheet, dataArray);
  
  // Restore director's highlighting (by Deal ID)
  restoreDirectorHighlighting(sheet, dataArray, highlightMap);
}

/**
 * Applies call quality color coding to consolidated pipeline
 * Uses conditional formatting rules (FAST - single operation per column)
 * @param {Sheet} sheet - Director's tab
 * @param {Array} dataArray - 2D data array
 */
function applyCallQualityFormattingToConsolidated(sheet, dataArray) {
  const headers = dataArray[0];
  const dataRowCount = dataArray.length - 1; // Exclude header
  if (dataRowCount < 1) return;
  
  // Find call quality column indices (ALL 13 fields)
  const callQualityHeaders = [
    'Questioning',
    'Trust',
    'Recap Needs',
    'Building Value',
    'Program Alignment',
    'Requirements',
    'Funding Needs',
    'Funding Solution',
    'Funding Commitment',
    'Objections',
    'Urgency',
    'Assume Sale',
    'Referral'
  ];
  
  const rules = [];
  
  callQualityHeaders.forEach(header => {
    const colIndex = headers.indexOf(header) + 1;
    if (colIndex === 0) return;
    
    const range = sheet.getRange(2, colIndex, dataRowCount, 1);
    
    // Gradient: Red (0) → Yellow (2.5) → Green (5)
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue('#F4C7C3', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setGradientMidpointWithValue('#FCE8B2', SpreadsheetApp.InterpolationType.NUMBER, '2.5')
      .setGradientMaxpointWithValue('#B7E1CD', SpreadsheetApp.InterpolationType.NUMBER, '5')
      .setRanges([range])
      .build();
    
    rules.push(rule);
  });
  
  if (rules.length > 0) {
    // Get existing rules and append new ones
    const existingRules = sheet.getConditionalFormatRules();
    sheet.setConditionalFormatRules(existingRules.concat(rules));
  }
}

/**
 * Restores director's highlighting by Deal ID
 * @param {Sheet} sheet - Director's tab
 * @param {Array} dataArray - 2D data array
 * @param {Map} highlightMap - Preserved highlighting
 */
function restoreDirectorHighlighting(sheet, dataArray, highlightMap) {
  if (highlightMap.size === 0) return;
  
  for (let row = 2; row <= dataArray.length; row++) {
    const dealId = sheet.getRange(row, 1).getValue()?.toString();
    if (!dealId) continue;
    
    const highlighting = highlightMap.get(dealId);
    if (!highlighting) continue;
    
    const rowRange = sheet.getRange(row, 1, 1, dataArray[0].length);
    
    if (highlighting.backgrounds) {
      rowRange.setBackgrounds([highlighting.backgrounds]);
    }
    
    if (highlighting.fontColors) {
      rowRange.setFontColors([highlighting.fontColors]);
    }
  }
  
  Logger.log(`    Restored highlighting for ${highlightMap.size} deals`);
}

