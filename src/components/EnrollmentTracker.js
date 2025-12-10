/**
 * Enrollment Tracker Component
 * 
 * Displays enrollment data from HubSpot with rolling monthly history
 * - Fetches current month + last month enrollments
 * - Preserves older historical data
 * - Splits enrollments from Easy Starts
 * 
 * Note: Uses TAB_ENROLLMENT constant from SheetProvisioner.js
 */

// ============================================================================
// FIELD CONFIGURATION
// ============================================================================

// Colors for tier progress bar
const TIER_COLORS = {
  BACKGROUND: '#fff3cd',         // Light yellow background for progress
  BACKGROUND_MAX: '#d4edda',     // Light green background for max tier
  TEXT: '#000000'                // Black text
};

// Enrollment fields (displayed for both Regular and Easy Start enrollments)
const ENROLLMENT_FIELDS = [
  { property: 'dealname', header: 'Deal Name', hyperlink: true, type: 'text' },
  { property: 'platform_email', header: 'Platform Email', type: 'text' },
  { property: 'program', header: 'Program', type: 'text' },
  { property: 'start_date', header: 'Cohort Start Date', type: 'date' }, // Note: property name is 'start_date' in HubSpot
  { property: 'closedate', header: 'Close Date', type: 'date' },
  { property: 'warm_handoff_scoring', header: 'Warm Handoff', colorCode: true, type: 'number' },
  { property: 's_closing_the_deal__a_ask_for_referral', header: 'Ask for Referral', colorCode: true, type: 'number' }
];

// Additional field for Easy Starts table
const EASY_START_FIELD = { property: 'easy_start_option', header: 'Easy Start Option', type: 'text' };

// Easy Start values that qualify as "Easy Start" (NOT regular enrollment)
const EASY_START_VALUES = ['Waiting for the start', 'Started', 'Didn\'t pay'];

/**
 * Gets all properties to fetch from HubSpot
 * @returns {Array<string>} Array of property names
 */
function getEnrollmentProperties() {
  const properties = [];
  
  // Add enrollment fields
  ENROLLMENT_FIELDS.forEach(field => {
    properties.push(field.property);
  });
  
  // Add easy_start_option for filtering
  properties.push('easy_start_option');
  
  // Add standard fields always needed
  if (!properties.includes('hubspot_owner_id')) {
    properties.push('hubspot_owner_id');
  }
  
  return properties;
}

/**
 * Gets column headers for regular enrollments
 * @returns {Array<string>} Array of header names
 */
function getRegularEnrollmentHeaders() {
  return ENROLLMENT_FIELDS.map(field => field.header);
}

/**
 * Gets column headers for Easy Starts (includes Easy Start Option)
 * @returns {Array<string>} Array of header names
 */
function getEasyStartHeaders() {
  const headers = ENROLLMENT_FIELDS.map(field => field.header);
  // Insert Easy Start Option after Deal Name
  headers.splice(1, 0, EASY_START_FIELD.header);
  return headers;
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Formats a percentage value for display
 * Handles both decimal (0.07) and percentage (7) formats
 * @param {number|string} value - Percentage value
 * @returns {string} Formatted percentage (e.g., "7%")
 */
function formatPercentDisplay(value) {
  if (!value && value !== 0) return '';
  
  let numValue = typeof value === 'number' ? value : parseFloat(value);
  
  if (isNaN(numValue)) return '';
  
  // If value is between 0 and 1, it's likely a decimal (0.07 = 7%)
  if (numValue > 0 && numValue < 1) {
    numValue = numValue * 100;
  }
  
  // Round to 1 decimal place if needed
  const rounded = Math.round(numValue * 10) / 10;
  
  // Remove .0 if it's a whole number
  return rounded % 1 === 0 ? `${Math.round(rounded)}%` : `${rounded}%`;
}

// ============================================================================
// TIER PROGRESS CALCULATION
// ============================================================================

/**
 * Calculates tier progress for a person based on their role and enrollment count
 * @param {string} role - Person's role (e.g., "SrAE", "AE", "CAE")
 * @param {number} enrollmentCount - Current enrollment count
 * @param {Map} tierLevels - Tier levels configuration
 * @returns {Object} Progress info {currentTier, progress, nextTierAt, percent, message}
 */
function calculateTierProgress(role, enrollmentCount, tierLevels) {
  if (!role || !tierLevels.has(role)) {
    return null; // No tier data for this role
  }
  
  const tiers = tierLevels.get(role);
  if (!tiers || tiers.length === 0) {
    return null;
  }
  
  // Find current tier
  let currentTier = null;
  let nextTier = null;
  
  for (let i = 0; i < tiers.length; i++) {
    const tier = tiers[i];
    
    if (enrollmentCount >= tier.min && enrollmentCount <= tier.max) {
      currentTier = tier;
      nextTier = i < tiers.length - 1 ? tiers[i + 1] : null;
      break;
    }
  }
  
  // If not in any tier, check if below first tier
  if (!currentTier && enrollmentCount < tiers[0].min) {
    nextTier = tiers[0];
    
    return {
      currentTier: 0,
      currentTierInfo: null,
      nextTierInfo: nextTier,
      progress: 0,
      nextTierAt: nextTier.min,
      remaining: nextTier.min - enrollmentCount,
      percent: 0,
      message: `${enrollmentCount} enrollments | ${nextTier.min - enrollmentCount} more to reach Tier 1 (${formatPercentDisplay(nextTier.percent)})`,
      allTiers: tiers,
      enrollmentCount: enrollmentCount
    };
  }
  
  // Calculate progress
  if (!currentTier) {
    return null; // Shouldn't happen
  }
  
  // If we're in the max tier (tier 3)
  if (!nextTier || currentTier.max === Infinity) {
    return {
      currentTier: currentTier.tier,
      currentTierInfo: currentTier,
      nextTierInfo: null,
      progress: 100,
      nextTierAt: null,
      remaining: 0,
      percent: 100,
      message: `${enrollmentCount} enrollments | MAX TIER ${currentTier.tier} 🎉 | ${formatPercentDisplay(currentTier.percent)} commission`,
      allTiers: tiers,
      enrollmentCount: enrollmentCount
    };
  }
  
  // Calculate progress within current tier towards next tier
  const tierRange = nextTier.min - currentTier.min;
  const progressInTier = enrollmentCount - currentTier.min;
  const progressPercent = Math.min(100, Math.round((progressInTier / tierRange) * 100));
  
  // For display, show progress towards next tier
  const progressToNext = enrollmentCount - currentTier.min;
  const totalToNext = nextTier.min - currentTier.min;
  const percentToNext = Math.min(100, Math.round((progressToNext / totalToNext) * 100));
  
  return {
    currentTier: currentTier.tier,
    currentTierInfo: currentTier,
    nextTierInfo: nextTier,
    progress: percentToNext,
    nextTierAt: nextTier.min,
    remaining: nextTier.min - enrollmentCount,
    percent: percentToNext,
    message: `Tier ${currentTier.tier} | ${enrollmentCount} enrollments | ${percentToNext}% to Tier ${nextTier.tier} 🎯 | +${nextTier.min - enrollmentCount} for ${formatPercentDisplay(nextTier.percent)} commission!`,
    allTiers: tiers,
    enrollmentCount: enrollmentCount
  };
}

// ============================================================================
// MAIN ENROLLMENT TRACKER FUNCTION
// ============================================================================

/**
 * Updates the Enrollment Tracker tab for a salesperson
 * @param {Spreadsheet} individualSheet - The salesperson's individual sheet
 * @param {Object} person - Person object {name, email}
 * @returns {Object} Update result
 */
function updateEnrollmentTracker(individualSheet, person) {
  try {
    Logger.log(`[Enrollment Tracker] Updating for ${person.name}...`);
    
    const startTime = new Date();
    const sheet = individualSheet.getSheetByName(TAB_ENROLLMENT);
    
    if (!sheet) {
      throw new Error(`Enrollment Tracker tab not found for ${person.name}`);
    }
    
    // Step 1: Capture historical data (older than last month)
    Logger.log('  Step 1: Capturing historical data...');
    const historicalData = captureHistoricalData(sheet);
    
    // Step 2: Fetch enrollment deals from HubSpot (current + last month)
    Logger.log('  Step 2: Fetching enrollment deals from HubSpot...');
    const properties = getEnrollmentProperties();
    const options = {};
    
    // Pass HubSpot User ID if available
    if (person.hubspotUserId && person.hubspotUserId !== '') {
      options.hubspotUserId = person.hubspotUserId;
    }
    
    const deals = fetchEnrollmentDeals(person.email, properties, options);
    Logger.log(`  Found ${deals.length} enrollment deals`);
    
    // Debug: Log properties of first deal to verify property names
    if (deals.length > 0) {
      Logger.log(`  DEBUG - First deal properties: ${Object.keys(deals[0].properties || {}).join(', ')}`);
      Logger.log(`  DEBUG - start_date value: ${deals[0].properties?.start_date || 'EMPTY'}`);
      Logger.log(`  DEBUG - cohort_start_date value: ${deals[0].properties?.cohort_start_date || 'EMPTY'}`);
    }
    
    // Step 3: Group deals by month and type
    Logger.log('  Step 3: Grouping deals by month...');
    const groupedData = groupDealsByMonth(deals);
    
    // Step 4: Calculate tier progress (if role is available)
    Logger.log('  Step 4: Calculating tier progress...');
    let tierProgress = null;
    if (person.role && person.role !== '') {
      const tierLevels = getTierLevels();
      
      // Count ONLY CURRENT MONTH enrollments (not Easy Starts)
      // Get current month key (YYYY-MM format)
      const now = new Date();
      const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
      
      // Get current month's regular enrollments (Easy Starts excluded)
      const currentMonthEnrollments = groupedData[currentMonthKey]?.regularEnrollments?.length || 0;
      
      Logger.log(`  Current month (${currentMonthKey}): ${currentMonthEnrollments} enrollments`);
      
      tierProgress = calculateTierProgress(person.role, currentMonthEnrollments, tierLevels);
      if (tierProgress) {
        Logger.log(`  Tier progress: ${tierProgress.message}`);
      }
    }
    
    // Step 5: Build sheet data
    Logger.log('  Step 5: Building sheet data...');
    const { dataArray, urlMap } = buildEnrollmentDataArray(groupedData, historicalData, tierProgress);
    
    // Step 6: Clear and write data
    Logger.log('  Step 6: Writing data to sheet...');
    sheet.clear();
    writeEnrollmentDataToSheet(sheet, dataArray, urlMap);
    
    // Step 7: Apply formatting
    Logger.log('  Step 7: Applying formatting...');
    applyEnrollmentFormatting(sheet, dataArray, tierProgress);
    
    const duration = (new Date() - startTime) / 1000;
    Logger.log(`[Enrollment Tracker] Complete for ${person.name} (${duration}s)`);
    
    return {
      success: true,
      enrollmentCount: deals.length,
      duration: duration
    };
    
  } catch (error) {
    Logger.log(`[Enrollment Tracker] Error for ${person.name}: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================================================
// DATA GROUPING
// ============================================================================

/**
 * Groups deals by month and splits into Regular vs Easy Start
 * @param {Array<Object>} deals - Raw deals from HubSpot
 * @returns {Object} Grouped data by month
 */
function groupDealsByMonth(deals) {
  const grouped = {};
  
  deals.forEach(deal => {
    const closeDateValue = extractDealProperty(deal, 'closedate');
    
    if (!closeDateValue || closeDateValue === '') {
      Logger.log(`  Warning: Deal ${deal.id} has no close date, skipping`);
      return;
    }
    
    // Parse close date
    let closeDate;
    try {
      if (typeof closeDateValue === 'string') {
        closeDate = new Date(closeDateValue);
      } else if (closeDateValue instanceof Date) {
        closeDate = closeDateValue;
      } else {
        const timestamp = parseInt(closeDateValue.toString());
        closeDate = new Date(timestamp);
      }
      
      if (isNaN(closeDate.getTime())) {
        Logger.log(`  Warning: Deal ${deal.id} has invalid close date: ${closeDateValue}`);
        return;
      }
    } catch (e) {
      Logger.log(`  Warning: Deal ${deal.id} close date parse error: ${e.message}`);
      return;
    }
    
    // Create month key (YYYY-MM)
    const monthKey = `${closeDate.getFullYear()}-${String(closeDate.getMonth() + 1).padStart(2, '0')}`;
    
    if (!grouped[monthKey]) {
      grouped[monthKey] = {
        year: closeDate.getFullYear(),
        month: closeDate.getMonth(),
        regularEnrollments: [],
        easyStarts: []
      };
    }
    
    // Determine if this is an Easy Start
    const easyStartOption = extractDealProperty(deal, 'easy_start_option');
    const isEasyStart = easyStartOption && EASY_START_VALUES.includes(easyStartOption);
    
    if (isEasyStart) {
      grouped[monthKey].easyStarts.push(deal);
    } else {
      grouped[monthKey].regularEnrollments.push(deal);
    }
  });
  
  // Sort months descending (newest first)
  const sortedMonths = Object.keys(grouped).sort((a, b) => b.localeCompare(a));
  
  const result = {};
  sortedMonths.forEach(key => {
    result[key] = grouped[key];
  });
  
  Logger.log(`  Grouped into ${sortedMonths.length} months`);
  return result;
}

// ============================================================================
// DATA BUILDING
// ============================================================================

/**
 * Builds the data array for Enrollment Tracker sheet
 * @param {Object} groupedData - Deals grouped by month
 * @param {Array} historicalData - Preserved historical data
 * @param {Object} tierProgress - Tier progress information (optional)
 * @returns {Object} {dataArray, urlMap}
 */
function buildEnrollmentDataArray(groupedData, historicalData, tierProgress) {
  const dataArray = [];
  const urlMap = {}; // Map of row index to URL for hyperlinks
  let currentRow = 1;
  
  // Add tier progress bar at the top if available
  if (tierProgress && tierProgress.allTiers && tierProgress.allTiers.length > 0) {
    // CREATE A SINGLE CONTINUOUS PROGRESS BAR WITH EMBEDDED TIER NUMBERS
    const allTiers = tierProgress.allTiers;
    const currentEnrollments = tierProgress.enrollmentCount || 0;
    
    // Find max tier threshold (use the last tier's min since max might be Infinity)
    const maxTier = allTiers[allTiers.length - 1];
    const maxThreshold = maxTier.min; // Use min of last tier as the max threshold
    
    // Build progress bar with embedded numbers
    const totalLength = 50; // Total character length for the bar
    
    // Calculate positions for each tier marker
    const tierPositions = allTiers.map(tier => ({
      tier: tier.tier,
      value: tier.min,
      position: Math.round((tier.min / maxThreshold) * totalLength),
      percent: formatPercentDisplay(tier.percent)
    }));
    
    // Build the progress bar character by character
    let progressBar = '[';
    const currentPosition = Math.round((currentEnrollments / maxThreshold) * totalLength);
    
    for (let i = 0; i <= totalLength; i++) {
      // Check if there's a tier marker at this position
      const tierAtPosition = tierPositions.find(t => {
        const numStr = t.value.toString();
        const numLength = numStr.length;
        // Check if this is the start position of this number
        return t.position === i;
      });
      
      if (tierAtPosition) {
        // Insert the tier number
        const numStr = tierAtPosition.value.toString();
        progressBar += numStr;
        i += numStr.length - 1; // Skip ahead by the length of the number
      } else {
        // Regular progress bar character
        if (i < currentPosition) {
          progressBar += '▓'; // Filled
        } else {
          progressBar += '░'; // Empty
        }
      }
    }
    
    progressBar += ']';
    
    // Determine current status message
    let statusMsg = '';
    if (tierProgress.currentTier === 0) {
      // Below first tier
      const nextTier = allTiers[0];
      const remaining = nextTier.min - currentEnrollments;
      statusMsg = `${currentEnrollments} enrollments ║ +${remaining} to reach Tier 1 (${formatPercentDisplay(nextTier.percent)})`;
    } else if (tierProgress.nextTierInfo) {
      // In a tier, working towards next
      const remaining = tierProgress.remaining;
      const nextTier = tierProgress.nextTierInfo.tier;
      const nextPercent = formatPercentDisplay(tierProgress.nextTierInfo.percent);
      statusMsg = `TIER ${tierProgress.currentTier} ║ ${currentEnrollments} enrollments ║ +${remaining} to Tier ${nextTier} (${nextPercent})`;
    } else {
      // Max tier achieved
      const commission = formatPercentDisplay(tierProgress.currentTierInfo.percent);
      statusMsg = `🏆 MAX TIER ${tierProgress.currentTier} ║ ${currentEnrollments} enrollments ║ ${commission} Commission`;
    }
    
    // Assemble full display
    const displayText = `🎯 ${statusMsg}  ${progressBar}`;
    
    dataArray.push([displayText]);
    currentRow++;
    
    // Empty row for spacing
    dataArray.push([]);
    currentRow++;
  }
  
  const months = Object.keys(groupedData);
  
  // Build current month and last month from HubSpot data
  months.forEach((monthKey, idx) => {
    const monthData = groupedData[monthKey];
    const monthName = getMonthName(monthData.month);
    const year = monthData.year;
    const enrollmentCount = monthData.regularEnrollments.length;
    const easyStartCount = monthData.easyStarts.length;
    
    // Month header for Regular Enrollments with total
    dataArray.push([`${monthName.toUpperCase()} ${year} - ENROLLMENTS (Total: ${enrollmentCount})`]);
    currentRow++;
    
    // Regular Enrollments Header
    dataArray.push(getRegularEnrollmentHeaders());
    currentRow++;
    
    // Regular Enrollments Data
    if (monthData.regularEnrollments.length > 0) {
      monthData.regularEnrollments.forEach(deal => {
        const { row, url } = buildEnrollmentRow(deal, false);
        dataArray.push(row);
        if (url) {
          urlMap[currentRow] = url;
        }
        currentRow++;
      });
    } else {
      dataArray.push(['No enrollments this month']);
      currentRow++;
    }
    
    // Empty row separator
    dataArray.push([]);
    currentRow++;
    
    // Month header for Easy Starts with total
    dataArray.push([`${monthName.toUpperCase()} ${year} - EASY STARTS (Total: ${easyStartCount})`]);
    currentRow++;
    
    // Easy Starts Header
    dataArray.push(getEasyStartHeaders());
    currentRow++;
    
    // Easy Starts Data
    if (monthData.easyStarts.length > 0) {
      monthData.easyStarts.forEach(deal => {
        const { row, url } = buildEnrollmentRow(deal, true);
        dataArray.push(row);
        if (url) {
          urlMap[currentRow] = url;
        }
        currentRow++;
      });
    } else {
      dataArray.push(['No easy starts this month']);
      currentRow++;
    }
    
    // Empty row separator between months
    dataArray.push([]);
    currentRow++;
    dataArray.push([]);
    currentRow++;
  });
  
  // Append historical data (if any)
  if (historicalData && historicalData.length > 0) {
    Logger.log(`  Appending ${historicalData.length} rows of historical data`);
    historicalData.forEach(row => {
      dataArray.push(row);
      currentRow++;
    });
  }
  
  return { dataArray, urlMap };
}

/**
 * Builds a single enrollment row
 * @param {Object} deal - Deal object from HubSpot
 * @param {boolean} isEasyStart - Whether this is for Easy Starts table
 * @returns {Object} {row, url}
 */
function buildEnrollmentRow(deal, isEasyStart) {
  const row = [];
  
  // Deal Name
  const dealName = extractDealProperty(deal, 'dealname');
  row.push(dealName);
  const url = buildDealUrl(deal.id);
  
  // If Easy Start, add Easy Start Option after Deal Name
  if (isEasyStart) {
    row.push(extractDealProperty(deal, 'easy_start_option'));
  }
  
  // Rest of the fields (skip dealname)
  for (let i = 1; i < ENROLLMENT_FIELDS.length; i++) {
    const field = ENROLLMENT_FIELDS[i];
    
    if (field.type === 'date') {
      row.push(extractDateProperty(deal, field.property));
    } else if (field.type === 'number') {
      row.push(extractNumericProperty(deal, field.property));
    } else {
      row.push(extractDealProperty(deal, field.property));
    }
  }
  
  return { row, url };
}

// ============================================================================
// SHEET WRITING
// ============================================================================

/**
 * Writes data to the Enrollment Tracker sheet
 * @param {Sheet} sheet - The sheet to write to
 * @param {Array<Array>} dataArray - 2D array of data
 * @param {Object} urlMap - Map of row index to URL for hyperlinks
 */
function writeEnrollmentDataToSheet(sheet, dataArray, urlMap) {
  if (dataArray.length === 0) {
    return;
  }
  
  // Write all data at once
  const maxCols = Math.max(...dataArray.map(row => row.length));
  const range = sheet.getRange(1, 1, dataArray.length, maxCols);
  
  // Pad rows to have consistent column count
  const paddedData = dataArray.map(row => {
    const paddedRow = [...row];
    while (paddedRow.length < maxCols) {
      paddedRow.push('');
    }
    return paddedRow;
  });
  
  range.setValues(paddedData);
  
  // Apply hyperlinks to Deal Name column (Column A) - batch operation
  if (Object.keys(urlMap).length > 0) {
    Object.keys(urlMap).forEach(rowIndex => {
      const row = parseInt(rowIndex);
      const url = urlMap[rowIndex];
      const dealName = paddedData[row - 1][0]; // -1 for 0-based array
      
      if (url && dealName && dealName !== '' && dealName !== 'No enrollments this month' && dealName !== 'No easy starts this month') {
        const richText = SpreadsheetApp.newRichTextValue()
          .setText(dealName)
          .setLinkUrl(url)
          .build();
        sheet.getRange(row, 1).setRichTextValue(richText);
      }
    });
  }
}

// ============================================================================
// FORMATTING
// ============================================================================

/**
 * Applies formatting to Enrollment Tracker sheet
 * @param {Sheet} sheet - The sheet to format
 * @param {Array<Array>} dataArray - The data array for context
 * @param {Object} tierProgress - Tier progress information (optional)
 */
function applyEnrollmentFormatting(sheet, dataArray, tierProgress) {
  if (dataArray.length === 0) {
    return;
  }
  
  // Auto-resize all columns
  const maxCols = Math.max(...dataArray.map(row => row.length));
  for (let col = 1; col <= maxCols; col++) {
    sheet.autoResizeColumn(col);
  }
  
  let progressBarRow = 0;
  
  // Apply formatting to each row based on content
  for (let i = 0; i < dataArray.length; i++) {
    const rowIndex = i + 1;
    const firstCell = dataArray[i][0];
    
    if (!firstCell || firstCell === '') {
      // Empty row - skip
      continue;
    }
    
    // Tier progress bar (starts with 🎯 or 🏆)
    if (typeof firstCell === 'string' && (firstCell.startsWith('🎯') || firstCell.startsWith('🏆'))) {
      progressBarRow = rowIndex;
      const range = sheet.getRange(rowIndex, 1, 1, maxCols);
      
      // Determine color based on max tier achievement
      const isMaxTier = firstCell.startsWith('🏆');
      const bgColor = isMaxTier ? '#d4edda' : '#fff3cd'; // Green for max tier, yellow for progress
      
      range
        .setFontWeight('bold')
        .setFontSize(12) // Bigger font
        .setBackground(bgColor)
        .setFontColor('#000000')
        .setWrap(false)
        .setHorizontalAlignment('center') // Center the progress bar
        .setVerticalAlignment('middle')
        .merge();
      
      // Set row height for better visibility
      sheet.setRowHeight(rowIndex, 40); // Taller row
      
      continue;
    }
    
    // Month section headers (e.g., "NOVEMBER 2025 - ENROLLMENTS")
    if (typeof firstCell === 'string' && (firstCell.includes('ENROLLMENTS') || firstCell.includes('EASY STARTS')) && firstCell.includes('-')) {
      const range = sheet.getRange(rowIndex, 1, 1, maxCols);
      range
        .setFontWeight('bold')
        .setFontSize(12)
        .setBackground('#1155CC')
        .setFontColor('#FFFFFF')
        .merge();
      continue;
    }
    
    // Column headers (deal name, platform email, etc.)
    if (typeof firstCell === 'string' && firstCell === 'Deal Name') {
      const range = sheet.getRange(rowIndex, 1, 1, dataArray[i].length);
      range
        .setFontWeight('bold')
        .setBackground('#4285F4')
        .setFontColor('#FFFFFF')
        .setHorizontalAlignment('center');
      continue;
    }
  }
  
  // Freeze the progress bar row if it exists
  if (progressBarRow > 0) {
    sheet.setFrozenRows(progressBarRow);
  }
  
  // Apply color-coding to score columns
  applyEnrollmentColorCoding(sheet, dataArray);
}

/**
 * Applies red-yellow-green color coding to score columns
 * @param {Sheet} sheet - The sheet to format
 * @param {Array<Array>} dataArray - The data array
 */
function applyEnrollmentColorCoding(sheet, dataArray) {
  // Find all header rows and apply color coding to their score columns
  for (let i = 0; i < dataArray.length; i++) {
    const firstCell = dataArray[i][0];
    
    // Check if this is a header row
    if (firstCell === 'Deal Name') {
      const headers = dataArray[i];
      const dataRowStart = i + 2; // +1 for next row, +1 for 1-based indexing
      
      // Find where the data ends for this section (next empty row or section header)
      let dataRowEnd = dataRowStart;
      for (let j = i + 1; j < dataArray.length; j++) {
        const nextCell = dataArray[j][0];
        if (!nextCell || nextCell === '' || 
            (typeof nextCell === 'string' && (nextCell.includes('ENROLLMENTS') || nextCell.includes('EASY STARTS')))) {
          dataRowEnd = j; // Don't include this row
          break;
        }
        dataRowEnd = j + 1; // Include this row (1-based)
      }
      
      const dataRowCount = dataRowEnd - dataRowStart + 1;
      
      if (dataRowCount > 0) {
        // Apply color coding to Warm Handoff and Ask for Referral columns
        ENROLLMENT_FIELDS.forEach((field, idx) => {
          if (field.colorCode) {
            let colIndex = headers.indexOf(field.header) + 1; // +1 for 1-based
            
            if (colIndex > 0) {
              const range = sheet.getRange(dataRowStart, colIndex, dataRowCount, 1);
              
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
              
              // Apply the rule
              const existingRules = sheet.getConditionalFormatRules();
              existingRules.push(rule);
              sheet.setConditionalFormatRules(existingRules);
            }
          }
        });
      }
    }
  }
}

// ============================================================================
// HISTORICAL DATA PRESERVATION
// ============================================================================

/**
 * Captures historical data (older than last month) from existing sheet
 * @param {Sheet} sheet - The Enrollment Tracker sheet
 * @returns {Array<Array>} Historical data rows (raw values)
 */
function captureHistoricalData(sheet) {
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 1) {
    Logger.log('  No historical data to preserve (sheet is empty)');
    return [];
  }
  
  // Read all data
  const allData = sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues();
  
  // Find where historical data starts
  // Historical data is anything after the second month section
  let monthSectionCount = 0;
  let historicalStartRow = -1;
  
  for (let i = 0; i < allData.length; i++) {
    const firstCell = allData[i][0];
    
    // Check if this is a month section header
    if (typeof firstCell === 'string' && (firstCell.includes('ENROLLMENTS') || firstCell.includes('EASY STARTS')) && 
        firstCell.includes('-') && !firstCell.includes('No enrollments') && !firstCell.includes('No easy starts')) {
      
      // Count ENROLLMENT headers only (not EASY STARTS)
      if (firstCell.includes('ENROLLMENTS') && !firstCell.includes('EASY STARTS')) {
        monthSectionCount++;
        
        // After we've seen 2 months of enrollment headers, everything after is historical
        if (monthSectionCount >= 2) {
          // Skip forward past this month's easy starts section
          for (let j = i + 1; j < allData.length; j++) {
            const cell = allData[j][0];
            if (typeof cell === 'string' && cell.includes('EASY STARTS')) {
              // Find the end of this Easy Starts section
              for (let k = j + 1; k < allData.length; k++) {
                const nextCell = allData[k][0];
                if (!nextCell || nextCell === '') {
                  // Skip empty rows
                  if (k + 1 < allData.length && allData[k + 1][0]) {
                    historicalStartRow = k + 1;
                    break;
                  }
                }
              }
              break;
            }
          }
          break;
        }
      }
    }
  }
  
  // If we found historical data, extract it
  if (historicalStartRow > 0 && historicalStartRow < allData.length) {
    const historicalData = allData.slice(historicalStartRow);
    Logger.log(`  Preserved ${historicalData.length} rows of historical data (starting from row ${historicalStartRow + 1})`);
    return historicalData;
  }
  
  Logger.log('  No historical data to preserve');
  return [];
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Gets month name from month number (0-11)
 * @param {number} month - Month number (0 = January, 11 = December)
 * @returns {string} Month name
 */
function getMonthName(month) {
  const months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  return months[month] || 'Unknown';
}
