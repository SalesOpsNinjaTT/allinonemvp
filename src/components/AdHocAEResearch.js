/**
 * Ad Hoc AE Research (30-day Pipeline)
 * Creates a standalone spreadsheet with separate tabs per AE,
 * mirroring Pipeline Review formatting, fields, and manual columns,
 * but limited to the last 30 days by deal createdate.
 */

/**
 * Entry point: generate ad hoc research for a list of AEs.
 * @param {Array<Object>} aeList - [{ email, hubspotUserId?, tabName? }]
 * @returns {Object} { success, sheetUrl, sheetId, tabs }
 */
function generateAdHocAEResearch(aeList) {
  // Fallback to default two AEs if none provided
  if (!Array.isArray(aeList) || aeList.length === 0) {
    aeList = [
      { email: 'zoya.white@tripleten.com', hubspotUserId: '30293278', tabName: 'Zoya White (30d)' },
      { email: 'maureen.bonthuys@tripleten.com', hubspotUserId: '30293277', tabName: 'Maureen Bonthuys (30d)' }
    ];
  }

  const todayStr = new Date().toISOString().slice(0, 10);
  const ssName = `AE Research - ${todayStr}`;
  const ss = SpreadsheetApp.create(ssName);

  // Keep default Sheet1 until we create the first tab to avoid removing all sheets
  const defaultSheet = ss.getSheetByName('Sheet1');

  const tabs = [];
  let firstTabUsed = false;

  aeList.forEach(ae => {
    // Reuse the default sheet for the first AE to avoid a zero-sheet state
    const sheetOverride = !firstTabUsed && defaultSheet ? defaultSheet : null;
    const result = upsertAdHocAETab(ss, ae, sheetOverride);
    firstTabUsed = true;
    tabs.push(result);
  });

  // If defaultSheet was not reused (edge case), remove it now
  if (defaultSheet && !firstTabUsed) {
    // Shouldn't happen, but safety check
    ss.deleteSheet(defaultSheet);
  }

  Logger.log(`Ad Hoc AE Research created: ${ss.getUrl()}`);

  return {
    success: true,
    sheetUrl: ss.getUrl(),
    sheetId: ss.getId(),
    tabs
  };
}

/**
 * Creates or refreshes a tab for a single AE with 30-day deals.
 * @param {Spreadsheet} ss - Target standalone spreadsheet
 * @param {Object} ae - { email, hubspotUserId?, tabName? }
 * @param {Sheet|null} sheetOverride - Optional pre-existing sheet to reuse/rename
 * @returns {Object} { tabName, dealCount }
 */
function upsertAdHocAETab(ss, ae, sheetOverride) {
  if (!ae || !ae.email) {
    throw new Error('AE entry requires email');
  }

  const tabName = sanitizeTabName(ae.tabName || `${ae.email} - 30d`);
  let sheet = sheetOverride || ss.getSheetByName(tabName);

  if (!sheet) {
    sheet = ss.insertSheet(tabName);
  } else {
    // If reusing an existing sheet (e.g., default Sheet1), rename it
    if (sheet.getName() !== tabName) {
      sheet.setName(tabName);
    }
  }

  Logger.log(`[AdHoc AE] Updating ${tabName}...`);

  // Preserve manual columns/formatting if tab already existed
  const preserved = capturePreservedData(sheet);

  // Fetch deals with same properties as Pipeline Review, but 30-day window
  const properties = getPipelineReviewProperties();
  const options = { dateRangeDays: 30 };
  if (ae.hubspotUserId && ae.hubspotUserId !== '') {
    options.hubspotUserId = ae.hubspotUserId;
  }

  const deals = fetchDealsByOwner(ae.email, properties, options);
  Logger.log(`  Found ${deals.length} deals for ${ae.email}`);

  // Build data and write
  const { dataArray, urlMap, dealIdMap } = buildPipelineDataArray(deals);
  sheet.clear();
  writeDataToSheet(sheet, dataArray, urlMap);
  applyPipelineFormatting(sheet, dataArray.length - 1);
  restorePreservedData(sheet, preserved, dealIdMap);

  Logger.log(`[AdHoc AE] Complete ${tabName} (${deals.length} deals)`);

  return { tabName, dealCount: deals.length };
}

/**
 * Basic sheet name sanitizer for readability and Sheets compatibility.
 * @param {string} name - Desired sheet name
 * @returns {string} Safe sheet name
 */
function sanitizeTabName(name) {
  if (!name) return 'AE';
  // Remove characters Sheets disallows: : \ / ? * [ ]
  const cleaned = name.replace(/[:\\/?*\\[\\]]/g, ' ');
  return cleaned.substring(0, 99); // Sheets limit
}

/**
 * Convenience runner with the two specified AEs and 30-day window.
 * Run this directly from Apps Script UI (no arguments needed).
 */
function runAdHocAEResearchDefault() {
  return generateAdHocAEResearch([
    { email: 'zoya.white@tripleten.com', hubspotUserId: '30293278', tabName: 'Zoya White (30d)' },
    { email: 'maureen.bonthuys@tripleten.com', hubspotUserId: '30293277', tabName: 'Maureen Bonthuys (30d)' }
  ]);
}

