/**
 * Quick Sync Component
 * 
 * Fast, on-demand sync operations for real-time collaboration:
 * - AE → Director: Push notes only (no full refresh)
 * - Director → AE: Push highlighting only (no full refresh)
 */

/**
 * DIRECTOR BUTTON: Pull notes from team AE sheets
 * Called from Control Sheet
 * Duration: 2-3 seconds
 */
function pullNotesFromMyTeam() {
  const lock = LockService.getScriptLock();
  const lockAcquired = lock.tryLock(5000);
  
  if (!lockAcquired) {
    showToast('⏳ System is busy, please wait and try again', 'Busy', 5);
    return;
  }
  
  try {
    const startTime = new Date();
    showToast('📤 Pulling notes from your team...', 'Syncing', 3);
    
    // Step 1: Detect which director tab is active
    const controlSheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = controlSheet.getActiveSheet();
    const activeTabName = activeSheet.getName();
    
    Logger.log(`[Quick Sync] Pulling notes from "${activeTabName}"...`);
    
    // Step 2: Find which director owns this tab
    const directors = getDirectorConfig();
    const director = directors.find(d => d.tabName === activeTabName);
    
    if (!director) {
      showToast(`ℹ️ "${activeTabName}" is not a director tab. Please switch to a director's tab and try again.`, 'Wrong Tab', 8);
      return;
    }
    
    Logger.log(`  Detected director: ${director.name} (${director.team} team)`);
    
    // Step 3: Find all AEs on this team
    const config = loadConfiguration();
    const { salespeople } = config;
    const teamAEs = salespeople.filter(p => 
      p.team && p.team.toLowerCase() === director.team.toLowerCase() && p.sheetId
    );
    
    if (teamAEs.length === 0) {
      throw new Error(`No AEs found for ${director.team} team.`);
    }
    
    Logger.log(`  Found ${teamAEs.length} AE(s) on ${director.team} team`);
    
    // Step 4: Collect notes from all team AE sheets
    const notesMap = new Map(); // Map: Deal ID -> notes
    
    teamAEs.forEach(person => {
      try {
        const aeSheet = SpreadsheetApp.openById(person.sheetId);
        const pipelineSheet = aeSheet.getSheetByName(TAB_PIPELINE);
        
        if (!pipelineSheet) {
          Logger.log(`    No Pipeline Review tab for ${person.name}, skipping`);
          return;
        }
        
        const lastRow = pipelineSheet.getLastRow();
        if (lastRow < 2) return;
        
        // Read Deal IDs (column A) and Notes (column G)
        const NOTES_COLUMN_INDEX = 6; // Column G (0-based)
        const data = pipelineSheet.getRange(2, 1, lastRow - 1, Math.max(7, pipelineSheet.getLastColumn())).getValues();
        
        data.forEach(row => {
          const dealId = row[0]?.toString();
          const notes = row[NOTES_COLUMN_INDEX] || '';
          
          if (dealId && notes) {
            notesMap.set(dealId, notes);
          }
        });
        
      } catch (e) {
        Logger.log(`    Error collecting from ${person.name}: ${e.message}`);
      }
    });
    
    Logger.log(`  Collected notes for ${notesMap.size} deals`);
    
    if (notesMap.size === 0) {
      showToast('ℹ️ No notes found in team AE sheets.', 'No Notes', 5);
      return;
    }
    
    // Step 5: Update notes in director's sheet
    const directorLastRow = activeSheet.getLastRow();
    if (directorLastRow < 2) {
      showToast('ℹ️ Director sheet is empty. Please run full dashboard generation first.', 'No Data', 5);
      return;
    }
    
    // Read director sheet (Deal ID in column A, Notes in column H)
    const DIRECTOR_NOTES_COLUMN = 8; // Column H (1-based)
    const directorDealIds = activeSheet.getRange(2, 1, directorLastRow - 1, 1).getValues(); // Just Deal IDs
    
    let updatedCount = 0;
    const updates = []; // Store updates for batch operation
    
    directorDealIds.forEach((row, index) => {
      const dealId = row[0]?.toString();
      if (dealId && notesMap.has(dealId)) {
        const rowIndex = index + 2; // +2 for header and 0-based
        updates.push({
          row: rowIndex,
          notes: notesMap.get(dealId)
        });
        updatedCount++;
      }
    });
    
    // Batch update notes (FAST!)
    if (updates.length > 0) {
      updates.forEach(update => {
        activeSheet.getRange(update.row, DIRECTOR_NOTES_COLUMN).setValue(update.notes);
      });
      Logger.log(`  Updated ${updatedCount} notes in director sheet`);
    }
    
    const duration = (new Date() - startTime) / 1000;
    Logger.log(`[Quick Sync] Complete (${duration}s)`);
    
    if (updatedCount > 0) {
      showToast(`✅ Pulled notes for ${updatedCount} deal(s) from ${teamAEs.length} AE(s) (${duration.toFixed(1)}s)`, 'Success', 5);
    } else {
      showToast('ℹ️ No matching deals found with notes.', 'No Updates', 5);
    }
    
  } catch (error) {
    Logger.log(`[Quick Sync] ERROR: ${error.message}`);
    Logger.log(error.stack);
    
    // Show toast AND send email
    showToast(`❌ Failed: ${error.message}`, 'Error', 10);
    sendErrorEmail('Pull Notes from Team', error);
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * DIRECTOR BUTTON: Sync highlighting from active director tab to team AEs
 * Called from Control Sheet
 * Duration: 3-5 seconds
 */
function syncHighlightingToMyTeam() {
  const lock = LockService.getScriptLock();
  const lockAcquired = lock.tryLock(5000);
  
  if (!lockAcquired) {
    showToast('⏳ System is busy, please wait and try again', 'Busy', 5);
    return;
  }
  
  try {
    const startTime = new Date();
    showToast('📥 Syncing highlighting to your team...', 'Syncing', 3);
    
    // Step 1: Detect which director tab is active
    const controlSheet = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = controlSheet.getActiveSheet();
    const activeTabName = activeSheet.getName();
    
    Logger.log(`[Quick Sync] Syncing highlighting from "${activeTabName}"...`);
    
    // Step 2: Find which director owns this tab
    const directors = getDirectorConfig();
    const director = directors.find(d => d.tabName === activeTabName);
    
    if (!director) {
      showToast(`ℹ️ "${activeTabName}" is not a director tab. Please switch to a director's tab and try again.`, 'Wrong Tab', 8);
      return;
    }
    
    Logger.log(`  Detected director: ${director.name} (${director.team} team)`);
    
    // Step 3: Collect highlighting from this director's sheet
    const directorHighlighting = captureDirectorHighlighting(activeSheet);
    Logger.log(`  Captured highlighting for ${directorHighlighting.size} deals`);
    
    if (directorHighlighting.size === 0) {
      showToast('ℹ️ No highlighting detected. Highlight some cells first, then try again.', 'No Highlighting', 5);
      return;
    }
    
    // Step 4: Find all AEs on this team
    const config = loadConfiguration();
    const { salespeople } = config;
    const teamAEs = salespeople.filter(p => 
      p.team && p.team.toLowerCase() === director.team.toLowerCase() && p.sheetId
    );
    
    if (teamAEs.length === 0) {
      throw new Error(`No AEs found for ${director.team} team.`);
    }
    
    Logger.log(`  Found ${teamAEs.length} AE(s) on ${director.team} team`);
    
    // Step 5: Apply highlighting to each AE's sheet
    let totalDealsHighlighted = 0;
    let aesUpdated = 0;
    
    teamAEs.forEach(person => {
      try {
        const aeSheet = SpreadsheetApp.openById(person.sheetId);
        const pipelineSheet = aeSheet.getSheetByName(TAB_PIPELINE);
        
        if (!pipelineSheet) {
          Logger.log(`    No Pipeline Review tab for ${person.name}, skipping`);
          return;
        }
        
        const lastRow = pipelineSheet.getLastRow();
        if (lastRow < 2) return;
        
        const dataRowCount = lastRow - 1;
        const aeColumnCount = pipelineSheet.getLastColumn();
        
        // Read Deal IDs from AE sheet (BATCH)
        const aeDealIds = pipelineSheet.getRange(2, 1, dataRowCount, 1).getValues().flat();
        
        // Read current backgrounds and font colors (BATCH)
        const dataRange = pipelineSheet.getRange(2, 1, dataRowCount, aeColumnCount);
        const currentBackgrounds = dataRange.getBackgrounds();
        const currentFontColors = dataRange.getFontColors();
        
        // Update highlighting for matched deals
        let appliedCount = 0;
        aeDealIds.forEach((dealId, index) => {
          if (!dealId) return;
          
          const dealIdStr = dealId.toString();
          const highlighting = directorHighlighting.get(dealIdStr);
          
          if (highlighting) {
            if (index >= currentBackgrounds.length) return;
            if (!highlighting.backgrounds || !highlighting.fontColors) return;
            
            // Director sheets have Owner column (22 cols), AE sheets don't (21 cols)
            // Remove Owner column (index 2) from highlighting arrays
            let adjustedBackgrounds = [...highlighting.backgrounds];
            let adjustedFontColors = [...highlighting.fontColors];
            
            if (adjustedBackgrounds.length > aeColumnCount) {
              adjustedBackgrounds.splice(2, 1); // Remove Owner column
              adjustedFontColors.splice(2, 1);
            }
            
            // Ensure arrays match AE column count
            while (adjustedBackgrounds.length < aeColumnCount) {
              adjustedBackgrounds.push('#ffffff');
              adjustedFontColors.push('#000000');
            }
            while (adjustedBackgrounds.length > aeColumnCount) {
              adjustedBackgrounds.pop();
              adjustedFontColors.pop();
            }
            
            if (adjustedBackgrounds.length !== aeColumnCount) return;
            
            currentBackgrounds[index] = adjustedBackgrounds;
            currentFontColors[index] = adjustedFontColors;
            appliedCount++;
          }
        });
        
        // Apply all highlighting changes in ONE batch operation (FAST!)
        if (appliedCount > 0) {
          dataRange.setBackgrounds(currentBackgrounds);
          dataRange.setFontColors(currentFontColors);
          Logger.log(`    ${person.name}: Applied highlighting to ${appliedCount} deal(s)`);
          totalDealsHighlighted += appliedCount;
          aesUpdated++;
        }
        
      } catch (e) {
        Logger.log(`    Error syncing to ${person.name}: ${e.message}`);
      }
    });
    
    const duration = (new Date() - startTime) / 1000;
    Logger.log(`[Quick Sync] Complete (${duration}s)`);
    
    if (totalDealsHighlighted > 0) {
      showToast(`✅ Synced highlighting for ${totalDealsHighlighted} deal(s) across ${aesUpdated} AE(s) (${duration.toFixed(1)}s)`, 'Success', 5);
    } else {
      showToast('ℹ️ No matching deals found in AE sheets. Make sure AE sheets are up to date.', 'No Matches', 5);
    }
    
  } catch (error) {
    Logger.log(`[Quick Sync] ERROR: ${error.message}`);
    Logger.log(error.stack);
    
    // Show toast AND send email
    showToast(`❌ Failed: ${error.message}`, 'Error', 10);
    sendErrorEmail('Sync Highlighting to Team', error);
    
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper: Show toast notification
 */
function showToast(message, title, duration) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    if (sheet) {
      sheet.toast(message, title, duration);
    }
  } catch (e) {
    Logger.log(`Toast failed: ${e.message}`);
  }
}

/**
 * Helper: Send error email
 */
function sendErrorEmail(operation, error) {
  try {
    const user = Session.getActiveUser().getEmail();
    const timestamp = new Date().toISOString();
    
    const subject = `[All-in-One Dashboard] Quick Sync Error: ${operation}`;
    const body = `
Error occurred during Quick Sync operation.

Operation: ${operation}
User: ${user}
Timestamp: ${timestamp}

Error Message:
${error.message}

Stack Trace:
${error.stack || 'No stack trace available'}

Please check the logs for more details.
    `;
    
    // Send to you (system admin)
    MailApp.sendEmail({
      to: 'konstantin.gevorkov@tripleten.com',
      subject: subject,
      body: body
    });
    
    Logger.log('Error email sent to admin');
    
  } catch (e) {
    Logger.log(`Failed to send error email: ${e.message}`);
  }
}

