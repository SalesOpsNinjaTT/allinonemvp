/**
 * Director Menu Component
 * 
 * Custom menu for directors to flag deals in Director Hub
 * Provides Hot/Cold/Attention flags with optional notes
 * 
 * Note: TAB_DIRECTOR_HUB is defined in DirectorHub.js
 * Note: TAB_PIPELINE is defined in SheetProvisioner.js
 */

// Flag colors for row backgrounds
const FLAG_COLORS = {
  HOT: '#d9ead3',      // Light green
  COLD: '#f4cccc',     // Light red
  ATTENTION: '#fff2cc', // Light yellow
  NONE: '#ffffff'      // White
};

/**
 * Creates custom menu when Control Sheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 Directives')
    .addItem('🟢 Mark as Hot', 'markAsHot')
    .addItem('🔴 Mark as Cold', 'markAsCold')
    .addItem('🟡 Mark as Needs Attention', 'markAsAttention')
    .addSeparator()
    .addItem('⚪ Clear Flag', 'clearFlag')
    .addToUi();
}

/**
 * Mark selected deal as Hot (low-hanging fruit)
 */
function markAsHot() {
  applyDirectorFlag('🟢', 'Hot', FLAG_COLORS.HOT);
}

/**
 * Mark selected deal as Cold (dead deal)
 */
function markAsCold() {
  applyDirectorFlag('🔴', 'Cold', FLAG_COLORS.COLD);
}

/**
 * Mark selected deal as Needs Attention
 */
function markAsAttention() {
  applyDirectorFlag('🟡', 'Attention', FLAG_COLORS.ATTENTION);
}

/**
 * Clear flag from selected deal
 */
function clearFlag() {
  applyDirectorFlag('', 'Clear', FLAG_COLORS.NONE);
}

/**
 * Applies a director flag to the selected deal
 * @param {string} flag - Flag emoji
 * @param {string} flagName - Name for UI prompts
 * @param {string} color - Background color for row
 */
function applyDirectorFlag(flag, flagName, color) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const ui = SpreadsheetApp.getUi();
    
    // Validate we're in Director Hub
    if (sheet.getName() !== TAB_DIRECTOR_HUB) {
      ui.alert('⚠️ Wrong Sheet', 'This action only works in the Director Hub tab.', ui.ButtonSet.OK);
      return;
    }
    
    // Get selected range
    const selection = sheet.getActiveRange();
    if (!selection) {
      ui.alert('⚠️ No Selection', 'Please select a deal row first.', ui.ButtonSet.OK);
      return;
    }
    
    const selectedRow = selection.getRow();
    
    // Validate row (must be > 1, i.e., not header)
    if (selectedRow === 1) {
      ui.alert('⚠️ Invalid Selection', 'Please select a deal row, not the header.', ui.ButtonSet.OK);
      return;
    }
    
    // Get column indices
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dealNameCol = headers.indexOf('Deal Name') + 1;
    const priorityCol = headers.indexOf('Director Priority') + 1;
    const noteCol = headers.indexOf('Director Note') + 1;
    
    if (dealNameCol === 0 || priorityCol === 0 || noteCol === 0) {
      ui.alert('⚠️ Error', 'Could not find required columns.', ui.ButtonSet.OK);
      return;
    }
    
    // Get deal name for confirmation
    const dealName = sheet.getRange(selectedRow, dealNameCol).getValue();
    if (!dealName) {
      ui.alert('⚠️ Empty Row', 'This row does not contain a deal.', ui.ButtonSet.OK);
      return;
    }
    
    // Prompt for optional note (skip if clearing)
    let note = '';
    if (flag !== '') {
      const response = ui.prompt(
        `Mark as ${flagName}`,
        `Deal: ${dealName}\n\nAdd optional note for AE:`,
        ui.ButtonSet.OK_CANCEL
      );
      
      if (response.getSelectedButton() === ui.Button.CANCEL) {
        return; // User cancelled
      }
      
      note = response.getResponseText();
    }
    
    // Apply flag and note
    sheet.getRange(selectedRow, priorityCol).setValue(flag);
    sheet.getRange(selectedRow, noteCol).setValue(note);
    
    // Apply background color to entire row
    const rowRange = sheet.getRange(selectedRow, 1, 1, headers.length);
    rowRange.setBackground(color);
    
    // Success message
    const action = flag === '' ? 'cleared' : `marked as ${flagName}`;
    ui.toast(`✅ Deal ${action}: ${dealName}`, 'Success', 3);
    
    Logger.log(`[Director Flag] Row ${selectedRow}: ${dealName} → ${flag} ${flagName} "${note}"`);
    
  } catch (error) {
    Logger.log(`[Director Flag] Error: ${error.message}`);
    SpreadsheetApp.getUi().alert('❌ Error', `Failed to apply flag: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Syncs director flags from Director Hub to individual AE sheets
 * Called after Director Hub refresh to propagate flags to AE sheets
 * @param {Spreadsheet} controlSheet - The control spreadsheet
 * @param {Array<Object>} salespeople - Array of salesperson configs
 */
function syncDirectorFlagsToAESheets(controlSheet, salespeople) {
  try {
    Logger.log('[Director Sync] Syncing flags to AE sheets...');
    
    const directorSheet = controlSheet.getSheetByName(TAB_DIRECTOR_HUB);
    if (!directorSheet) {
      Logger.log('[Director Sync] No Director Hub found, skipping sync');
      return;
    }
    
    const lastRow = directorSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('[Director Sync] No deals in Director Hub, skipping sync');
      return;
    }
    
    // Read Director Hub data
    const headers = directorSheet.getRange(1, 1, 1, directorSheet.getLastColumn()).getValues()[0];
    const ownerCol = headers.indexOf('Owner') + 1;
    const dealNameCol = headers.indexOf('Deal Name') + 1;
    const priorityCol = headers.indexOf('Director Priority') + 1;
    const noteCol = headers.indexOf('Director Note') + 1;
    
    if (ownerCol === 0 || dealNameCol === 0 || priorityCol === 0 || noteCol === 0) {
      Logger.log('[Director Sync] Missing required columns');
      return;
    }
    
    const data = directorSheet.getRange(2, 1, lastRow - 1, directorSheet.getLastColumn()).getValues();
    const backgrounds = directorSheet.getRange(2, 1, lastRow - 1, directorSheet.getLastColumn()).getBackgrounds();
    
    // Build map of flags by owner and deal name
    const flagsByOwner = {};
    
    for (let i = 0; i < data.length; i++) {
      const owner = data[i][ownerCol - 1];
      const dealName = data[i][dealNameCol - 1];
      const priority = data[i][priorityCol - 1];
      const note = data[i][noteCol - 1];
      const background = backgrounds[i];
      
      if (owner && dealName && (priority || note)) {
        if (!flagsByOwner[owner]) {
          flagsByOwner[owner] = {};
        }
        
        flagsByOwner[owner][dealName] = {
          priority: priority,
          note: note,
          background: background
        };
      }
    }
    
    const ownerCount = Object.keys(flagsByOwner).length;
    Logger.log(`[Director Sync] Found flags for ${ownerCount} owners`);
    
    // Apply flags to each AE's sheet
    let totalSynced = 0;
    
    salespeople.forEach(person => {
      if (!flagsByOwner[person.name]) {
        return; // No flags for this AE
      }
      
      const flags = flagsByOwner[person.name];
      const flagCount = Object.keys(flags).length;
      
      Logger.log(`[Director Sync] Syncing ${flagCount} flags to ${person.name}...`);
      
      try {
        if (!person.sheetId || person.sheetId === '') {
          Logger.log(`[Director Sync] No sheet ID for ${person.name}, skipping`);
          return;
        }
        
        const aeSheet = SpreadsheetApp.openById(person.sheetId);
        const pipelineSheet = aeSheet.getSheetByName(TAB_PIPELINE);
        
        if (!pipelineSheet) {
          Logger.log(`[Director Sync] No Pipeline Review tab for ${person.name}, skipping`);
          return;
        }
        
        const aeLastRow = pipelineSheet.getLastRow();
        if (aeLastRow < 2) {
          return;
        }
        
        const aeHeaders = pipelineSheet.getRange(1, 1, 1, pipelineSheet.getLastColumn()).getValues()[0];
        const aeDealNameCol = aeHeaders.indexOf('Deal Name') + 1;
        const aePriorityCol = aeHeaders.indexOf('Director Priority') + 1;
        const aeNoteCol = aeHeaders.indexOf('Director Note') + 1;
        
        if (aeDealNameCol === 0 || aePriorityCol === 0 || aeNoteCol === 0) {
          Logger.log(`[Director Sync] Missing columns in ${person.name}'s sheet, skipping`);
          return;
        }
        
        const aeDealNames = pipelineSheet.getRange(2, aeDealNameCol, aeLastRow - 1, 1).getValues();
        
        let syncedForAE = 0;
        
        for (let i = 0; i < aeDealNames.length; i++) {
          const dealName = aeDealNames[i][0];
          const rowIndex = i + 2;
          
          if (dealName && flags[dealName]) {
            const flag = flags[dealName];
            
            // Set priority and note
            pipelineSheet.getRange(rowIndex, aePriorityCol).setValue(flag.priority);
            pipelineSheet.getRange(rowIndex, aeNoteCol).setValue(flag.note);
            
            // Set row background
            const rowRange = pipelineSheet.getRange(rowIndex, 1, 1, aeHeaders.length);
            rowRange.setBackgrounds([flag.background]);
            
            syncedForAE++;
          }
        }
        
        totalSynced += syncedForAE;
        Logger.log(`[Director Sync] Synced ${syncedForAE} flags to ${person.name}`);
        
      } catch (error) {
        Logger.log(`[Director Sync] Error syncing to ${person.name}: ${error.message}`);
      }
    });
    
    Logger.log(`[Director Sync] Complete: ${totalSynced} total flags synced`);
    
  } catch (error) {
    Logger.log(`[Director Sync] Error: ${error.message}`);
  }
}

