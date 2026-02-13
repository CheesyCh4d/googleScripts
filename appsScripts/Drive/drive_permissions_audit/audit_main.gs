/**
 * Initiates the audit process.
 * Sets up the sheet and queue, then hands off to resumeAudit().
 */
function startAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  deleteTriggers(); // Clean up old triggers (from Utilities.gs)

  const response = ui.prompt('Start New Audit', 'Paste the Folder URL here:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const folderUrl = response.getResponseText();
  if (!folderUrl || !folderUrl.includes('/folders/')) {
    ui.alert('Invalid URL. Please enter a valid Google Drive Folder URL.');
    return;
  }

  const folderId = folderUrl.split('/folders/')[1].split('?')[0];

  // Setup Main Sheet
  const sheet = ss.getActiveSheet();
  sheet.setName('Audit Results');
  sheet.clear();
  const headers = ['Type', 'Path/Folder', 'Item Name', 'Link', 'User Email', 'Current Role', 'Adjusted Permission', 'File ID'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9d9d9');
  sheet.setFrozenRows(1);

  // Setup Queue Sheet
  let queueSheet = ss.getSheetByName('_AuditQueue');
  if (queueSheet) {
    queueSheet.clear();
  } else {
    queueSheet = ss.insertSheet('_AuditQueue');
    queueSheet.hideSheet();
  }
  queueSheet.appendRow([folderId, '/']);

  ss.toast('Audit started. This may take a while...', 'System Started');
  resumeAudit();
}

/**
 * Stops the audit and cleans up.
 */
function stopAudit() {
  deleteTriggers();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = ss.getSheetByName('_AuditQueue');
  if (queueSheet) ss.deleteSheet(queueSheet);
  ss.toast('Audit process stopped.', 'Stopped');
}

/**
 * The main loop. Processes the queue and manages the time limit.
 */
function resumeAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Audit Results') || ss.getSheets()[0];
  const queueSheet = ss.getSheetByName('_AuditQueue');

  if (!queueSheet || queueSheet.getLastRow() === 0) {
    finishAudit(sheet, queueSheet);
    return;
  }

  const startTime = Date.now();
  let rowBuffer = [];
  let lastFlushTime = Date.now();
  let lastToastTime = 0;

  while (queueSheet.getLastRow() > 0) {
    // Autopilot: Restart if time is running out
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      flushBuffer(sheet, rowBuffer);
      ScriptApp.newTrigger('resumeAudit').timeBased().after(5000).create();
      ss.toast('Time limit reached. Auto-restarting in 5s...', 'Autopilot Engaged');
      return;
    }

    const range = queueSheet.getRange(1, 1, 1, 2);
    const [currentFolderId, currentPath] = range.getValues()[0];

    try {
      const folder = DriveApp.getFolderById(currentFolderId);

      if (Date.now() - lastToastTime > 10000) {
        ss.toast(`Scanning: ${currentPath}`, 'Working...');
        lastToastTime = Date.now();
      }

      // Scan Files
      const files = folder.getFiles();
      while (files.hasNext()) {
        collectItemRows(files.next(), 'File', currentPath, rowBuffer);
        if (rowBuffer.length >= BATCH_SIZE || Date.now() - lastFlushTime > FLUSH_INTERVAL) {
          flushBuffer(sheet, rowBuffer);
          rowBuffer = [];
          lastFlushTime = Date.now();
        }
      }

      // Scan Subfolders
      const subfolders = folder.getFolders();
      while (subfolders.hasNext()) {
        const sub = subfolders.next();
        queueSheet.appendRow([sub.getId(), currentPath + sub.getName() + '/']);
        collectItemRows(sub, 'Folder', currentPath, rowBuffer);
        
        if (rowBuffer.length >= BATCH_SIZE || Date.now() - lastFlushTime > FLUSH_INTERVAL) {
          flushBuffer(sheet, rowBuffer);
          rowBuffer = [];
          lastFlushTime = Date.now();
        }
      }
      queueSheet.deleteRow(1);
    } catch (e) {
      console.log(`Error accessing ${currentFolderId}: ${e.message}`);
      queueSheet.deleteRow(1);
    }
  }

  flushBuffer(sheet, rowBuffer);
  finishAudit(sheet, queueSheet);
}