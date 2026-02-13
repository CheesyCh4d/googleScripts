/**
 * Collects permission data for a single file/folder and pushes to buffer.
 */
function collectItemRows(item, type, path, buffer) {
  const fileId = item.getId();
  const name = item.getName();
  const url = item.getUrl();

  const makeRow = (email, role, adjustedRole) => {
    return [type, path, name, url, email, role, adjustedRole, fileId];
  };

  // Owner
  try {
    const owner = item.getOwner();
    if (owner) buffer.push(makeRow(owner.getEmail(), 'Owner', 'Owner'));
  } catch (e) {}

  // Editors
  try {
    const editors = item.getEditors();
    for (let i = 0; i < editors.length; i++) {
      buffer.push(makeRow(editors[i].getEmail(), 'Editor', 'Editor'));
    }
  } catch (e) {}

  // Viewers & Commenters
  try {
    const viewers = item.getViewers();
    for (let i = 0; i < viewers.length; i++) {
      const email = viewers[i].getEmail();
      let role = 'Viewer';
      try {
        if (item.getAccess(email) === DriveApp.Permission.COMMENT) {
          role = 'Commenter';
        }
      } catch (e) {}
      buffer.push(makeRow(email, role, role));
    }
  } catch (e) {}
}

/**
 * Writes the buffer to the sheet in one batch.
 */
function flushBuffer(sheet, buffer) {
  if (buffer.length === 0) return;
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, buffer.length, buffer[0].length).setValues(buffer);
}

/**
 * Cleanup function when audit completes.
 */
function finishAudit(sheet, queueSheet) {
  deleteTriggers();
  applyDropdowns(sheet);
  if (queueSheet) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(queueSheet);
  SpreadsheetApp.getUi().alert('Audit Finished Successfully!');
}