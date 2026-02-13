/**
 * Reads the sheet and applies changes to Drive.
 */
function syncPermissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Audit Results') || ss.getSheets()[0];
  const ui = SpreadsheetApp.getUi();
  const data = sheet.getDataRange().getValues();

  let updatesCount = 0;
  let errorCount = 0;

  for (let i = 1; i < data.length; i++) {
    const [type, path, name, link, email, currentRole, newRole, fileId] = data[i];

    if (currentRole === newRole || !fileId || !email) continue;
    if (currentRole === 'Owner') continue;
    if (currentRole === '---' && newRole === 'Remove Access') continue;

    try {
      const isFolder = (type === 'Folder' || type === 'Root Folder');
      const item = isFolder ? DriveApp.getFolderById(fileId) : DriveApp.getFileById(fileId);

      // Remove Access
      if (newRole === 'Remove Access') {
        removeOldRole(item, email, currentRole);
        sheet.getRange(i + 1, 6).setValue('Removed');
        sheet.getRange(i + 1, 7).setValue('Removed');
        updatesCount++;
        continue;
      }

      // Handle Commenter on Folder edge case
      if (newRole === 'Commenter' && isFolder) {
        removeOldRole(item, email, currentRole);
        item.addViewer(email);
        sheet.getRange(i + 1, 6).setValue('Viewer (Folder Restriction)');
        sheet.getRange(i + 1, 7).setValue('Viewer');
        updatesCount++;
        continue;
      }

      // Standard Change
      if (currentRole !== '---') {
        removeOldRole(item, email, currentRole);
      }

      if (newRole === 'Editor') item.addEditor(email);
      else if (newRole === 'Viewer') item.addViewer(email);
      else if (newRole === 'Commenter') item.addCommenter(email);

      sheet.getRange(i + 1, 6).setValue(newRole);
      updatesCount++;

    } catch (e) {
      console.log(`Error updating "${name}" for ${email}: ${e.message}`);
      errorCount++;
    }
  }

  let message = `Sync Complete!\nUpdated permissions for ${updatesCount} entries.`;
  if (errorCount > 0) message += `\n${errorCount} entries had errors.`;
  ui.alert(message);
}

/**
 * Helper to remove existing permissions before adding new ones.
 */
function removeOldRole(item, email, currentRole) {
  if (currentRole === 'Editor') item.removeEditor(email);
  else if (currentRole === 'Viewer' || currentRole === 'Commenter') item.removeViewer(email);
}