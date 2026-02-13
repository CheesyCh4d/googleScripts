/**
 * Creates the custom menu when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Drive Audit')
    .addItem('1. Start New Audit', 'startAudit')
    .addItem('2. Force Resume (If stuck)', 'resumeAudit')
    .addItem('3. Stop/Cancel Auto-Audit', 'stopAudit')
    .addSeparator()
    .addItem('4. Add New User to Selected File', 'addNewUser')
    .addItem('5. Sync Changes (Update Drive)', 'syncPermissions')
    .addToUi();
}