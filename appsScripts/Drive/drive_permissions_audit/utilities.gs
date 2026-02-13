/**
 * Removes triggers associated with the audit loop.
 */
function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'resumeAudit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**
 * Applies Data Validation dropdowns to the sheet.
 */
function applyDropdowns(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Editor', 'Viewer', 'Commenter', 'Remove Access'], true)
    .setAllowInvalid(false)
    .build();

  const roles = sheet.getRange(2, 6, lastRow - 1, 1).getValues();

  for (let i = 0; i < roles.length; i++) {
    // Don't put dropdowns on Owner rows
    if (roles[i][0] === 'Owner') {
      sheet.getRange(i + 2, 7).clearDataValidations();
    } else {
      sheet.getRange(i + 2, 7).setDataValidation(rule);
    }
  }

  sheet.hideColumns(8);
  sheet.autoResizeColumns(1, 7);
}