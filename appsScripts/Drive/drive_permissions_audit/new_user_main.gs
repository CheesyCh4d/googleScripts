/**
 * Adds a new row for a user to the sheet (does not sync to Drive yet).
 */
function addNewUser() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const range = sheet.getActiveRange();
  const row = range.getRow();

  if (row <= 1 || range.getNumRows() > 1) {
    ui.alert('Please select a single row containing the file you want to share.');
    return;
  }

  const rowValues = sheet.getRange(row, 1, 1, 8).getValues()[0];
  const fileId = rowValues[7]; 

  if (!fileId) {
    ui.alert('No File ID found in this row.');
    return;
  }

  const emailResp = ui.prompt('Add New User', `Enter email for: ${rowValues[2]}`, ui.ButtonSet.OK_CANCEL);
  if (emailResp.getSelectedButton() !== ui.Button.OK) return;

  const newEmail = emailResp.getResponseText().trim();
  if (!newEmail) {
    ui.alert('Email cannot be empty.');
    return;
  }

  sheet.insertRowAfter(row);
  const newRowIndex = row + 1;

  sheet.getRange(newRowIndex, 1, 1, 8).setValues([[
    rowValues[0], rowValues[1], rowValues[2], rowValues[3], 
    newEmail, '---', 'Viewer', fileId
  ]]);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Editor', 'Viewer', 'Commenter', 'Remove Access'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(newRowIndex, 7).setDataValidation(rule);
  sheet.getRange(newRowIndex, 1, 1, 7).setBackground('#fff2cc');
}