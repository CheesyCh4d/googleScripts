/**
 * Adds a new user to ONE or MANY selected files.
 * * Workflow:
 * 1. Select rows.
 * 2. Enter Email.
 * 3. Select Permission (Editor, Viewer, Commenter).
 * 4. Script inserts rows and sets everything up instantly.
 */
function addNewUser() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // 1. Get the selection
  const rangeList = sheet.getActiveRangeList();
  if (!rangeList) {
    ui.alert('Please select at least one file row.');
    return;
  }
  
  const ranges = rangeList.getRanges();
  let uniqueRows = new Set();

  // Extract unique row numbers
  ranges.forEach(range => {
    const start = range.getRow();
    const count = range.getNumRows();
    for (let i = 0; i < count; i++) {
      uniqueRows.add(start + i);
    }
  });

  // Sort DESCENDING (High to Low) to keep indices valid during insertion
  const rowsToProcess = Array.from(uniqueRows).sort((a, b) => b - a);
  const validRows = rowsToProcess.filter(r => r > 1); // Skip header

  if (validRows.length === 0) {
    ui.alert('Please select valid file rows (not the header).');
    return;
  }

  // 2. Ask for the Email
  const emailResp = ui.prompt(
    'Add New User',
    `Enter email to add to these ${validRows.length} items:`,
    ui.ButtonSet.OK_CANCEL
  );

  if (emailResp.getSelectedButton() !== ui.Button.OK) return;
  const newEmail = emailResp.getResponseText().trim();

  if (!newEmail) {
    ui.alert('Email cannot be empty.');
    return;
  }

  // 3. Ask for the Permission Level (NEW STEP)
  const roleResp = ui.prompt(
    'Select Permission Level',
    `What role should ${newEmail} have?\n\n` +
    'Type one letter:\n' +
    'E = Editor\n' +
    'V = Viewer\n' +
    'C = Commenter',
    ui.ButtonSet.OK_CANCEL
  );

  if (roleResp.getSelectedButton() !== ui.Button.OK) return;
  
  const roleInput = roleResp.getResponseText().toUpperCase().trim();
  let newRole = 'Viewer'; // Default fallback

  switch (roleInput) {
    case 'E': newRole = 'Editor'; break;
    case 'V': newRole = 'Viewer'; break;
    case 'C': newRole = 'Commenter'; break;
    default:
      ui.alert('Invalid input. Defaulting to "Viewer".');
      newRole = 'Viewer';
  }

  // 4. Loop through rows and insert
  validRows.forEach(rowIndex => {
    // Get file info from the current row
    // Columns: 1=Type, 2=Path, 3=Name, 4=Link, 5=Email, 6=Role, 7=Adj, 8=ID
    const rowValues = sheet.getRange(rowIndex, 1, 1, 8).getValues()[0];
    const fileId = rowValues[7]; // Column H (Index 7)

    if (!fileId) return;

    // Insert new row AFTER the current one
    sheet.insertRowAfter(rowIndex);
    const newRowIndex = rowIndex + 1;

    // Set values
    sheet.getRange(newRowIndex, 1, 1, 8).setValues([[
      rowValues[0], // Type
      rowValues[1], // Path
      rowValues[2], // Name
      rowValues[3], // Link
      newEmail,     // New Email
      '---',        // Current Role (none)
      newRole,      // <-- PRE-FILLED PERMISSION
      fileId        // File ID
    ]]);

    // Add Dropdown to the new row
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Editor', 'Viewer', 'Commenter', 'Remove Access'], true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(newRowIndex, 7).setDataValidation(rule);

    // Color it yellow to highlight the change
    sheet.getRange(newRowIndex, 1, 1, 7).setBackground('#fff2cc');
  });

  ui.alert(`Success!\nAdded "${newEmail}" as "${newRole}" to ${validRows.length} items.\n\nRun "Sync Changes" to apply.`);
}

/**
 * Changes the "Adjusted Permission" for all selected rows at once.
 */
function bulkEditPermissions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const rangeList = sheet.getActiveRangeList();
  
  if (!rangeList) {
    ui.alert('Please select at least one cell in the rows you want to edit.');
    return;
  }

  const response = ui.prompt(
    'Bulk Permission Editor',
    'Enter the letter for the permission you want to apply to ALL selected rows:\n\n' +
    'E = Editor\n' +
    'V = Viewer\n' +
    'C = Commenter\n' +
    'R = Remove Access',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().toUpperCase().trim();
  let newRole = '';

  switch (input) {
    case 'E': newRole = 'Editor'; break;
    case 'V': newRole = 'Viewer'; break;
    case 'C': newRole = 'Commenter'; break;
    case 'R': newRole = 'Remove Access'; break;
    default:
      ui.alert('Invalid input. Please type E, V, C, or R.');
      return;
  }

  const ranges = rangeList.getRanges();
  let updateCount = 0;

  ranges.forEach(range => {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    if (startRow === 1) return;

    for (let i = 0; i < numRows; i++) {
      const currentRowIndex = startRow + i;
      const currentRole = sheet.getRange(currentRowIndex, 6).getValue();

      if (currentRole === 'Owner') continue;

      sheet.getRange(currentRowIndex, 7).setValue(newRole);
      updateCount++;
    }
  });

  sheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  ui.alert(`Success! Updated ${updateCount} rows to "${newRole}".\n\nRun "Sync Changes" to apply these to Drive.`);
}

/**
 * Scans the sheet for a specific email and updates ONLY the matching rows.
 */
function bulkEditByEmail() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  const emailResponse = ui.prompt(
    'Bulk Edit by Email',
    'Enter the exact email address you want to find (e.g., crash.bandicoot@naughtydog.com):',
    ui.ButtonSet.OK_CANCEL
  );

  if (emailResponse.getSelectedButton() !== ui.Button.OK) return;
  const targetEmail = emailResponse.getResponseText().trim().toLowerCase();
  
  if (!targetEmail) {
    ui.alert('Please enter a valid email address.');
    return;
  }

  const roleResponse = ui.prompt(
    'Select New Role',
    `What permission should ${targetEmail} have?\n\n` +
    'Type one letter:\n' +
    'R = Remove Access\n' +
    'V = Viewer\n' +
    'E = Editor\n' +
    'C = Commenter',
    ui.ButtonSet.OK_CANCEL
  );

  if (roleResponse.getSelectedButton() !== ui.Button.OK) return;

  const input = roleResponse.getResponseText().toUpperCase().trim();
  let newRole = '';

  switch (input) {
    case 'R': newRole = 'Remove Access'; break;
    case 'V': newRole = 'Viewer'; break;
    case 'E': newRole = 'Editor'; break;
    case 'C': newRole = 'Commenter'; break;
    default:
      ui.alert('Invalid code. Please type R, V, E, or C.');
      return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const range = sheet.getRange(2, 5, lastRow - 1, 2);
  const data = range.getValues();
  
  let rangesToUpdate = [];

  for (let i = 0; i < data.length; i++) {
    const rowEmail = String(data[i][0]).toLowerCase().trim();
    const currentRole = data[i][1];

    if (rowEmail === targetEmail) {
      if (currentRole === 'Owner') continue;
      rangesToUpdate.push(`G${i + 2}`);
    }
  }

  if (rangesToUpdate.length === 0) {
    ui.alert(`No editable files found for "${targetEmail}".\n(They might be the Owner, or not in the list.)`);
    return;
  }

  sheet.getRangeList(rangesToUpdate).setValue(newRole);
  ui.alert(`Success!\nUpdated ${rangesToUpdate.length} rows for ${targetEmail} to "${newRole}".\n\nRun "Sync Changes" to apply these to Drive.`);
}