/**
 * Project 2: Registrar Record Validator & Email Dispatcher
 * Adds custom menu to audit student IDs and emails, and sends batch updates.
 */

// Add custom menu to Google Sheets toolbar on open
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Registrar Tools")
    .addItem("Validate Records", "validateStudentRecords")
    .addItem("Send Placement Emails", "sendBatchPlacementEmails")
    .addToUi();
}

// Audits student IDs and email formats, flagging errors in red
function validateStudentRecords() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No student data found to validate.');
    return;
  }

// Grab the data 
  const range = sheet.getRange(2, 1, lastRow - 1, 2);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();

// RegEx: Exactly 9 digits
  const idPattern = /^\d{9}$/;
  const emailPattern = /^[a-zA-Z0-9._%+-]+@torontomu\.ca$/;

  let errorCount = 0;

  for (let i = 0; i < values.length; i++) {
    const studentId = String(values[i][0]).trim();
    const email = String(values[i][1]).trim();

    // Check 9-digit Student ID
    if (!idPattern.test(studentId)) {
      backgrounds[i][0] = '#F4CCCC'; // Light red
      errorCount++;
    } else {
      backgrounds[i][0] = '#FFFFFF'; // Reset to white
    }

    // Check Institutional Email
    if (!emailPattern.test(email)) {
      backgrounds[i][1] = '#F4CCCC';
      errorCount++;
    } else {
      backgrounds[i][1] = '#FFFFFF';
    }
  }

  // Batch update colors in one call instead of cell-by-cell
  range.setBackgrounds(backgrounds);

  SpreadsheetApp.getUi().alert(
    'Validation Complete',
    `Audit finished. Found ${errorCount} formatting issue(s).`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Quick helper to reset sheet colors
function clearValidationHighlighting() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 2).setBackground('#FFFFFF');
  }
}
