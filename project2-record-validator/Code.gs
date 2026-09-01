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
