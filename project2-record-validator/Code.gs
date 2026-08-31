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
