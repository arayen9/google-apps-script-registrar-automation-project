// Google Doc Template and Destination Folder IDs
// Replace these placeholders with your actual Google Drive IDs
const TEMPLATE_ID = "YOUR_GOOGLE_DOC_TEMPLATE_ID_HERE";
const FOLDER_ID = "YOUR_DESTINATION_FOLDER_ID_HERE";

/**
 * Triggered on Google Form submission.
 * Extracts student responses, populates the doc template,
 * converts to PDF, emails the student, and cleans up temp files.
 */
function generateAndEmailStudentDocument(e) {
  try {
    // 1. Grab values submitted from Google Form
    // Order: [Timestamp, Student Name, Student ID, Program, Email, Term]
    const responses = e.values;
    const timestamp = responses[0];
    const studentName = responses[1];
    const studentId = responses[2];
    const programTitle = responses[3];
    const studentEmail = responses[4];
    const academicTerm = responses[5];

    // Quick check to ensure recipient email exists
    if (!studentEmail) {
      Logger.log("No student email found. Exiting script.");
      return;
    }

    // 2. Open template and make a copy for this specific student
    const templateFile = DriveApp.getFileById(TEMPLATE_ID);
    const targetFolder = DriveApp.getFolderById(FOLDER_ID);
    const copyTitle = "Enrolment Letter - " + studentName + " (" + studentId + ")";
    const newDocFile = templateFile.makeCopy(copyTitle, targetFolder);

    // 3. Open the newly copied doc and replace all the {{placeholders}}
    const newDoc = DocumentApp.openById(newDocFile.getId());
    const docBody = newDoc.getBody();

    docBody.replaceText("{{Name}}", studentName);
    docBody.replaceText("{{StudentNumber}}", studentId);
    docBody.replaceText("{{Program}}", programTitle);
    docBody.replaceText("{{Email}}", studentEmail);
    docBody.replaceText("{{Term}}", academicTerm);
    docBody.replaceText("{{Timestamp}}", timestamp);

    // saveAndClose is required so Google saves changes before converting to PDF
    newDoc.saveAndClose();

    // 4. Convert the Google Doc to a PDF attachment
    const pdfAttachment = newDocFile.getAs(MimeType.PDF);
    pdfAttachment.setName(studentName + "_Enrolment_Letter.pdf");

    // 5. Send confirmation email with PDF attached
    const subject = "Official Document: Enrolment Verification - " + academicTerm;
    const emailBody = "Hi " + studentName + ",\n\n" +
      "Your request for an official enrolment letter has been processed. " +
      "Please find your verification document attached to this email as a PDF.\n\n" +
      "Details on file:\n" +
      "- Student ID: " + studentId + "\n" +
      "- Program: " + programTitle + "\n" +
      "- Term: " + academicTerm + "\n\n" +
      "Best regards,\n" +
      "Office of the Registrar";

    GmailApp.sendEmail(studentEmail, subject, emailBody, {
      name: "Office of the Registrar",
      attachments: [pdfAttachment]
    });

    // 6. Delete temporary Google Doc copy to save Drive space (PDF is already sent)
    newDocFile.setTrashed(true);

    Logger.log("Successfully created and emailed PDF to: " + studentEmail);

  } catch (err) {
    Logger.log("Error in generateAndEmailStudentDocument: " + err.toString());
  }
}

/**
 * Test function to run directly in the Apps Script editor without submitting a form.
 */
function test_generateAndEmailStudentDocument() {
  const fakeEvent = {
    values: [
      new Date().toLocaleString("en-CA"),
      "Jane Doe",
      "501234567",
      "Computer Science (BSc)",
      "your_test_email@domain.com", // Replace with your actual email when testing locally
      "Fall 2026"
    ]
  };

  Logger.log("Running manual test...");
  generateAndEmailStudentDocument(fakeEvent);
  Logger.log("Test finished.");
}
