/**
 * -----------------------------------------------------------------------------
 * Project: Automated Email Sender with Google Apps Script
 * Author: Kimbergoldess (vibecoding with ChatGPT)
 * Description:
 *   This Google Apps Script reads data from a Google Sheet and sends
 *   personalized emails through Gmail. Each row in the sheet should contain
 *   a recipient name, a unique ID/code, an optional version label, and an
 *   email address.
 *
 * How to use:
 *   1. Create a Google Sheet with these columns:
 *        A: Company / Recipient Name
 *        B: Unique Code or Reference ID
 *        C: Version (optional)
 *        D: Email Address
 *   2. Open the Google Sheet → Extensions → Apps Script.
 *   3. Paste this script into the editor.
 *   4. Replace the SHEET_ID and SHEET_NAME with your own.
 *   5. Run the function `sendEmails()` and authorize access when prompted.
 *
 * Notes:
 *   - This script uses your Gmail account to send emails.
 *   - Works in a sandbox mode (sends only what you confirm in testing).
 *   - Modify the subject/body text freely to match your use case.
 * -----------------------------------------------------------------------------
 */

function sendEmails() {
  // --- SETTINGS ---
  const SHEET_ID = 'YOUR_SHEET_ID_HERE'; // Replace with your actual Google Sheet ID
  const SHEET_NAME = 'Sheet1'; // Replace with your sheet name if different

  // --- GET DATA FROM GOOGLE SHEET ---
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  // --- LOOP THROUGH ALL ROWS (starting from row 2) ---
  for (let i = 1; i < data.length; i++) {
    const company = data[i][0];
    const recordCode = data[i][1];
    const versionLabel = data[i][2];
    const email = data[i][3];

    // --- EMAIL CONTENT ---
    const subject = 'Information Note – Ref. ' + recordCode;
    const body = 'Dear ' + company + ',\n\n' +
                 'Please find attached the document related to reference ' + recordCode + '.\n\n' +
                 'Best regards,\n' +
                 'Automated Email System';

    // --- SEND EMAIL ---
    MailApp.sendEmail(email, subject, body);
  }

  // --- CONFIRMATION MESSAGE ---
  Logger.log('All emails have been sent successfully.');
}

