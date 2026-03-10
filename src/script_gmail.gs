/**
 * Project:     Automated Email Sender with Google Drive Attachments
 * Description: Reads data from a Google Sheet, finds the corresponding PDF
 *              in a Drive folder (using the canonical filename written by the
 *              VBA script in column G), and sends it as an email attachment.
 *
 * Sheet Column Setup:
 *   A: Company Name | B: Record Code | C: Version | D: Email | E: Date | F: Status (VBA) | G: PDF Filename ← source of truth
 *
 * Column H is used by this script to log email send status independently
 * from the VBA status in column F.
 *
 * Setup:
 *   1. Replace SHEET_ID with your Google Sheet ID (from its URL).
 *   2. Replace FOLDER_ID with your Google Drive folder ID.
 *   3. Run sendEmailsWithAttachments() from the Apps Script editor.
 */

// ---------------------------------------------------------------------------
// SETTINGS — edit these before running
// ---------------------------------------------------------------------------
const SHEET_ID   = 'YOUR_SHEET_ID_HERE';
const SHEET_NAME = 'Sheet1';
const FOLDER_ID  = 'YOUR_DRIVE_FOLDER_ID_HERE';

const MAX_EMAILS_PER_RUN = 90; // Stay safely under Gmail daily quota (100 free / 1500 Workspace)

// Column indices (0-based)
const COL_COMPANY      = 0; // A
const COL_CODE         = 1; // B
const COL_EMAIL        = 3; // D
const COL_PDF_FILENAME = 6; // G  ← written by VBA, read here
const COL_EMAIL_STATUS = 7; // H  ← written by this script

// ---------------------------------------------------------------------------
// MAIN FUNCTION
// ---------------------------------------------------------------------------
function sendEmailsWithAttachments() {

  const sheet  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data   = sheet.getDataRange().getValues();
  const folder = DriveApp.getFolderById(FOLDER_ID);

  let sent    = 0;
  let skipped = 0;
  let errors  = 0;

  // Loop over rows starting from row 2 (index 1)
  for (let i = 1; i < data.length; i++) {

    // --- Quota guard ---
    if (sent >= MAX_EMAILS_PER_RUN) {
      Logger.log('Quota limit reached (' + MAX_EMAILS_PER_RUN + '). Stopping. Re-run to continue.');
      break;
    }

    const company     = data[i][COL_COMPANY]      ? String(data[i][COL_COMPANY]).trim()      : '';
    const recordCode  = data[i][COL_CODE]          ? String(data[i][COL_CODE]).trim()          : '';
    const email       = data[i][COL_EMAIL]         ? String(data[i][COL_EMAIL]).trim()         : '';
    const pdfFileName = data[i][COL_PDF_FILENAME]  ? String(data[i][COL_PDF_FILENAME]).trim()  : '';
    const emailStatus = data[i][COL_EMAIL_STATUS]  ? String(data[i][COL_EMAIL_STATUS]).trim()  : '';

    // --- Skip if required fields are missing ---
    if (!email || !recordCode || !pdfFileName) {
      Logger.log('Row ' + (i + 1) + ' skipped: missing email, record code, or PDF filename (col G).');
      skipped++;
      continue;
    }

    // --- Skip if email already sent (duplicate prevention) ---
    if (emailStatus.toLowerCase().startsWith('sent')) {
      Logger.log('Row ' + (i + 1) + ' skipped: email already sent (' + emailStatus + ').');
      skipped++;
      continue;
    }

    // --- Find PDF in Drive using the canonical filename from column G ---
    let attachment = null;

    try {
      const files = folder.getFilesByName(pdfFileName);
      if (files.hasNext()) {
        attachment = files.next();
      }
    } catch (e) {
      Logger.log('Row ' + (i + 1) + ' — Drive search error: ' + e.toString());
    }

    if (!attachment) {
      const msg = 'WARNING: PDF not found in Drive — "' + pdfFileName + '". Email not sent.';
      Logger.log('Row ' + (i + 1) + ' — ' + msg);
      sheet.getRange(i + 1, COL_EMAIL_STATUS + 1).setValue('ERROR - PDF not found: ' + pdfFileName);
      errors++;
      continue;
    }

    // --- Build email content ---
    const subject = 'Documentation – Ref. ' + recordCode;
    const body    =
      'Dear ' + company + ',\n\n' +
      'Please find enclosed the document relating to reference ' + recordCode + '.\n\n' +
      'Kind regards,\n' +
      'Automated System';

    // --- Send email ---
    try {
      MailApp.sendEmail({
        to:          email,
        subject:     subject,
        body:        body,
        attachments: [attachment.getAs(MimeType.PDF)]
      });

      // Mark as sent with timestamp (column H)
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
      sheet.getRange(i + 1, COL_EMAIL_STATUS + 1).setValue('Sent - ' + timestamp);

      Logger.log('Row ' + (i + 1) + ' — Email sent to ' + email + ' with attachment: ' + pdfFileName);
      sent++;

    } catch (e) {
      const errMsg = 'ERROR - Send failed: ' + e.toString();
      Logger.log('Row ' + (i + 1) + ' — ' + errMsg);
      sheet.getRange(i + 1, COL_EMAIL_STATUS + 1).setValue(errMsg);
      errors++;
    }
  }

  // --- Summary log ---
  Logger.log(
    '\n========== RUN COMPLETE ==========\n' +
    '  Sent    : ' + sent    + '\n' +
    '  Skipped : ' + skipped + '\n' +
    '  Errors  : ' + errors  + '\n' +
    '=================================='
  );
}
