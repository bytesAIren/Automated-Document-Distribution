/**
 * Project: Automated Email Sender with Google Drive Attachments
 * Descrizione: Cerca un file PDF in una cartella specifica di Google Drive
 * basandosi sul codice record e lo invia come allegato.
 */

function sendEmailsWithAttachments() {
  // --- IMPOSTAZIONI ---
  const SHEET_ID = 'IL_TUO_ID_FOGLIO_QUI'; 
  const SHEET_NAME = 'Sheet1';
  // Inserisci l'ID della cartella di Google Drive dove hai caricato i PDF
  const FOLDER_ID = 'IL_TUO_ID_CARTELLA_DRIVE_QUI'; 

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const folder = DriveApp.getFolderById(FOLDER_ID);

  // Ciclo sulle righe (partendo dalla seconda)
  for (let i = 1; i < data.length; i++) {
    const company = data[i][0];
    const recordCode = data[i][1];
    const email = data[i][3];

    if (!email || !recordCode) continue;

    // --- RICERCA ALLEGATO ---
    // Cerca file che iniziano con il recordCode nella cartella specificata
    const files = folder.getFilesByName(recordCode + "_" + company + ".pdf");
    let attachment = null;
    
    if (files.hasNext()) {
      attachment = files.next();
    } else {
      // Prova ricerca parziale se il nome azienda è diverso
      const partialFiles = folder.searchFiles("title contains '" + recordCode + "'");
      if (partialFiles.hasNext()) {
        attachment = partialFiles.next();
      }
    }

    // --- INVIO EMAIL ---
    const subject = 'Documentazione – Rif. ' + recordCode;
    const body = 'Gentile ' + company + ',\n\n' +
                 'In allegato alla presente inviamo il documento relativo al riferimento ' + recordCode + '.\n\n' +
                 'Cordiali saluti,\n' +
                 'Sistema Automatico';

    try {
      if (attachment) {
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: body,
          attachments: [attachment.getAs(MimeType.PDF)]
        });
        Logger.log('Inviata a ' + email + ' con allegato: ' + attachment.getName());
      } else {
        Logger.log('ATTENZIONE: Allegato non trovato per ' + recordCode + '. Email non inviata.');
      }
    } catch (e) {
      Logger.log('Errore riga ' + (i+1) + ': ' + e.toString());
    }
  }
}
