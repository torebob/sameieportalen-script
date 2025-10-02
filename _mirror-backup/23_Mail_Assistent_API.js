/**
 * @OnlyCurrentDoc
 *
 * FILE: 23_Mail_Assistent_API.js
 * VERSION: 1.0.0
 * AUTHOR: Jules
 * DATE: 2025-09-28
 *
 * DESCRIPTION:
 * Backend API for the Email Assistant UI. Handles data fetching and actions
 * for managing categorized emails.
 */

function openEmailAssistant() {
  const html = HtmlService.createHtmlOutputFromFile('40_Mail_Assistent_UI.html')
    .setWidth(1100)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, 'E-postassistent');
}

function getEmailsForAssistant() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.EPOST_LOGG);
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const headers = SHEET_HEADERS[SHEETS.EPOST_LOGG];

    const cId = headers.indexOf('Epost-ID');
    const cReceived = headers.indexOf('Mottatt-Dato');
    const cSender = headers.indexOf('Avsender');
    const cSubject = headers.indexOf('Emne');
    const cCategory = headers.indexOf('Kategori');
    const cStatus = headers.indexOf('Status');
    const cSuggestion = headers.indexOf('Svar-Forslag');

    const emails = data.map(row => ({
      id: row[cId],
      receivedDate: row[cReceived],
      sender: row[cSender],
      subject: row[cSubject],
      category: row[cCategory],
      status: row[cStatus],
      replySuggestion: row[cSuggestion]
    })).filter(email => email.status === 'Ny'); // Only show new emails

    return emails;
  } catch (e) {
    safeLog('getEmailsForAssistant_Fail', e.message);
    throw new Error('Kunne ikke hente e-poster: ' + e.message);
  }
}

function handleEmailAction(payload) {
  try {
    const { action, emailId, replyText } = payload;

    if (!action || !emailId) {
      throw new Error('Mangler påkrevd data (handling eller e-post-ID).');
    }

    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.EPOST_LOGG);
    const finder = sheet.createTextFinder(emailId).findNext();
    if (!finder) {
      throw new Error('Fant ikke den spesifiserte e-posten i loggen.');
    }

    const row = finder.getRow();
    const headers = SHEET_HEADERS[SHEETS.EPOST_LOGG];
    const cStatus = headers.indexOf('Status') + 1;
    const cSubject = headers.indexOf('Emne') + 1;
    const cSender = headers.indexOf('Avsender') + 1;
    const cThreadId = headers.indexOf('Tråd-ID') + 1;

    if (action === 'approve') {
      const subject = sheet.getRange(row, cSubject).getValue();
      const sender = sheet.getRange(row, cSender).getValue();
      const threadId = sheet.getRange(row, cThreadId).getValue();

      const thread = GmailApp.getThreadById(threadId);
      if (thread) {
        thread.reply(replyText, {
          name: APP.NAME
        });
      } else {
        // Fallback if thread not found, send new email
        GmailApp.sendEmail(sender, `Re: ${subject}`, replyText, { name: APP.NAME });
      }

      sheet.getRange(row, cStatus).setValue('Behandlet');
      safeLog('Email_Action_Approve', `Godkjente og svarte på e-post ${emailId}`);
      return { ok: true, message: 'E-post godkjent og svar sendt.' };
    }
    else if (action === 'delete') {
      sheet.getRange(row, cStatus).setValue('Slettet');
      safeLog('Email_Action_Delete', `Slettet e-postlogg ${emailId}`);
      return { ok: true, message: 'E-postlogg merket som slettet.' };
    }

    throw new Error('Ukjent handling.');

  } catch (e) {
    safeLog('handleEmailAction_Fail', e.message);
    return { ok: false, message: e.message };
  }
}