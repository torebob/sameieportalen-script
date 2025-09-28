/* ====================== Kommunikasjon (Oppslag) - API ======================
 * FILE: 21_Kommunikasjon_API.gs | VERSION: 2.0.0 | UPDATED: 2025-09-26
 * FORMÅL: Backend for å lage, sende og spore digitale oppslag.
 * ENDRINGER v2.0.0:
 *  - Modernisert til let/const og arrow functions.
 *  - Forbedret kodestruktur og lesbarhet.
 * ========================================================================== */

function openNyttOppslagUI() {
  const html = HtmlService.createHtmlOutputFromFile('36_NyttOppslag.html')
    .setWidth(600).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send nytt oppslag');
}

function getRecipientGroups() {
  return ['Alle beboere', 'Kun eiere', 'Kun styret'];
}

function sendOppslag(payload) {
  try {
    requirePermission('SEND_OPPSLAG');

    if (!payload?.tittel || !payload.innhold || !payload.maalgruppe) {
      throw new Error("Mangler tittel, innhold eller målgruppe.");
    }

    const personerSheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.PERSONER);
    if (!personerSheet || personerSheet.getLastRow() < 2) {
      throw new Error("Finner ingen personer å sende til.");
    }

    const data = personerSheet.getDataRange().getValues();
    const headers = data.shift();
    const cEpost = headers.indexOf('epost');
    const cRolle = headers.indexOf('rolle');
    const cPersonId = headers.indexOf('person_id');

    const recipients = data.filter(row => {
      const rolle = String(row[cRolle] || '').toLowerCase();
      if (payload.maalgruppe === 'Alle beboere') return true;
      if (payload.maalgruppe === 'Kun eiere') return rolle === 'eier';
      if (payload.maalgruppe === 'Kun styret') return ['styremedlem', 'kjernebruker'].includes(rolle);
      return false;
    });

    if (recipients.length === 0) {
      throw new Error("Fant ingen mottakere for den valgte målgruppen.");
    }

    const oppslagSheet = _ensureSheetWithHeaders_(SHEETS.OPPSLAG, SHEET_HEADERS[SHEETS.OPPSLAG]);
    const oppslagId = `OPP-${Utilities.getUuid().slice(0, 8)}`;
    const forfatter = Session.getActiveUser().getEmail();

    oppslagSheet.appendRow([oppslagId, payload.tittel, payload.innhold, forfatter, new Date(), payload.maalgruppe, recipients.length, 0]);

    const webAppUrl = ScriptApp.getService().getUrl();

    recipients.forEach(recipient => {
      const personId = recipient[cPersonId];
      const personEmail = recipient[cEpost];

      if (personId && personEmail && webAppUrl) {
        const trackingUrl = `${webAppUrl}?page=tracking&oppslagId=${oppslagId}&personId=${personId}`;
        const htmlBody = `
          <html><body>
            <h2>${payload.tittel}</h2>
            <div style="white-space: pre-wrap; font-size: 14px;">${payload.innhold}</div>
            <p>Med vennlig hilsen,<br>Styret</p>
            <img src="${trackingUrl}" width="1" height="1" alt="">
          </body></html>`;

        GmailApp.sendEmail(personEmail, payload.tittel, "", {
          htmlBody: htmlBody,
          name: APP.NAME
        });
      }
    });

    _safeLog_('Kommunikasjon', `Sendte oppslag "${oppslagId}" til ${recipients.length} mottakere.`);
    return { ok: true, message: `Oppslag sendt til ${recipients.length} mottakere.` };

  } catch (e) {
    _safeLog_('Kommunikasjon_Feil', e.message);
    throw e;
  }
}

function handleTrackingPixelRequest(e) {
  try {
    const { oppslagId, personId } = e.parameter;

    if (oppslagId && personId) {
      const sporingSheet = _ensureSheetWithHeaders_(SHEETS.OPPSLAG_SPORING, SHEET_HEADERS[SHEETS.OPPSLAG_SPORING]);
      const data = sporingSheet.getDataRange().getValues();
      const headers = data.shift();
      const cOppslagId = headers.indexOf('Oppslag-ID');
      const cPersonId = headers.indexOf('Person-ID');

      const alreadyLogged = data.some(row => row[cOppslagId] === oppslagId && row[cPersonId] === personId);

      if (!alreadyLogged) {
        const sporingId = `SPOR-${Utilities.getUuid().slice(0,8)}`;
        sporingSheet.appendRow([sporingId, oppslagId, personId, new Date()]);

        const oppslagSheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.OPPSLAG);
        const oppslagFinder = oppslagSheet.createTextFinder(oppslagId).findNext();
        if (oppslagFinder) {
          const row = oppslagFinder.getRow();
          const cAntallApnet = SHEET_HEADERS[SHEETS.OPPSLAG].indexOf('Antall-Åpnet');
          const cell = oppslagSheet.getRange(row, cAntallApnet + 1);
          cell.setValue((cell.getValue() || 0) + 1);
        }
      }
    }
  } catch (e) {
    // Fail silently
  }

  const gif = Utilities.base64Decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
  return ContentService.createImage(gif).setMimeType(ContentService.MimeType.GIF);
}