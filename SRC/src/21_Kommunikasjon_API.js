/* ====================== Kommunikasjon (Oppslag) - API ======================
 * FILE: 21_Kommunikasjon_API.gs | VERSION: 1.0.0 | UPDATED: 2025-09-14
 * FORMÅL: Backend for å lage, sende og spore digitale oppslag.
 * Inkluderer en web app (doGet) som fungerer som sporingspiksel.
 * ========================================================================== */

/**
 * Åpner brukergrensesnittet for å sende et nytt oppslag.
 */
function openNyttOppslagUI() {
  const html = HtmlService.createHtmlOutputFromFile('36_NyttOppslag.html')
    .setWidth(600).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Send nytt oppslag');
}

/**
 * Henter tilgjengelige målgrupper for oppslag.
 * @returns {string[]} En liste med målgruppe-navn.
 */
function getRecipientGroups() {
  // Denne kan gjøres mer avansert senere ved å lese fra Seksjoner-arket
  return ['Alle beboere', 'Kun eiere', 'Kun styret'];
}

/**
 * Oppretter og sender et nytt oppslag til en valgt målgruppe.
 * @param {object} payload - Data fra HTML-skjemaet: { tittel, innhold, maalgruppe }.
 * @returns {object} Et suksess- eller feilobjekt.
 */
function sendOppslag(payload) {
  try {
    requirePermission('SEND_OPPSLAG');
    
    if (!payload || !payload.tittel || !payload.innhold || !payload.maalgruppe) {
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

    // Filtrer mottakere basert på målgruppe
    const recipients = data.filter(row => {
      const rolle = String(row[cRolle] || '').toLowerCase();
      if (payload.maalgruppe === 'Alle beboere') return true;
      if (payload.maalgruppe === 'Kun eiere') return rolle === 'eier';
      if (payload.maalgruppe === 'Kun styret') return rolle === 'styremedlem' || rolle === 'kjernebruker';
      return false;
    });

    if (recipients.length === 0) {
      throw new Error("Fant ingen mottakere for den valgte målgruppen.");
    }

    const oppslagSheet = _ensureSheetWithHeaders_(SHEETS.OPPSLAG, SHEET_HEADERS[SHEETS.OPPSLAG]);
    const oppslagId = `OPP-${Utilities.getUuid().slice(0, 8)}`;
    const forfatter = Session.getActiveUser().getEmail();

    // Logg oppslaget FØR utsending
    oppslagSheet.appendRow([oppslagId, payload.tittel, payload.innhold, forfatter, new Date(), payload.maalgruppe, recipients.length, 0]);

    // Send personlige e-poster med sporingspiksel
    const webAppUrl = ScriptApp.getService().getUrl();
    
    recipients.forEach(recipient => {
      const personId = recipient[cPersonId];
      const personEmail = recipient[cEpost];

      if (personId && personEmail && webAppUrl) {
        const trackingUrl = `${webAppUrl}?oppslagId=${oppslagId}&personId=${personId}`;
        const htmlBody = `
          <html><body>
            <h2>${payload.tittel}</h2>
            <div style="white-space: pre-wrap; font-size: 14px;">${payload.innhold}</div>
            <p>Med vennlig hilsen,<br>Styret</p>
            <img src="${trackingUrl}" width="1" height="1" alt="">
          </body></html>`;
        
        GmailApp.sendEmail(personEmail, payload.tittel, "", {
          htmlBody: htmlBody,
          name: APP.NAME // Sender fra et hyggelig navn
        });
      }
    });

    _logEvent('Kommunikasjon', `Sendte oppslag "${oppslagId}" til ${recipients.length} mottakere.`);
    return { ok: true, message: `Oppslag sendt til ${recipients.length} mottakere.` };

  } catch (e) {
    _logEvent('Kommunikasjon_Feil', e.message);
    throw e;
  }
}

/**
 * Web App som fungerer som sporingspiksel. Kjøres når en mottaker åpner e-posten.
 * @param {object} e - Event-objektet fra web-forespørselen.
 */
function doGet(e) {
  try {
    const { oppslagId, personId } = e.parameter;

    if (oppslagId && personId) {
      const sporingSheet = _ensureSheetWithHeaders_(SHEETS.OPPSLAG_SPORING, SHEET_HEADERS[SHEETS.OPPSLAG_SPORING]);
      const data = sporingSheet.getDataRange().getValues();
      const headers = data.shift();
      const cOppslagId = headers.indexOf('Oppslag-ID');
      const cPersonId = headers.indexOf('Person-ID');

      // Sjekk om denne åpningen allerede er logget for å unngå duplikater
      const alreadyLogged = data.some(row => row[cOppslagId] === oppslagId && row[cPersonId] === personId);
      
      if (!alreadyLogged) {
        const sporingId = `SPOR-${Utilities.getUuid().slice(0,8)}`;
        sporingSheet.appendRow([sporingId, oppslagId, personId, new Date()]);

        // Oppdater teller i hoved-oppslagsarket (litt tregt, kan forbedres)
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
    // Feiler stille for å ikke krasje pikselen. Feil logges i Apps Script-dashboardet.
  }
  
  // Returner en 1x1 transparent GIF
  const gif = Utilities.base64Decode("R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");
  return ContentService.createImage(gif).setMimeType(ContentService.MimeType.GIF);
}

/**
 * Sender innkalling til årsmøte eller annet møte basert på en mal.
 * @param {string} moteId - ID-en til møtet det skal sendes innkalling for.
 * @returns {object} Et suksess- eller feilobjekt.
 */
function sendInnkalling(moteId) {
  try {
    if (!moteId) throw new Error("Møte-ID er påkrevd.");

    // 1. Hent møtedetaljer og saksliste
    const meeting = global.listMeetings_({ scope: 'all' }).find(m => m.id === moteId);
    if (!meeting) throw new Error(`Fant ikke møte med ID: ${moteId}`);

    const agendaItems = global.listAgenda(moteId);

    // 2. Hent mottakere (kun eiere)
    const personerSheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.PERSONER);
    if (!personerSheet || personerSheet.getLastRow() < 2) {
      throw new Error("Finner ingen personer å sende til.");
    }
    const data = personerSheet.getDataRange().getValues();
    const headers = data.shift();
    const cEpost = headers.indexOf('epost');
    const cRolle = headers.indexOf('rolle');

    const recipients = data
      .filter(row => String(row[cRolle] || '').toLowerCase() === 'eier' && row[cEpost])
      .map(row => row[cEpost]);

    if (recipients.length === 0) {
      throw new Error("Fant ingen eiere med registrert e-post.");
    }

    // 3. Bygg HTML-innhold for e-posten
    const meetingDate = Utilities.formatDate(new Date(meeting.dato), Session.getScriptTimeZone(), "dd. MMMM yyyy");
    const subject = `Innkalling til ${meeting.type}: ${meeting.tittel}`;

    let agendaHtml = "<ul>";
    if (agendaItems.length > 0) {
      agendaItems.forEach(item => {
        agendaHtml += `<li><b>${item.saksnr}: ${item.tittel}</b><p>${item.forslag || 'Ingen forslag'}</p></li>`;
      });
    } else {
      agendaHtml += "<li>Saksliste er ikke klar.</li>";
    }
    agendaHtml += "</ul>";

    const htmlBody = `
      <html><body>
        <h2>Innkalling til ${meeting.type}</h2>
        <p>Det kalles inn til ${meeting.type} <b>${meetingDate} kl. ${meeting.start || ''}</b>.</p>
        <p><b>Sted:</b> ${meeting.sted || 'Ikke spesifisert'}</p>
        <hr>
        <h3>Saksliste</h3>
        ${agendaHtml}
        <hr>
        <p>Med vennlig hilsen,<br>Styret</p>
      </body></html>`;

    // 4. Send e-post til alle eiere
    GmailApp.sendEmail(recipients.join(','), subject, "", {
      htmlBody: htmlBody,
      name: APP.NAME
    });

    _logEvent('Kommunikasjon', `Sendte innkalling for møte "${moteId}" til ${recipients.length} eiere.`);
    return { ok: true, message: `Innkalling sendt til ${recipients.length} eiere.` };

  } catch (e) {
    _logEvent('Kommunikasjon_Feil', `Feil ved sending av innkalling for ${moteId}: ${e.message}`);
    return { ok: false, message: e.message };
  }
}

// Eksporter funksjonen for å gjøre den tilgjengelig for UI
global.sendInnkalling = sendInnkalling;
