// =============================================================================
// Eierskifte – UI & API (Profesjonell)
// FILE: 02_Eierskifte_API.gs
// VERSION: 2.0.1
// UPDATED: 2025-09-15
// FORMÅL: Vise og prosessere eierskifte via en robust, transaksjonell modell.
// ENDRINGER v2.0.1:
//  - Fjernet redeklarasjon av PROPS (idempotent global)
//  - La til _readAllDataForOwnershipChange_()
//  - Rettet bruk av Seksjonsnr vs Seksjon-ID i transaksjoner og e-post
//  - Rettet dato-parsing (_normalizeDate_)
//  - La til _getSheetHeaders_() og ryddet i _applyOwnershipChanges_()
//  - Smårobuste forbedringer og logging
// AVHENGIGHETER: SHEETS + getSheetData() fra 01_Setup_og_Vedlikehold.gs
// =============================================================================

// ----------------------------- Standardisert modell --------------------------

const COLUMN_MAPPINGS = Object.freeze({
  PERSONER: {
    ID: 'Person-ID', NAVN: 'Navn', EPOST: 'Epost', TELEFON: 'Telefon',
    ROLLE: 'Rolle', AKTIV: 'Aktiv', OPPRETTET_AV: 'Opprettet-Av', OPPRETTET_DATO: 'Opprettet-Dato'
  },
  EIERSKAP: {
    ID: 'Eierskap-ID', SEKSJON_ID: 'Seksjon-ID', PERSON_ID: 'Person-ID',
    FRA_DATO: 'Fra-Dato', TIL_DATO: 'Til-Dato', STATUS: 'Status'
  },
  SEKSJONER: {
    ID: 'Seksjon-ID', NUMMER: 'Nummer', BESKRIVELSE: 'Beskrivelse'
  }
});

// ----------------------------- UI (skjema) -----------------------------------

function openOwnershipForm() {
  const ui = _ui();
  try {
    const html = HtmlService.createHtmlOutputFromFile('EierskifteSkjema').setWidth(560).setHeight(680);
    ui.showModalDialog(html, 'Registrer eierskifte');
  } catch (e) {
    const fallback = HtmlService.createHtmlOutput(
      '<html><body><h3>Eierskifte</h3><p>HTML-filen <code>EierskifteSkjema.html</code> mangler.</p></body></html>'
    ).setWidth(420).setHeight(150);
    ui.showModalDialog(fallback, 'Feil ved åpning av skjema');
  }
}

function getSeksjonerForForm() {
  const data = getSheetData(SHEETS.SEKSJONER);
  return data.map(s => ({
    id: s[COLUMN_MAPPINGS.SEKSJONER.NUMMER],              // vis nummer i UI
    beskrivelse: s[COLUMN_MAPPINGS.SEKSJONER.BESKRIVELSE] // valgfritt felt
  }));
}

// ----------------------------- API (lagring) ---------------------------------

/**
 * Prosesserer eierskifte fra skjema.
 * payload: { seksjonsnr, navn, epost, fraDato, vedlegg_url?, kommentar? }
 */
function processOwnershipForm(payload) {
  const tx = Utilities.getUuid().slice(0,8);
  _logEvent('Transaksjon', `Start eierskifte [${tx}] for seksjon ${payload ? payload.seksjonsnr : 'UKJENT'}`);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActive();
    const allData = _readAllDataForOwnershipChange_(ss);

    const sanitized = _sanitizeAndValidatePayload_(payload, allData);
    _validateOwnershipConsistency_(sanitized, allData);
    const fraDato = _normalizeDate_(sanitized.fraDato);

    const operations = _prepareOwnershipChanges_(sanitized, fraDato, allData);
    _applyOwnershipChanges_(ss, operations, tx);

    _logEvent('Transaksjon', `Fullført eierskifte [${tx}]`);
    _sendConfirmationEmails_(sanitized, fraDato, allData);

    return { ok: true, message: `Eierskifte registrert for seksjon ${sanitized.seksjonsnr}` };

  } catch (error) {
    const isValidationError = /^VALIDERING:/.test(error.message);
    const userMessage = isValidationError ? error.message : 'En uventet teknisk feil oppstod. Kontakt administrator.';
    _logEvent('Transaksjon Feil', `Eierskifte feilet: ${error.message}`);
    throw new Error(userMessage);
  } finally {
    lock.releaseLock();
  }
}

// ----------------------------- Validering ------------------------------------

function _sanitizeAndValidatePayload_(payload, data) {
  if (!payload || typeof payload !== 'object') throw new Error('VALIDERING: Tomt eller ugyldig dataformat.');
  const seksjonsnr = String(payload.seksjonsnr || '').trim();
  const navn = String(payload.navn || '').trim();
  const epost = String(payload.epost || '').trim();

  if (!seksjonsnr) throw new Error('VALIDERING: Seksjonsnummer mangler.');
  if (!navn) throw new Error('VALIDERING: Navn mangler.');
  if (!epost || !VALIDATION_RULES.EMAIL_PATTERN.test(epost)) throw new Error('VALIDERING: Ugyldig e-postadresse.');
  if (!payload.fraDato) throw new Error('VALIDERING: Overtakelsesdato mangler.');

  const exists = data.seksjoner.some(s => String(s[COLUMN_MAPPINGS.SEKSJONER.NUMMER] || '').trim() === seksjonsnr);
  if (!exists) throw new Error(`VALIDERING: Seksjon ${seksjonsnr} finnes ikke.`);

  return { ...payload, seksjonsnr, navn, epost };
}

function _validateOwnershipConsistency_(payload, data) {
  const M_E = COLUMN_MAPPINGS.EIERSKAP;
  const M_S = COLUMN_MAPPINGS.SEKSJONER;

  const seksjon = data.seksjoner.find(s => String(s[M_S.NUMMER] || '').trim() === payload.seksjonsnr);
  if (!seksjon) throw new Error(`Teknisk feil: Fant ikke Seksjon-ID for nummer ${payload.seksjonsnr}`);
  const seksjonId = seksjon[M_S.ID];

  const existingOwnerships = data.eierskap.filter(row => row[M_E.SEKSJON_ID] === seksjonId);
  const activeOwnerships = existingOwnerships.filter(row => !row[M_E.TIL_DATO]);
  if (activeOwnerships.length > 1) {
    throw new Error(`VALIDERING: Flere aktive eierskap funnet for seksjon ${payload.seksjonsnr}. Rydd i data før registrering.`);
  }

  const fraDato = _normalizeDate_(payload.fraDato);
  const latestOwnership = existingOwnerships
    .filter(row => row[M_E.TIL_DATO])
    .sort((a, b) => new Date(b[M_E.TIL_DATO]) - new Date(a[M_E.TIL_DATO]))[0];

  if (latestOwnership && new Date(latestOwnership[M_E.TIL_DATO]) >= fraDato) {
    throw new Error(`VALIDERING: Overtakelsesdato (${payload.fraDato}) må være etter forrige eierskaps sluttdato.`);
  }
}

// ------------------------ Transaksjonshåndtering -----------------------------

function _prepareOwnershipChanges_(payload, fraDato, data) {
  const ops = [];
  const M_P = COLUMN_MAPPINGS.PERSONER;
  const M_E = COLUMN_MAPPINGS.EIERSKAP;
  const M_S = COLUMN_MAPPINGS.SEKSJONER;

  // Finn Seksjon-ID fra seksjonsnummer
  var seksjon = data.seksjoner.find(function(s){ return String(s[M_S.NUMMER] || '').trim() === payload.seksjonsnr; });
  if (!seksjon) throw new Error(`Teknisk feil: Fant ikke Seksjon-ID for nummer ${payload.seksjonsnr}`);
  var seksjonId = seksjon[M_S.ID];

  // Finn/opprett person
  var person = data.personer.find(function(p){ return String(p[M_P.EPOST] || '').toLowerCase() === String(payload.epost).toLowerCase(); });
  var personId = person ? person[M_P.ID] : null;
  if (!personId) {
    personId = 'PER-' + Utilities.getUuid().slice(0,4);
    var newPersonData = {};
    newPersonData[M_P.ID] = personId;
    newPersonData[M_P.NAVN] = payload.navn;
    newPersonData[M_P.EPOST] = payload.epost;
    newPersonData[M_P.ROLLE] = 'Eier';
    newPersonData[M_P.AKTIV] = 'Aktiv';
    newPersonData[M_P.OPPRETTET_AV] = _currentEmail_();
    newPersonData[M_P.OPPRETTET_DATO] = new Date();
    ops.push({ type: 'INSERT', sheetName: SHEETS.PERSONER, data: newPersonData });
  }

  // Lukk aktivt eierskap (om finnes) for denne Seksjon-ID
  var aktivtEierskap = data.eierskap.find(function(e){ return e[M_E.SEKSJON_ID] === seksjonId && !e[M_E.TIL_DATO]; });
  if (aktivtEierskap) {
    var tilDato = new Date(fraDato); tilDato.setDate(tilDato.getDate() - 1);
    var updates = {}; updates[M_E.TIL_DATO] = tilDato;
    ops.push({
      type: 'UPDATE',
      sheetName: SHEETS.EIERSKAP,
      rowId: aktivtEierskap[M_E.ID],
      updates: updates
    });
  }

  // Opprett nytt eierskap (bruk Seksjon-ID, ikke seksjonsnr)
  var newOwnership = {};
  newOwnership[M_E.ID] = 'EIE-' + Utilities.getUuid().slice(0,4);
  newOwnership[M_E.SEKSJON_ID] = seksjonId;
  newOwnership[M_E.PERSON_ID] = personId;
  newOwnership[M_E.FRA_DATO] = fraDato;
  newOwnership[M_E.STATUS] = 'Aktiv';
  ops.push({ type: 'INSERT', sheetName: SHEETS.EIERSKAP, data: newOwnership });

  return ops;
}

function _applyOwnershipChanges_(ss, operations, tx) {
  operations.forEach(function(op){
    _logEvent('Transaksjon', `[${tx}] ${op.type} → ${op.sheetName}`);
    var sheet = ss.getSheetByName(op.sheetName);
    if (!sheet) throw new Error('Mangler ark: ' + op.sheetName);

    var headers = _getSheetHeaders_(op.sheetName);
    if (!headers || !headers.length) throw new Error('Fant ikke headers for ' + op.sheetName);

    if (op.type === 'INSERT') {
      var row = headers.map(function(h){ return op.data.hasOwnProperty(h) ? op.data[h] : ''; });
      sheet.appendRow(row);
    } else if (op.type === 'UPDATE') {
      var data = getSheetData(op.sheetName);        // array av objekter
      var idColumn = headers[0];                    // antar ID i kolonne 1 (i tråd med skjemaet)
      var idx = data.findIndex(function(r){ return String(r[idColumn]) === String(op.rowId); });
      if (idx >= 0) {
        Object.keys(op.updates || {}).forEach(function(h){
          var col = headers.indexOf(h);
          if (col >= 0) sheet.getRange(idx + 2, col + 1).setValue(op.updates[h]);
        });
      } else {
        _logEvent('Transaksjon', `Advarsel: Fant ikke rad for oppdatering (ID=${op.rowId}) i ${op.sheetName}`);
      }
    } else {
      _logEvent('Transaksjon', `Ukjent operasjonstype: ${op.type}`);
    }
  });
}

// ----------------------------- E-postbekreftelser ----------------------------

function _sendConfirmationEmails_(payload, fraDato, allData) {
  try {
    var config = _getEmailConfig_();
    if (!config.enabled) { _logEvent('E-post', 'E-postvarsler er deaktivert.'); return; }

    var M_P = COLUMN_MAPPINGS.PERSONER;
    var M_E = COLUMN_MAPPINGS.EIERSKAP;
    var M_S = COLUMN_MAPPINGS.SEKSJONER;

    // Finn Seksjon-ID
    var seksjon = allData.seksjoner.find(function(s){ return String(s[M_S.NUMMER] || '').trim() === payload.seksjonsnr; });
    var seksjonId = seksjon ? seksjon[M_S.ID] : null;

    var newOwner = { navn: payload.navn, epost: payload.epost };
    var currentOwner = null;

    if (seksjonId) {
      var aktivt = allData.eierskap.find(function(e){ return e[M_E.SEKSJON_ID] === seksjonId && !e[M_E.TIL_DATO]; });
      if (aktivt) {
        currentOwner = allData.personer.find(function(p){ return p[M_P.ID] === aktivt[M_E.PERSON_ID]; }) || null;
      }
    }

    var recipients = [newOwner.epost];
    if (currentOwner && currentOwner[M_P.EPOST]) recipients.push(currentOwner[M_P.EPOST]);
    recipients = recipients.filter(function(x){ return !!x; });
    if (!recipients.length) {
      _logEvent('E-post', `Ingen gyldige mottakere for seksjon ${payload.seksjonsnr}.`);
      return;
    }

    var template = _buildOwnershipChangeEmailTemplate_(payload, fraDato, newOwner, currentOwner);
    _sendEmailWithRetry_(recipients, template, config);

  } catch (error) {
    _logEvent('E-post Feil', `Klarte ikke sende e-post for seksjon ${payload.seksjonsnr}: ${error.message}`);
  }
}

function _buildOwnershipChangeEmailTemplate_(payload, fraDato, newOwner, currentOwner) {
  var subject = 'Bekreftelse på eierskifte for seksjon ' + payload.seksjonsnr;
  var prevOwnerTxt = currentOwner ? ('<li><b>Forrige eier:</b> ' + currentOwner[COLUMN_MAPPINGS.PERSONER.NAVN] + '</li>') : '';
  var body =
    '<p>Hei,</p>' +
    '<p>Dette er en bekreftelse på at eierskifte for <b>seksjon ' + payload.seksjonsnr + '</b> er registrert.</p>' +
    '<ul>' +
      '<li><b>Ny eier:</b> ' + newOwner.navn + ' (' + newOwner.epost + ')</li>' +
      '<li><b>Overtakelsesdato:</b> ' + Utilities.formatDate(fraDato, _tz_(), 'dd.MM.yyyy') + '</li>' +
      prevOwnerTxt +
    '</ul>' +
    '<p>Med vennlig hilsen<br>Styret</p>';
  return { subject: subject, body: body };
}

// ----------------------------- Avhengigheter/stubs ---------------------------

// PROPS: idempotent global (unngår "Identifier 'PROPS' has already been declared")
(function (glob) {
  glob.PROPS = glob.PROPS || PropertiesService.getScriptProperties();
})(globalThis);

function _getEmailConfig_() {
  return {
    enabled: true,                        // kan styres via PROPS/SHEETS.KONFIG
    fromAddress: PROPS.getProperty('MAIL_FROM') || Session.getActiveUser().getEmail(),
    replyTo: PROPS.getProperty('MAIL_REPLYTO') || Session.getActiveUser().getEmail()
  };
}

function _sendEmailWithRetry_(recipients, template, config, maxRetries) {
  maxRetries = Number(maxRetries || 3);
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: template.subject,
        htmlBody: template.body,
        from: config.fromAddress,
        replyTo: config.replyTo
      });
      _logEvent('E-post', 'E-post sendt til: ' + recipients.join(',') + ' (forsøk ' + attempt + ')');
      return;
    } catch (error) {
      _logEvent('E-post Feil', 'Forsøk ' + attempt + ' feilet: ' + error.message);
      if (attempt < maxRetries) Utilities.sleep(1000 * attempt);
      else throw error;
    }
  }
}

function _readAllDataForOwnershipChange_(ss) {
  return {
    personer: getSheetData(SHEETS.PERSONER),
    eierskap: getSheetData(SHEETS.EIERSKAP),
    seksjoner: getSheetData(SHEETS.SEKSJONER)
  };
}

function _getSheetHeaders_(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastColumn() < 1) return [];
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

// Små stubs – forventes allerede i 01_*.gs, men trygg å ha her
function _logEvent(type, message) { try { Logger.log('['+type+'] ' + message); } catch(e){} }
function _ui() { return SpreadsheetApp.getUi(); }
function _tz_() { return Session.getScriptTimeZone() || 'Europe/Oslo'; }
function _currentEmail_(){ return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || ''; }

// Robust dato-parser: yyyy-MM-dd, dd.MM.yyyy, eller Date
function _normalizeDate_(value) {
  if (value instanceof Date) return value;
  var s = String(value || '').trim();
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  throw new Error('VALIDERING: Ugyldig datoformat');
}
