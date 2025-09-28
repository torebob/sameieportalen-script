// =============================================================================
// Eierskifte – UI & API (Profesjonell)
// FILE: 02_Eierskifte_API.gs
// VERSION: 3.0.0
// UPDATED: 2025-09-26
// FORMÅL: Vise og prosessere eierskifte via en robust, transaksjonell modell.
// ENDRINGER v3.0.0:
//  - Modernisert til let/const og arrow functions.
//  - Fjernet lokale hjelpefunksjoner, bruker nå 00b_Utils.js.
//  - Forbedret lesbarhet og kodestruktur.
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
function getSeksjonerForForm() {
  const data = getSheetData(SHEETS.SEKSJONER);
  return data.map(s => ({
    id: s[COLUMN_MAPPINGS.SEKSJONER.NUMMER],
    beskrivelse: s[COLUMN_MAPPINGS.SEKSJONER.BESKRIVELSE]
  }));
}

// ----------------------------- API (lagring) ---------------------------------
function processOwnershipForm(payload) {
  const tx = Utilities.getUuid().slice(0, 8);
  safeLog('Transaksjon', `Start eierskifte [${tx}] for seksjon ${payload?.seksjonsnr || 'UKJENT'}`);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActive();
    const allData = _readAllDataForOwnershipChange_(ss);
    const sanitized = _sanitizeAndValidatePayload_(payload, allData);
    _validateOwnershipConsistency_(sanitized, allData);
    const fraDato = normalizeDate(sanitized.fraDato);

    if (!fraDato) {
      throw new Error('VALIDERING: Ugyldig overtakelsesdato.');
    }

    const operations = _prepareOwnershipChanges_(sanitized, fraDato, allData);
    _applyOwnershipChanges_(ss, operations, tx);

    safeLog('Transaksjon', `Fullført eierskifte [${tx}]`);
    _sendConfirmationEmails_(sanitized, fraDato, allData);

    return { ok: true, message: `Eierskifte registrert for seksjon ${sanitized.seksjonsnr}` };

  } catch (error) {
    const isValidationError = /^VALIDERING:/.test(error.message);
    const userMessage = isValidationError ? error.message : 'En uventet teknisk feil oppstod. Kontakt administrator.';
    safeLog('Transaksjon Feil', `Eierskifte feilet: ${error.message}`);
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

  const fraDato = normalizeDate(payload.fraDato);
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
  const { PERSONER: M_P, EIERSKAP: M_E, SEKSJONER: M_S } = COLUMN_MAPPINGS;

  const seksjon = data.seksjoner.find(s => String(s[M_S.NUMMER] || '').trim() === payload.seksjonsnr);
  if (!seksjon) throw new Error(`Teknisk feil: Fant ikke Seksjon-ID for nummer ${payload.seksjonsnr}`);
  const seksjonId = seksjon[M_S.ID];

  let person = data.personer.find(p => String(p[M_P.EPOST] || '').toLowerCase() === String(payload.epost).toLowerCase());
  let personId = person ? person[M_P.ID] : null;

  if (!personId) {
    personId = `PER-${Utilities.getUuid().slice(0, 4)}`;
    const newPersonData = {
      [M_P.ID]: personId,
      [M_P.NAVN]: payload.navn,
      [M_P.EPOST]: payload.epost,
      [M_P.ROLLE]: 'Eier',
      [M_P.AKTIV]: 'Aktiv',
      [M_P.OPPRETTET_AV]: getCurrentEmail(),
      [M_P.OPPRETTET_DATO]: new Date(),
    };
    ops.push({ type: 'INSERT', sheetName: SHEETS.PERSONER, data: newPersonData });
  }

  const aktivtEierskap = data.eierskap.find(e => e[M_E.SEKSJON_ID] === seksjonId && !e[M_E.TIL_DATO]);
  if (aktivtEierskap) {
    const tilDato = new Date(fraDato);
    tilDato.setDate(tilDato.getDate() - 1);
    const updates = { [M_E.TIL_DATO]: tilDato };
    ops.push({ type: 'UPDATE', sheetName: SHEETS.EIERSKAP, rowId: aktivtEierskap[M_E.ID], updates });
  }

  const newOwnership = {
    [M_E.ID]: `EIE-${Utilities.getUuid().slice(0, 4)}`,
    [M_E.SEKSJON_ID]: seksjonId,
    [M_E.PERSON_ID]: personId,
    [M_E.FRA_DATO]: fraDato,
    [M_E.STATUS]: 'Aktiv',
  };
  ops.push({ type: 'INSERT', sheetName: SHEETS.EIERSKAP, data: newOwnership });

  return ops;
}

function _applyOwnershipChanges_(ss, operations, tx) {
  operations.forEach(op => {
    safeLog('Transaksjon', `[${tx}] ${op.type} → ${op.sheetName}`);
    const sheet = ss.getSheetByName(op.sheetName);
    if (!sheet) throw new Error(`Mangler ark: ${op.sheetName}`);

    const headers = _getSheetHeaders_(op.sheetName);
    if (!headers || !headers.length) throw new Error(`Fant ikke headers for ${op.sheetName}`);

    if (op.type === 'INSERT') {
      const row = headers.map(h => op.data[h] ?? '');
      sheet.appendRow(row);
    } else if (op.type === 'UPDATE') {
      const data = getSheetData(op.sheetName);
      const idColumn = headers[0];
      const idx = data.findIndex(r => String(r[idColumn]) === String(op.rowId));
      if (idx >= 0) {
        Object.entries(op.updates || {}).forEach(([h, val]) => {
          const col = headers.indexOf(h);
          if (col >= 0) sheet.getRange(idx + 2, col + 1).setValue(val);
        });
      } else {
        safeLog('Transaksjon', `Advarsel: Fant ikke rad for oppdatering (ID=${op.rowId}) i ${op.sheetName}`);
      }
    } else {
      safeLog('Transaksjon', `Ukjent operasjonstype: ${op.type}`);
    }
  });
}

// ----------------------------- E-postbekreftelser ----------------------------
function _sendConfirmationEmails_(payload, fraDato, allData) {
  try {
    const config = _getEmailConfig_();
    if (!config.enabled) {
      safeLog('E-post', 'E-postvarsler er deaktivert.');
      return;
    }

    const { PERSONER: M_P, EIERSKAP: M_E, SEKSJONER: M_S } = COLUMN_MAPPINGS;
    const seksjon = allData.seksjoner.find(s => String(s[M_S.NUMMER] || '').trim() === payload.seksjonsnr);
    const seksjonId = seksjon?.[M_S.ID];
    const newOwner = { navn: payload.navn, epost: payload.epost };
    let currentOwner = null;

    if (seksjonId) {
      const aktivt = allData.eierskap.find(e => e[M_E.SEKSJON_ID] === seksjonId && !e[M_E.TIL_DATO]);
      if (aktivt) {
        currentOwner = allData.personer.find(p => p[M_P.ID] === aktivt[M_E.PERSON_ID]);
      }
    }

    const recipients = [newOwner.epost, currentOwner?.[M_P.EPOST]].filter(Boolean);
    if (recipients.length === 0) {
      safeLog('E-post', `Ingen gyldige mottakere for seksjon ${payload.seksjonsnr}.`);
      return;
    }

    const template = _buildOwnershipChangeEmailTemplate_(payload, fraDato, newOwner, currentOwner);
    _sendEmailWithRetry_(recipients, template, config);

  } catch (error) {
    safeLog('E-post Feil', `Klarte ikke sende e-post for seksjon ${payload.seksjonsnr}: ${error.message}`);
  }
}

function _buildOwnershipChangeEmailTemplate_(payload, fraDato, newOwner, currentOwner) {
  const subject = `Bekreftelse på eierskifte for seksjon ${payload.seksjonsnr}`;
  const prevOwnerTxt = currentOwner ? `<li><b>Forrige eier:</b> ${currentOwner[COLUMN_MAPPINGS.PERSONER.NAVN]}</li>` : '';
  const body = `
    <p>Hei,</p>
    <p>Dette er en bekreftelse på at eierskifte for <b>seksjon ${payload.seksjonsnr}</b> er registrert.</p>
    <ul>
      <li><b>Ny eier:</b> ${newOwner.navn} (${newOwner.epost})</li>
      <li><b>Overtakelsesdato:</b> ${Utilities.formatDate(fraDato, getScriptTimezone(), 'dd.MM.yyyy')}</li>
      ${prevOwnerTxt}
    </ul>
    <p>Med vennlig hilsen<br>Styret</p>`;
  return { subject, body };
}

// ----------------------------- Avhengigheter/stubs ---------------------------
((glob) => {
  glob.PROPS = glob.PROPS || PropertiesService.getScriptProperties();
})(globalThis);

function _getEmailConfig_() {
  return {
    enabled: true,
    fromAddress: PROPS.getProperty('MAIL_FROM') || Session.getActiveUser().getEmail(),
    replyTo: PROPS.getProperty('MAIL_REPLYTO') || Session.getActiveUser().getEmail()
  };
}

function _sendEmailWithRetry_(recipients, template, config, maxRetries = 3) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      MailApp.sendEmail({
        to: recipients.join(','),
        subject: template.subject,
        htmlBody: template.body,
        from: config.fromAddress,
        replyTo: config.replyTo
      });
      safeLog('E-post', `E-post sendt til: ${recipients.join(',')} (forsøk ${attempt})`);
      return;
    } catch (error) {
      safeLog('E-post Feil', `Forsøk ${attempt} feilet: ${error.message}`);
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
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastColumn() < 1) return [];
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

// MERK: Hjelpefunksjoner er flyttet til 00b_Utils.js.
