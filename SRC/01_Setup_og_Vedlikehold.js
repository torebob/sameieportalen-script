// =============================================================================
// Oppsett, Vedlikehold & Kvalitet
// FILE: 01_Setup_og_Vedlikehold.gs
// VERSION: 4.0.0
// UPDATED: 2025-09-26
// FORMÅL:
// - Modernisert med let/const og arrow functions for bedre lesbarhet.
// - Idempotent global konfig (SHEETS/SHEET_HEADERS/VALIDATION_RULES).
// - Dynamisk oppdaging av eksisterende ark-navn.
// - Enkle hjelpefunksjoner for kjerneoppgaver.
// =============================================================================

// ---------------------- SHEETS (kollisjonsfri + dynamisk) --------------------
((glob) => {
  const existing = glob.SHEETS || {};
  const defaults = {
    KONFIG: 'Konfig',
    LOGG: 'Logg',
    REPORT: 'Integritetsrapport',
    PERSONER: 'Personer',
    SEKSJONER: 'Seksjoner',
    EIERSKAP: 'Eierskap',
    TASKS: 'Oppgaver',
    MOTER: 'Møter',
    MOTE_SAKER: 'Møtesaker',
    MOTE_KOMMENTARER: 'Kommentarer',
    MOTE_STEMMER: 'Stemmer',
    PROTOKOLL_GODKJENNING: 'ProtokollGodkjenning'
  };

  const dyn = {};
  try {
    const ss = SpreadsheetApp.getActive();
    if (ss.getSheetByName('Hendelseslogg')) dyn.LOGG = 'Hendelseslogg';
    else if (ss.getSheetByName('Logg')) dyn.LOGG = 'Logg';

    if (ss.getSheetByName('Oppgaver')) dyn.TASKS = 'Oppgaver';
    else if (ss.getSheetByName('TASKS')) dyn.TASKS = 'TASKS';
  } catch (e) {
    // Ignorer feil hvis vi ikke har tilgang til SpreadsheetApp
  }

  glob.SHEETS = { ...defaults, ...existing, ...dyn };
})(globalThis);

// ----------------------- SHEET_HEADERS (flettes inn trygt) -------------------
((glob) => {
  const S = glob.SHEETS;
  const add = {
    [S.KONFIG]: ['Nøkkel', 'Verdi', 'Beskrivelse'],
    [S.LOGG]: ['Tidsstempel', 'Type', 'Bruker', 'Beskrivelse'],
    [S.REPORT]: ['Kj.Dato', 'Kategori', 'Nøkkel', 'Status', 'Detaljer'],
    [S.PERSONER]: ['Person-ID', 'Navn', 'Epost', 'Telefon', 'Rolle', 'Aktiv', 'Opprettet-Av', 'Opprettet-Dato', 'Sist-Endret'],
    [S.SEKSJONER]: ['Seksjon-ID', 'Nummer', 'Beskrivelse', 'Areal', 'Status', 'Opprettet-Av', 'Opprettet-Dato', 'Sist-Endret'],
    [S.EIERSKAP]: ['Eierskap-ID', 'Seksjon-ID', 'Person-ID', 'Fra-Dato', 'Til-Dato', 'Eierandel', 'Status', 'Sist-Endret'],
    [S.TASKS]: ['Oppgave-ID', 'Tittel', 'Beskrivelse', 'Kategori', 'Prioritet', 'Opprettet', 'Frist', 'Status', 'Ansvarlig-ID', 'Seksjonsnr'],
    [S.MOTER]: ['Møte-ID', 'Type', 'Dato', 'Tittel', 'Agenda-URL', 'Protokoll-URL', 'Status', 'Opprettet-Av', 'Opprettet-Dato', 'Sist-Endret'],
    [S.MOTE_SAKER]: ['Sak-ID', 'Møte-ID', 'Tittel', 'Bakgrunn', 'Status', 'Opprettet-Av', 'Opprettet-Dato', 'Sist-Endret'],
    [S.MOTE_KOMMENTARER]: ['Kommentar-ID', 'Sak-ID', 'Bruker-ID', 'Kommentar', 'Opprettet-Dato', 'Sist-Endret'],
    [S.MOTE_STEMMER]: ['Stemme-ID', 'Sak-ID', 'Bruker-ID', 'Votum', 'Begrunnelse', 'Tidspunkt', 'Vekt'],
    [S.PROTOKOLL_GODKJENNING]: ['Godkjenning-ID', 'Møte-ID', 'Protokoll-URL', 'Status', 'Utsendt-Dato', 'Frist-Dato', 'Godkjent-Av-ID', 'Godkjent-Dato', 'Kommentarer', 'Opprettet-Av', 'Sist-Endret'],
  };
  glob.SHEET_HEADERS = { ...(glob.SHEET_HEADERS || {}), ...add };
})(globalThis);

// ---------------------- VALIDATION_RULES (flettes inn trygt) -----------------
((glob) => {
  const add = {
    EMAIL_PATTERN: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    PHONE_PATTERN: /^\+?\d{8,}$/,
    ID_PATTERN: /^[A-Z]{3,4}-\d{4,}$/,
    MEETING_STATUS: ['Planlagt', 'Avholdt', 'Avlyst', 'Utsatt'],
    VOTE_OPTIONS: ['For', 'Mot', 'Blankt', 'Fraværende'],
    APPROVAL_STATUS: ['Sendt', 'Godkjent', 'Avvist', 'Utløpt'],
    TASK_PRIORITY: ['Lav', 'Normal', 'Høy', 'Kritisk'],
    TASK_STATUS: ['Ny', 'Pågår', 'Venter', 'Fullført', 'Kansellert'],
    PERSON_ROLES: ['Eier', 'Leietaker', 'Styremedlem', 'Leder', 'Administrator'],
  };
  glob.VALIDATION_RULES = { ...(glob.VALIDATION_RULES || {}), ...add };
})(globalThis);

// ========================= HOVED: Opprett/valider ark ========================

/** Opprett kjernefaner (kun de som mangler) m/korrekte headere. */
function setupWorkbook() {
  const ss = SpreadsheetApp.getActive();
  const headersMap = globalThis.SHEET_HEADERS || {};
  Object.keys(headersMap).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      const h = headersMap[name];
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
      sh.freezeRows(1);
      Logger.log(`Opprettet ark: ${name}`);
    }
  });
  Logger.log('Oppsett av arbeidsbok er fullført.');
}

/** Full header-sjekk (tillater små variasjoner på Oppgaver). */
function validateSheetHeaders(sheetName) {
  const headersMap = globalThis.SHEET_HEADERS || {};
  if (!headersMap[sheetName]) throw new Error(`Ingen header-definisjon for: ${sheetName}`);

  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return false;

  const expected = headersMap[sheetName];
  const actual = sh.getRange(1, 1, 1, Math.min(sh.getLastColumn(), expected.length)).getValues()[0];

  const exact = expected.every((h, i) => String(actual[i] || '') === String(h));
  if (exact) return true;

  if (sheetName === globalThis.SHEETS.TASKS) {
    const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    const want = expected.map(norm);
    const have = (sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || []).map(norm);
    const ok = want.every(w =>
      have.includes(w) ||
      (w === 'oppgaveid' && have.includes('oppgaveid')) ||
      (w === 'ansvarligid' && (have.includes('ansvarlig') || have.includes('ansvarligid'))) ||
      (w === 'opprettetav' && (have.includes('opprettet') || have.includes('opprettetav'))) ||
      (w === 'seksjonsnr' && (have.includes('seksjonsnr') || have.includes('seksjon')))
    );
    if (ok) {
      Logger.log(`Header OK (loose) for ${sheetName}`);
      return true;
    }
  }

  Logger.log(`Header-mismatch i ${sheetName}. Forventet: [${expected.join(', ')}], faktisk: [${actual.join(', ')}]`);
  return false;
}

/** Aggregert sjekk av sentrale ark + enkle integritetstester. */
function validateAllSheets() {
  const names = Object.keys(globalThis.SHEET_HEADERS || {});
  const res = {};
  names.forEach(n => {
    const sh = SpreadsheetApp.getActive().getSheetByName(n);
    res[n] = { exists: !!sh, headersValid: sh ? validateSheetHeaders(n) : false };
  });
  const issues = runDataIntegrityChecks();
  const overall = Object.keys(res).every(k => res[k].exists && res[k].headersValid) && issues.length === 0;
  Logger.log(`Validering: ${overall ? 'OK' : 'FEIL'} – issues: ${issues.length}`);
  return { sheetValidation: res, integrityIssues: issues, overallStatus: overall };
}

/** Lettvekts integritetssjekk (ID-referanser + datoer). */
function runDataIntegrityChecks() {
  const S = globalThis.SHEETS;
  const persons = getSheetData(S.PERSONER);
  const sections = getSheetData(S.SEKSJONER);
  const ownerships = getSheetData(S.EIERSKAP);
  const meetings = getSheetData(S.MOTER);
  const cases = getSheetData(S.MOTE_SAKER);

  const pIds = new Set(persons.map(p => p['Person-ID']));
  const sIds = new Set(sections.map(s => s['Seksjon-ID']));
  const mIds = new Set(meetings.map(m => m['Møte-ID']));

  const issues = [];
  (cases || []).forEach(c => {
    if (c['Møte-ID'] && !mIds.has(c['Møte-ID'])) {
      issues.push(`Møtesak [${c['Sak-ID'] || '?'}] peker til ukjent møte [${c['Møte-ID']}]`);
    }
  });
  (ownerships || []).forEach(o => {
    if (o['Person-ID'] && !pIds.has(o['Person-ID'])) {
      issues.push(`Eierskap [${o['Eierskap-ID'] || '?'}] ukjent person [${o['Person-ID']}]`);
    }
    if (o['Seksjon-ID'] && !sIds.has(o['Seksjon-ID'])) {
      issues.push(`Eierskap [${o['Eierskap-ID'] || '?'}] ukjent seksjon [${o['Seksjon-ID']}]`);
    }
    const f = o['Fra-Dato'] ? new Date(o['Fra-Dato']) : null;
    const t = o['Til-Dato'] ? new Date(o['Til-Dato']) : null;
    if (f && t && f >= t) {
      issues.push(`Eierskap [${o['Eierskap-ID'] || '?'}] Til-Dato før/lik Fra-Dato`);
    }
  });
  return issues;
}

// ============================== HJELPERE (data) ==============================
function getSheetData(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return values.map(row =>
    headers.reduce((obj, header, index) => {
      obj[header] = row[index];
      return obj;
    }, {})
  );
}

// ==================== PRIORITET 1–4: SNARVEISFUNKSJONER =====================

/** (1) Beskytt hendelseslogg (Hendelseslogg/Logg) med kun eier som editor. */
function protectEventLogSheet() {
  const S = globalThis.SHEETS;
  const ss = SpreadsheetApp.getActive();
  const name = ss.getSheetByName(S.LOGG) ? S.LOGG : (ss.getSheetByName('Hendelseslogg') ? 'Hendelseslogg' : (ss.getSheetByName('Logg') ? 'Logg' : null));
  if (!name) throw new Error('Fant ikke ark for Hendelseslogg/Logg.');

  const sh = ss.getSheetByName(name);
  const prot = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || sh.protect();
  prot.setDescription('Beskyttet av Setup v4.0.0').setWarningOnly(false);

  const me = Session.getEffectiveUser().getEmail();
  try {
    prot.removeEditors(prot.getEditors());
    prot.addEditor(me);
  } catch (e) {
    // Ignorer feil hvis det ikke er noen editorer å fjerne
  }
  if (prot.canDomainEdit()) prot.setDomainEdit(false);

  return { ok: true, sheet: name };
}

/** (2) Installer tidstriggere for HMS (uten duplikater). */
function hmsInstallTriggersSafe() {
  const have = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  const ensureTrigger = (fn, triggerBuilder) => {
    if (have.includes(fn)) return;
    if (typeof globalThis[fn] === 'function') {
      triggerBuilder(ScriptApp.newTrigger(fn)).create();
    }
  };

  ensureTrigger('hmsNotifyUpcomingTasks', trigger => trigger.timeBased().atHour(7).everyDays(1));
  ensureTrigger('hmsGenerateTasks', trigger => trigger.timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6));
  ensureTrigger('runFormsDailyScheduler', trigger => trigger.timeBased().atHour(5).everyDays(1));

  return { ok: true };
}

/** (3a) Kjør budsjett-validering (finner ark BUDSJETT/Budsjett automatisk). */
function budgetAuditQuick() {
  const ss = SpreadsheetApp.getActive();
  const target = ss.getSheetByName('BUDSJETT') || ss.getSheetByName('Budsjett');
  if (!target) return { ok: false, error: 'Fant ikke BUDSJETT/Budsjett-ark.' };
  if (typeof globalThis.auditBudgetTemplate !== 'function') {
    return { ok: false, error: 'auditBudgetTemplate() mangler (se 55_Budsjett_Audit.gs).' };
  }
  return globalThis.auditBudgetTemplate(target);
}

/** (3b) Åpne budsjett-UI hvis funksjon finnes i budsjettmodulen. */
function openBudgetUIQuick() {
  if (typeof globalThis.openBudgetWebapp !== 'function') {
    throw new Error('Budsjett-UI mangler (openBudgetWebapp). Sjekk 50–57-filene.');
  }
  return globalThis.openBudgetWebapp();
}

/** (4) Kjør prosjektoversikt (97__Project_Overview) om tilgjengelig. */
function projectOverviewQuick() {
  if (typeof globalThis.projectOverview !== 'function') {
    throw new Error('projectOverview() mangler (97__Project_Overview).');
  }
  return globalThis.projectOverview();
}

/** One-click som kjører 1→4 i riktig rekkefølge. */
function runPrioritizedSetup() {
  const out = {};
  setupWorkbook();
  out.protectLog = protectEventLogSheet();
  out.triggers = hmsInstallTriggersSafe();
  out.budgetAudit = budgetAuditQuick();
  out.projOverview = (() => {
    try {
      return projectOverviewQuick();
    } catch (e) {
      return { ok: false, error: e.message };
    }
  })();
  return out;
}
