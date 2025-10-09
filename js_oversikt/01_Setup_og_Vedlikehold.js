// =============================================================================
// Oppsett, Vedlikehold & Kvalitet
// FILE: 01_Setup_og_Vedlikehold.gs
// VERSION: 3.1.0
// UPDATED: 2025-09-15
// FORMÅL:
// - Idempotent global konfig (SHEETS/SHEET_HEADERS/VALIDATION_RULES) uten
//   "Identifier ... has already been declared"
// - Dynamisk oppdaging av eksisterende ark-navn (Hendelseslogg/Logg, Oppgaver/TASKS)
// - Enkle hjelpefunksjoner for: kjernefaner, beskytte logg, HMS-triggere,
//   budsjett-validering og prosjektoversikt (til 97__Project_Overview)
// =============================================================================

// ---------------------- SHEETS (kollisjonsfri + dynamisk) --------------------
(function (glob) {
  var existing = glob.SHEETS || {};
  var defaults = {
    KONFIG: 'Konfig',
    LOGG: 'Logg',
    REPORT: 'Integritetsrapport',
    PERSONER: 'Personer',
    SEKSJONER: 'Seksjoner',
    EIERSKAP: 'Eierskap',
    TASKS: 'Oppgaver',           // norsk som standard
    MOTER: 'Møter',
    MOTE_SAKER: 'Møtesaker',
    MOTE_KOMMENTARER: 'Kommentarer',
    MOTE_STEMMER: 'Stemmer',
    PROTOKOLL_GODKJENNING: 'ProtokollGodkjenning'
  };

  // Dynamisk autodeteksjon hvis vi allerede har ark i dokumentet
  var dyn = {};
  try {
    var ss = SpreadsheetApp.getActive();
    if (ss.getSheetByName('Hendelseslogg'))      dyn.LOGG  = 'Hendelseslogg';
    else if (ss.getSheetByName('Logg'))         dyn.LOGG  = 'Logg';

    if (ss.getSheetByName('Oppgaver'))          dyn.TASKS = 'Oppgaver';
    else if (ss.getSheetByName('TASKS'))        dyn.TASKS = 'TASKS';
  } catch (_) {}

  glob.SHEETS = Object.assign({}, defaults, existing, dyn);
})(globalThis);

// ----------------------- SHEET_HEADERS (flettes inn trygt) -------------------
(function (glob) {
  var S = glob.SHEETS;
  var add = {};
  add[S.KONFIG]   = ['Nøkkel','Verdi','Beskrivelse'];
  add[S.LOGG]     = ['Tidsstempel','Type','Bruker','Beskrivelse'];
  add[S.REPORT]   = ['Kj.Dato','Kategori','Nøkkel','Status','Detaljer'];
  add[S.PERSONER] = ['Person-ID','Navn','Epost','Telefon','Rolle','Aktiv','Opprettet-Av','Opprettet-Dato','Sist-Endret'];
  add[S.SEKSJONER]= ['Seksjon-ID','Nummer','Beskrivelse','Areal','Status','Opprettet-Av','Opprettet-Dato','Sist-Endret'];
  add[S.EIERSKAP] = ['Eierskap-ID','Seksjon-ID','Person-ID','Fra-Dato','Til-Dato','Eierandel','Status','Sist-Endret'];
  // NB: Oppgaver er “loose validated” (tillater OppgaveID/Ansvarlig uten -ID)
  add[S.TASKS]    = ['Oppgave-ID','Tittel','Beskrivelse','Kategori','Prioritet','Opprettet','Frist','Status','Ansvarlig-ID','Seksjonsnr'];
  add[S.MOTER]    = ['Møte-ID','Type','Dato','Tittel','Agenda-URL','Protokoll-URL','Status','Opprettet-Av','Opprettet-Dato','Sist-Endret'];
  add[S.MOTE_SAKER]=['Sak-ID','Møte-ID','Tittel','Bakgrunn','Status','Opprettet-Av','Opprettet-Dato','Sist-Endret'];
  add[S.MOTE_KOMMENTARER]=['Kommentar-ID','Sak-ID','Bruker-ID','Kommentar','Opprettet-Dato','Sist-Endret'];
  add[S.MOTE_STEMMER]=['Stemme-ID','Sak-ID','Bruker-ID','Votum','Begrunnelse','Tidspunkt','Vekt'];
  add[S.PROTOKOLL_GODKJENNING]=['Godkjenning-ID','Møte-ID','Protokoll-URL','Status','Utsendt-Dato','Frist-Dato','Godkjent-Av-ID','Godkjent-Dato','Kommentarer','Opprettet-Av','Sist-Endret'];

  glob.SHEET_HEADERS = Object.assign({}, glob.SHEET_HEADERS || {}, add);
})(globalThis);

// ---------------------- VALIDATION_RULES (flettes inn trygt) -----------------
(function (glob) {
  var add = {
    EMAIL_PATTERN: /^[^\s@]+@[^\s@]+\.[^\s@]+$/,
    PHONE_PATTERN: /^\+?\d{8,}$/,
    ID_PATTERN: /^[A-Z]{3,4}-\d{4,}$/,
    MEETING_STATUS: ['Planlagt','Avholdt','Avlyst','Utsatt'],
    VOTE_OPTIONS: ['For','Mot','Blankt','Fraværende'],
    APPROVAL_STATUS: ['Sendt','Godkjent','Avvist','Utløpt'],
    TASK_PRIORITY: ['Lav','Normal','Høy','Kritisk'],
    TASK_STATUS: ['Ny','Pågår','Venter','Fullført','Kansellert'],
    PERSON_ROLES: ['Eier','Leietaker','Styremedlem','Leder','Administrator']
  };
  glob.VALIDATION_RULES = Object.assign({}, glob.VALIDATION_RULES || {}, add);
})(globalThis);

// ========================= HOVED: Opprett/valider ark ========================

/** Opprett kjernefaner (kun de som mangler) m/korrekte headere. */
function setupWorkbook() {
  var ss = SpreadsheetApp.getActive();
  var headersMap = globalThis.SHEET_HEADERS || {};
  Object.keys(headersMap).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      var h = headersMap[name];
      sh.getRange(1,1,1,h.length).setValues([h]).setFontWeight('bold');
      sh.freezeRows(1);
      Logger.log('Opprettet ark: ' + name);
    }
  });
  Logger.log('Oppsett av arbeidsbok er fullført.');
}

/** Full header-sjekk (tillater små variasjoner på Oppgaver). */
function validateSheetHeaders(sheetName) {
  var headersMap = globalThis.SHEET_HEADERS || {};
  if (!headersMap[sheetName]) throw new Error('Ingen header-definisjon for: ' + sheetName);

  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return false;

  var expected = headersMap[sheetName];
  var actual   = sh.getRange(1,1,1,Math.min(sh.getLastColumn(), expected.length)).getValues()[0];

  // 1) Eksakt match?
  var exact = expected.every(function (h,i){ return String(actual[i]||'') === String(h); });
  if (exact) return true;

  // 2) “Loose” match for Oppgaver: normaliser og sjekk at alle nøkkelkolonner finnes
  if (sheetName === globalThis.SHEETS.TASKS) {
    var norm = function (s){ return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,''); };
    var want = expected.map(norm);
    var have = (sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]||[]).map(norm);
    var ok = want.every(function (w){ return have.indexOf(w) >= 0 || // eksakt
      // tolerer varianter:
      (w==='oppgaveid'      && have.indexOf('oppgaveid')>=0) ||
      (w==='ansvarligid'    && (have.indexOf('ansvarlig')>=0 || have.indexOf('ansvarligid')>=0)) ||
      (w==='opprettetav'    && (have.indexOf('opprettet')>=0 || have.indexOf('opprettetav')>=0)) ||
      (w==='seksjonsnr'     && (have.indexOf('seksjonsnr')>=0 || have.indexOf('seksjon')>=0))
    });
    if (ok) { Logger.log('Header OK (loose) for ' + sheetName); return true; }
  }

  Logger.log('Header-mismatch i ' + sheetName + '. Forventet: [' + expected.join(', ') + '], faktisk: [' + actual.join(', ') + ']');
  return false;
}

/** Aggregert sjekk av sentrale ark + enkle integritetstester. */
function validateAllSheets() {
  var names = Object.keys(globalThis.SHEET_HEADERS || {});
  var res = {};
  names.forEach(function (n) {
    var sh = SpreadsheetApp.getActive().getSheetByName(n);
    res[n] = { exists: !!sh, headersValid: sh ? validateSheetHeaders(n) : false };
  });
  var issues = runDataIntegrityChecks();
  var overall = Object.keys(res).every(function(k){ return res[k].exists && res[k].headersValid; }) && issues.length===0;
  Logger.log('Validering: ' + (overall?'OK':'FEIL') + ' – issues: ' + issues.length);
  return { sheetValidation: res, integrityIssues: issues, overallStatus: overall };
}

/** Lettvekts integritetssjekk (ID-referanser + datoer). */
function runDataIntegrityChecks() {
  var S = globalThis.SHEETS;
  var persons = getSheetData(S.PERSONER);
  var sections = getSheetData(S.SEKSJONER);
  var ownerships = getSheetData(S.EIERSKAP);
  var meetings = getSheetData(S.MOTER);
  var cases = getSheetData(S.MOTE_SAKER);

  var pIds = new Set(persons.map(function(p){return p['Person-ID'];}));
  var sIds = new Set(sections.map(function(s){return s['Seksjon-ID'];}));
  var mIds = new Set(meetings.map(function(m){return m['Møte-ID'];}));

  var issues = [];
  (cases||[]).forEach(function(c){
    if (c['Møte-ID'] && !mIds.has(c['Møte-ID'])) issues.push('Møtesak ['+(c['Sak-ID']||'?')+'] peker til ukjent møte ['+c['Møte-ID']+']');
  });
  (ownerships||[]).forEach(function(o){
    if (o['Person-ID'] && !pIds.has(o['Person-ID'])) issues.push('Eierskap ['+(o['Eierskap-ID']||'?')+'] ukjent person ['+o['Person-ID']+']');
    if (o['Seksjon-ID'] && !sIds.has(o['Seksjon-ID'])) issues.push('Eierskap ['+(o['Eierskap-ID']||'?')+'] ukjent seksjon ['+o['Seksjon-ID']+']');
    var f = o['Fra-Dato'] ? new Date(o['Fra-Dato']) : null;
    var t = o['Til-Dato'] ? new Date(o['Til-Dato']) : null;
    if (f && t && f >= t) issues.push('Eierskap ['+(o['Eierskap-ID']||'?')+'] Til-Dato før/lik Fra-Dato');
  });
  return issues;
}

// ============================== HJELPERE (data) ==============================
function getSheetData(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getDataRange().getValues();
  var headers = values.shift();
  return values.map(function (row) {
    var o = {}; for (var i=0;i<headers.length;i++) o[headers[i]] = row[i]; return o;
  });
}

// ==================== PRIORITET 1–4: SNARVEISFUNKSJONER =====================

/** (1) Beskytt hendelseslogg (Hendelseslogg/Logg) med kun eier som editor. */
function protectEventLogSheet() {
  var S = globalThis.SHEETS;
  var ss = SpreadsheetApp.getActive();
  var name = ss.getSheetByName(S.LOGG) ? S.LOGG : (ss.getSheetByName('Hendelseslogg') ? 'Hendelseslogg' : (ss.getSheetByName('Logg') ? 'Logg' : null));
  if (!name) throw new Error('Fant ikke ark for Hendelseslogg/Logg.');
  var sh = ss.getSheetByName(name);

  var prot = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0] || sh.protect();
  prot.setDescription('Beskyttet av Setup v3.1.0').setWarningOnly(false);
  // kun skripteier/effektiv bruker som editor
  var me = Session.getEffectiveUser().getEmail();
  try { prot.removeEditors(prot.getEditors()); } catch(_){}
  try { prot.addEditor(me); } catch(_){}
  if (prot.canDomainEdit()) prot.setDomainEdit(false);

  return { ok:true, sheet:name };
}

/** (2) Installer tidstriggere for HMS (uten duplikater). */
function hmsInstallTriggersSafe() {
  var have = ScriptApp.getProjectTriggers().map(function(t){ return t.getHandlerFunction(); });
  function ensureDaily(fn, hour){
    if (have.indexOf(fn) >= 0) return;
    if (typeof globalThis[fn] !== 'function') return;
    ScriptApp.newTrigger(fn).timeBased().atHour(hour).everyDays(1).create();
  }
  function ensureWeekly(fn, dow, hour){
    if (have.indexOf(fn) >= 0) return;
    if (typeof globalThis[fn] !== 'function') return;
    var tb = ScriptApp.newTrigger(fn).timeBased().onWeekDay(dow).atHour(hour);
    tb.create();
  }
  // Prøv “beste kjente” handlers hvis de finnes
  ensureDaily('hmsNotifyUpcomingTasks', 7);                         // daglig varsel
  ensureWeekly('hmsGenerateTasks', ScriptApp.WeekDay.MONDAY, 6);    // generér ukeplan
  // Fallback: forms/kalender hvis dere bruker dem
  if (typeof globalThis['runFormsDailyScheduler'] === 'function') ensureDaily('runFormsDailyScheduler', 5);

  return { ok:true };
}

/** (3a) Kjør budsjett-validering (finner ark BUDSJETT/Budsjett automatisk). */
function budgetAuditQuick() {
  var ss = SpreadsheetApp.getActive();
  var target = ss.getSheetByName('BUDSJETT') ? 'BUDSJETT' :
               (ss.getSheetByName('Budsjett') ? 'Budsjett' : null);
  if (!target) return { ok:false, error:'Fant ikke BUDSJETT/Budsjett-ark.' };
  if (typeof globalThis['auditBudgetTemplate'] !== 'function') return { ok:false, error:'auditBudgetTemplate() mangler (se 55_Budsjett_Audit.gs).' };
  return globalThis['auditBudgetTemplate'](target);
}

/** (3b) Åpne budsjett-UI hvis funksjon finnes i budsjettmodulen. */
function openBudgetUIQuick() {
  if (typeof globalThis['openBudgetWebapp'] === 'function') return globalThis['openBudgetWebapp']();
  throw new Error('Budsjett-UI mangler (openBudgetWebapp). Sjekk 50–57-filene.');
}

/** (4) Kjør prosjektoversikt (97__Project_Overview) om tilgjengelig. */
function projectOverviewQuick() {
  if (typeof globalThis['projectOverview'] === 'function') return globalThis['projectOverview']();
  throw new Error('projectOverview() mangler (97__Project_Overview).');
}

/** One-click som kjører 1→4 i riktig rekkefølge. */
function runPrioritizedSetup() {
  var out = {};
  setupWorkbook();
  out.protectLog   = protectEventLogSheet();
  out.triggers     = hmsInstallTriggersSafe();
  out.budgetAudit  = budgetAuditQuick();
  out.projOverview = (function(){ try { return projectOverviewQuick(); } catch(e){ return {ok:false, error:e.message}; }})();
  return out;
}
