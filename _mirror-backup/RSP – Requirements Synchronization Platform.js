/**
 * RSP – Requirements Synchronization Platform (All-in-One) v1.4.0
 * Single-file bundle: Config + Utilities + Wizard + Validation + Menu + Lean Sync (Sheet→Doc)
 *
 * WHAT'S INCLUDED
 * - Config (RSP_CFG) with chapter templates + validation settings
 * - Utilities (logging, properties, header resolver, locking helpers)
 * - Full backup (duplicate spreadsheet) + chunked sheet backup
 * - First-run wizard with global locking, progress, baseline v9.3 document seeding
 * - Validation framework (structured issues, config/system/doc/sheet validation)
 * - Menu with health indicators + "first run", validate, open doc, push/pull stubs
 * - Lean Sheet→Doc sync with chunking, per-chunk doc open/close, version log updates
 *
 * ASSUMPTIONS
 * - Your requirements sheet is named "Requirements"
 * - Version log sheet is named "Version_Log"
 * - Sync log sheet is named "Sync_Log"
 * - Headings supported (Norwegian + English aliases)
 *
 * HOW TO USE
 * 1) Open the spreadsheet, then Extensions → Apps Script, paste this file, save.
 * 2) Set RSP_CFG.DOC_ID (below) to your kravdokument ID (or use menu action to set it).
 * 3) Reload spreadsheet → menu “Krav Sync” appears.
 * 4) Run “0) Førstegangsoppsett (wizard)” (creates sheets and baseline v9.3 in Doc if needed).
 * 5) Use “2) Valider system” to confirm healthy state.
 * 6) Use “3) Push (Sheet → Doc)” to synchronize requirements into the Doc with version history.
 */

/* ----------------------------- Configuration ------------------------------ */

var RSP_CFG = (typeof RSP_CFG === 'object' && RSP_CFG) ? RSP_CFG : {
  VERSION: '1.4.0',

  // Set your kravdokument ID here OR set via menu (stored in Script Properties)
  DOC_ID: 'SETT-DIN-DOC-ID-HER',

  SHEET_REQ_NAME: 'Requirements',
  SHEET_VERLOG_NAME: 'Version_Log',
  SHEET_SYNCLOG_NAME: 'Sync_Log',

  // Header mapping (aliases supported)
  HEADERS: {
    id:    ['kravid','krav id','krav-id','id'],
    krav:  ['krav','beskrivelse','tekst','requirement','description','text'],
    prio:  ['prioritet','prio','priority'],
    prog:  ['fremdrift %','fremdrift%','fremdrift','progress','progress %','%'],
    kap:   ['kapittel','kap','chapter'],
    vers:  ['versjon','version'],
    kom:   ['kommentar','comment']
  },

  // Validation module
  VALIDATION: {
    MAX_CHAPTERS: 39,
    MAX_VALIDATION_ROWS: 15000,
    REQUIRED_SHEETS: ['Requirements','Version_Log','Sync_Log'],
    DOC_MARKERS: ['[[RSP_V93_BASELINE_MARKER]]']
  },

  // Seeding & performance
  SEEDING: {
    BATCH_SIZE: 12,
    YIELD_EVERY_BATCHES: 2,
    YIELD_MS: 50
  },

  // Sync (Sheet → Doc)
  SYNC: {
    CHUNK_SIZE: 40,     // rows per doc open/close
    RATE_LIMIT_MS: 25,  // small yield to avoid quotas
    MAX_TEXT_LEN_IN_HEADING: 120
  },

  // Chapter templates (config-driven, editable)
  CHAPTER_TEMPLATES: [
    { n: 0,  title: 'Innledning' },
    { n: 1,  title: 'Mål og suksesskriterier' },
    { n: 2,  title: 'Roller, grupper, tilgang (RBAC)' },
    { n: 3,  title: 'Navigasjon og menyer' },
    { n: 4,  title: 'Datamodell' },
    { n: 5,  title: 'Kjerneflyter og Arbeidsprosesser' },
    { n: 6,  title: 'Skjemaer' },
    { n: 7,  title: 'Styremodul (oversikt)' },
    { n: 8,  title: 'Økonomimodul (Integrasjon)' },
    { n: 9,  title: 'Kommunikasjon' },
    { n: 10, title: 'Dokumentarkiv' },
    { n: 11, title: 'Søk' },
    { n: 12, title: 'Validering & feilmeldinger' },
    { n: 13, title: 'Logging, tidsstempler og robusthet' },
    { n: 14, title: 'Tilgjengelighet og UX' },
    { n: 15, title: 'Sikkerhet / rettigheter' },
    { n: 16, title: 'Teknisk – Lagring og Integrasjoner' },
    { n: 17, title: 'Rapporter og Dashboards' },
    { n: 18, title: 'Administrativt krav' },
    { n: 19, title: 'Akseptansekriterier (Testscenarioer)' },
    { n: 20, title: 'MVP og faser' },
    { n: 21, title: 'Detaljerte UI-krav (Mikrointeraksjoner)' },
    { n: 23, title: 'Import og Eksport' },
    { n: 28, title: 'Brukeradopsjon, Opplæring og Endringsledelse' },
    { n: 29, title: 'Forvaltning og Datakvalitet (GDPR)' },
    { n: 30, title: 'Fremtidig potensial og videreutvikling' },
    { n: 31, title: 'Teknisk veikart og skalerbarhet' },
    { n: 32, title: 'Standard for oppsett av skjemaer' },
    { n: 33, title: 'Kvalitetssikring og Release' },
    { n: 34, title: 'Overvåkning og Stabilitet' },
    { n: 35, title: 'Testmiljø og "Sandkasse"' },
    { n: 36, title: 'Sentralisert Varslingssystem' },
    { n: 37, title: 'Brukerlivssyklus (On/Offboarding)' },
    { n: 38, title: 'Sikkerhetskopi og Gjenoppretting' },
    { n: 39, title: 'Dokumentstandard' }
  ]
};

/* ------------------------------ Utilities ---------------------------------- */

function _rsp_log_(level, fn, msg, data) {
  try {
    console.log('['+level+']', fn, msg, data||{});
  } catch(_) {}
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(RSP_CFG.SHEET_SYNCLOG_NAME) || ss.insertSheet(RSP_CFG.SHEET_SYNCLOG_NAME);
    if (sh.getLastRow() === 0) sh.appendRow(['Tid','Level','Funksjon','Melding','Detaljer']);
    sh.appendRow([new Date(), level, fn, msg, JSON.stringify(data||{})]);
  } catch(_) {}
}

function _rsp_effectiveDocId_() {
  try {
    var prop = PropertiesService.getScriptProperties().getProperty('RSP_DOC_ID');
    if (prop && String(prop).trim()) return prop;
  } catch(_) {}
  if (RSP_CFG.DOC_ID && RSP_CFG.DOC_ID !== 'SETT-DIN-DOC-ID-HER') return RSP_CFG.DOC_ID;
  return null;
}

function _rsp_safeSetProperty_(key, value, retries) {
  retries = retries || 3;
  for (var i=0;i<retries;i++) {
    try {
      PropertiesService.getScriptProperties().setProperty(key, value);
      return true;
    } catch(e) {
      if (i === retries - 1) throw e;
      Utilities.sleep(500 * (i + 1));
    }
  }
  return false;
}

function _safeString_(v, d) {
  return (v !== null && v !== undefined) ? String(v).trim() : (d || '');
}

function _hash_(s) {
  // simple, deterministic (not crypto) hash
  s = String(s || '');
  var h = 0, i = 0, len = s.length;
  while (i < len) { h = (h << 5) - h + s.charCodeAt(i++) | 0; }
  return String(h >>> 0);
}

function _rsp_getSheetData_(name, maxRows) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) return null;
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return { headers: [], data: [] };
  var limit = Math.min(lastRow, Number(RSP_CFG.VALIDATION.MAX_VALIDATION_ROWS) || 15000);
  if (maxRows) limit = Math.min(limit, maxRows);
  var vals = sh.getRange(1,1,limit,lastCol).getValues();
  return { headers: vals[0] || [], data: (vals.length > 1) ? vals.slice(1) : [] };
}

function _HeaderResolver_(headers) {
  if (!Array.isArray(headers)) throw new Error('Headers må være en array');
  this._headers = Object.freeze(headers.map(function(h){return String(h||'').trim().toLowerCase();}));
}
_HeaderResolver_.prototype.find = function(aliases) {
  var idx = -1;
  var low = this._headers;
  (aliases||[]).some(function(a){
    var pos = low.indexOf(String(a||'').trim().toLowerCase());
    if (pos >= 0) { idx = pos; return true; }
    return false;
  });
  return idx;
};

function _rsp_acquireGlobalLock_(ms) {
  var lock = LockService.getScriptLock();
  lock.waitLock(Math.max(1, ms || 30000));
  return lock;
}

function _rsp_releaseLockSafe_(lock) {
  try { if (lock) lock.releaseLock(); } catch(_) {}
}

function _rsp_toast_(msg, title, secs) {
  try { SpreadsheetApp.getActive().toast(msg, title || 'Status', secs || 3); } catch(_) {}
}

/* --------------------------- Validation Models ----------------------------- */

var VALIDATION_CATEGORIES = {
  STRUCTURE: 'Struktur/Konfig',
  PLACEHOLDERS: 'Plassholder/Dekning',
  HEADERS: 'Header-validering',
  DOC: 'Dokument-struktur'
};
var SEVERITY_LEVELS = { ERROR: 'ERROR', WARN: 'WARN', INFO: 'INFO' };

function ValidationIssue(category, severity, message, context) {
  this.category = category;
  this.severity = severity;
  this.message = message;
  this.context = context || {};
  this.timestamp = new Date();
}
ValidationIssue.prototype.toString = function(){ return this.message; };

/* ------------------------------- Backups ----------------------------------- */

function _rsp_fullBackupToSpreadsheet_() {
  var ss = SpreadsheetApp.getActive();
  var name = ss.getName() + ' [RSP Backup ' + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss') + ']';
  var file = DriveApp.getFileById(ss.getId());
  var copy = file.makeCopy(name);
  _rsp_log_('INFO','_rsp_fullBackupToSpreadsheet_','Backup created', { newFileId: copy.getId(), name: name });
  return copy.getId();
}

function _rsp_backupSheetChunked_(sourceSheet, targetSheet, chunkSize) {
  var lastRow = sourceSheet.getLastRow();
  var lastCol = sourceSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;
  chunkSize = chunkSize || 1000;
  for (var start = 1; start <= lastRow; start += chunkSize) {
    var end = Math.min(start + chunkSize - 1, lastRow);
    var rows = end - start + 1;
    var chunk = sourceSheet.getRange(start, 1, rows, lastCol).getValues();
    targetSheet.getRange(start, 1, rows, lastCol).setValues(chunk);
    Utilities.sleep(100);
  }
}

/* ---------------------------- Document Helpers ----------------------------- */

function _rsp_validateDocAccess_(docId) {
  try {
    var doc = DocumentApp.openById(docId);
    var name = doc.getName();
    var editors = doc.getEditors();
    var lastMod = DriveApp.getFileById(docId).getLastUpdated();
    doc.saveAndClose();
    return { accessible:true, name:name, editors:editors.map(function(u){return u.getEmail();}), lastModified:lastMod };
  } catch(e) {
    return { accessible:false, error:e.message };
  }
}

function _rsp_findHeading_(body, text) {
  for (var i=0;i<body.getNumChildren();i++) {
    var el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var p = el.asParagraph();
      var h = p.getHeading();
      if (h === DocumentApp.ParagraphHeading.HEADING1 || h === DocumentApp.ParagraphHeading.HEADING2) {
        if (String(p.getText()).trim() === String(text).trim()) return p;
      }
    }
  }
  return null;
}

function _rsp_detectExistingBaseline_(body) {
  // Method 1: exact marker
  for (var i=0;i<RSP_CFG.VALIDATION.DOC_MARKERS.length;i++) {
    var m = RSP_CFG.VALIDATION.DOC_MARKERS[i];
    if (body.findText(m)) return true;
  }
  // Method 2: H1 headings expected
  var expectedH1 = ['Kapittel 0: Innledning','Innholdsfortegnelse','Endringslogg'];
  var found = [];
  for (var j=0;j<body.getNumChildren();j++) {
    var c = body.getChild(j);
    if (c.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var para = c.asParagraph();
      if (para.getHeading() === DocumentApp.ParagraphHeading.HEADING1) {
        found.push(String(para.getText()).trim());
      }
    }
  }
  var all = expectedH1.every(function(t){ return found.indexOf(t) >= 0; });
  return all;
}

function _rsp_seedBaselineV93_(doc, body) {
  // Title
  body.appendParagraph('Sameieportalen – Kravdokument v9.3 (Gjeldende)').setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph('[[RSP_V93_BASELINE_MARKER]]').setHeading(DocumentApp.ParagraphHeading.NORMAL);

  // Intro + How auto-update works
  body.appendParagraph('Kapittel 0: Innledning').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Hvorfor en Sameieportal? (Motivasjon)').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(
    'Dette dokumentet er koblet til Google Sheets ved hjelp av RSP (Requirements Synchronization Platform). ' +
    'Når krav oppdateres i arket, kan de pushes hit (Sheet→Doc). Versjonshistorikk lagres i "Version_Log"-arket. ' +
    'Førstegangsoppsettet opprettet dokumentstruktur, og fremtidige endringer skjer kontrollert med validering og låsing.'
  );

  // TOC
  body.appendParagraph('Innholdsfortegnelse').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Oppdater manuelt via Sett inn → Innholdsfortegnelse, eller bruk Doc-funksjoner.');

  // Chapters (config-driven)
  RSP_CFG.CHAPTER_TEMPLATES.forEach(function(ch) {
    var title = 'Kapittel ' + ch.n + ': ' + ch.title;
    body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    // Optional default placeholder
    if (ch.n === 1) {
      body.appendParagraph('Krav K1-01: Sporbarhet i kjerneprosesser').setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph('Dette er et hovedmål: alle prosesser er sporbare i én portal med hendelseslogg.');
    }
  });

  // Central "Krav" anchor for sync where we place H2 entries
  body.appendParagraph('Krav (K1–K39)').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Alle genererte/oppdaterte kravseksjoner plasseres under denne overskriften.');

  // Changelog
  body.appendParagraph('Endringslogg').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('Her kan viktige endringer beskrives. Automatisk versjonshistorikk ligger i arket "Version_Log".');
}

/* ------------------------------- Validation -------------------------------- */

function _rsp_validateConfiguration_() {
  var issues = [];
  // Required numeric configs
  ['MAX_CHAPTERS','MAX_VALIDATION_ROWS'].forEach(function(k){
    var v = RSP_CFG.VALIDATION[k];
    if (!(typeof v === 'number' && v > 0)) {
      issues.push(new ValidationIssue(VALIDATION_CATEGORIES.STRUCTURE, SEVERITY_LEVELS.ERROR,
        'Konfigurasjon '+k+' må være et positivt heltall, fikk: '+v));
    }
  });
  // Required sheets array
  if (!Array.isArray(RSP_CFG.VALIDATION.REQUIRED_SHEETS)) {
    issues.push(new ValidationIssue(VALIDATION_CATEGORIES.STRUCTURE, SEVERITY_LEVELS.ERROR,
      'REQUIRED_SHEETS må være en array'));
  }

  // DOC_ID format + access
  var docId = _rsp_effectiveDocId_();
  if (!docId) {
    issues.push(new ValidationIssue(VALIDATION_CATEGORIES.STRUCTURE, SEVERITY_LEVELS.ERROR,
      'DOC_ID ikke konfigurert. Bruk menyen "Sett/Endre DOC_ID".'));
  } else if (!/^[a-zA-Z0-9-_]{20,}$/.test(docId)) {
    issues.push(new ValidationIssue(VALIDATION_CATEGORIES.STRUCTURE, SEVERITY_LEVELS.WARN,
      'DOC_ID format ser mistenkelig ut. Sjekk at ID er korrekt.'));
  } else {
    var access = _rsp_validateDocAccess_(docId);
    if (!access.accessible) {
      issues.push(new ValidationIssue(VALIDATION_CATEGORIES.STRUCTURE, SEVERITY_LEVELS.ERROR,
        'Får ikke åpnet dokumentet: '+access.error));
    }
  }
  return issues;
}

function _rsp_validateSystemState_() {
  var issues = [];
  var ss = SpreadsheetApp.getActive();
  (RSP_CFG.VALIDATION.REQUIRED_SHEETS || []).forEach(function(name){
    if (!ss.getSheetByName(name)) {
      issues.push(new ValidationIssue(VALIDATION_CATEGORIES.STRUCTURE, SEVERITY_LEVELS.ERROR, 'Mangler ark: '+name));
    }
  });

  var data = _rsp_getSheetData_(RSP_CFG.SHEET_REQ_NAME);
  if (!data) {
    issues.push(new ValidationIssue(VALIDATION_CATEGORIES.STRUCTURE, SEVERITY_LEVELS.ERROR, 'Mangler Requirements-arket.'));
    return issues;
  }
  var resolver = new _HeaderResolver_(data.headers);
  var idxId = resolver.find(RSP_CFG.HEADERS.id);
  var idxKrav = resolver.find(RSP_CFG.HEADERS.krav);
  var idxPrio = resolver.find(RSP_CFG.HEADERS.prio);
  var idxKap = resolver.find(RSP_CFG.HEADERS.kap);
  var idxProg = resolver.find(RSP_CFG.HEADERS.prog);
  if (idxId < 0 || idxKrav < 0 || idxPrio < 0 || idxKap < 0) {
    issues.push(new ValidationIssue(VALIDATION_CATEGORIES.HEADERS, SEVERITY_LEVELS.ERROR,
      'Requirements-headers mangler (trenger id/krav/prioritet/kapittel).'));
  }

  // Chapter coverage check (best-effort)
  var seen = Object.create(null);
  data.data.forEach(function(row){
    var kap = _safeString_(row[idxKap], '');
    var n = parseInt(kap, 10);
    if (!isNaN(n) && n >= 0 && n <= RSP_CFG.VALIDATION.MAX_CHAPTERS) seen[n] = true;
  });
  for (var i=1;i<=RSP_CFG.VALIDATION.MAX_CHAPTERS;i++) {
    if (!seen[i]) {
      issues.push(new ValidationIssue(VALIDATION_CATEGORIES.PLACEHOLDERS, SEVERITY_LEVELS.WARN,
        'Ingen krav registrert i kapittel '+i+'.'));
    }
  }
  return issues;
}

function _rsp_validateDocumentStructure_(body) {
  var issues = [];
  // Title present?
  var hasTitle = false;
  for (var i=0;i<body.getNumChildren();i++) {
    var el = body.getChild(i);
    if (el.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var p = el.asParagraph();
      if (p.getHeading() === DocumentApp.ParagraphHeading.TITLE && _safeString_(p.getText())) {
        hasTitle = true; break;
      }
    }
  }
  if (!hasTitle) {
    issues.push(new ValidationIssue(VALIDATION_CATEGORIES.DOC, SEVERITY_LEVELS.WARN, 'Dokumentet mangler tittel.'));
  }

  // Baseline detection
  var ok = _rsp_detectExistingBaseline_(body);
  if (!ok) {
    issues.push(new ValidationIssue(VALIDATION_CATEGORIES.DOC, SEVERITY_LEVELS.WARN,
      'Baseline v9.3 ikke tydelig oppdaget (markør eller H1-struktur mangler).'));
  }
  return issues;
}

/* --------------------------------- Wizard ---------------------------------- */

function rsp_menu_firstRunWizard() {
  var ui = SpreadsheetApp.getUi();
  var lock;
  try {
    lock = _rsp_acquireGlobalLock_(30000);
  } catch(e) {
    ui.alert('En annen prosess kjører allerede. Prøv igjen om et øyeblikk.');
    return;
  }

  var ss = SpreadsheetApp.getActive();
  try {
    _rsp_toast_('Starter førstegangsoppsett...', 'Wizard');
    _rsp_trackWizardProgress_('init');

    // 1) Validate config (allows setting DOC_ID later)
    var cfgIssues = _rsp_validateConfiguration_();
    if (cfgIssues.length) {
      _rsp_log_('WARN','rsp_menu_firstRunWizard','Konfig validering', { count: cfgIssues.length });
    }

    // 2) Backup metadata (fast duplicate)
    _rsp_toast_('Tar sikkerhetskopi av regnearket...', 'Wizard');
    var backupId = _rsp_fullBackupToSpreadsheet_();
    _rsp_trackWizardProgress_('backup');

    // 3) Ensure sheets & headers
    _rsp_toast_('Sikrer ark og headere...', 'Wizard');
    _rsp_ensureSheetAndHeaders_();
    _rsp_trackWizardProgress_('sheets');

    // 4) Initialize document structure if needed
    var docId = _rsp_effectiveDocId_();
    if (!docId || docId === 'SETT-DIN-DOC-ID-HER') {
      var resp = ui.prompt('DOC_ID mangler. Lim inn dokument-ID (fra URL til Google Doc):', ui.ButtonSet.OK_CANCEL);
      if (resp.getSelectedButton() !== ui.Button.OK) throw new Error('Wizard avbrutt (DOC_ID ikke satt).');
      var input = _safeString_(resp.getResponseText(), '');
      if (!input) throw new Error('DOC_ID kan ikke være tom.');
      _rsp_safeSetProperty_('RSP_DOC_ID', input, 3);
      docId = input;
    }

    _rsp_toast_('Initierer dokument...', 'Wizard');
    var doc = DocumentApp.openById(docId);
    var body = doc.getBody();
    if (!_rsp_detectExistingBaseline_(body)) {
      _rsp_seedBaselineV93_(doc, body);
      _rsp_log_('INFO','rsp_menu_firstRunWizard','Seeded v9.3 baseline', {});
    }
    doc.saveAndClose();
    _rsp_trackWizardProgress_('doc');

    // 5) Validate structure
    _rsp_toast_('Validerer system...', 'Wizard');
    var sysIssues = _rsp_validateSystemState_();
    var docIssues = (function(){
      try {
        var d = DocumentApp.openById(docId);
        var b = d.getBody();
        var res = _rsp_validateDocumentStructure_(b);
        d.saveAndClose();
        return res;
      } catch(e) {
        return [new ValidationIssue(VALIDATION_CATEGORIES.DOC, SEVERITY_LEVELS.ERROR, 'Kunne ikke åpne dokument for validering: '+e.message)];
      }
    })();

    var all = cfgIssues.concat(sysIssues).concat(docIssues);
    if (all.length) {
      var msg = 'Førstegangsoppsett fullført med advarsler/feil:\n\n• ' + all.map(function(i){return i.message;}).join('\n• ');
      ui.alert(msg);
    } else {
      ui.alert('Førstegangsoppsett fullført uten avvik. System klart ✅');
    }

    _rsp_trackWizardProgress_('done');
  } catch(e) {
    ui.alert('Førstegangsoppsett feilet:\n\n' + (e && e.message) + '\n\nSjekk "Sync_Log" for detaljer.');
    _rsp_log_('ERROR','rsp_menu_firstRunWizard','Wizard error',{ error:e && e.message, stack:e && e.stack });
  } finally {
    _rsp_releaseLockSafe_(lock);
  }
}

function _rsp_trackWizardProgress_(step) {
  try {
    PropertiesService.getScriptProperties().setProperty('RSP_WIZARD_PROGRESS',
      JSON.stringify({ step: step, ts: new Date().toISOString(), version: RSP_CFG.VERSION }));
  } catch(_) {}
}

function _rsp_ensureSheetAndHeaders_() {
  var ss = SpreadsheetApp.getActive();
  var req = ss.getSheetByName(RSP_CFG.SHEET_REQ_NAME) || ss.insertSheet(RSP_CFG.SHEET_REQ_NAME);
  var ver = ss.getSheetByName(RSP_CFG.SHEET_VERLOG_NAME) || ss.insertSheet(RSP_CFG.SHEET_VERLOG_NAME);
  var log = ss.getSheetByName(RSP_CFG.SHEET_SYNCLOG_NAME) || ss.insertSheet(RSP_CFG.SHEET_SYNCLOG_NAME);

  if (req.getLastRow() === 0) {
    req.appendRow(['KravID','Krav','Prioritet','Fremdrift %','Kapittel','Versjon','Kommentar','SistEndret']);
    req.setFrozenRows(1);
  }
  if (ver.getLastRow() === 0) {
    ver.appendRow(['KravID','Versjon','Hash','Tid','Kommentar']);
    ver.setFrozenRows(1);
  }
  if (log.getLastRow() === 0) {
    log.appendRow(['Tid','Level','Funksjon','Melding','Detaljer']);
    log.setFrozenRows(1);
  }
}

/* ---------------------------------- Sync ----------------------------------- */
/**
 * Lean push: Sheet → Doc, chunked, with version log updates.
 * - Creates/updates H2 sections under "Krav (K1–K39)".
 * - Uses simple hash of key fields to skip unchanged rows.
 * - Increments version in Version_Log when content changes.
 */
function rsp_syncSheetToDoc(opts) {
  opts = opts || {};
  var dryRun = !!opts.dryRun;
  var onProgress = typeof opts.onProgress === 'function' ? opts.onProgress : function(){};

  var lock;
  try { lock = _rsp_acquireGlobalLock_(30000); } catch(e) {
    throw new Error('En annen sync kjører. Prøv igjen snart.');
  }

  try {
    var docId = _rsp_effectiveDocId_();
    if (!docId) throw new Error('DOC_ID ikke satt.');

    var data = _rsp_getSheetData_(RSP_CFG.SHEET_REQ_NAME);
    if (!data) throw new Error('Mangler Requirements-arket.');

    var resolver = new _HeaderResolver_(data.headers);
    var idx = {
      id: resolver.find(RSP_CFG.HEADERS.id),
      krav: resolver.find(RSP_CFG.HEADERS.krav),
      prio: resolver.find(RSP_CFG.HEADERS.prio),
      prog: resolver.find(RSP_CFG.HEADERS.prog),
      kap: resolver.find(RSP_CFG.HEADERS.kap),
      vers: resolver.find(RSP_CFG.HEADERS.vers),
      kom: resolver.find(RSP_CFG.HEADERS.kom)
    };
    if (idx.id<0 || idx.krav<0 || idx.prio<0 || idx.kap<0) {
      throw new Error('Headers mangler (trenger id/krav/prioritet/kapittel).');
    }

    // Load latest versions per KravID (recent-only)
    var vmap = _rsp_readLatestVersionMap_();

    var total = data.data.length;
    var chunkSize = Math.max(1, Number(RSP_CFG.SYNC.CHUNK_SIZE) || 40);
    var processed = 0;
    for (var start=0; start<total; start+=chunkSize) {
      var end = Math.min(start + chunkSize, total);
      var doc = DocumentApp.openById(docId);
      var body = doc.getBody();
      var anchor = _rsp_findHeading_(body, 'Krav (K1–K39)');
      if (!anchor) {
        // if missing, create it near end
        body.appendParagraph('Krav (K1–K39)').setHeading(DocumentApp.ParagraphHeading.HEADING1);
        anchor = _rsp_findHeading_(body, 'Krav (K1–K39)');
      }

      for (var r=start; r<end; r++) {
        var row = data.data[r];
        var kravId = _safeString_(row[idx.id], '');
        if (!kravId) { processed++; continue; }
        var kravTxt = _safeString_(row[idx.krav], '');
        var prio = _safeString_(row[idx.prio], '');
        var kap = _safeString_(row[idx.kap], '');
        var prog = (idx.prog >= 0) ? _safeString_(row[idx.prog], '') : '';
        var vers = (idx.vers >= 0) ? Number(row[idx.vers] || 0) : 0;
        var kom  = (idx.kom  >= 0) ? _safeString_(row[idx.kom], '') : '';

        var headingText = _buildKravHeading_(kravId, kravTxt);
        var hash = _hash_([kravId,kravTxt,prio,kap,prog,kom].join('|'));

        var last = vmap[kravId];
        if (last && last.hash === hash) {
          // unchanged → skip
        } else {
          if (!dryRun) {
            _renderOrUpdateKravSection_(body, anchor, headingText, { kravId:kravId, krav:kravTxt, prio:prio, kap:kap, prog:prog, kom:kom });
            _rsp_upsertVersionLog_(kravId, (last ? last.ver+1 : 1), hash, 'Oppdatert via Sheet→Doc');
          }
        }

        processed++;
        if (r % 10 === 0) {
          onProgress({ processed: processed, total: total, current: kravId });
        }
      }

      doc.saveAndClose();
      Utilities.sleep(RSP_CFG.SYNC.RATE_LIMIT_MS);
    }

    _rsp_log_('INFO','rsp_syncSheetToDoc','Sync fullført',{ total: total, dryRun: dryRun });
  } catch(e) {
    _rsp_log_('ERROR','rsp_syncSheetToDoc','Sync feilet',{ error:e && e.message, stack:e && e.stack });
    throw e;
  } finally {
    _rsp_releaseLockSafe_(lock);
  }
}

function _rsp_readLatestVersionMap_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSP_CFG.SHEET_VERLOG_NAME);
  var out = Object.create(null);
  if (!sh || sh.getLastRow() < 2) return out;
  var vals = sh.getDataRange().getValues();
  for (var r=vals.length-1; r>=1; r--) {
    var id = _safeString_(vals[r][0], '');
    if (!id) continue;
    if (!out[id]) {
      out[id] = { ver: Number(vals[r][1]||0) || 0, hash: _safeString_(vals[r][2], '') };
    }
  }
  return out;
}

function _rsp_upsertVersionLog_(kravId, ver, hash, kommentar) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSP_CFG.SHEET_VERLOG_NAME) || ss.insertSheet(RSP_CFG.SHEET_VERLOG_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(['KravID','Versjon','Hash','Tid','Kommentar']);
  sh.appendRow([kravId, ver, hash, new Date(), kommentar||'']);
}

function _buildKravHeading_(kravId, kravTxt) {
  var clean = (kravTxt || '').replace(/\s+/g, ' ').trim();
  if (clean.length > RSP_CFG.SYNC.MAX_TEXT_LEN_IN_HEADING) {
    clean = clean.slice(0, RSP_CFG.SYNC.MAX_TEXT_LEN_IN_HEADING - 1) + '…';
  }
  return 'Krav ' + kravId + ': ' + clean;
}

function _renderOrUpdateKravSection_(body, anchorH1, headingText, kv) {
  // Find existing H2 for this kravId (by startsWith "Krav {id}:")
  var target = null;
  var prefix = 'Krav ' + kv.kravId + ':';
  for (var i=0;i<body.getNumChildren();i++) {
    var el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var p = el.asParagraph();
    if (p.getHeading() === DocumentApp.ParagraphHeading.HEADING2) {
      var txt = _safeString_(p.getText(), '');
      if (txt.indexOf(prefix) === 0) { target = p; break; }
    }
  }
  if (!target) {
    // insert after anchorH1
    var insertIndex = body.getChildIndex(anchorH1) + 1;
    target = body.insertParagraph(insertIndex, headingText).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    body.insertParagraph(insertIndex+1, 'Prioritet: ' + (kv.prio || ''));
    body.insertParagraph(insertIndex+2, 'Kapittel: ' + (kv.kap || ''));
    body.insertParagraph(insertIndex+3, 'Fremdrift: ' + (kv.prog || ''));
    if (kv.kom) body.insertParagraph(insertIndex+4, 'Kommentar: ' + kv.kom);
  } else {
    // update existing section: rewrite heading and the next few lines
    target.setText(headingText);
    var idx = body.getChildIndex(target);
    // Overwrite up to 4 lines of details (add if missing)
    var details = [
      'Prioritet: ' + (kv.prio || ''),
      'Kapittel: ' + (kv.kap || ''),
      'Fremdrift: ' + (kv.prog || ''),
    ];
    if (kv.kom) details.push('Kommentar: ' + kv.kom);

    for (var k=0;k<details.length;k++) {
      var pos = idx + 1 + k;
      if (pos < body.getNumChildren() && body.getChild(pos).getType() === DocumentApp.ElementType.PARAGRAPH) {
        body.getChild(pos).asParagraph().setText(details[k]);
      } else {
        body.insertParagraph(pos, details[k]);
      }
    }
  }
}

/**
 * Pull Doc → Sheet (scaffold).
 * Intentionally conservative: logs a message and returns to keep this single-file focused.
 * Extend here if you need full Doc→Sheet parsing.
 */
function rsp_syncDocToSheet(opts) {
  opts = opts || {};
  SpreadsheetApp.getUi().alert(
    'Doc → Sheet synk er ikke aktivert i denne kompakte versjonen.\n' +
    'Push (Sheet → Doc) er fullstendig implementert.'
  );
}

/* ---------------------------------- Menu ----------------------------------- */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var docId = _rsp_effectiveDocId_();
  var configured = docId ? '✅' : '❌';
  var health = _rsp_getSystemHealthEmoji_();

  ui.createMenu('Krav Sync')
    .addItem('0) Førstegangsoppsett (wizard) '+configured, 'rsp_menu_firstRunWizard')
    .addSeparator()
    .addItem('1) Sett/Endre DOC_ID', 'rsp_menu_setDocId')
    .addItem('2) Valider system '+health, 'rsp_menu_validateSystemState')
    .addSeparator()
    .addItem('3) Push (Sheet → Doc)', 'rsp_menu_pushRun')
    .addItem('4) Pull (Doc → Sheet) [scaffold]', 'rsp_menu_pullRun')
    .addSeparator()
    .addItem('Åpne kravdokument', 'rsp_menu_openDoc')
    .addToUi();
}

function _rsp_getSystemHealthEmoji_() {
  try {
    var issues = _rsp_validateSystemState_();
    if (issues.length === 0) return '✅';
    if (issues.some(function(i){return /DOC_ID|dokument/i.test(i.message);} )) return '❌';
    return '⚠️';
  } catch(_) { return '⚠️'; }
}

function rsp_menu_setDocId() {
  var ui = SpreadsheetApp.getUi();
  var current = _rsp_effectiveDocId_() || '(ikke satt)';
  var resp = ui.prompt('Sett/Endre DOC_ID', 'Nåværende: ' + current + '\nLim inn dokument-ID (fra Google Doc URL):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var id = _safeString_(resp.getResponseText(), '');
  if (!id) { ui.alert('DOC_ID kan ikke være tom.'); return; }
  _rsp_safeSetProperty_('RSP_DOC_ID', id, 3);
  ui.alert('DOC_ID oppdatert.');
}

function rsp_menu_validateSystemState() {
  var ui = SpreadsheetApp.getUi();
  var steps = ['Konfigurasjon','Ark og headers','Dokument-struktur'];
  _rsp_toast_('Validerer: ' + steps[0], 'Validering');
  var cfg = _rsp_validateConfiguration_();
  _rsp_toast_('Validerer: ' + steps[1], 'Validering');
  var sys = _rsp_validateSystemState_();
  _rsp_toast_('Validerer: ' + steps[2], 'Validering');

  var docIssues = [];
  var docId = _rsp_effectiveDocId_();
  if (docId) {
    try {
      var doc = DocumentApp.openById(docId);
      docIssues = _rsp_validateDocumentStructure_(doc.getBody());
      doc.saveAndClose();
    } catch(e) {
      docIssues = [new ValidationIssue(VALIDATION_CATEGORIES.DOC, SEVERITY_LEVELS.ERROR, 'Kunne ikke åpne dokument: '+e.message)];
    }
  } else {
    docIssues = [new ValidationIssue(VALIDATION_CATEGORIES.STRUCTURE, SEVERITY_LEVELS.ERROR, 'DOC_ID ikke satt')];
  }

  var all = cfg.concat(sys).concat(docIssues);
  if (!all.length) {
    ui.alert('Validering OK ✅');
  } else {
    var msg = 'Validering fant ' + all.length + ' funn:\n\n• ' + all.map(function(i){return i.message;}).join('\n• ');
    ui.alert(msg);
  }
}

function rsp_menu_pushRun() {
  // pre-flight
  var issues = _rsp_validateSystemState_();
  if (issues.length > 0) {
    SpreadsheetApp.getUi().alert('System ikke klart:\n\n• ' + issues.map(function(i){return i.message;}).join('\n• ') + '\n\nKjør "Førstegangsoppsett" først.');
    return;
  }
  rsp_syncSheetToDoc({ dryRun: false, onProgress: function(p){ Logger.log(p); } });
}

function rsp_menu_pullRun() {
  rsp_syncDocToSheet({ dryRun: true });
}

function rsp_menu_openDoc() {
  var docId = _rsp_effectiveDocId_();
  if (!docId) {
    SpreadsheetApp.getUi().alert('DOC_ID ikke satt.');
    return;
  }
  var url = 'https://docs.google.com/document/d/' + docId + '/edit';
  SpreadsheetApp.getActive().toast('Åpner: ' + url, 'Kravdokument', 5);
  // Note: Apps Script cannot auto-open URLs from menu; show toast for user to click/open manually.
}
