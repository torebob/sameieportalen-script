/* ============================================================================
 * Diagnose & Repair Toolkit (kollisjonsfri)
 * FILE: 96_Repair_Tools.gs
 * VERSION: 1.2.0
 * UPDATED: 2025-09-28
 *
 * HVA DEN GJØR (sikkert og idempotent):
 *  - Oppretter kjerneark, slår sammen ark, migrerer data, og kjører sjekker.
 *
 * SIKKERHET:
 *  - Alle offentlige funksjoner som utfører endringer er nå beskyttet av _requireAdmin_().
 *  - Funksjoner som kan kalles fra UI eller direkte via google.script.run er sikret.
 * ========================================================================== */

// ============================================================================
//  PUBLIC FACING FUNCTIONS - SECURED
// ============================================================================

/** Kjør ALT i ett (trygt å kjøre flere ganger). Krever admin. */
function repair96_RunAll() {
  _requireAdmin_();
  return repair96_RunAll_();
}

/** Legg til en liten meny nå. Selve menyen er ikke sikret, men funksjonene den kaller er det. */
function repair96_ShowMenu() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Diagnose/Repair 96')
      .addItem('Kjør ALT (RunAll)', 'repair96_RunAll')
      .addSeparator()
      .addItem('Opprett kjerneark', 'repair96_EnsureCoreSheets')
      .addItem('Fusjoner stemme-ark', 'repair96_MergeVoteSheets')
      .addItem('Migrer Ark 9 → Logg', 'repair96_MigrateArk9ToStyringsdokLogg')
      .addItem('Migrer Ark 10 → Checks', 'repair96_MigrateArk10ToDiagChecks')
      .addSeparator()
      .addItem('Diag quick fix', 'repair96_DiagQuickFix')
      .addToUi();
  } catch (e) {
    Logger.log('repair96_ShowMenu: ' + (e && e.message || e));
  }
}

/** Opprett TILGANG & LEVERANDØRER. Krever admin. */
function repair96_EnsureCoreSheets() {
  _requireAdmin_();
  return repair96_EnsureCoreSheets_();
}

/** Slå sammen MøteStemmer → MøteSakStemmer. Krever admin. */
function repair96_MergeVoteSheets() {
  _requireAdmin_();
  return repair96_MergeVoteSheets_();
}

/** Migrer Ark 9 → Styringsdokumenter_Logg. Krever admin. */
function repair96_MigrateArk9ToStyringsdokLogg() {
  _requireAdmin_();
  return repair96_MigrateArk9ToStyringsdokLogg_();
}

/** Migrer Ark 10 → DIAG_CHECKS. Krever admin. */
function repair96_MigrateArk10ToDiagChecks() {
  _requireAdmin_();
  return repair96_MigrateArk10ToDiagChecks_();
}

/** Liten “quick fix”. Krever admin. */
function repair96_DiagQuickFix() {
  _requireAdmin_();
  return repair96_DiagQuickFix_();
}

/** En enkel “helsesjekk” som kan kalles fra UI. Krever admin. */
function runAllChecks() {
  _requireAdmin_();
  return runAllChecks_();
}

/** Enkel "helsesjekk" for triggere. Ingen auth-sjekk, da triggeren må installeres av en admin. */
function runAllChecksTriggered() {
  return runAllChecks_();
}

/** Installerer en daglig trigger. Krever admin. */
function installDailyDiagChecksTrigger() {
  _requireAdmin_();
  return installDailyDiagChecksTrigger_();
}


// ============================================================================
//  INTERNAL IMPLEMENTATION FUNCTIONS
// ============================================================================

function repair96_RunAll_() {
  var logs = [];
  logs.push(repair96_EnsureCoreSheets_());
  logs.push(repair96_MergeVoteSheets_());
  logs.push(repair96_MigrateArk9ToStyringsdokLogg_());
  logs.push(repair96_MigrateArk10ToDiagChecks_());
  logs.push(repair96_DiagQuickFix_());
  try { if (typeof projectOverview === 'function') projectOverview(); } catch (e) {}
  return logs.join(' | ');
}

function repair96_EnsureCoreSheets_() {
  var ss = SpreadsheetApp.getActive();
  var changes = [];
  (function () {
    var name = 'TILGANG';
    if (!ss.getSheetByName(name)) {
      var sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, 2).setValues([['Email', 'Rolle']]).setFontWeight('bold');
      try { sh.setFrozenRows(1); } catch (_) {}
      changes.push('Opprettet TILGANG');
    }
  })();
  (function () {
    var name = 'LEVERANDØRER';
    if (!ss.getSheetByName(name)) {
      var sh = ss.insertSheet(name);
      var headers = ['LeverandørID','Navn','Kontakt','E-post','Telefon','Fagområde','AvtaleNr','Rating','Notat','SistOppdatert'];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      try { sh.setFrozenRows(1); } catch (_) {}
      changes.push('Opprettet LEVERANDØRER');
    }
  })();
  return changes.length ? changes.join(' / ') : 'Kjerneark OK';
}

function repair96_MergeVoteSheets_() {
  var ss = SpreadsheetApp.getActive();
  var dst = ss.getSheetByName('MøteSakStemmer');
  var src = ss.getSheetByName('MøteStemmer');
  if (!dst || !src) return 'Ingen fusjon nødvendig';
  var srcRows = src.getLastRow();
  if (srcRows > 1) {
    var vals = src.getRange(2, 1, srcRows - 1, src.getLastColumn()).getValues();
    if (vals.length) dst.getRange(dst.getLastRow() + 1, 1, vals.length, vals[0].length).setValues(vals);
  }
  ss.deleteSheet(src);
  return 'Fusjonerte MøteStemmer → MøteSakStemmer';
}

function repair96_MigrateArk9ToStyringsdokLogg_() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName('Ark 9');
  if (!src) return 'Ark 9: ingenting å migrere';
  var dstName = 'Styringsdokumenter_Logg';
  src.setName(dstName);
  return 'Omdøpt Ark 9 → ' + dstName;
}

function repair96_MigrateArk10ToDiagChecks_() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName('Ark 10');
  if (!src) return 'Ark 10: ingenting å migrere';
  var dstName = 'DIAG_CHECKS';
  var header = ['Tid','Kilde','Test','OK','WARN','FAIL','Melding'];
  var dst = ss.getSheetByName(dstName) || ss.insertSheet(dstName);
  if (dst.getLastRow() === 0) {
    dst.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
    try { dst.setFrozenRows(1); } catch(_) {}
  }
  var rows = src.getLastRow();
  if (rows > 0) {
    var vals = src.getRange(1, 1, rows, src.getLastColumn()).getValues();
    dst.getRange(dst.getLastRow() + 1, 1, vals.length, vals[0].length).setValues(vals);
  }
  ss.deleteSheet(src);
  return 'Migrerte ' + rows + ' rader fra Ark 10 → ' + dstName;
}

function repair96_DiagQuickFix_() {
  var ss = SpreadsheetApp.getActive();
  var fixed = [];
  if (!ss.getSheetByName('DIAG_PROJECT')) {
    ss.insertSheet('DIAG_PROJECT');
    fixed.push('Opprettet DIAG_PROJECT');
  }
  var candidates = ['TILGANG','LEVERANDØRER','Oppgaver','HMS_PLAN','MøteSakStemmer','Styringsdokumenter_Logg','DIAG_CHECKS','DIAG_PROJECT','BEBOERE','Seksjoner','Eierskap','Møter'];
  candidates.forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh) { try { sh.setFrozenRows(1); } catch (_) {} }
  });
  fixed.push('Frosset headere');
  try { if (typeof projectOverview === 'function') { projectOverview(); fixed.push('Oppdatert DIAG_PROJECT'); } } catch (_) {}
  return fixed.length ? fixed.join(' / ') : 'Quick fix: ingenting å endre';
}

function runAllChecks_() {
  var ss = SpreadsheetApp.getActive();
  var ok=0, warn=0, fail=0, notes=[];
  [['HMS_PLAN', true], ['Oppgaver', true]].forEach(function(spec){
    (ss.getSheetByName(spec[0]) ? ok++ : fail++);
    notes.push((ss.getSheetByName(spec[0]) ? 'OK ' : 'Mangler ') + spec[0]);
  });
  ['TILGANG','LEVERANDØRER'].forEach(function(name){
    (ss.getSheetByName(name) ? ok++ : warn++);
    notes.push((ss.getSheetByName(name) ? 'OK ' : 'Mangler (anbefalt) ') + name);
  });
  if (ss.getSheetByName('BEBOERE')) { ok++; notes.push('OK BEBOERE'); }
  var msg = 'runAllChecks: OK='+ok+', WARN='+warn+', FAIL='+fail+' | '+notes.join('; ');
  diagChecks_Log('Checks', 'runAllChecks', ok, warn, fail, msg);
  return msg;
}

function installDailyDiagChecksTrigger_() {
  const triggerHandler = 'runAllChecksTriggered';
  // Rydder gamle triggere for denne funksjonen først
  ScriptApp.getProjectTriggers().forEach(function(t){
    const handler = t.getHandlerFunction();
    if (handler === triggerHandler || handler === 'runAllChecks') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger(triggerHandler).timeBased().atHour(7).nearMinute(30).everyDays(1).create();
  return 'Installert daglig helsesjekk-trigger.';
}

/* ============================== Små helpers (uendret) =============================== */

function repair96_parseDateNor_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/); // dd.MM.yyyy
  if (m) {
    var d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return isNaN(d) ? s : d;
  }
  var d2 = new Date(s);
  return isNaN(d2) ? s : d2;
}

function repair96_parseNorDateTime_(s) {
  if (s instanceof Date && !isNaN(s)) return s;
  var str = String(s || '').trim();
  var rx = /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s*(?:kl\.?|@)?\s*)?(\d{1,2})[.:](\d{2})(?:[.:](\d{2}))?$/i;
  var m = str.match(rx);
  if (m) {
    var d = Number(m[1]), mo = Number(m[2]) - 1, y = Number(m[3]);
    var H = Number(m[4]), M = Number(m[5]), S = m[6] ? Number(m[6]) : 0;
    var dt = new Date(y, mo, d, H, M, S);
    if (!isNaN(dt)) return dt;
  }
  var d2 = new Date(str);
  return isNaN(d2) ? null : d2;
}

function repair96_intFrom_(text, re) {
  var m = String(text || '').match(re);
  return m ? parseInt(m[1], 10) : 0;
}

function diagChecks_Log(source, test, ok, warn, fail, message) {
  var ss = SpreadsheetApp.getActive();
  var name = 'DIAG_CHECKS', sh = ss.getSheetByName(name);
  var header = ['Tid','Kilde','Test','OK','WARN','FAIL','Melding'];
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
    try { sh.setFrozenRows(1); } catch(_) {}
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
    try { sh.setFrozenRows(1); } catch(_) {}
  }
  var row = [new Date(), String(source||''), String(test||''), Number(ok||0), Number(warn||0), Number(fail||0), String(message||'')];
  sh.getRange(sh.getLastRow()+1, 1, 1, header.length).setValues([row]);
  sh.getRange(2,1,Math.max(1, sh.getLastRow()-1),1).setNumberFormat('dd.MM.yyyy HH:mm:ss');
}