/* FILE: 96_Repair_Tools.gs
 * VERSION: 1.1.0
 * UPDATED: 2025-09-15
 *
 * HVA DEN GJØR (sikkert og idempotent):
 *  - Oppretter kjerneark hvis de mangler: TILGANG, LEVERANDØRER, (valgfritt) Oppgaver
 *  - Slår sammen MøteStemmer → MøteSakStemmer
 *  - Migrerer Ark 9 → Styringsdokumenter_Logg (ryddig loggformat)
 *  - Migrerer Ark 10 → DIAG_CHECKS (standard sjekklogg)
 *  - Liten "quick fix": fryser header-rad på utvalgte ark, lager DIAG_PROJECT om mangler
 *  - Kaller projectOverview() hvis den finnes
 *
 * SIKKERHET:
 *  - Ingen globale const-er. Kun toppnivå-funksjoner med prefix repair96_.
 *  - Tåler at andre filer definerer ACCESS_SHEET, MONTHS, osv.
 * ========================================================================== */

/** Kjør ALT i ett (trygt å kjøre flere ganger) */
function repair96_RunAll() {
  var logs = [];
  logs.push(repair96_EnsureCoreSheets());
  logs.push(repair96_MergeVoteSheets());
  logs.push(repair96_MigrateArk9ToStyringsdokLogg());
  logs.push(repair96_MigrateArk10ToDiagChecks());
  logs.push(repair96_addCategoryValidation());
  logs.push(repair96_DiagQuickFix());
  try { if (typeof projectOverview === 'function') projectOverview(); } catch (e) {}
  return logs.join(' | ');
}

/** Legg til en liten meny nå (kan også kalles fra onOpen andre steder) */
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
    .addItem('Legg til kategorivalidering', 'repair96_addCategoryValidation')
    .addItem('Diag quick fix', 'repair96_DiagQuickFix')
    .addToUi();
  } catch (e) {
    Logger.log('repair96_ShowMenu: ' + (e && e.message || e));
  }
}

/** Opprett TILGANG & LEVERANDØRER (+ valgfritt Oppgaver hvis blankt) */
function repair96_EnsureCoreSheets() {
  var ss = SpreadsheetApp.getActive();
  var changes = [];

  // TILGANG
  (function () {
    var name = 'TILGANG';
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, 2).setValues([['Email', 'Rolle']]).setFontWeight('bold');
      try { sh.setFrozenRows(1); } catch (_) {}
      changes.push('Opprettet TILGANG (Email|Rolle)');
    }
  })();

  // LEVERANDØRER
  (function () {
    var name = 'LEVERANDØRER';
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      var headers = ['LeverandørID','Navn','Kontakt','E-post','Telefon','Fagområde','AvtaleNr','Rating','Notat','SistOppdatert'];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      try { sh.setFrozenRows(1); } catch (_) {}
      changes.push('Opprettet LEVERANDØRER');
    }
  })();

  // Oppgaver (bare hvis ikke finnes – eller finnes men er helt tomt)
  (function () {
    var name = 'Oppgaver';
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    if (sh.getLastRow() === 0) {
      var headers = ['OppgaveID','Tittel','Beskrivelse','Kategori','Prioritet','Opprettet','Frist','Status','Ansvarlig','Seksjonsnr','Kostnad','Lenke'];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      try { sh.setFrozenRows(1); } catch (_) {}
      changes.push('Initierte header på Oppgaver');
    }
  })();

  return changes.length ? changes.join(' / ') : 'Kjerneark OK';
}

/** Slå sammen MøteStemmer → MøteSakStemmer (beholder dest-header) */
function repair96_MergeVoteSheets() {
  var ss = SpreadsheetApp.getActive();
  var dst = ss.getSheetByName('MøteSakStemmer');
  var src = ss.getSheetByName('MøteStemmer');
  if (!dst || !src) return 'Ingen fusjon nødvendig';

  var srcRows = src.getLastRow(), srcCols = src.getLastColumn();
  if (srcRows > 1) {
    var header = dst.getRange(1, 1, 1, Math.max(1, dst.getLastColumn())).getValues()[0];
    var cols = Math.min(srcCols, header.length);
    var vals = src.getRange(2, 1, srcRows - 1, cols).getValues();
    if (vals.length) dst.getRange(dst.getLastRow() + 1, 1, vals.length, cols).setValues(vals);
  }
  ss.deleteSheet(src);
  return 'Fusjonerte MøteStemmer → MøteSakStemmer';
}

/** Migrer Ark 9 → Styringsdokumenter_Logg (ryddig loggformat) */
function repair96_MigrateArk9ToStyringsdokLogg() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName('Ark 9');
  if (!src) return 'Ark 9: ingenting å migrere';

  var rows = src.getLastRow(), cols = src.getLastColumn();
  var dstName = 'Styringsdokumenter_Logg';
  var header = ['Dato','Kategori','Mål','Status','Kommentar','AnsvarligRolle','SisteVersjon','SistEndret','MasterURL','PDF_URL'];

  if (rows < 2) {
    src.setName(dstName);
    try { src.setFrozenRows(1); } catch (_) {}
    if (src.getLastRow() === 0) {
      src.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
      try { src.setFrozenRows(1); } catch(_) {}
    }
    return 'Ark 9 var tom – døpte om til ' + dstName;
  }

  var dst = ss.getSheetByName(dstName);
  if (!dst) {
    dst = ss.insertSheet(dstName);
    dst.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
    try { dst.setFrozenRows(1); } catch (_) {}
  }

  var srcHeader = src.getRange(1,1,1,cols).getValues()[0].map(String);
  function idx(name){ return srcHeader.indexOf(name) + 1; }
  var c = {
    dokument_id: idx('dokument_id'),
    navn: idx('navn'),
    master: idx('masterdokument_url'),
    pdf: idx('gjeldende_pdf_url'),
    ansvarlig: idx('ansvarlig_rolle'),
    versjon: idx('siste_versjon'),
    endret: idx('sist_endret')
  };

  if (!c.navn || !c.master || !c.pdf || !c.ansvarlig) {
    src.setName(dstName);
    try { src.setFrozenRows(1); } catch (_) {}
    return 'Ark 9 hadde uventet struktur – ga nytt navn til ' + dstName + ' (ingen migrasjon)';
  }

  var vals = src.getRange(2, 1, rows - 1, cols).getValues();
  var out = vals.map(function (r) {
    var datoRaw = c.dokument_id ? r[c.dokument_id - 1] : '';
    var dato = repair96_parseDateNor_(datoRaw);
    return [
      dato,
      String(r[c.navn - 1] || ''),
                     String(r[c.master - 1] || ''),
                     String(r[c.pdf - 1] || ''),
                     String(r[c.ansvarlig - 1] || ''),
                     String(r[c.ansvarlig - 1] || ''),
                     String(c.versjon ? r[c.versjon - 1] : ''),
                     String(c.endret ? r[c.endret - 1] : ''),
                     String(r[c.master - 1] || ''),
                     String(r[c.pdf - 1] || '')
    ];
  });

  if (out.length) dst.getRange(dst.getLastRow() + 1, 1, out.length, header.length).setValues(out);
  ss.deleteSheet(src);
  return 'Migrerte ' + out.length + ' rader fra Ark 9 → ' + dstName;
}

/** Migrer Ark 10 → DIAG_CHECKS (tid/kilde/test/OK/WARN/FAIL/melding) */
function repair96_MigrateArk10ToDiagChecks() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName('Ark 10');
  if (!src) return 'Ark 10: ingenting å migrere';

  var rows = src.getLastRow(), cols = Math.max(1, src.getLastColumn());
  var dstName = 'DIAG_CHECKS';
  var header = ['Tid','Kilde','Test','OK','WARN','FAIL','Melding'];

  var dst = ss.getSheetByName(dstName);
  if (!dst) {
    dst = ss.insertSheet(dstName);
    dst.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
    try { dst.setFrozenRows(1); } catch (_) {}
  } else if (dst.getLastRow() === 0) {
    dst.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
    try { dst.setFrozenRows(1); } catch (_) {}
  }

  if (rows < 1) {
    ss.deleteSheet(src);
    return 'Ark 10 var tom – slettet';
  }

  var vals = src.getRange(1, 1, rows, cols).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var rawTs = String(r[0] || '').trim();
    var source = String(r[1] || '').trim() || 'Checks';
    var message = String(r[2] || '').trim();
    if (!rawTs && !source && !message) continue;

    var ts = repair96_parseNorDateTime_(rawTs) || new Date();
    var test = (message.split(':', 2)[0] || '').trim();
    var ok   = repair96_intFrom_(message, /OK\s*=\s*(\d+)/i);
    var warn = repair96_intFrom_(message, /WARN\s*=\s*(\d+)/i);
    var fail = repair96_intFrom_(message, /FAIL\s*=\s*(\d+)/i);

    out.push([ts, source, test, ok, warn, fail, message]);
  }

  if (out.length) dst.getRange(dst.getLastRow() + 1, 1, out.length, header.length).setValues(out);
  ss.deleteSheet(src);
  return 'Migrerte ' + out.length + ' rader fra Ark 10 → ' + dstName;
}

/** Liten "quick fix": fryser headere, sørger for DIAG_PROJECT, prøver projectOverview() */
function repair96_DiagQuickFix() {
  var ss = SpreadsheetApp.getActive();
  var fixed = [];

  (function () {
    var name = 'DIAG_PROJECT';
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      fixed.push('Opprettet DIAG_PROJECT');
    }
  })();

  var candidates = ['TILGANG','LEVERANDØRER','Oppgaver','HMS_PLAN','MøteSakStemmer','Styringsdokumenter_Logg','DIAG_CHECKS','DIAG_PROJECT','BEBOERE','Seksjoner','Eierskap','Møter'];
  candidates.forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (!sh) return;
    try { sh.setFrozenRows(1); fixed.push('Frosset header på ' + n); } catch (_) {}
  });

  try { if (typeof projectOverview === 'function') { projectOverview(); fixed.push('Oppdatert DIAG_PROJECT'); } } catch (_) {}

  return fixed.length ? fixed.join(' / ') : 'Quick fix: ingenting å endre';
}

/** Legg til datavalidering for "Kategori" i Styringsdokumenter_Logg */
function repair96_addCategoryValidation() {
  var ss = SpreadsheetApp.getActive();
  var sheetName = 'Styringsdokumenter_Logg';
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return 'Ark ' + sheetName + ' finnes ikke. Hoppet over validering.';
  }

  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var categoryIndex = header.indexOf('Kategori');

  if (categoryIndex === -1) {
    return 'Fant ikke kolonnen "Kategori" i ' + sheetName + '.';
  }

  var categories = ["Vedtekter", "Protokoller", "Regnskap", "Kontrakter", "Tegninger", "Annet"];
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(categories).build();

  var range = sheet.getRange(2, categoryIndex + 1, sheet.getMaxRows() - 1);
  range.setDataValidation(rule);

  return 'La til kategorivalidering i ' + sheetName + '.';
}

/* ============================== Små helpers =============================== */

function repair96_parseDateNor_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
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

/** Append én sjekk-linje til DIAG_CHECKS (sikker og idempotent på header) */
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

/** En enkel "helsesjekk" som logger til DIAG_CHECKS */
function runAllChecks() {
  var ss = SpreadsheetApp.getActive();
  var ok=0, warn=0, fail=0, notes=[];

  [['HMS_PLAN', true], ['Oppgaver', true]].forEach(function(spec){
    var exists = !!ss.getSheetByName(spec[0]);
    if (exists) { ok++; notes.push('OK '+spec[0]); }
    else { fail++; notes.push('Mangler '+spec[0]); }
  });

  ['TILGANG','LEVERANDØRER'].forEach(function(name){
    var exists = !!ss.getSheetByName(name);
    if (exists) { ok++; notes.push('OK '+name); }
    else { warn++; notes.push('Mangler (anbefalt) '+name); }
  });

  if (ss.getSheetByName('BEBOERE')) { ok++; notes.push('OK BEBOERE'); }

  var msg = 'runAllChecks: OK='+ok+', WARN='+warn+', FAIL='+fail+' | '+notes.join('; ');
  diagChecks_Log('Checks', 'runAllChecks', ok, warn, fail, msg);
  return msg;
}

/** (Valgfritt) Daglig trigger kl 07:30 som kjører runAllChecks og logger til DIAG_CHECKS */
function installDailyDiagChecksTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'runAllChecks') ScriptApp.deleteTrigger(t);
  });
    ScriptApp.newTrigger('runAllChecks').timeBased().atHour(7).nearMinute(30).everyDays(1).create();
}
