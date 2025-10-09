// =============================================================================
// TestData – BAD-regler (kollisjonsfri)
// FILE: 93_TestData_BadRegler.gs
// VERSION: 1.1.0
// UPDATED: 2025-09-15
// NOTE: Ingen globale const-er; alt er namespacet under globalThis.TESTDATA
// PURPOSE: Legge inn både gyldig og bevisst "dårlig" testdata for å trigge
//          valideringer uten å redeklarere SHEETS/HEADERS/etc.
// =============================================================================

// Namespace (idempotent)
(function (glob) {
  glob.TESTDATA = Object.assign(glob.TESTDATA || {}, {
    VERSION: '1.1.0',
    UPDATED: '2025-09-15'
  });
})(globalThis);

// ------------------------------- Public API ----------------------------------

// 1) Gyldig grunnsett: Personer, Seksjoner, Eierskap (ett aktivt)
function testdataSeedBasic() {
  _tdEnsureSheets_();
  var S = _tdSheets_();

  _tdUpsertById_(S.PERSONER, 'Person-ID', [
    { 'Person-ID':'TST-PER-001', 'Navn':'Test Eier 1', 'Epost':'eier1@example.com', 'Telefon':'+4711111111', 'Rolle':'Eier', 'Aktiv':'Aktiv', 'Opprettet-Av':_tdUser_(), 'Opprettet-Dato':new Date(), 'Sist-Endret':new Date() },
    { 'Person-ID':'TST-PER-002', 'Navn':'Test Eier 2', 'Epost':'eier2@example.com', 'Telefon':'+4722222222', 'Rolle':'Eier', 'Aktiv':'Aktiv', 'Opprettet-Av':_tdUser_(), 'Opprettet-Dato':new Date(), 'Sist-Endret':new Date() }
  ]);

  _tdUpsertById_(S.SEKSJONER, 'Seksjon-ID', [
    { 'Seksjon-ID':'TST-SX-101', 'Nummer':'101', 'Beskrivelse':'Testleilighet 101', 'Areal':70, 'Status':'Aktiv', 'Opprettet-Av':_tdUser_(), 'Opprettet-Dato':new Date(), 'Sist-Endret':new Date() },
    { 'Seksjon-ID':'TST-SX-102', 'Nummer':'102', 'Beskrivelse':'Testleilighet 102', 'Areal':65, 'Status':'Aktiv', 'Opprettet-Av':_tdUser_(), 'Opprettet-Dato':new Date(), 'Sist-Endret':new Date() }
  ]);

  _tdUpsertById_(S.EIERSKAP, 'Eierskap-ID', [
    { 'Eierskap-ID':'TST-EIE-101A', 'Seksjon-ID':'TST-SX-101', 'Person-ID':'TST-PER-001', 'Fra-Dato':_tdD_(2023,1,1), 'Til-Dato':'', 'Status':'Aktiv', 'Sist-Endret':new Date() }
  ]);

  return 'OK: testdata (grunnsett) lagt inn.';
}

// 2) Bevisst "dårlig" data: ugyldig e-post, foreldreløs sak, overlappende eierskap, ugyldig dato
function testdataInsertBadRules() {
  _tdEnsureSheets_();
  var S = _tdSheets_();

  // Ugyldig e-post
  _tdAppend_(S.PERSONER, { 'Person-ID':'TST-PER-BADMAIL', 'Navn':'Feilformat Epost', 'Epost':'not-an-email', 'Rolle':'Eier', 'Aktiv':'Aktiv', 'Opprettet-Av':_tdUser_(), 'Opprettet-Dato':new Date(), 'Sist-Endret':new Date() });

  // Foreldreløs møtesak (refererer til Møte-ID som ikke finnes)
  _tdEnsureHeader_(S.MOTER);
  _tdEnsureHeader_(S.MOTE_SAKER);
  _tdAppend_(S.MOTE_SAKER, { 'Sak-ID':'TST-SAK-ORPHAN', 'Møte-ID':'TST-MO-404', 'Tittel':'Foreldreløs sak', 'Bakgrunn':'Skal trigge validering', 'Status':'Planlagt', 'Opprettet-Av':_tdUser_(), 'Opprettet-Dato':new Date(), 'Sist-Endret':new Date() });

  // Overlappende eierskap for samme Seksjon-ID
  _tdAppend_(S.EIERSKAP, { 'Eierskap-ID':'TST-EIE-101B', 'Seksjon-ID':'TST-SX-101', 'Person-ID':'TST-PER-002', 'Fra-Dato':_tdD_(2024,1,1), 'Til-Dato':'', 'Status':'Aktiv', 'Sist-Endret':new Date() });

  // Ugyldig datorekkefølge (Til-Dato før Fra-Dato)
  _tdAppend_(S.EIERSKAP, { 'Eierskap-ID':'TST-EIE-BADDATE', 'Seksjon-ID':'TST-SX-102', 'Person-ID':'TST-PER-001', 'Fra-Dato':_tdD_(2024,5,10), 'Til-Dato':_tdD_(2024,5,1), 'Status':'Historisk', 'Sist-Endret':new Date() });

  return 'OK: Dårlige testregler lagt inn (ugyldig e-post, foreldreløs sak, overlapp/feil dato i eierskap).';
}

// 3) Budsjett – legg inn "dårlige" rader for å trigge validering/audit
function testdataInsertBudgetBad() {
  var B = globalThis.BUDGET || {};
  var sheetName = B.SHEET || 'BUDSJETT';
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  var header = ['År','Versjon','Konto','Navn','Kostnadssted','Prosjekt','MVA','Type','Måned','Beløp','Kommentar'];
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');

  // Dårlige eksempler (long form):
  var rows = [
    // Non-numerisk beløp
    [2025,'main','4000','Strøm felles','FD','', '25','Kostnad','Jan','1 234,5x','feil beløpformat'],
    // Manglende konto
    [2025,'main','',    'Renhold','FD','', '25','Kostnad','Feb','2500','mangler konto'],
    // Ugyldig MVA
    [2025,'main','6100','Heis service','FD','', '27','Kostnad','Mar','5000','ugyldig mva'],
    // Duplikat (samme År|Versjon|Konto|Måned som neste rad)
    [2025,'main','4000','Strøm felles','FD','', '25','Kostnad','Apr','3000','dupe A'],
    [2025,'main','4000','Strøm felles','FD','', '25','Kostnad','Apr','3000','dupe B']
  ];

  sh.getRange(sh.getLastRow()+1, 1, rows.length, header.length).setValues(rows);
  return 'OK: BUDSJETT – “dårlige” linjer lagt inn (' + rows.length + ' rader).';
}

// 4) Rydd opp testdata (sletter rader der ID starter med "TST-")
function testdataCleanup() {
  var S = _tdSheets_();
  ['PERSONER','SEKSJONER','EIERSKAP','MOTER','MOTE_SAKER'].forEach(function(key){
    _tdDeleteRowsStartingWithId_(S[key], ['Person-ID','Seksjon-ID','Eierskap-ID','Møte-ID','Sak-ID']);
  });

  // Rydd budsjett-testrader (kommentar inneholder 'dupe' eller 'feil')
  var B = globalThis.BUDGET || {};
  var budgetName = B.SHEET || 'BUDSJETT';
  var sh = SpreadsheetApp.getActive().getSheetByName(budgetName);
  if (sh && sh.getLastRow() > 1) {
    var values = sh.getDataRange().getValues();
    var header = values.shift();
    var cComment = header.indexOf('Kommentar');
    var cNavn = header.indexOf('Navn');
    var toDelete = [];
    for (var r=0;r<values.length;r++){
      var cm = String(values[r][cComment]||'').toLowerCase();
      var nm = String(values[r][cNavn]||'').toLowerCase();
      if (cm.indexOf('dupe')>=0 || cm.indexOf('feil')>=0 || nm.indexOf('strøm felles')>=0 || nm.indexOf('heis service')>=0 || nm.indexOf('renhold')>=0) {
        toDelete.push(r+2);
      }
    }
    for (var i=toDelete.length-1;i>=0;i--) sh.deleteRow(toDelete[i]);
  }

  return 'OK: testdata fjernet.';
}

// ------------------------------ Intern helpers -------------------------------

function _tdSheets_() {
  var S = globalThis.SHEETS || {};
  return {
    PERSONER: S.PERSONER || 'Personer',
    SEKSJONER: S.SEKSJONER || 'Seksjoner',
    EIERSKAP: S.EIERSKAP || 'Eierskap',
    MOTER: S.MOTER || 'Møter',
    MOTE_SAKER: S.MOTE_SAKER || 'Møtesaker'
  };
}

function _tdEnsureSheets_() {
  // Bruk eksisterende setup hvis finnes
  try { if (typeof setupWorkbook === 'function') { setupWorkbook(); return; } } catch(e) {}

  // Minimal fallback-setup
  var ss = SpreadsheetApp.getActive();
  var S = _tdSheets_();
  function mk(name, headers) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    sh.freezeRows(1);
  }
  mk(S.PERSONER, ['Person-ID','Navn','Epost','Telefon','Rolle','Aktiv','Opprettet-Av','Opprettet-Dato','Sist-Endret']);
  mk(S.SEKSJONER, ['Seksjon-ID','Nummer','Beskrivelse','Areal','Status','Opprettet-Av','Opprettet-Dato','Sist-Endret']);
  mk(S.EIERSKAP, ['Eierskap-ID','Seksjon-ID','Person-ID','Fra-Dato','Til-Dato','Status','Sist-Endret']);
  mk(S.MOTER, ['Møte-ID','Type','Dato','Tittel','Agenda-URL','Protokoll-URL','Status','Opprettet-Av','Opprettet-Dato','Sist-Endret']);
  mk(S.MOTE_SAKER, ['Sak-ID','Møte-ID','Tittel','Bakgrunn','Status','Opprettet-Av','Opprettet-Dato','Sist-Endret']);

  // Budsjettark (til audit)
  var B = globalThis.BUDGET || {};
  var budgetName = B.SHEET || 'BUDSJETT';
  var bh = ss.getSheetByName(budgetName) || ss.insertSheet(budgetName);
  if (bh.getLastRow() === 0) {
    var header = ['År','Versjon','Konto','Navn','Kostnadssted','Prosjekt','MVA','Type','Måned','Beløp','Kommentar'];
    bh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
    bh.freezeRows(1);
  }
}

function _tdGetHeaders_(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastColumn() < 1) return [];
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(function(h){ return String(h||''); });
}

function _tdUpsertById_(sheetName, idHeader, rows) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheetName); if (!sh) return;
  var headers = _tdGetHeaders_(sheetName);
  var idCol = headers.indexOf(idHeader);
  if (idCol < 0) throw new Error('Header mangler: ' + idHeader + ' i ' + sheetName);

  var data = sh.getDataRange().getValues();
  var map = {};
  for (var r=1;r<data.length;r++){ map[String(data[r][idCol])] = r+1; } // 1-basert
  rows.forEach(function(obj){
    var row = new Array(headers.length).fill('');
    for (var i=0;i<headers.length;i++) {
      var h=headers[i];
      if (Object.prototype.hasOwnProperty.call(obj, h)) row[i] = obj[h];
    }
    var id = String(obj[idHeader] || '');
    if (map[id]) {
      sh.getRange(map[id],1,1,headers.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  });
}

function _tdAppend_(sheetName, objByHeader) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName); if (!sh) return;
  var headers = _tdGetHeaders_(sheetName);
  var row = new Array(headers.length).fill('');
  for (var i=0;i<headers.length;i++) {
    var h=headers[i];
    if (Object.prototype.hasOwnProperty.call(objByHeader, h)) row[i] = objByHeader[h];
  }
  sh.appendRow(row);
}

function _tdEnsureHeader_(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return;
  if (sh.getLastRow() === 0) return;
}

function _tdDeleteRowsStartingWithId_(sheetName, idHeaderCandidates) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName); if (!sh) return;
  var headers = _tdGetHeaders_(sheetName);
  var idIdx = -1;
  for (var i=0;i<idHeaderCandidates.length;i++){
    var pos = headers.indexOf(idHeaderCandidates[i]);
    if (pos >= 0) { idIdx = pos; break; }
  }
  if (idIdx < 0) return;

  var values = sh.getDataRange().getValues(); if (values.length < 2) return;
  var toDelete = [];
  for (var r=1;r<values.length;r++){
    var idVal = String(values[r][idIdx] || '');
    if (idVal.indexOf('TST-') === 0) toDelete.push(r+1);
  }
  for (var j=toDelete.length-1;j>=0;j--) sh.deleteRow(toDelete[j]);
}

function _tdD_(yyyy, mm, dd) { return new Date(yyyy, mm-1, dd); }
function _tdUser_(){ return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'system@example.com'; }
