// =============================================================================
// HMS – Vedlikeholdsplan (plan → TASKS)  [UTVIDET]
// FILE: 60_HMS_Vedlikeholdsplan.gs
// VERSION: 1.1.0
// UPDATED: 2025-09-15
// CHANGES:
//  - Utvidet HMS_PLAN: SesongAvhengig, LeverandørKontakt, SistUtført, HistoriskKost,
//    Byggnummer, Garantistatus
//  - Utvidet TASKS: Hasteprioritering, Værforhold, BeboerVarsling
//  - Sesonglogikk + “takrenne om vinter” eksempel + byggfilter + beboervarelse hooks
//  - markTaskCompleted(): oppdaterer TASKS + HMS_PLAN.SistUtført + HistoriskKost (glidende snitt 3 år)
// =============================================================================

var HMS_PLAN_SHEET = 'HMS_PLAN';
var TASKS_SHEET    = 'TASKS';
var SUPPLIERS_SHEET = 'LEVERANDØRER'; // (valgfritt) leverandørdatabase

// --- Nye/utvidede headere ---
var PLAN_HEADER = [
  'PlanID','System','Komponent','Oppgave','Beskrivelse','Frekvens',
  'PreferertMåned','NesteStart','AnsvarligRolle','Leverandør','LeverandørKontakt',
  'Myndighetskrav','Standard/Referanse','Kritikalitet(1-5)',
  'EstTidTimer','EstKost','HistoriskKost','BudsjettKonto',
  'DokumentasjonURL','SjekklisteURL','Lokasjon','Byggnummer','Garantistatus',
  'SesongAvhengig','SistUtført','Kommentar','Aktiv'
];

var TASKS_HEADER = [
  'Tittel','Kategori','Status','Frist','Opprettet','Ansvarlig',
  'Seksjonsnr','PlanID','AutoKey','System','Komponent','Lokasjon','Byggnummer',
  'Myndighetskrav','Kritikalitet','Hasteprioritering',
  'EstKost','BudsjettKonto','FaktiskKost',
  'DokumentasjonURL','SjekklisteURL','Garantistatus','BeboerVarsling','Værforhold','Leverandør','LeverandørKontakt',
  'Kommentar','OppdatertAv','Oppdatert'
];

// =============================================================================
// Skjema/migrering
// =============================================================================

function hmsMigrateSchema_v1_1() {
  var ss = SpreadsheetApp.getActive();

  // HMS_PLAN
  var plan = ss.getSheetByName(HMS_PLAN_SHEET) || ss.insertSheet(HMS_PLAN_SHEET);
  if (plan.getLastRow() === 0) {
    plan.getRange(1,1,1,PLAN_HEADER.length).setValues([PLAN_HEADER]).setFontWeight('bold');
    plan.setFrozenRows(1);
  } else {
    _ensureColumns_(plan, PLAN_HEADER);
  }

  // TASKS
  var tasks = ss.getSheetByName(TASKS_SHEET) || ss.insertSheet(TASKS_SHEET);
  if (tasks.getLastRow() === 0) {
    tasks.getRange(1,1,1,TASKS_HEADER.length).setValues([TASKS_HEADER]).setFontWeight('bold');
    tasks.setFrozenRows(1);
  } else {
    _ensureColumns_(tasks, TASKS_HEADER);
  }

  // Leverandører (valgfritt)
  var sup = ss.getSheetByName(SUPPLIERS_SHEET);
  if (!sup) {
    sup = ss.insertSheet(SUPPLIERS_SHEET);
    sup.getRange(1,1,1,6).setValues([['Kategori/System','Komponent','Navn','Telefon','Epost','Notat']]).setFontWeight('bold');
    sup.setFrozenRows(1);
  }

  return 'HMS v1.1 migrering ok';
}

function _ensureColumns_(sheet, desiredHeader) {
  var rng = sheet.getRange(1,1,1, sheet.getLastColumn() || desiredHeader.length);
  var cur = rng.getValues()[0];
  var map = {}; for (var i=0;i<cur.length;i++) map[cur[i]] = i+1;
  // append manglende til slutt
  var missing = desiredHeader.filter(function(h){ return !map[h]; });
  if (missing.length) {
    var start = (cur.filter(String).length || 0) + 1;
    sheet.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
  // rekkefølge lar vi være – vi jobber med indeksoppslag på navn
}

function hmsEnsurePlanSheet() {
  return hmsMigrateSchema_v1_1();
}

// =============================================================================
// Generator (HMS_PLAN → TASKS)
// =============================================================================

function hmsGenerateTasks(options) {
  options = options || {};
  var monthsAhead    = Number(options.monthsAhead || 12);
  var startDate      = options.startDate ? new Date(options.startDate) : new Date();
  var kategori       = options.kategori || 'HMS';
  var statusDefault  = options.status || 'Åpen';
  var buildingFilter = options.byggnummer || null; // f.eks. 1..6 eller array
  var replaceExisting = options.replaceExisting !== false; // default true

  var ss = SpreadsheetApp.getActive();
  var plan = ss.getSheetByName(HMS_PLAN_SHEET); if (!plan) return { ok:false, error:'Mangler ark: ' + HMS_PLAN_SHEET };
  var tasks = ss.getSheetByName(TASKS_SHEET) || ss.insertSheet(TASKS_SHEET);
  if (tasks.getLastRow() === 0) tasks.getRange(1,1,1,TASKS_HEADER.length).setValues([TASKS_HEADER]).setFontWeight('bold');

  var planValues = plan.getDataRange().getValues(); if (planValues.length < 2) return { ok:false, error:'Planen er tom.' };
  var planHeader = planValues.shift(); var pidx = _byName_(planHeader);

  var tHeader = tasks.getRange(1,1,1,tasks.getLastColumn()).getValues()[0];
  var tidx = _byName_(tHeader);
  var existing = tasks.getLastRow() > 1 ? tasks.getRange(2,1,tasks.getLastRow()-1,tasks.getLastColumn()).getValues() : [];

  // Bygg dupe-sett
  var autoKeySet = {};
  if (existing.length && tidx.AutoKey) {
    var akCol = tidx.AutoKey-1;
    for (var i=0;i<existing.length;i++) {
      var ak = String(existing[i][akCol]||'').trim();
      if (ak) autoKeySet[ak] = true;
    }
  }

  // Valgfri ryddejobb i tidsvinduet
  if (replaceExisting && existing.length && tidx.PlanID && tidx.Frist) {
    var endDate = new Date(startDate); endDate.setMonth(endDate.getMonth()+monthsAhead);
    var delRows = [];
    var frCol = tidx.Frist-1, pidCol = tidx.PlanID-1;

    for (var r=0;r<existing.length;r++) {
      var d = existing[r][frCol], pid = String(existing[r][pidCol]||'').trim();
      if (d instanceof Date && d >= startDate && d <= endDate && pid) delRows.push(r+2);
    }
    delRows.sort(function(a,b){ return b-a; });
    for (var j=0;j<delRows.length;j++) tasks.deleteRow(delRows[j]);

    existing = tasks.getLastRow() > 1 ? tasks.getRange(2,1,tasks.getLastRow()-1,tasks.getLastColumn()).getValues() : [];
    autoKeySet = {};
    if (existing.length && tidx.AutoKey) {
      var akCol2 = tidx.AutoKey-1;
      for (var k=0;k<existing.length;k++) {
        var ak2 = String(existing[k][akCol2]||'').trim();
        if (ak2) autoKeySet[ak2] = true;
      }
    }
  }

  var out = [];
  var created = 0;
  var end = new Date(startDate); end.setMonth(end.getMonth()+monthsAhead);

  for (var r=0;r<planValues.length;r++) {
    var row = planValues[r];

    // Aktiv?
    var aktiv = _str(row[pidx.Aktiv-1] || 'Ja').toLowerCase();
    if (aktiv === 'nei' || aktiv === '0' || aktiv === 'false') continue;

    var planId = _str(row[pidx.PlanID-1]);
    if (!planId) continue;

    // Byggfilter?
    var byggnr = pidx.Byggnummer ? row[pidx.Byggnummer-1] : '';
    if (buildingFilter) {
      if (Array.isArray(buildingFilter)) {
        if (buildingFilter.indexOf(Number(byggnr)) < 0) continue;
      } else {
        if (Number(byggnr) !== Number(buildingFilter)) continue;
      }
    }

    var system  = pidx.System ? _str(row[pidx.System-1]) : '';
    var komponent = pidx.Komponent ? _str(row[pidx.Komponent-1]) : '';
    var oppgave  = pidx.Oppgave ? _str(row[pidx.Oppgave-1]) : '';
    var beskrivelse = pidx.Beskrivelse ? _str(row[pidx.Beskrivelse-1]) : '';
    var frek    = _str(row[pidx.Frekvens-1]);
    var pref    = pidx.PreferertMåned ? _str(row[pidx.PreferertMåned-1]) : (pidx['PreferertM\u00e5ned'] ? _str(row[pidx['PreferertM\u00e5ned']-1]) : '');
    var nextStart = pidx.NesteStart ? _parseDateSafe_(row[pidx.NesteStart-1]) : null;
    var ansvarlig = pidx.AnsvarligRolle ? _str(row[pidx.AnsvarligRolle-1]) : '';
    var lokasjon = pidx.Lokasjon ? _str(row[pidx.Lokasjon-1]) : '';
    var mynd = pidx.Myndighetskrav ? _str(row[pidx.Myndighetskrav-1]) : '';
    var krit = pidx['Kritikalitet(1-5)'] ? row[pidx['Kritikalitet(1-5)']-1] : '';
    var estKost = pidx.EstKost ? _num(row[pidx.EstKost-1]) : '';
    var konto = pidx.BudsjettKonto ? _str(row[pidx.BudsjettKonto-1]) : '';
    var dok = pidx.DokumentasjonURL ? _str(row[pidx.DokumentasjonURL-1]) : '';
    var sjekkl = pidx.SjekklisteURL ? _str(row[pidx.SjekklisteURL-1]) : '';
    var garanti = pidx.Garantistatus ? _str(row[pidx.Garantistatus-1]) : '';
    var sesongAvh = pidx.SesongAvhengig ? _str(row[pidx.SesongAvhengig-1]).toLowerCase()==='ja' : false;
    var lever = pidx.Leverandør ? _str(row[pidx.Leverandør-1]) : '';
    var leverKontakt = pidx.LeverandørKontakt ? _str(row[pidx.LeverandørKontakt-1]) : '';

    // Hasteprioritering: default ut fra kritikalitet (kan overstyres i plan ved å legge egen kolonne senere)
    var haste = _derivePriorityFromCriticality_(krit);

    // Finn forekomster
    var anchor = nextStart || startDate;
    var prefMonths = _parsePreferredMonths_(pref);
    var occurrences = _expandOccurrences_(anchor, end, frek, prefMonths);

    for (var i=0;i<occurrences.length;i++) {
      var due = new Date(occurrences[i]);

      // Sameie-spesifikke regler:
      // - Takrenne/renner i vintermåned → flytt til april
      if (_containsIgnoreCase_(oppgave, 'takrenne') && _isWinterMonth_(due)) {
        due = _moveToMonth_(due, 4); // april
      }
      // - SesongAvhengig utearbeid: vintermåned → flytt til april (kan utvides)
      if (sesongAvh && _isWinterMonth_(due)) {
        due = _moveToMonth_(due, 4);
      }

      var dueStr = Utilities.formatDate(due, Session.getScriptTimeZone() || 'Europe/Oslo','yyyy-MM-dd');
      var autoKey = planId + '::' + dueStr + (byggnr ? ('::' + byggnr) : '');
      if (autoKeySet[autoKey]) continue; // unngå duplikater
      autoKeySet[autoKey] = true;

      var title = (system ? system+': ' : '') + (komponent ? komponent+' – ' : '') + (oppgave || 'Oppgave');
      var opprettet = new Date();

      var rowOut = _emptyRow_(tHeader.length);
      _set(rowOut, tidx, 'Tittel', title);
      _set(rowOut, tidx, 'Kategori', kategori);
      _set(rowOut, tidx, 'Status', statusDefault);
      _set(rowOut, tidx, 'Frist', due);
      _set(rowOut, tidx, 'Opprettet', opprettet);
      _set(rowOut, tidx, 'Ansvarlig', ansvarlig);
      _set(rowOut, tidx, 'Seksjonsnr', ''); // felles
      _set(rowOut, tidx, 'PlanID', planId);
      _set(rowOut, tidx, 'AutoKey', autoKey);
      _set(rowOut, tidx, 'System', system);
      _set(rowOut, tidx, 'Komponent', komponent);
      _set(rowOut, tidx, 'Lokasjon', lokasjon);
      _set(rowOut, tidx, 'Byggnummer', byggnr);
      _set(rowOut, tidx, 'Myndighetskrav', mynd);
      _set(rowOut, tidx, 'Kritikalitet', krit);
      _set(rowOut, tidx, 'Hasteprioritering', haste);
      _set(rowOut, tidx, 'EstKost', estKost);
      _set(rowOut, tidx, 'BudsjettKonto', konto);
      _set(rowOut, tidx, 'FaktiskKost', ''); // settes ved ferdigstillelse
      _set(rowOut, tidx, 'DokumentasjonURL', dok);
      _set(rowOut, tidx, 'SjekklisteURL', sjekkl);
      _set(rowOut, tidx, 'Garantistatus', garanti);
      // BeboerVarsling: initieres fra planlogikk (ute/inn, myndighetskrav, bygg), her “Behov?” som standard
      _set(rowOut, tidx, 'BeboerVarsling', _suggestResidentNotice_(lokasjon, oppgave, mynd, byggnr));
      // Værforhold: kreves for utendørs – la feltet stå tomt ("TBD") så utfører fyller inn
      _set(rowOut, tidx, 'Værforhold', (lokasjon && /ute|utendørs/i.test(lokasjon)) ? '' : 'N/A');
      _set(rowOut, tidx, 'Leverandør', lever);
      _set(rowOut, tidx, 'LeverandørKontakt', leverKontakt || _lookupSupplierContact_(system, komponent));
      _set(rowOut, tidx, 'Kommentar', beskrivelse);

      out.push(rowOut);
      created++;
    }
  }

  if (out.length) tasks.getRange(tasks.getLastRow()+1, 1, out.length, tHeader.length).setValues(out);
  return { ok:true, created: created };
}

// =============================================================================
// Ferdigstilling: marker utført + oppdater Plan.SistUtført + HistoriskKost
// =============================================================================

/**
 * Marker oppgave som utført basert på AutoKey ELLER radnummer i TASKS.
 * @param {Object} options { autoKey?: string, row?: number, faktiskKost?: number }
 */
function markTaskCompleted(options) {
  options = options || {};
  var ss = SpreadsheetApp.getActive();
  var tasks = ss.getSheetByName(TASKS_SHEET);
  if (!tasks) return { ok:false, error:'Mangler ark: ' + TASKS_SHEET };

  var tHeader = tasks.getRange(1,1,1,tasks.getLastColumn()).getValues()[0];
  var tidx = _byName_(tHeader);

  var rowIndex = options.row || null;
  var autoKey = options.autoKey || null;
  if (!rowIndex && !autoKey) return { ok:false, error:'Oppgi row eller autoKey' };

  if (!rowIndex) {
    var values = tasks.getDataRange().getValues(); values.shift();
    var akCol = tidx.AutoKey-1;
    for (var i=0;i<values.length;i++) {
      if (String(values[i][akCol]||'').trim() === autoKey) { rowIndex = i+2; break; }
    }
    if (!rowIndex) return { ok:false, error:'Fant ikke oppgave med AutoKey' };
  }

  var now = new Date();
  if (tidx.Status) tasks.getRange(rowIndex, tidx.Status, 1, 1).setValue('Utført');
  if (tidx.OppdatertAv) tasks.getRange(rowIndex, tidx.OppdatertAv, 1, 1).setValue(Session.getEffectiveUser().getEmail());
  if (tidx.Oppdatert) tasks.getRange(rowIndex, tidx.Oppdatert, 1, 1).setValue(now);
  if (tidx.FaktiskKost && options.faktiskKost != null) tasks.getRange(rowIndex, tidx.FaktiskKost, 1, 1).setValue(Number(options.faktiskKost)||0);

  // Oppdater HMS_PLAN.SistUtført + HistoriskKost (glidende snitt 3 år)
  var planId = tidx.PlanID ? tasks.getRange(rowIndex, tidx.PlanID, 1, 1).getValue() : '';
  if (planId) _updatePlanAfterCompletion_(String(planId), now, Number(options.faktiskKost)||null);

  // Hooks: beboer-tavle, Sites/Docs osv. kan trigges her
  // _postToNoticeBoard_(rowIndex); // implementér hvis ønskelig

  return { ok:true, row: rowIndex };
}

function _updatePlanAfterCompletion_(planId, dateDone, actualCost) {
  var ss = SpreadsheetApp.getActive();
  var plan = ss.getSheetByName(HMS_PLAN_SHEET);
  if (!plan) return;
  var values = plan.getDataRange().getValues(); var header = values.shift();
  var pidx = _byName_(header);
  var pidCol = pidx.PlanID-1, sistCol = pidx.SistUtført-1, histCol = pidx.HistoriskKost ? (pidx.HistoriskKost-1) : null, estCol = pidx.EstKost ? (pidx.EstKost-1) : null;

  for (var i=0;i<values.length;i++) {
    if (String(values[i][pidCol]||'').trim() === planId) {
      // SistUtført
      plan.getRange(i+2, sistCol+1, 1, 1).setValue(dateDone);
      // HistoriskKost: glidende snitt (3 års verdi) – bruker Faktisk om tilgjengelig, ellers Est
      if (histCol != null) {
        var prev = Number(values[i][histCol] || 0);
        var base = (actualCost != null && !isNaN(actualCost)) ? actualCost
                 : (estCol != null ? Number(values[i][estCol]||0) : 0);
        var updated = prev ? Math.round(((prev*2 + base)/3)*100)/100 : base;
        plan.getRange(i+2, histCol+1, 1, 1).setValue(updated);
      }
      break;
    }
  }
}

// =============================================================================
// Leverandør-database (enkelt oppslag)
// =============================================================================
function _lookupSupplierContact_(system, komponent) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SUPPLIERS_SHEET);
  if (!sh) return '';
  var values = sh.getDataRange().getValues(); values.shift();
  for (var i=0;i<values.length;i++) {
    var cat = _str(values[i][0]);
    var comp = _str(values[i][1]);
    if ((cat && _equalsIgnoreCase_(cat, system)) && (!comp || _equalsIgnoreCase_(comp, komponent))) {
      var navn = _str(values[i][2]), tlf = _str(values[i][3]), ep = _str(values[i][4]);
      var txt = navn; if (tlf) txt += ' ' + tlf; if (ep) txt += ' ' + ep;
      return txt;
    }
  }
  return '';
}

// =============================================================================
// Varsling (enkel): Beboervarsling-anbefaling og e-posthooks
// =============================================================================
function _suggestResidentNotice_(lokasjon, oppgave, myndighetskrav, byggnr) {
  var out = 'Nei';
  // typiske oppgaver som påvirker beboere
  if (/heis|vann|sprinkler|brannalarm|garasjeport/i.test(oppgave)) out = 'Ja';
  if (lokasjon && /ute|utendørs/i.test(lokasjon)) out = 'Vurder';
  if (_str(myndighetskrav).toLowerCase()==='ja') out = 'Vurder';
  if (byggnr) out += ' (Bygg ' + byggnr + ')';
  return out;
}

function hmsNotifyUpcomingTasks(daysAhead) {
  var ss = SpreadsheetApp.getActive();
  var tasks = ss.getSheetByName(TASKS_SHEET); if (!tasks) return { ok:true, notified:0 };
  var values = tasks.getDataRange().getValues(); var header = values.shift();
  var idx = _byName_(header);

  var today = new Date();
  var ahead = Number(daysAhead || 14);
  var limit = new Date(today.getFullYear(), today.getMonth(), today.getDate() + ahead);

  var res = [];
  for (var i=0;i<values.length;i++) {
    var r = values[i];
    var status = _str(r[idx.Status-1]||'');
    var due = r[idx.Frist-1];
    if (status !== 'Åpen') continue;
    if (!(due instanceof Date)) continue;
    if (due >= today && due <= limit) res.push(r);
  }

  if (!res.length) return { ok:true, notified:0 };

  var to = PropertiesService.getScriptProperties().getProperty('HMS_NOTIFY_EMAIL') ||
           Session.getEffectiveUser().getEmail();
  var lines = res.slice(0,50).map(function(r){
    var t = r[idx.Tittel-1], d = Utilities.formatDate(r[idx.Frist-1], Session.getScriptTimeZone()||'Europe/Oslo', 'yyyy-MM-dd');
    var b = r[idx.Byggnummer-1]||'', h = r[idx.Hasteprioritering-1]||'';
    return '• ['+(h||'Normal')+'] ' + t + (b?(' (Bygg '+b+')'):'') + ' – frist ' + d;
  });
  MailApp.sendEmail({ to: to, subject: 'HMS: kommende oppgaver ('+res.length+')', body: lines.join('\n') });
  return { ok:true, notified: res.length, to: to };
}

// =============================================================================
// Hjelpere
// =============================================================================

function _byName_(header) {
  var map = {};
  for (var i=0;i<header.length;i++) {
    var h = String(header[i]||'').trim();
    if (h) map[h] = i+1;
  }
  return map;
}
function _emptyRow_(n){ var a=[]; for (var i=0;i<n;i++) a.push(''); return a; }
function _set(arr, idx, name, val){ if (idx[name]) arr[idx[name]-1] = val; }
function _str(v){ return String(v==null?'':v).trim(); }
function _equalsIgnoreCase_(a,b){ return String(a||'').trim().toLowerCase() === String(b||'').trim().toLowerCase(); }
function _containsIgnoreCase_(s, needle){ return String(s||'').toLowerCase().indexOf(String(needle||'').toLowerCase()) >= 0; }

function _parseDateSafe_(v) {
  if (v instanceof Date) return v;
  var s = _str(v); if (!s) return null;
  var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) { var d = new Date(Number(m[1]), Number(m[2])-1, Number(m[3])); if (!isNaN(d.getTime())) return d; }
  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

function _parsePreferredMonths_(s) {
  s = _str(s); if (!s) return [];
  var tokens = s.split(/[;,\/\s]+/).filter(function(x){return x;});
  var out = [];
  var map = { jan:1,januar:1,feb:2,februar:2,mar:3,mars:3,apr:4,april:4,mai:5,jun:6,juni:6,jul:7,juli:7,aug:8,august:8,sep:9,september:9,okt:10,oktober:10,nov:11,november:11,des:12,desember:12,
              'vår':-1,'var':-1,'sommer':-2,'høst':-3,'host':-3,'vinter':-4 };
  for (var i=0;i<tokens.length;i++){
    var t = tokens[i].toLowerCase();
    var m = map[t];
    if (m) {
      if (m>0) out.push(m);
      else {
        if (m===-1) out = out.concat([3,4,5]);
        if (m===-2) out = out.concat([6,7,8]);
        if (m===-3) out = out.concat([9,10,11]);
        if (m===-4) out = out.concat([12,1,2]);
      }
    } else {
      var n = Number(t); if (Number.isInteger(n) && n>=1 && n<=12) out.push(n);
    }
  }
  // uniq + sort
  var uniq = {}, res = [];
  for (var j=0;j<out.length;j++){ var mm=out[j]; if (!uniq[mm]){ uniq[mm]=true; res.push(mm);} }
  res.sort(function(a,b){return a-b;}); return res;
}

function _expandOccurrences_(anchorDate, endDate, freq, preferredMonths) {
  var occ = [];
  var a = new Date(anchorDate);
  var e = new Date(endDate);
  var f = _str(freq).toUpperCase().replace('Å','A').replace('Ø','O').replace('Æ','AE');

  function makeDate(y,m,d){ return new Date(y, m-1, d||15); }

  var interval = 0;
  if (f.indexOf('MND')>=0 || f.indexOf('MANED')>=0 || f==='MÅNEDLIG' || f==='MANEDLIG') interval = 1;
  else if (f.indexOf('KVART')>=0) interval = 3;
  else if (f.indexOf('HALV')>=0) interval = 6;
  else if (f.indexOf('2AAR')>=0 || f==='2ÅR') interval = 24;
  else if (f.indexOf('3AAR')>=0 || f==='3ÅR') interval = 36;
  else if (f.indexOf('5AAR')>=0 || f==='5ÅR') interval = 60;
  else if (f.indexOf('10AAR')>=0 || f==='10ÅR') interval = 120;
  else if (f.indexOf('AAR')>=0 || f==='ÅRLIG' || f==='AARLIG') interval = 12;
  else interval = 12;

  if (preferredMonths && preferredMonths.length) {
    var sy=a.getFullYear(), sm=a.getMonth()+1; var ey=e.getFullYear(), em=e.getMonth()+1;
    for (var y=sy; y<=ey; y++){
      for (var m=1; m<=12; m++){
        if ((y===sy && m<sm) || (y===ey && m>em)) continue;
        if (preferredMonths.indexOf(m)>=0) occ.push(makeDate(y,m,a.getDate()||15));
      }
    }
  } else {
    var d = new Date(a);
    while (d <= e) { occ.push(new Date(d)); d.setMonth(d.getMonth()+interval); }
  }
  return occ.filter(function(dt){ return dt >= a && dt <= e; });
}

function _isWinterMonth_(date) {
  var m = (date.getMonth()+1);
  return (m===11 || m===12 || m<=3);
}
function _moveToMonth_(date, month1to12) {
  return new Date(date.getFullYear(), month1to12-1, 15);
}

function _derivePriorityFromCriticality_(krit) {
  var n = Number(krit||0);
  if (n >= 5) return 'Kritisk';
  if (n >= 4) return 'Høy';
  if (n >= 3) return 'Normal';
  return 'Lav';
}
function _num(v){
  if (v === '' || v == null) return '';
  var s = String(v).trim().replace(/\s/g,'');
  var hasC = s.indexOf(',')>=0, hasD = s.indexOf('.')>=0;
  if (hasC && !hasD) s = s.replace(/\./g,'').replace(',','.');
  else if (hasC && hasD && s.lastIndexOf(',')>s.lastIndexOf('.')) s = s.replace(/\./g,'').replace(',','.');
  var n = Number(s); return isNaN(n) ? '' : n;
}
