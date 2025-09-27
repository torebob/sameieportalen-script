// =============================================================================
// HMS – Dashboard & Varsling (Sameieportalen)
// FILE: 61_HMS_Dashboard.gs
// VERSION: 1.0.0
// UPDATED: 2025-09-15
// REQUIRES: 60_HMS_Vedlikeholdsplan.gs (HMS_PLAN → TASKS), ark: TASKS, TILGANG, (valgfritt) BEBOERE
// ROLES: LEDER/KASSERER kan sende varsler / kalender-synk
// =============================================================================

var TASKS_SHEET = 'TASKS';
var ACCESS_SHEET = 'TILGANG';
var BEBOER_SHEET = 'BEBOERE'; // valgfritt: Kolonner anbefalt: Email,Navn,Bygg,Leil,TillatVarsling(Ja/Nei)
var EDIT_ROLES = { 'LEDER': true, 'KASSERER': true };

// ---------- Dashboard-data (for widget/tile) ----------
function getHMSDashboardData(daysAhead) {
  var ahead = Number(daysAhead || 30);
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(TASKS_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, total: 0, open: 0, dueSoon: 0, items: [] };

  var values = sh.getDataRange().getValues(); 
  var header = values.shift();
  var idx = _byName_(header);

  var now = _startOfDay_(new Date());
  var limit = new Date(now.getFullYear(), now.getMonth(), now.getDate() + ahead);

  var total = 0, open = 0, dueSoon = 0;
  var items = [];

  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var kategori = _s(r[idx.Kategori-1] || '');
    if (kategori !== 'HMS') continue;
    total++;

    var status = _s(r[idx.Status-1] || '');
    if (status === 'Åpen') open++;

    var frist = r[idx.Frist-1];
    if (frist instanceof Date && frist >= now && frist <= limit && status === 'Åpen') {
      dueSoon++;
      items.push({
        title: _s(r[idx.Tittel-1] || ''),
        due: Utilities.formatDate(frist, Session.getScriptTimeZone() || 'Europe/Oslo', 'yyyy-MM-dd'),
        building: r[idx.Byggnummer-1] || '',
        haste: _s(r[idx.Hasteprioritering-1] || (_derivePriorityFromCriticality_(_s(r[idx.Kritikalitet-1] || '')))),
        planId: _s(r[idx.PlanID-1] || ''),
        autoKey: _s(r[idx.AutoKey-1] || ''),
        mustNotify: _s(r[idx.BeboerVarsling-1] || 'Nei')
      });
    }
  }

  // sorter viktigst først
  var priRank = { 'Kritisk': 3, 'Høy': 2, 'Normal': 1, 'Lav': 0 };
  items.sort(function(a,b){
    var p = (priRank[b.haste]||0) - (priRank[a.haste]||0);
    if (p !== 0) return p;
    return a.due.localeCompare(b.due);
  });

  return { ok: true, total: total, open: open, dueSoon: dueSoon, items: items.slice(0, 20) };
}

// ---------- Beboervarsling (e-post) ----------
function hmsNotifyResidents(options) {
  _ensureCanEdit_();
  options = options || {};
  var daysAhead = Number(options.daysAhead || 7);
  var buildingFilter = options.byggnummer || null; // tall eller array
  var dryRun = options.dryRun === true;

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(TASKS_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, sent: 0, preview: [] };

  var values = sh.getDataRange().getValues(); 
  var header = values.shift();
  var idx = _byName_(header);

  var now = _startOfDay_(new Date());
  var limit = new Date(now.getFullYear(), now.getMonth(), now.getDate() + daysAhead);

  var perBuilding = {}; // bygg -> liste av tekster
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (_s(r[idx.Kategori-1] || '') !== 'HMS') continue;
    if (_s(r[idx.Status-1] || '') !== 'Åpen') continue;

    var frist = r[idx.Frist-1];
    if (!(frist instanceof Date) || frist < now || frist > limit) continue;

    var notify = _s(r[idx.BeboerVarsling-1] || 'Nei');
    if (/^nei$/i.test(notify)) continue; // skal ikke varsles

    var bygg = r[idx.Byggnummer-1] || 'FELLES';
    if (buildingFilter) {
      if (Array.isArray(buildingFilter) && buildingFilter.indexOf(Number(bygg)) < 0) continue;
      if (!Array.isArray(buildingFilter) && Number(bygg) !== Number(buildingFilter)) continue;
    }

    var title = _s(r[idx.Tittel-1] || '');
    var when = Utilities.formatDate(frist, Session.getScriptTimeZone() || 'Europe/Oslo', 'EEEE d. MMMM', 'nb_NO');
    var where = _s(r[idx.Lokasjon-1] || '');
    var info = _s(r[idx.Kommentar-1] || '');
    var line = '• ' + title + ' — ' + when + (where ? (' — ' + where) : '') + (info ? ('\n  ' + info) : '');
    (perBuilding[bygg] = perBuilding[bygg] || []).push(line);
  }

  var preview = [];
  var totalSent = 0;
  var sendToAll = _loadResidentsByBuilding_();

  Object.keys(perBuilding).forEach(function(bygg){
    var lines = perBuilding[bygg];
    var recipients = sendToAll[bygg] && sendToAll[bygg].length ? sendToAll[bygg] : sendToAll['*'] || [];
    if (!recipients.length) {
      preview.push({ bygg: bygg, to: [], body: lines.join('\n') });
      return;
    }
    var subject = 'Varsel: planlagt vedlikehold (' + (bygg === 'FELLES' ? 'felles' : ('Bygg ' + bygg)) + ')';
    var body = 'Hei,\n\nFølgende planlagte oppgaver gjennomføres de neste dagene:\n\n' + lines.join('\n') +
               '\n\nMer info på Sameieportalen (HMS) eller på oppslagstavla.\n\nVennlig hilsen\nStyret';
    preview.push({ bygg: bygg, to: recipients, body: body });

    if (!dryRun) {
      var chunks = _chunk_(recipients, 30); // hold e-postliste moderat per sending
      for (var c = 0; c < chunks.length; c++) {
        MailApp.sendEmail({
          bcc: chunks[c].join(','),
          subject: subject,
          body: body
        });
        totalSent++;
      }
    }
  });

  return { ok: true, sent: dryRun ? 0 : totalSent, preview: preview };
}

// ---------- Kalender-synk (valgfritt) ----------
function hmsSyncCalendar(daysAhead) {
  _ensureCanEdit_();
  var ahead = Number(daysAhead || 60);
  var calName = 'HMS Sameiet';
  var cal = _ensureCalendar_(calName);

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(TASKS_SHEET);
  if (!sh || sh.getLastRow() < 2) return { ok: true, events: 0 };

  var values = sh.getDataRange().getValues(); 
  var header = values.shift();
  var idx = _byName_(header);

  var now = _startOfDay_(new Date());
  var limit = new Date(now.getFullYear(), now.getMonth(), now.getDate() + ahead);

  var count = 0;
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (_s(r[idx.Kategori-1] || '') !== 'HMS') continue;
    if (_s(r[idx.Status-1] || '') !== 'Åpen') continue;
    var due = r[idx.Frist-1];
    if (!(due instanceof Date) || due < now || due > limit) continue;

    var title = _s(r[idx.Tittel-1] || 'HMS-oppgave');
    var desc = (_s(r[idx.System-1] || '') ? ('System: ' + r[idx.System-1] + '\n') : '') +
               (_s(r[idx.Komponent-1] || '') ? ('Komponent: ' + r[idx.Komponent-1] + '\n') : '') +
               (_s(r[idx.Lokasjon-1] || '') ? ('Lokasjon: ' + r[idx.Lokasjon-1] + '\n') : '') +
               (_s(r[idx.DokumentasjonURL-1] || '') ? ('Dok: ' + r[idx.DokumentasjonURL-1] + '\n') : '') +
               (_s(r[idx.SjekklisteURL-1] || '') ? ('Sjekkliste: ' + r[idx.SjekklisteURL-1] + '\n') : '') +
               (_s(r[idx.BeboerVarsling-1] || '') ? ('Beboervarsling: ' + r[idx.BeboerVarsling-1] + '\n') : '');

    // Unik tittel for å unngå duplikater samme dato
    cal.createAllDayEvent(title + ' [' + _s(r[idx.AutoKey-1] || '') + ']', due, { description: desc });
    count++;
  }
  return { ok: true, events: count };
}

// ---------- Hjelp: beboerregister, roller, utils ----------
function _loadResidentsByBuilding_() {
  var out = { '*': [] }; // '*' = fallback felles
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(BEBOER_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;

  var values = sh.getDataRange().getValues(); 
  var header = values.shift();
  var map = _byName_(header);

  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var ok = String(r[map['TillatVarsling(Ja/Nei)']-1] || 'Ja').toLowerCase() !== 'nei';
    if (!ok) continue;
    var email = _s(r[map.Email-1] || ''); if (!email) continue;
    var bygg = map.Bygg ? r[map.Bygg-1] : (map.Byggnummer ? r[map.Byggnummer-1] : '');
    bygg = bygg ? String(bygg).trim() : '*';
    (out[bygg] = out[bygg] || []).push(email);
  }
  return out;
}

function _ensureCanEdit_() {
  var role = _getRoleForEmail_(_getUserEmail_());
  if (!EDIT_ROLES[role]) throw new Error('Tilgang nektet: bare Leder/Kasserer kan utføre dette.');
}
function _getUserEmail_() {
  return String(Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').trim();
}
function _getRoleForEmail_(email) {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(ACCESS_SHEET);
    if (!sh || sh.getLastRow() < 2) return 'LESER';
    var values = sh.getDataRange().getValues(); values.shift();
    for (var i=0;i<values.length;i++){
      if (String(values[i][0]||'').trim().toLowerCase() === String(email||'').trim().toLowerCase()) {
        var r = String(values[i][1]||'LESER').toUpperCase().trim();
        return EDIT_ROLES[r] || ['STYRE','LESER'].indexOf(r)>=0 ? r : 'LESER';
      }
    }
    return 'LESER';
  } catch(e){ return 'LESER'; }
}

/*
 * MERK: _byName_() er fjernet fra denne filen for å unngå konflikter.
 * Funksjonen er nå definert sentralt i 60_HMS_Vedlikeholdsplan.js.
 */
function _s(v){ return String(v==null?'':v).trim(); }
function _startOfDay_(d){ return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function _chunk_(arr, n){ var out=[]; for (var i=0;i<arr.length;i+=n) out.push(arr.slice(i,i+n)); return out; }
/*
 * MERK: _derivePriorityFromCriticality_() er fjernet fra denne filen.
 * Funksjonen er nå definert sentralt i 60_HMS_Vedlikeholdsplan.js.
 */
function _ensureCalendar_(name) {
  var cals = CalendarApp.getCalendarsByName(name);
  return cals && cals.length ? cals[0] : CalendarApp.createCalendar(name);
}
