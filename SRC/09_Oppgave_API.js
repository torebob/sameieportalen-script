/* ====================== Oppgave-API for Vaktmester (stabil) ======================
 * FILE: 09_Oppgave_API.gs | VERSION: 1.2.0 | UPDATED: 2025-09-14
 * FORMÅL: Hente oppgaver (aktive/historikk), oppdatere status m/kommentar,
 *         legge til kommentar og opprette ny sak fra Vaktmester-UI.
 * =============================================================================== */

var VM_ACTIVE_STATUSES = ['ny','påbegynt','paabegynt','venter'];
var VM_CLOSED_STATUSES = ['fullført','fullfort','avvist'];

function _hdrMap_(headers){
  var map = {};
  headers.forEach(function(h, i){ map[String(h||'').trim()] = i+1; });
  return map;
}
function _getTasksSheet_(){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.TASKS);
  if (!sh) throw new Error("Fant ikke arket '" + SHEETS.TASKS + "'.");
  if (sh.getLastRow() === 0){
    sh.appendRow(['OppgaveID','Tittel','Beskrivelse','Kategori','Prioritet','Opprettet','Frist','Status','Ansvarlig','Seksjonsnr','Relatert','Kommentarer']);
  }
  return sh;
}

function _currentUserProfile_(){
  var email = (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
              (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) || '';
  email = String(email||'').toLowerCase();
  var name = '';
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.BOARD);
    if (sh && sh.getLastRow() > 1){
      var data = sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
      for (var i=0;i<data.length;i++){
        var e = String(data[i][1]||'').toLowerCase();
        if (e === email){ name = String(data[i][0]||''); break; }
      }
    }
  } catch(_){}
  var keys = [email];
  if (name) keys.push(name.toLowerCase());
  return { email: email, name: name, keys: keys };
}

function _dateToNo_(d){
  return (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, _tz_SP_(), 'dd.MM.yyyy') : '';
}
function _tz_SP_(){
  try { if (typeof _tz_ === 'function') return _tz_(); } catch(_){}
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Europe/Oslo';
}

/**
 * Hent oppgaver for innlogget vaktmester.
 * @param {'active'|'history'} filter
 * @returns {{ok:boolean, items:object[]}}
 */
function getTasksForVaktmester(filter){
  try{
    var profile = _currentUserProfile_();
    if (!profile.email) throw new Error('Kunne ikke identifisere brukeren.');

    var sh = _getTasksSheet_();
    if (sh.getLastRow() < 2) return { ok:true, items:[] };

    var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var M = _hdrMap_(H);
    var vals = sh.getRange(2,1,sh.getLastRow()-1, sh.getLastColumn()).getValues();

    var wantClosed = (String(filter||'active').toLowerCase() === 'history');
    var items = [];
    var meKeys = profile.keys;

    for (var i=0;i<vals.length;i++){
      var r = vals[i];

      var ansvarlig = String(r[M['Ansvarlig']-1]||'').trim().toLowerCase();
      var kategori  = String(r[M['Kategori']-1]||'').trim().toLowerCase();
      var statusRaw = String(r[M['Status']-1]||'').trim().toLowerCase();

      var assignedToMe = meKeys.indexOf(ansvarlig) >= 0 ||
                         (ansvarlig === '' && kategori.indexOf('vaktmester') >= 0); // fallback

      if (!assignedToMe) continue;

      var inActive = VM_ACTIVE_STATUSES.indexOf(statusRaw) >= 0;
      var inClosed = VM_CLOSED_STATUSES.indexOf(statusRaw) >= 0;

      if (wantClosed && !inClosed) continue;
      if (!wantClosed && !inActive) continue;

      var frist = r[M['Frist']-1]; var opprettet = r[M['Opprettet']-1];
      items.push({
        id: r[M['OppgaveID']-1],
        tittel: r[M['Tittel']-1],
        beskrivelse: r[M['Beskrivelse']-1],
        seksjon: r[M['Seksjonsnr']-1],
        frist: _dateToNo_(frist),
        opprettet: _dateToNo_(opprettet),
        status: r[M['Status']-1],
        prioritet: r[M['Prioritet']-1] || '—'
      });
    }
    return { ok:true, items: items };
  } catch(e){
    _logEvent && _logEvent('VaktmesterAPI_Feil', 'getTasksForVaktmester: ' + e.message);
    return { ok:false, error: e.message, items: [] };
  }
}

/**
 * Sett status (Fullført/Avvist) – med valgfri kommentar.
 */
function updateTaskStatusByVaktmester(taskId, newStatus, comment){
  try{
    var profile = _currentUserProfile_();
    if (!taskId) throw new Error('Mangler OppgaveID.');
    newStatus = String(newStatus||'').trim();
    var valid = ['Fullført','Avvist'];
    if (valid.indexOf(newStatus) < 0) throw new Error('Ugyldig status.');

    var sh = _getTasksSheet_();
    var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var M = _hdrMap_(H);

    var range = sh.createTextFinder(taskId).matchEntireCell(true).findNext();
    if (!range) throw new Error('Fant ikke oppgave ' + taskId);
    var row = range.getRow();
    var rowVals = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];

    var ansvarlig = String(rowVals[M['Ansvarlig']-1]||'').trim().toLowerCase();
    var allowed = (profile.keys.indexOf(ansvarlig) >= 0);
    if (!allowed) throw new Error('Tilgang nektet. Du er ikke ansvarlig for denne oppgaven.');

    // Oppdater status
    sh.getRange(row, M['Status']).setValue(newStatus);

    // Kommentar-append
    if (comment && M['Kommentarer']){
      var cur = String(rowVals[M['Kommentarer']-1]||'');
      var stamp = Utilities.formatDate(new Date(), _tz_SP_(), 'yyyy-MM-dd HH:mm');
      var add = (cur ? (cur + '\n') : '') + '['+stamp+'] ' + (profile.email||'') + ': ' + String(comment||'');
      sh.getRange(row, M['Kommentarer']).setValue(add);
    }

    _logEvent && _logEvent('Oppgave_Status', 'Vaktmester ' + (profile.email||'') + ' endret ' + taskId + ' → ' + newStatus);
    return { ok:true, message: 'Status for ' + taskId + ' er satt til ' + newStatus + '.' };
  } catch(e){
    _logEvent && _logEvent('VaktmesterAPI_Feil', 'updateTaskStatusByVaktmester: ' + e.message);
    return { ok:false, error: e.message };
  }
}

/** Legg til kommentar på en oppgave (uten å endre status). */
function addTaskCommentByVaktmester(taskId, comment){
  try{
    if (!taskId) throw new Error('Mangler OppgaveID.');
    if (!comment) throw new Error('Skriv en kommentar.');

    var profile = _currentUserProfile_();
    var sh = _getTasksSheet_();
    var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var M = _hdrMap_(H);

    var range = sh.createTextFinder(taskId).matchEntireCell(true).findNext();
    if (!range) throw new Error('Fant ikke oppgave ' + taskId);
    var row = range.getRow();
    var rowVals = sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];

    // Valgfritt: sjekk eierskap – vi lar alle vaktmestere som "eier" saken gjøre dette
    var ansvarlig = String(rowVals[M['Ansvarlig']-1]||'').trim().toLowerCase();
    var allowed = _currentUserProfile_().keys.indexOf(ansvarlig) >= 0;
    if (!allowed) throw new Error('Tilgang nektet.');

    var cur = String(rowVals[M['Kommentarer']-1]||'');
    var stamp = Utilities.formatDate(new Date(), _tz_SP_(), 'yyyy-MM-dd HH:mm');
    var add = (cur ? (cur + '\n') : '') + '['+stamp+'] ' + (profile.email||'') + ': ' + String(comment||'');
    sh.getRange(row, M['Kommentarer']).setValue(add);

    return { ok:true, message: 'Kommentar lagt til.' };
  } catch(e){
    _logEvent && _logEvent('VaktmesterAPI_Feil', 'addTaskCommentByVaktmester: ' + e.message);
    return { ok:false, error: e.message };
  }
}

/**
 * Opprett ny oppgave (vaktmester-sak) fra UI.
 * payload: { tittel, beskrivelse, seksjonsnr, frist }  (frist: yyyy-mm-dd)
 */
function createVaktmesterTask(payload){
  try{
    var profile = _currentUserProfile_();
    var sh = _getTasksSheet_();
    var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var M = _hdrMap_(H);

    var title = String(payload && payload.tittel || '').trim();
    if (!title) throw new Error('Skriv tittel.');
    var desc = String(payload && payload.beskrivelse || '').trim();
    var seksjon = String(payload && payload.seksjonsnr || '').trim();
    var frist = payload && payload.frist ? new Date(payload.frist) : null;

    var id = (typeof _nextTaskId_ === 'function') ? _nextTaskId_() : ('TASK-' + new Date().getTime());

    var row = [];
    row[M['OppgaveID']-1] = id;
    row[M['Tittel']-1]    = title;
    row[M['Beskrivelse']-1]= desc;
    row[M['Kategori']-1]  = 'Vaktmester';
    row[M['Prioritet']-1] = 'Medium';
    row[M['Opprettet']-1] = new Date();
    row[M['Frist']-1]     = (frist && !isNaN(frist)) ? frist : '';
    row[M['Status']-1]    = 'Ny';
    // Lagre ansvarlig som e-post (robust matching i API støtter både navn og e-post)
    row[M['Ansvarlig']-1] = profile.email || profile.name || '';
    row[M['Seksjonsnr']-1]= seksjon;
    row[M['Relatert']-1]  = '';
    row[M['Kommentarer']-1]= '';

    // Fyll tomme celler
    for (var c=0;c<H.length;c++){ if (typeof row[c] === 'undefined') row[c]=''; }

    sh.appendRow(row);
    _logEvent && _logEvent('Oppgaver','Ny vaktmester-sak: ' + id + ' (' + (profile.email||'') + ')');
    return { ok:true, id:id, message:'Opprettet ny sak.' };
  } catch(e){
    _logEvent && _logEvent('VaktmesterAPI_Feil', 'createVaktmesterTask: ' + e.message);
    return { ok:false, error: e.message };
  }
}

/** Til UI: hent enkel profil (navn/epost). */
function getCurrentVaktmesterProfile(){
  return _currentUserProfile_();
}

/** Test: lag en oppgave tildelt meg selv, så Vaktmester-UI får noe å vise */
function _debugCreateTaskForMe(opts) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS.TASKS) || ss.insertSheet(SHEETS.TASKS);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['OppgaveID','Tittel','Ansvarlig','Status','Seksjonsnr','Frist']);
  }
  const email = (Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail() || '').toLowerCase();
  const id = (typeof _nextTaskId_==='function') ? _nextTaskId_() : `TASK-${Date.now()}`;
  const tittel = opts?.tittel || 'Skifte lyspære i oppgang';
  const seksjon = opts?.seksjon || '';
  const frist = opts?.frist || '';
  sh.appendRow([id, tittel, email, 'Ny', seksjon, frist]);
  return id;
}
