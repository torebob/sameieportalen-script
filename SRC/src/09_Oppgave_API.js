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

    var ctx = _getTaskRowCtx_(taskId);
var sh = ctx.sh, H = ctx.H, M = ctx.M, row = ctx.row, rowVals = ctx.rowVals;

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
    var ctx = _getTaskRowCtx_(taskId);
var sh = ctx.sh, H = ctx.H, M = ctx.M, row = ctx.row, rowVals = ctx.rowVals;

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
    var ctx = _getTaskRowCtx_(taskId);
var sh = ctx.sh, H = ctx.H, M = ctx.M, row = ctx.row, rowVals = ctx.rowVals;

  return { sh: sh, H: H, M: M, row: row, rowVals: rowVals };
}
