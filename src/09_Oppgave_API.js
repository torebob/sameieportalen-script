/* global _getTasksSheet_, _hdrMap_ */
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
  } catch(e){}
  var keys = [email];
  if (name) keys.push(name.toLowerCase());
  return { email: email, name: name, keys: keys };
}

function _dateToNo_(d){
  return (d instanceof Date && !isNaN(d)) ? Utilities.formatDate(d, _tz_SP_(), 'dd.MM.yyyy') : '';
}

function _tz_SP_(){
  try { if (typeof _tz_ === 'function') return _tz_(); } catch(e){}
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Europe/Oslo';
}

function getTasksForVaktmester(filter){
  try{
    var profile = _currentUserProfile_();
    if (!profile.email) throw new Error('Kunne ikke identifisere brukeren.');

    var sh = _getTasksSheet_();
    var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var M = _hdrMap_(H);

    if (sh.getLastRow() < 2) return { ok:true, items:[] };

    var data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
    var items = [];

    for (var i=0; i<data.length; i++){
      var row = data[i];
      var status = String(row[M['Status']-1]||'').toLowerCase();
      var ansvarlig = String(row[M['Ansvarlig']-1]||'').toLowerCase();

      var isOwner = profile.keys.indexOf(ansvarlig) >= 0;
      if (!isOwner) continue;

      var isActive = VM_ACTIVE_STATUSES.indexOf(status) >= 0;
      var isClosed = VM_CLOSED_STATUSES.indexOf(status) >= 0;

      if (filter === 'active' && !isActive) continue;
      if (filter === 'history' && !isClosed) continue;

      items.push({
        oppgaveID: row[M['OppgaveID']-1],
        tittel: row[M['Tittel']-1],
        beskrivelse: row[M['Beskrivelse']-1],
        status: status,
        frist: _dateToNo_(row[M['Frist']-1]),
                 prioritet: row[M['Prioritet']-1],
                 seksjonsnr: row[M['Seksjonsnr']-1]
      });
    }

    return { ok:true, items:items };
  }
  catch(e){
    return { ok:false, error: e.message };
  }
}

function _getTaskRowCtx_(taskId){
  var sh = _getTasksSheet_();
  var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var M = _hdrMap_(H);
  var data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();

  for (var i=0; i<data.length; i++){
    if (String(data[i][M['OppgaveID']-1]) === String(taskId)){
      return { sh:sh, H:H, M:M, row:i+2, rowVals:data[i] };
    }
  }
  throw new Error('Fant ikke oppgave: ' + taskId);
}

function addTaskCommentByVaktmester(taskId, comment){
  try{
    if (!taskId) throw new Error('Mangler OppgaveID.');
    if (!comment) throw new Error('Skriv en kommentar.');

    var profile = _currentUserProfile_();
    var ctx = _getTaskRowCtx_(taskId);
    var sh = ctx.sh, M = ctx.M, row = ctx.row, rowVals = ctx.rowVals;

    var ansvarlig = String(rowVals[M['Ansvarlig']-1]||'').trim().toLowerCase();
    var allowed = profile.keys.indexOf(ansvarlig) >= 0;
    if (!allowed) throw new Error('Tilgang nektet.');

    var cur = String(rowVals[M['Kommentarer']-1]||'');
    var stamp = Utilities.formatDate(new Date(), _tz_SP_(), 'yyyy-MM-dd HH:mm');
    var add = (cur ? (cur + '\n') : '') + '['+stamp+'] ' + (profile.email||'') + ': ' + String(comment||'');
    sh.getRange(row, M['Kommentarer']).setValue(add);

    return { ok:true, message: 'Kommentar lagt til.' };
  }
  catch(e){
    return { ok:false, error: e.message };
  }
}

function createVaktmesterTask(payload){
  try{
    var profile = _currentUserProfile_();
    if (!profile.email) throw new Error('Kunne ikke identifisere brukeren.');

    var sh = _getTasksSheet_();
    var nextId = 'TASK-' + Utilities.formatDate(new Date(), _tz_SP_(), 'yyyyMMddHHmmss');

    sh.appendRow([
      nextId,
      payload.tittel || '',
      payload.beskrivelse || '',
      'Vaktmester',
      payload.prioritet || 'Normal',
      new Date(),
                 payload.frist || '',
                 'ny',
                 profile.email,
                 payload.seksjonsnr || '',
                 '',
                 ''
    ]);

    return { ok:true, taskId:nextId, message:'Oppgave opprettet: ' + nextId };
  }
  catch(e){
    return { ok:false, error: e.message };
  }
}
