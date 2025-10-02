/* ================== Elektronisk Protokollgodkjenning (API) ==================
 * FILE: 11_Protokoll_API.gs | VERSION: 2.0.0 | UPDATED: 2025-09-13
 * FORMÅL: Sende, spore og motta godkjenninger/avvisninger for møteprotokoller.
 * - Støtter Møte-ID fra Møter-arket, én rad pr styremedlem m/unik token.
 * - doGet(e) validerer token og oppdaterer både Godkjenning og Møter-status.
 * ========================================================================== */

/* ---------- Konstanter (tilpasser seg SHEET_HEADERS hvis tilgjengelig) ---------- */
var _PG_HEADERS_ = (typeof SHEET_HEADERS !== 'undefined' && SHEET_HEADERS[SHEETS.PROTOKOLL_GODKJENNING])
  ? SHEET_HEADERS[SHEETS.PROTOKOLL_GODKJENNING]
  : ['Godkjenning-ID','Møte-ID','Navn','E-post','Token','Utsendt-Dato','Status','Svar-Dato','Kommentar','Protokoll-URL'];

var _MOTE_HEADERS_FALLBACK_ = ['Møte-ID','Type','Dato','Starttid','Sluttid','Sted','Tittel','Agenda','Protokoll-URL','Deltakere','Kalender-ID','Status'];

/* ---------- Små hjelpere (failsafe) ---------- */
function _pgLog_(type, msg){
  try { (typeof _logEvent === 'function') ? _logEvent(type, msg) : Logger.log(type + '> ' + msg); } catch(_){}
}
function _ensureSheetLocal_(name, headers){
  if (typeof ensureSheetWithHeaders === 'function') return ensureSheetWithHeaders(name, headers);
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var cur = sh.getRange(1,1,1,headers.length).getValues()[0];
  var Sameie.Sheets.ensureHeader(sh, headers);
  return sh;
}
function _hdrIdxMap_(headers, names){
  var m = {};
  for (var i=0;i<names.length;i++) m[names[i]] = headers.indexOf(names[i]);
  return m;
}
function _uuid8_(){ return Utilities.getUuid().replace(/-/g,'').slice(0,8); }
function _tzLocal_(){ try { return (typeof _tz_==='function') ? _tz_() : (SpreadsheetApp.getActive().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Europe/Oslo'); } catch(_){ return 'Europe/Oslo'; } }

/* ---------- Datatilgang ---------- */
function _getBoardList_(){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.BOARD);
  if (!sh || sh.getLastRow() < 2) return [];
  // Forventet headers: ['Navn','E-post','Rolle']
  var vals = sh.getRange(2,1,sh.getLastRow()-1,3).getValues();
  var out = [];
  for (var i=0;i<vals.length;i++){
    var navn = String(vals[i][0]||'').trim();
    var mail = String(vals[i][1]||'').trim();
    if (navn && mail) out.push({navn:navn, email:mail});
  }
  return out;
}

function _findMoteRow_(moteId){
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEETS.MOTER);
  if (!sh || sh.getLastRow() < 2) return null;

  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var H = _hdrIdxMap_(headers, _MOTE_HEADERS_FALLBACK_);
  if (H['Møte-ID'] === -1) return null;

  var idColRange = sh.getRange(2, H['Møte-ID']+1, sh.getLastRow()-1, 1);
  var finder = idColRange.createTextFinder(String(moteId)).matchEntireCell(true);
  var hit = finder.findNext();
  if (!hit) return null;

  return { sheet: sh, row: hit.getRow(), headers: headers, H: H };
}

/* ========================================================================== */
/* K7-10: START UTSENDING – én rad pr styremedlem, unik token                 */
/* ========================================================================== */
/**
 * Start godkjenning for en protokoll.
 * @param {string} moteId - Møte-ID (obligatorisk)
 * @param {string} [protokollUrl] - Protokollens Google Docs-URL (hentes fra Møter hvis utelatt)
 * @returns {{ok:boolean,message:string,gid:string,count:number}}
 */
function sendProtokollForGodkjenning(moteId, protokollUrl){
  try{
    // Tillat bare fra admin/styremeny om RBAC finnes
    if (typeof hasPermission === 'function' && !hasPermission('VIEW_ADMIN_MENU')) {
      throw new Error('Tilgang nektet. (Krever VIEW_ADMIN_MENU)');
    }

    if (!moteId) throw new Error('Mangler Møte-ID.');
    var mote = _findMoteRow_(moteId);
    if (!mote) throw new Error('Fant ikke møtet i "' + SHEETS.MOTER + '".');

    // Finn/valider protokoll-URL
    var url = String(protokollUrl || '').trim();
    if (!url){
      var cProt = mote.headers.indexOf('Protokoll-URL');
      if (cProt === -1) throw new Error('Møter-arket mangler kolonnen "Protokoll-URL".');
      url = String(mote.sheet.getRange(mote.row, cProt+1).getValue() || '').trim();
    }
    if (!/^https:\/\/docs\.google\.com\/document\//.test(url)){
      throw new Error('Ugyldig protokoll-URL (må være Google Docs).');
    }

    var godkjSh = _ensureSheetLocal_(SHEETS.PROTOKOLL_GODKJENNING, _PG_HEADERS_);
    var board = _getBoardList_();
    if (!board.length) throw new Error('Fant ingen styremedlemmer i "' + SHEETS.BOARD + '".');

    // Lag ny Godkjenning-ID
    var gid = 'G-' + String(moteId).replace(/[^A-Za-z0-9_-]/g,'') + '-' + _uuid8_();
    var now = new Date();
    var tz = _tzLocal_();

    // Skriv én rad pr mottaker
    var h = godkjSh.getRange(1,1,1,godkjSh.getLastColumn()).getValues()[0];
    var H = _hdrIdxMap_(h, _PG_HEADERS_);
    var rows = [];
    for (var i=0;i<board.length;i++){
      var token = _uuid8_() + _uuid8_(); // 16 hex
      var r = new Array(h.length);
      r[H['Godkjenning-ID']] = gid;
      r[H['Møte-ID']]       = moteId;
      r[H['Navn']]          = board[i].navn;
      r[H['E-post']]        = board[i].email;
      r[H['Token']]         = token;
      r[H['Utsendt-Dato']]  = now;
      r[H['Status']]        = 'Sendt';
      r[H['Protokoll-URL']] = url;
      rows.push(r);
    }
    if (rows.length){
      godkjSh.getRange(godkjSh.getLastRow()+1,1,rows.length,h.length).setValues(rows);
    }

    // Oppdater Møter-status → Til godkjenning
    if (mote.H['Status'] !== -1){
      mote.sheet.getRange(mote.row, mote.H['Status']+1).setValue('Til godkjenning');
    }

    // Send e-poster
    var webAppUrl;
    try { webAppUrl = ScriptApp.getService().getUrl(); }
    catch(_){ webAppUrl = ''; }

    for (var j=0;j<board.length;j++){
      var tokenCell = rows[j][H['Token']];
      var email = board[j].email;
      var approveUrl = webAppUrl ? (webAppUrl + '?gid=' + encodeURIComponent(gid) + '&token=' + encodeURIComponent(tokenCell) + '&action=approve') : url;
      var rejectUrl  = webAppUrl ? (webAppUrl + '?gid=' + encodeURIComponent(gid) + '&token=' + encodeURIComponent(tokenCell) + '&action=reject')  : url;

      var subject = '[Sameieportalen] Til godkjenning: Protokoll ' + moteId;
      var body =
        '<p>Hei ' + board[j].navn + ',</p>' +
        '<p>Protokollen for møtet <b>' + moteId + '</b> er klar for godkjenning.</p>' +
        '<p><a href="'+ url +'" target="_blank" rel="noopener"><b>Les protokollen her</b></a></p>' +
        (webAppUrl
          ? ('<p>Registrer ditt valg:</p>' +
             '<div style="margin:12px 0">' +
             '<a href="'+approveUrl+'" style="background:#16a34a;color:#fff;padding:10px 14px;border-radius:6px;text-decoration:none;margin-right:8px">Godkjenn</a>' +
             '<a href="'+rejectUrl+'"  style="background:#dc2626;color:#fff;padding:10px 14px;border-radius:6px;text-decoration:none">Avvis</a>' +
             '</div>')
          : '<p>(WebApp-URL mangler, kontakt administrator.)</p>'
        ) +
        '<p>— Sameieportalen</p>';

      try {
        MailApp.sendEmail({ to: email, subject: subject, htmlBody: body });
      } catch (mailErr) {
        _pgLog_('Protokoll_MailFeil', 'E-post til ' + email + ' feilet: ' + mailErr.message);
      }
    }

    _pgLog_('Protokoll', 'Sendte godkjenning ' + gid + ' for Møte-ID ' + moteId + ' til ' + board.length + ' mottakere.');
    return { ok:true, message:'Protokoll sendt til ' + board.length + ' styremedlemmer.', gid: gid, count: board.length };
  } catch (e){
    _pgLog_('Protokoll_Feil', 'sendProtokollForGodkjenning: ' + e.message);
    throw e;
  }
}

/* ========================================================================== */
/* doGet: mottar klikk fra e-post (Godkjenn / Avvis)                          */
/* ========================================================================== */
/**
 * WebApp-endepunkt for godkjenning/avvisning via lenke.
 * Parametre: gid, token, action(approve|reject), [comment]
 */
function doGet(e){
  var title = 'Protokoll';
  function page(msg){ return HtmlService.createHtmlOutput('<h3>'+title+'</h3><p>'+msg+'</p>'); }

  try {
    var p = e && e.parameter ? e.parameter : {};
    var gid   = String(p.gid||'').trim();
    var token = String(p.token||'').trim();
    var action = String(p.action||'').trim().toLowerCase();
    var comment = String(p.comment||'').trim();

    if (!gid || !token || (action !== 'approve' && action !== 'reject')) {
      title = 'Ugyldig forespørsel';
      return page('Lenken mangler nødvendig informasjon.');
    }

    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEETS.PROTOKOLL_GODKJENNING);
    if (!sh || sh.getLastRow() < 2) {
      title = 'Feil';
      return page('Godkjenningsarket mangler.');
    }
    var h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var H = _hdrIdxMap_(h, _PG_HEADERS_);
    if ([H['Godkjenning-ID'],H['Token'],H['Status'],H['Svar-Dato'],H['Kommentar'],H['Møte-ID']].some(function(i){return i===-1;})){
      title = 'Feil';
      return page('Godkjenningsarket har ikke forventede kolonner.');
    }

    // Finn raden for token + gid
    var tokenColRange = sh.getRange(2, H['Token']+1, sh.getLastRow()-1, 1);
    var finder = tokenColRange.createTextFinder(token).matchEntireCell(true);
    var hit = finder.findNext();
    if (!hit){
      title = 'Ugyldig lenke';
      return page('Fant ikke denne godkjenningsforespørselen (token).');
    }
    var row = hit.getRow();
    var rowVals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
    if (String(rowVals[H['Godkjenning-ID']]||'').trim() !== gid){
      title = 'Ugyldig lenke';
      return page('Godkjennings-ID passer ikke.');
    }

    // Hvis allerede svart – vis status, ikke dobbelregistrer
    var prev = String(rowVals[H['Status']]||'').trim();
    if (prev === 'Godkjent' || prev === 'Avvist'){
      title = 'Allerede behandlet';
      return page('Ditt svar er allerede registrert: <b>' + prev + '</b>.');
    }

    // Oppdater denne raden
    var newStatus = (action === 'approve') ? 'Godkjent' : 'Avvist';
    sh.getRange(row, H['Status']+1).setValue(newStatus);
    sh.getRange(row, H['Svar-Dato']+1).setValue(new Date());
    if (comment) sh.getRange(row, H['Kommentar']+1).setValue(comment);

    // Evaluer samlet status for GID
    var dataRange = sh.getDataRange().getValues();
    var header = dataRange.shift();
    var rowsForGid = [];
    for (var i=0;i<dataRange.length;i++){
      if (String(dataRange[i][H['Godkjenning-ID']]||'').trim() === gid){
        rowsForGid.push(dataRange[i]);
      }
    }
    var anyRejected = rowsForGid.some(function(r){ return String(r[H['Status']]||'').trim() === 'Avvist'; });
    var allApproved = rowsForGid.length > 0 && rowsForGid.every(function(r){ return String(r[H['Status']]||'').trim() === 'Godkjent'; });

    // Oppdater Møter-status
    var moteId = String(rowVals[H['Møte-ID']]||'').trim();
    var mote = moteId ? _findMoteRow_(moteId) : null;
    if (mote && mote.H['Status'] !== -1){
      if (anyRejected)      mote.sheet.getRange(mote.row, mote.H['Status']+1).setValue('Avvist');
      else if (allApproved) mote.sheet.getRange(mote.row, mote.H['Status']+1).setValue('Godkjent');
      else                  mote.sheet.getRange(mote.row, mote.H['Status']+1).setValue('Til godkjenning');
    }

    _pgLog_('Protokoll', 'Mottok ' + newStatus + ' for ' + gid + ' (Møte ' + (moteId||'?') + ').');

    title = (newStatus === 'Godkjent') ? 'Takk for godkjenningen!' : 'Avvisning registrert';
    var msg = (newStatus === 'Godkjent')
      ? 'Din godkjenning er registrert. '
      : 'Din avvisning er registrert. Referent/styret blir varslet ved behov.';
    return page(msg);

  } catch (err){
    _pgLog_('Protokoll_WebApp_Feil', err.message);
    title = 'En feil oppstod';
    return page(err.message);
  }
}
