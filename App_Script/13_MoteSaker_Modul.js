/* ================== MøteSaker – Innspill & Stemming (REN MODUL) ================== *
 * FILE: 13_MoteSaker_Modul.gs | VERSION: 1.0.3 | UPDATED: 2025-09-14
 * Denne filen inneholder KUN:
 *  - Innspill: appendInnspill(sakId, text), listInnspill(sakId, sinceISO)
 *  - Stemming: castVote(sakId, value), getVoteSummary(sakId)
 *
 * VIKTIG: Møte- og sakshåndtering (create/list/save saker) ligger i 12_Moter_Modul.gs.
 * Ikke lim inn kode fra 12_ i denne filen.
 * ================================================================================ */

/** ------- KONFIG / FALLBACKS ------- **/
var _SHEETS13 = (typeof SHEETS !== 'undefined') ? SHEETS : {};
// Ark-navn (bruker SHEETS.* hvis satt, ellers norske standarder)
var SAKER_SHEET_13    = _SHEETS13.MOTE_SAKER        || _SHEETS13.MØTE_SAKER        || 'MøteSaker';
var INNSPILL_SHEET_13 = _SHEETS13.MOTE_KOMMENTARER  || _SHEETS13.MØTE_KOMMENTARER  || 'MøteSakKommentarer';
var STEMMER_SHEET_13  = _SHEETS13.MOTE_STEMMER      || _SHEETS13.MØTE_STEMMER      || 'MøteSakStemmer';

// Kolonneheadere vi forventer på arkene over
var SAKER_HEADERS_13    = ['mote_id','sak_id','saksnr','tittel','forslag','vedtak','created_ts','updated_ts'];
var INNSPILL_HEADERS_13 = ['sak_id','ts','from','text'];
var STEMMER_HEADERS_13  = ['vote_id','sak_id','mote_id','email','name','vote','ts'];

/** ------- HELPERE ------- **/
function _ensureSheetWithHeaders13_(name, headers){
  if (typeof _ensureSheetWithHeaders_ === 'function') {
    // Bruk felles helper om den finnes i prosjektet
    return _ensureSheetWithHeaders_(name, headers);
  }
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  var first = sh.getRange(1,1,1,headers.length).getValues()[0];
  var need = false;
  for (var i=0;i<headers.length;i++){ if (String(first[i]||'') !== headers[i]) { need = true; break; } }
  if (need){
    sh.clearContents();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sh;
}
function _now13_(){ return new Date(); }
function _userEmail13_(){
  try {
    return (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
           (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) || '';
  } catch(_) { return ''; }
}
function _log13_(topic, msg){
  try { if (typeof _logEvent === 'function') _logEvent(topic, msg); } catch(_){}
}

/** ======================================================================
 *  INNSPILL
 *  - appendInnspill(sakId, text)
 *  - listInnspill(sakId, sinceISO)
 * ====================================================================== */

/**
 * Legg til innspill på en sak.
 * @param {string} sakId
 * @param {string} text
 * @returns {{ok:boolean}}
 */
function appendInnspill(sakId, text){
  if (!sakId) throw new Error('Mangler sakId.');
  var txt = String(text||'').trim();
  if (!txt) return { ok:false, message:'Tomt innspill.' };

  var sh = _ensureSheetWithHeaders13_(INNSPILL_SHEET_13, INNSPILL_HEADERS_13);
  var email = String(_userEmail13_()||'').toLowerCase();

  sh.appendRow([sakId, _now13_(), email, txt]);
  _log13_('Innspill', email + ' -> ' + sakId);
  return { ok:true };
}

/**
 * Hent innspill for en sak, opsjonelt bare nye etter sinceISO.
 * @param {string} sakId
 * @param {string?} sinceISO
 * @returns {Array<{sakId:string, ts:Date, from:string, text:string}>}
 */
function listInnspill(sakId, sinceISO){
  if (!sakId) return [];
  var sh = _ensureSheetWithHeaders13_(INNSPILL_SHEET_13, INNSPILL_HEADERS_13);
  if (sh.getLastRow() < 2) return [];

  var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  var idx = { sak_id:H.indexOf('sak_id'), ts:H.indexOf('ts'), from:H.indexOf('from'), text:H.indexOf('text') };
  var since = sinceISO ? new Date(sinceISO) : null;

  var out = [];
  for (var i=0;i<data.length;i++){
    var r = data[i];
    if (String(r[idx.sak_id]) !== String(sakId)) continue;
    var ts = r[idx.ts] instanceof Date ? r[idx.ts] : (r[idx.ts] ? new Date(r[idx.ts]) : null);
    if (since && ts && ts <= since) continue;
    out.push({ sakId: r[idx.sak_id], ts: ts, from: r[idx.from], text: r[idx.text] });
  }
  out.sort(function(a,b){ return (a.ts&&b.ts)?(a.ts.getTime()-b.ts.getTime()):0; });
  return out;
}

/** ======================================================================
 *  STEMMING
 *  - castVote(sakId, value)   value ∈ {'JA','NEI','BLANK'}
 *  - getVoteSummary(sakId)    -> {JA,NEI,BLANK,total}
 * ====================================================================== */

/**
 * Lagre / oppdater stemme for innlogget bruker på en sak.
 * @param {string} sakId
 * @param {'JA'|'NEI'|'BLANK'} value
 */
function castVote(sakId, value){
  if (!sakId) throw new Error('Mangler sakId.');
  var v = String(value||'').toUpperCase();
  if (['JA','NEI','BLANK'].indexOf(v) === -1) throw new Error('Ugyldig stemme. Bruk JA/NEI/BLANK.');

  // Finn møtereferanse (mote_id) fra saker-arket
  var saker = _ensureSheetWithHeaders13_(SAKER_SHEET_13, SAKER_HEADERS_13);
  var Hs = saker.getRange(1,1,1,saker.getLastColumn()).getValues()[0];
  var di = saker.getRange(2,1,Math.max(0,saker.getLastRow()-1), saker.getLastColumn()).getValues();
  var sIdx = { sak_id:Hs.indexOf('sak_id'), mote_id:Hs.indexOf('mote_id') };

  var moteId = '';
  for (var i=0;i<di.length;i++){
    if (String(di[i][sIdx.sak_id]) === String(sakId)){ moteId = di[i][sIdx.mote_id]; break; }
  }
  if (!moteId) throw new Error('Fant ikke tilhørende møte for sak ' + sakId);

  var email = String(_userEmail13_()||'').toLowerCase();
  var name = ''; // valgfritt: kan fylles fra eget kontaktark

  var sh = _ensureSheetWithHeaders13_(STEMMER_SHEET_13, STEMMER_HEADERS_13);
  var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var rows = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues() : [];
  var idx = { vote_id:H.indexOf('vote_id'), sak_id:H.indexOf('sak_id'), mote_id:H.indexOf('mote_id'),
              email:H.indexOf('email'), name:H.indexOf('name'), vote:H.indexOf('vote'), ts:H.indexOf('ts') };

  // Sjekk om brukeren har stemt før på samme sak → oppdater
  var foundRow = -1;
  for (var r=0;r<rows.length;r++){
    if (String(rows[r][idx.sak_id]) === String(sakId) && String(rows[r][idx.email]).toLowerCase() === email){
      foundRow = r + 2; // data starter på rad 2
      break;
    }
  }

  if (foundRow > -1){
    var cur = sh.getRange(foundRow, 1, 1, sh.getLastColumn()).getValues()[0];
    cur[idx.vote] = v;
    cur[idx.ts]   = _now13_();
    sh.getRange(foundRow, 1, 1, sh.getLastColumn()).setValues([cur]);
  } else {
    sh.appendRow(['V-'+Utilities.getUuid().slice(0,8), sakId, String(moteId), email, name, v, _now13_()]);
  }

  _log13_('Stemme', email + ' -> ' + sakId + ' = ' + v);
  return { ok:true };
}

/**
 * Oppsummer stemmer for en sak.
 * @param {string} sakId
 * @returns {{JA:number,NEI:number,BLANK:number,total:number}}
 */
function getVoteSummary(sakId){
  if (!sakId) return {JA:0,NEI:0,BLANK:0,total:0};
  var sh = _ensureSheetWithHeaders13_(STEMMER_SHEET_13, STEMMER_HEADERS_13);
  if (sh.getLastRow() < 2) return {JA:0,NEI:0,BLANK:0,total:0};

  var H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  var idx = { sak_id:H.indexOf('sak_id'), vote:H.indexOf('vote') };

  var c = {JA:0,NEI:0,BLANK:0};
  for (var i=0;i<data.length;i++){
    if (String(data[i][idx.sak_id]) !== String(sakId)) continue;
    var v = String(data[i][idx.vote]||'').toUpperCase();
    if (c[v] != null) c[v]++;
  }
  return { JA:c.JA, NEI:c.NEI, BLANK:c.BLANK, total:(c.JA + c.NEI + c.BLANK) };
}
