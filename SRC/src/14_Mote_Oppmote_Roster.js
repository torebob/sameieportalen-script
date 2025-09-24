/* ===================== Møte – Oppmøte & Stemmerett (Roster) ===================== */
/* FILE: 14_Mote_Oppmote_Roster.gs | VERSION: 1.0.0 | UPDATED: 2025-09-14
 * FORMÅL: Per-møte deltakerliste (oppmøte) og stemmerett. Brukes av voter-modulen.
 * MERK: Prefererer roster til å bestemme stemmerett; faller tilbake til Styret-arket hvis ingen roster-data.
 */

function _attSheetName_(){
  return (typeof SHEETS !== 'undefined' && SHEETS.MOTE_TILSTEDE) ? SHEETS.MOTE_TILSTEDE : 'MoteTilstede';
}
function _attHeaders_(){ return ['MoteID','Email','Name','Role','Present','Voting','ProxyFor','Timestamp']; }

/** Sørger for at arket finnes og har headere. */
function _ensureAttendanceSheet_(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(_attSheetName_());
  if (!sh){ sh = ss.insertSheet(_attSheetName_()); sh.appendRow(_attHeaders_()); }
  if (sh.getLastRow() === 0) sh.appendRow(_attHeaders_());
  return sh;
}

/** Seeder roster for et møte fra Styret-arket (hvis tomt). */
function _seedRosterFromBoard_(moteId){
  const ss = SpreadsheetApp.getActive();
  const shB = ss.getSheetByName(SHEETS.BOARD);
  if (!shB || shB.getLastRow() < 2) return;
  const vals = shB.getRange(2, 1, shB.getLastRow()-1, 3).getValues(); // Navn, E-post, Rolle
  const rows = vals
    .filter(r => String(r[1]||'').includes('@'))
    .map(r => [moteId, String(r[1]).toLowerCase().trim(), String(r[0]||''), String(r[2]||''), false, false, '', new Date()]);
  if (!rows.length) return;
  const sh = _ensureAttendanceSheet_();
  sh.getRange(sh.getLastRow()+1, 1, rows.length, rows[0].length).setValues(rows);
}

/** Henter roster for et møte. Seeder fra Styret-arket første gang. */
function getMeetingRoster(moteId){
  if (!moteId) throw new Error('Mangler MoteID.');
  const sh = _ensureAttendanceSheet_();
  const hdrs = _attHeaders_();
  const last = sh.getLastRow();
  let data = last > 1 ? sh.getRange(2,1,last-1,hdrs.length).getValues() : [];
  let rows = data.filter(r => String(r[0]) === String(moteId));
  if (!rows.length){ _seedRosterFromBoard_(moteId);
    const last2 = sh.getLastRow();
    data = last2 > 1 ? sh.getRange(2,1,last2-1,hdrs.length).getValues() : [];
    rows = data.filter(r => String(r[0]) === String(moteId));
  }
  return rows.map(r => ({
    moteId: r[0],
    email: String(r[1]||'').toLowerCase(),
    name: String(r[2]||''),
    role: String(r[3]||''),
    present: (r[4] === true || String(r[4]).toLowerCase()==='true' || String(r[4])==='1'),
    voting:  (r[5] === true || String(r[5]).toLowerCase()==='true' || String(r[5])==='1'),
    proxyFor: String(r[6]||'')
  }));
}

/** Bulk-oppdater roster (tilstede/stemmerett/proxy). */
function updateMeetingRoster(moteId, entries){
  if (!moteId) throw new Error('Mangler MoteID.');
  if (!Array.isArray(entries)) throw new Error('Entries må være en liste.');
  const sh = _ensureAttendanceSheet_();
  const hdrs = _attHeaders_();
  const last = sh.getLastRow();
  const data = last>1 ? sh.getRange(2,1,last-1,hdrs.length).getValues() : [];
  const idx = {}; // key = moteId|email -> row#
  for (let i=0;i<data.length;i++){
    idx[String(data[i][0])+'|'+String(data[i][1]).toLowerCase()] = i+2;
  }
  const ts = new Date();
  let updated=0, inserted=0;
  entries.forEach(e=>{
    const email = String(e.email||'').toLowerCase().trim();
    if (!email) return;
    const name = String(e.name||'');
    const role = String(e.role||'');
    const present = !!e.present;
    const voting  = !!e.voting;
    const proxyFor = String(e.proxyFor||'');
    const row = [moteId, email, name, role, present, voting, proxyFor, ts];
    const key = moteId + '|' + email;
    const r = idx[key];
    if (r){ sh.getRange(r,1,row.length).setValues([row]); updated++; }
    else { sh.appendRow(row); inserted++; }
  });
  if (typeof _logEvent === 'function'){
    _logEvent('Roster', `Roster oppdatert for ${moteId} (oppdatert=${updated}, nye=${inserted}).`);
  }
  return { ok:true, updated, inserted };
}

/** Hent stemmeberettigede for møte – foretrekker roster (Present && Voting). Fallback: Styret-arket. */
function getEligibleVotersForMeeting(moteId){
  const sh = _ensureAttendanceSheet_();
  const hdrs = _attHeaders_();
  const last = sh.getLastRow();
  const data = last>1 ? sh.getRange(2,1,last-1,hdrs.length).getValues() : [];
  const rows = data.filter(r => String(r[0])===String(moteId));
  if (rows.length){
    return rows
      .filter(r => (r[4]===true || String(r[4]).toLowerCase()==='true' || String(r[4])==='1') &&
                   (r[5]===true || String(r[5]).toLowerCase()==='true' || String(r[5])==='1'))
      .map(r => ({ name:String(r[2]||''), email:String(r[1]||'').toLowerCase() }));
  }
  // Fallback – hele styret.
  const ss = SpreadsheetApp.getActive();
  const shB = ss.getSheetByName(SHEETS.BOARD);
  if (!shB || shB.getLastRow()<2) return [];
  const vals = shB.getRange(2,1, shB.getLastRow()-1, 2).getValues(); // Navn, Epost
  return vals
    .map(r => ({ name:String(r[0]||''), email:String(r[1]||'').toLowerCase() }))
    .filter(v => v.email.includes('@'));
}
