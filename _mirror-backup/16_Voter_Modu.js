/* ====================== Møte – Avstemming (Voter) ======================
 * FILE: 16_Voter_Modul.gs | VERSION: 1.0.1 | UPDATED: 2025-09-14
 * FORMÅL: Backend for avstemming pr. sak i et møte (Apps Script).
 * Funksjoner (globale):
 *   - voterEnsureSheets_()
 *   - voterSaveVote(moteId, saksnr, vote)            // vote = 'JA' | 'NEI' | 'BLANK'
 *   - voterGetStatus(moteId, saksnr)
 *   - voterLockDecision(moteId, saksnr, vedtakTekst)  // låser saken og lagrer endelig vedtak
 * Avhenger IKKE av andre hjelpefiler – kjører “standalone”.
 * Fallback-navn hvis SHEETS.* ikke finnes:
 *   - STEMME_SHEET: 'MøteSakStemmer'
 *   - SAKER_SHEET : 'MøteSaker' (ikke strengt nødvendig her, men satt for fremtidig bruk)
 * ====================================================================== */

(function(){

  // ----------- KONFIG (med trygge fallbacks) -----------
  const STEMME_SHEET =
    (typeof SHEETS !== 'undefined' && SHEETS && SHEETS.MOTE_STEMMER) ? SHEETS.MOTE_STEMMER : 'MøteSakStemmer';
  const SAKER_SHEET =
    (typeof SHEETS !== 'undefined' && SHEETS && SHEETS.MOTE_SAKER) ? SHEETS.MOTE_SAKER : 'MøteSaker';

  const VOTES = ['JA','NEI','BLANK'];
  const LOCK_ROW_USER = '_LOCK';  // spesial “bruker” for lås/vedtak-rad

  // Standard-header vi forventer i STEMME_SHEET
  const HEADER = ['MoteId','Saksnr','Bruker','Epost','Stemme','Tid','Vedtak','Låst'];

  // ----------- Utils -----------
  function _log(msg){ try{ console.log('[Voter] ' + msg); }catch(_){ Logger.log('[Voter] ' + msg); } }
  function _now(){ return new Date(); }
  function _email(){
    const a = Session.getActiveUser()?.getEmail?.() || '';
    const e = a || Session.getEffectiveUser()?.getEmail?.() || '';
    return String(e||'').toLowerCase();
  }
  function _openOrCreateSheet_(name){
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    return sh;
  }
  function _ensureHeader_(sh, wantHeader){
    const last = sh.getLastRow();
    if (last === 0){
      sh.getRange(1,1,1,wantHeader.length).setValues([wantHeader]);
      return;
    }
    const cur = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x => String(x||'').trim());
    const same = (cur.length >= wantHeader.length) && wantHeader.every((h,i)=>cur[i]===h);
    if (!same){
      // Skriv ønsket header på rad 1 (ikke destruktiv mot data under)
      sh.getRange(1,1,1,wantHeader.length).setValues([wantHeader]);
    }
  }
  /*
   * MERK: _headerMap_() er fjernet fra denne filen for å unngå konflikter.
   * Den globale versjonen fra 000_Utils.js brukes i stedet.
   */
  function _getRows_(sh){
    const last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
  }

  function _ensureSheets(){
    const stem = _openOrCreateSheet_(STEMME_SHEET);
    _ensureHeader_(stem, HEADER);
    return { stemmeSheet: STEMME_SHEET, sakerSheet: SAKER_SHEET };
  }

  // ----------- Offentlig: create/verify sheets -----------
  globalThis.voterEnsureSheets_ = function voterEnsureSheets_(){
    const res = _ensureSheets();
    _log('Sheets ok: ' + JSON.stringify(res));
    return res;
  };

  // ----------- Offentlig: lagre/oppdatere stemme -----------
  globalThis.voterSaveVote = function voterSaveVote(moteId, saksnr, vote){
    if (!moteId || !saksnr) throw new Error('Mangler Møte-ID eller Saksnr.');
    const v = String(vote||'').toUpperCase().trim();
    if (!VOTES.includes(v)) throw new Error('Ugyldig stemme. Bruk JA, NEI eller BLANK.');

    _ensureSheets();
    const sh = SpreadsheetApp.getActive().getSheetByName(STEMME_SHEET);
    const H = _headerMap_(sh);
    const rows = _getRows_(sh);

    // Låst?
    const locked = rows.some(r =>
      String(r[(H['moteid']||1)-1])===moteId &&
      String(r[(H['saksnr']||2)-1])===saksnr &&
      String(r[(H['bruker']||3)-1])===LOCK_ROW_USER &&
      String(r[(H['låst']||8)-1]).toString().toLowerCase()==='true'
    );
    if (locked) throw new Error('Saken er låst. Stemming avsluttet.');

    const email = _email();
    const userName = email ? email.split('@')[0] : '(ukjent)';

    // Finn eksisterende rad for (moteId,saksnr,email)
    const idx = rows.findIndex(r =>
      String(r[(H['moteid']||1)-1])===moteId &&
      String(r[(H['saksnr']||2)-1])===saksnr &&
      String(r[(H['epost']||4)-1]).toLowerCase()===email
    );

    const now = _now();
    if (idx >= 0){
      // oppdater eksisterende
      const r = idx+2; // 1 header + 1 offset
      sh.getRange(r, H['stemme']).setValue(v);
      sh.getRange(r, H['tid']).setValue(now);
    } else {
      // legg til ny
      sh.appendRow([
        moteId,
        saksnr,
        userName,         // Bruker
        email,            // Epost
        v,                // Stemme
        now,              // Tid
        '',               // Vedtak
        false             // Låst
      ]);
    }

    _log(`vote saved: ${moteId}/${saksnr} ${email} -> ${v}`);
    return { ok:true, message:'Stemme registrert.', moteId, saksnr, vote:v };
  };

  // ----------- Offentlig: hent status for sak -----------
  globalThis.voterGetStatus = function voterGetStatus(moteId, saksnr){
    if (!moteId || !saksnr) throw new Error('Mangler Møte-ID eller Saksnr.');
    _ensureSheets();
    const sh = SpreadsheetApp.getActive().getSheetByName(STEMME_SHEET);
    const H = _headerMap_(sh);
    const rows = _getRows_(sh);
    const me = _email();

    const filtered = rows.filter(r =>
      String(r[(H['moteid']||1)-1])===moteId && String(r[(H['saksnr']||2)-1])===saksnr
    );

    const meta = filtered.find(r => String(r[(H['bruker']||3)-1])===LOCK_ROW_USER);
    const locked = !!(meta && (String(meta[(H['låst']||8)-1]).toLowerCase()==='true'));
    const vedtak = meta ? String(meta[(H['vedtak']||7)-1]||'') : '';

    let JA=0, NEI=0, BLANK=0, myVote=null;
    filtered.forEach(r=>{
      const bruker = String(r[(H['bruker']||3)-1]||'');
      const epost  = String(r[(H['epost']||4)-1]||'').toLowerCase();
      if (bruker===LOCK_ROW_USER) return;
      const v = String(r[(H['stemme']||5)-1]||'').toUpperCase();
      if (VOTES.includes(v)){
        if (v==='JA') JA++;
        else if (v==='NEI') NEI++;
        else BLANK++;
        if (epost===me) myVote = v;
      }
    });

    return {
      ok:true,
      moteId, saksnr,
      counts: { JA, NEI, BLANK },
      myVote,
      locked,
      vedtak: vedtak || null
    };
  };

  // ----------- Offentlig: lås vedtak -----------
  globalThis.voterLockDecision = function voterLockDecision(moteId, saksnr, vedtakTekst){
    if (!moteId || !saksnr) throw new Error('Mangler Møte-ID eller Saksnr.');
    _ensureSheets();
    const sh = SpreadsheetApp.getActive().getSheetByName(STEMME_SHEET);
    const H = _headerMap_(sh);
    const rows = _getRows_(sh);

    const now = _now();

    // Finn LOCK-rad
    let lockRow = rows.findIndex(r =>
      String(r[(H['moteid']||1)-1])===moteId &&
      String(r[(H['saksnr']||2)-1])===saksnr &&
      String(r[(H['bruker']||3)-1])===LOCK_ROW_USER
    );

    if (lockRow >= 0){
      const r = lockRow + 2;
      if (H['vedtak']) sh.getRange(r, H['vedtak']).setValue(String(vedtakTekst||'').trim());
      if (H['låst'])   sh.getRange(r, H['låst']).setValue(true);
      if (H['tid'])    sh.getRange(r, H['tid']).setValue(now);
    } else {
      // Legg til ny meta-rad
      const rec = [
        moteId,                // MoteId
        saksnr,                // Saksnr
        LOCK_ROW_USER,         // Bruker
        '',                    // Epost
        '',                    // Stemme
        now,                   // Tid
        String(vedtakTekst||'').trim(), // Vedtak
        true                   // Låst
      ];
      sh.appendRow(rec);
    }

    _log(`locked decision: ${moteId}/${saksnr}`);
    return { ok:true, message:'Vedtak låst.' };
  };

})();
