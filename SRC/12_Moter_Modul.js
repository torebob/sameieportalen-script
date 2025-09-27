/* ======================= Møter & Agenda (Komplett UI backend) =======================
 * FILE: 12_Moter_Modul.gs | VERSION: 1.4.2 | UPDATED: 2025-09-14
 * Fokus: HØY YTELSE + “nesten sanntid” støtte for HTML-klienten
 * - Indekser (ID -> radnr) i ScriptProperties for O(1) oppslag
 * - Én lesing – modifiser i minnet – én skriving
 * - Låsing på alle skriv (LockService)
 * - Kompatibel med HTML: uiBootstrap, listMeetings_, newAgendaItem, saveAgenda,
 *   listAgenda, appendInnspill, listInnspill, rtServerNow, rtGetChanges
 * - Beholder alias: moterSaveFromUI(payload) -> upsertMeeting(payload)
 * ================================================================================== */

(function () {
  // -------------------- KONFIG --------------------
  const PROPS = PropertiesService.getScriptProperties();

  const SHEETNAMES =
    (typeof SHEETS !== 'undefined' && SHEETS) || {
      MOTER: 'Møter',
      MØTER: 'Møter',
      MEETINGS: 'Møter',

      MOTE_SAKER: 'MøteSaker',
      MØTE_SAKER: 'MøteSaker',

      MOTE_KOMMENTARER: 'MøteSakKommentarer',
      MØTE_KOMMENTARER: 'MøteSakKommentarer',
    };

  const MEETINGS_SHEET = SHEETNAMES.MOTER || SHEETNAMES.MØTER || SHEETNAMES.MEETINGS;
  const SAKER_SHEET    = SHEETNAMES.MOTE_SAKER || SHEETNAMES.MØTE_SAKER || 'MøteSaker';
  const INNSPILL_SHEET = SHEETNAMES.MOTE_KOMMENTARER || SHEETNAMES.MØTE_KOMMENTARER || 'MøteSakKommentarer';

  const MEETINGS_HEADERS = [
    'id','type','dato','start','slutt','sted','tittel','agenda','status','created_ts','updated_ts'
  ];
  const SAKER_HEADERS = [
    'mote_id','sak_id','saksnr','tittel','forslag','vedtak','created_ts','updated_ts'
  ];
  const INNSPILL_HEADERS = ['sak_id','ts','from','text'];

  // -------------------- HELPERE --------------------
  function tz_() { return Session.getScriptTimeZone() || 'Europe/Oslo'; }
  function pad(n, w) { n = String(n); while (n.length < w) n = '0' + n; return n; }
  function log_(topic, msg, extra) { try { if (typeof _logEvent === 'function') _logEvent(topic, msg, extra||{}); } catch(_){} }

  // Godtar Date, ISO, og dd.mm.yyyy -> Date | '' (falsy)
  function parseISODate_(val) {
    if (!val) return '';
    try {
      let d = val instanceof Date ? val : new Date(val);
      if (isNaN(d) && typeof val === 'string') {
        const m = val.trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
        if (m) d = new Date(`${m[3]}-${('0'+m[2]).slice(-2)}-${('0'+m[1]).slice(-2)}T00:00:00`);
      }
      return isNaN(d) ? '' : d;
    } catch (_) { return ''; }
  }

  // Finn/lag ark med riktige headere
  function ensureSheetWithHeaders_(name, headers) {
    if (typeof _ensureSheetWithHeaders_ === 'function') {
      return _ensureSheetWithHeaders_(name, headers);
    }
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    } else {
      const first = sh.getRange(1,1,1,headers.length).getValues()[0];
      let mismatch = false;
      for (let i=0;i<headers.length;i++){
        if (String(first[i]||'') !== headers[i]) { mismatch = true; break; }
      }
      if (mismatch) {
        sh.clearContents();
        sh.getRange(1,1,1,headers.length).setValues([headers]);
        sh.setFrozenRows(1);
      }
    }
    return sh;
  }

  // -------------------- RAD-INDEKSER (ID -> radnr) --------------------
  // PROPS key: IDX::<sheetName> => {"h":"idHeader","m":{"<id>":<rowNumber>}}
  function idxKey_(sheetName){ return 'IDX::' + sheetName; }

  function buildIndex_(sheetName, headers, idHeader) {
    const sh = ensureSheetWithHeaders_(sheetName, headers);
    const H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const idCol = H.indexOf(idHeader);
    if (idCol < 0) throw new Error(`Fant ikke kolonne '${idHeader}' i ${sheetName}`);
    const last = sh.getLastRow();
    const map = {};
    if (last > 1) {
      const ids = sh.getRange(2, idCol+1, last-1, 1).getValues();
      for (let i=0;i<ids.length;i++){
        const id = String(ids[i][0]||'').trim();
        if (id) map[id] = i+2; // 1-basert rad
      }
    }
    PROPS.setProperty(idxKey_(sheetName), JSON.stringify({h:idHeader,m:map}));
    return map;
  }

  function getIndex_(sheetName, headers, idHeader) {
    const raw = PROPS.getProperty(idxKey_(sheetName));
    if (!raw) return buildIndex_(sheetName, headers, idHeader);
    try {
      const parsed = JSON.parse(raw);
      if (!parsed || parsed.h !== idHeader || typeof parsed.m !== 'object') {
        return buildIndex_(sheetName, headers, idHeader);
      }
      return parsed.m || {};
    } catch(_) {
      return buildIndex_(sheetName, headers, idHeader);
    }
  }

  function indexSet_(sheetName, headers, idHeader, id, row){
    const key = idxKey_(sheetName);
    const parsed = JSON.parse(PROPS.getProperty(key) || '{"h":"","m":{}}');
    if (!parsed.h) parsed.h = idHeader;
    if (!parsed.m) parsed.m = {};
    parsed.m[String(id)] = row;
    PROPS.setProperty(key, JSON.stringify(parsed));
  }

  function indexRebuild_(sheetName, headers, idHeader){
    buildIndex_(sheetName, headers, idHeader);
  }

  // -------------------- UI BOOTSTRAP / ÅPNER --------------------
  function uiBootstrap(){
    const email =
      (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) || '';
    return { user: { email } };
  }

  /*
   * MERK: openMeetingsUI() er fjernet fra denne filen for å unngå konflikter.
   * Funksjonen kalles nå fra 00_App_Core.js, som bruker den sentrale UI_FILES-mappingen.
   */

  // ==================== MØTER ====================
  /**
   * Opprett/oppdater møte (ytelsesoptimalisert).
   * payload = {moteId?, type, datoISO, start, slutt, sted, tittel, agenda}
   */
  function upsertMeeting(payload){
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(15000);

      // Valider
      if (!payload || typeof payload !== 'object') return { ok:false, message:'Ugyldig payload' };
      const title = (payload.tittel || '').trim();
      if (!title) return { ok:false, message:'Møtetittel er påkrevd' };
      const dato = parseISODate_(payload.datoISO);
      if (!dato) return { ok:false, message:'Ugyldig/manglende møtedato' };

      const sh = ensureSheetWithHeaders_(MEETINGS_SHEET, MEETINGS_HEADERS);
      const lastCol = sh.getLastColumn();
      const H = sh.getRange(1,1,1,lastCol).getValues()[0];
      const idx = {
        id:H.indexOf('id'), type:H.indexOf('type'), dato:H.indexOf('dato'),
        start:H.indexOf('start'), slutt:H.indexOf('slutt'),
        sted:H.indexOf('sted'), tittel:H.indexOf('tittel'), agenda:H.indexOf('agenda'),
        status:H.indexOf('status'), created:H.indexOf('created_ts'), updated:H.indexOf('updated_ts')
      };

      const now = new Date();
      const index = getIndex_(MEETINGS_SHEET, MEETINGS_HEADERS, 'id');
      let row = payload.moteId ? index[String(payload.moteId)] : null;

      let id = payload.moteId;
      if (!row) {
        // NY rad -> append én gang
        id = id || ('M-' + Utilities.formatDate(now, tz_(), 'yyyyMMdd-HHmmss'));
        const newRow = Array(lastCol).fill('');
        newRow[idx.id] = id;
        newRow[idx.type] = payload.type || 'Styremøte';
        newRow[idx.dato] = dato;
        newRow[idx.start] = payload.start || '';
        newRow[idx.slutt] = payload.slutt || '';
        newRow[idx.sted] = payload.sted || '';
        newRow[idx.tittel] = title;
        newRow[idx.agenda] = payload.agenda || '';
        newRow[idx.status] = 'Planlagt';
        newRow[idx.created] = now;
        newRow[idx.updated] = now;
        sh.appendRow(newRow);
        indexSet_(MEETINGS_SHEET, MEETINGS_HEADERS, 'id', id, sh.getLastRow());
      } else {
        // OPPDATER -> les én rad, skriv én rad
        const range = sh.getRange(row, 1, 1, lastCol);
        const cur = range.getValues()[0];
        cur[idx.type]   = payload.type   || cur[idx.type] || 'Styremøte';
        cur[idx.dato]   = dato           || cur[idx.dato];
        cur[idx.start]  = payload.start  || cur[idx.start];
        cur[idx.slutt]  = payload.slutt  || cur[idx.slutt];
        cur[idx.sted]   = payload.sted   || cur[idx.sted];
        cur[idx.tittel] = title          || cur[idx.tittel];
        cur[idx.agenda] = payload.agenda || cur[idx.agenda];
        cur[idx.updated]= now;
        range.setValues([cur]);
        id = cur[idx.id] || id;
      }

      log_('Møte', `Lagring OK (${id})`);
      return { ok:true, id, message:`Møte lagret (${id})` };
    } catch (e) {
      log_('Møte_FEIL', e.message, { stack:e.stack });
      return { ok:false, message:e.message };
    } finally {
      try { lock.releaseLock(); } catch(_){}
    }
  }

  /** Liste møter (planlagte/avholdte) – les én gang, filtrer i minnet. */
  function listMeetings_(args){
    const scope = (args && args.scope) || 'planned';
    const sh = ensureSheetWithHeaders_(MEETINGS_SHEET, MEETINGS_HEADERS);
    const last = sh.getLastRow();
    if (last < 2) return [];
    const H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const data = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();

    const i = {
      id:H.indexOf('id'), type:H.indexOf('type'), dato:H.indexOf('dato'),
      start:H.indexOf('start'), slutt:H.indexOf('slutt'),
      sted:H.indexOf('sted'), tittel:H.indexOf('tittel'), status:H.indexOf('status')
    };

    const today = new Date(); today.setHours(0,0,0,0);

    return data.map(r => {
      const d = (r[i.dato] instanceof Date) ? r[i.dato] : (r[i.dato] ? new Date(r[i.dato]) : '');
      return {
        id:r[i.id],
        type:r[i.type] || 'Styremøte',
        dato:d,
        iso: (d && d instanceof Date) ? Utilities.formatDate(d, tz_(), 'yyyy-MM-dd') : '',
        start:r[i.start] || '',
        slutt:r[i.slutt] || '',
        sted:r[i.sted] || '',
        tittel:r[i.tittel] || '',
        status:r[i.status] || 'Planlagt'
      };
    })
    .filter(m => m.status !== 'Slettet' && m.status !== 'Arkivert')
    .filter(m => {
      if (!m.dato) return scope === 'planned';
      return scope === 'past' ? m.dato < today : m.dato >= today;
    })
    .sort((a,b)=>{
      const av = a.dato ? a.dato.getTime() : 0;
      const bv = b.dato ? b.dato.getTime() : 0;
      return av - bv;
    });
  }

  // ==================== AGENDA / SAKER ====================
  /**
   * Genererer saksnummer pr møte og år.
   * Persistens i ScriptProperties for O(1).
   * Fallback (én gang): hvis teller mangler, skanner for å finne max.
   */
  function nextSaksnr_(moteId){
    const year = new Date().getFullYear();
    const key = `SAKSSEQ::${moteId}::${year}`;
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try{
      let seq = parseInt(PROPS.getProperty(key) || '0', 10);
      if (!seq) {
        // Fallback: finn max eksisterende i arket
        const sh = ensureSheetWithHeaders_(SAKER_SHEET, SAKER_HEADERS);
        const last = sh.getLastRow();
        if (last > 1) {
          const H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
          const all = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
          const i_mid = H.indexOf('mote_id');
          const i_snr = H.indexOf('saksnr');
          let max = 0;
          for (let i=0;i<all.length;i++){
            if (String(all[i][i_mid]) !== String(moteId)) continue;
            const sn = String(all[i][i_snr] || '');
            const yy = sn.slice(-4);
            if (yy === String(year)) {
              const n = parseInt(sn.slice(2,5), 10);
              if (!isNaN(n)) max = Math.max(max, n);
            }
          }
          seq = max;
        }
      }
      seq++;
      PROPS.setProperty(key, String(seq));
      return `S-${pad(seq,3)}${year}`;
    } finally {
      try{ lock.releaseLock(); }catch(_){}
    }
  }

  /** Opprett en ny sak for et møte. */
  function newAgendaItem(moteId){
    const lock = LockService.getScriptLock();
    try{
      lock.waitLock(15000);

      const saker = ensureSheetWithHeaders_(SAKER_SHEET, SAKER_HEADERS);
      const lastCol = saker.getLastColumn();
      const H = saker.getRange(1,1,1,lastCol).getValues()[0];
      const idx = {
        mote_id:H.indexOf('mote_id'), sak_id:H.indexOf('sak_id'),
        saksnr:H.indexOf('saksnr'), tittel:H.indexOf('tittel'),
        created:H.indexOf('created_ts'), updated:H.indexOf('updated_ts')
      };

      const sakId = 'SAK-' + Utilities.getUuid().slice(0,8);
      const saksnr = nextSaksnr_(moteId);
      const now = new Date();

      const row = Array(lastCol).fill('');
      row[idx.mote_id] = moteId;
      row[idx.sak_id] = sakId;
      row[idx.saksnr] = saksnr;
      row[idx.tittel] = '';
      row[idx.created] = now;
      row[idx.updated] = now;

      saker.appendRow(row);
      indexSet_(SAKER_SHEET, SAKER_HEADERS, 'sak_id', sakId, saker.getLastRow());

      log_('Sak', `Ny sak ${sakId} (${saksnr}) for møte ${moteId}`);
      return { ok:true, sakId, saksnr };
    } catch(e){
      log_('Sak_FEIL', e.message, { stack:e.stack });
      return { ok:false, message:e.message };
    } finally {
      try{ lock.releaseLock(); }catch(_){}
    }
  }

  /** Lagre endringer på en sak. payload = {sakId, tittel?, forslag?, vedtak?} */
  function saveAgenda(payload){
    const lock = LockService.getScriptLock();
    try{
      lock.waitLock(15000);

      const saker = ensureSheetWithHeaders_(SAKER_SHEET, SAKER_HEADERS);
      const lastCol = saker.getLastColumn();
      const H = saker.getRange(1,1,1,lastCol).getValues()[0];
      const idx = {
        sak_id:H.indexOf('sak_id'),
        tittel:H.indexOf('tittel'),
        forslag:H.indexOf('forslag'),
        vedtak:H.indexOf('vedtak'),
        updated:H.indexOf('updated_ts')
      };

      const index = getIndex_(SAKER_SHEET, SAKER_HEADERS, 'sak_id');
      const row = index[String(payload.sakId)];
      if (!row) return { ok:false, message:'Fant ikke sak ' + payload.sakId };

      const range = saker.getRange(row,1,1,lastCol);
      const cur = range.getValues()[0];

      if (payload.tittel  != null) cur[idx.tittel]  = payload.tittel;
      if (payload.forslag != null) cur[idx.forslag] = payload.forslag;
      if (payload.vedtak  != null) cur[idx.vedtak]  = payload.vedtak;
      cur[idx.updated] = new Date();

      range.setValues([cur]);
      return { ok:true, message:'Sak lagret' };
    } catch(e){
      log_('Sak_FEIL', e.message, { stack:e.stack });
      return { ok:false, message:e.message };
    } finally {
      try{ lock.releaseLock(); }catch(_){}
    }
  }

  /** Hent alle saker til et møte (inkl. updated_ts for polling). */
  function listAgenda(moteId){
    const saker = ensureSheetWithHeaders_(SAKER_SHEET, SAKER_HEADERS);
    const last = saker.getLastRow();
    if (last < 2) return [];
    const H = saker.getRange(1,1,1,saker.getLastColumn()).getValues()[0];
    const data = saker.getRange(2,1,last-1,saker.getLastColumn()).getValues();
    const idx = {
      mote_id:H.indexOf('mote_id'), sak_id:H.indexOf('sak_id'),
      saksnr:H.indexOf('saksnr'), tittel:H.indexOf('tittel'),
      forslag:H.indexOf('forslag'), vedtak:H.indexOf('vedtak'),
      updated:H.indexOf('updated_ts')
    };
    return data.filter(r => String(r[idx.mote_id]) === String(moteId))
      .map(r => ({
        moteId:r[idx.mote_id],
        sakId:r[idx.sak_id],
        saksnr:r[idx.saksnr],
        tittel:r[idx.tittel],
        forslag:r[idx.forslag],
        vedtak:r[idx.vedtak],
        updated_ts:r[idx.updated] || null
      }))
      .sort((a,b)=> String(a.saksnr).localeCompare(String(b.saksnr)));
  }

  /** Legg til innspill (fra innlogget bruker) på en sak – append én rad. */
  function appendInnspill(sakId, text){
    const sh = ensureSheetWithHeaders_(INNSPILL_SHEET, INNSPILL_HEADERS);
    const email =
      (Session.getActiveUser() && Session.getActiveUser().getEmail()) ||
      (Session.getEffectiveUser() && Session.getEffectiveUser().getEmail()) || '';
    sh.appendRow([sakId, new Date(), email, text]);
    return { ok:true };
  }

  /** Hent innspill (valgfritt filtrert “siden” tidsstempel). */
  function listInnspill(sakId, sinceISO){
    const sh = ensureSheetWithHeaders_(INNSPILL_SHEET, INNSPILL_HEADERS);
    if (sh.getLastRow() < 2) return [];
    const H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const data = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
    const i = { sak_id:H.indexOf('sak_id'), ts:H.indexOf('ts'), from:H.indexOf('from'), text:H.indexOf('text') };
    const since = sinceISO ? new Date(sinceISO) : null;

    const out = [];
    for (const r of data){
      if (String(r[i.sak_id]) !== String(sakId)) continue;
      const ts = r[i.ts] instanceof Date ? r[i.ts] : (r[i.ts] ? new Date(r[i.ts]) : null);
      if (since && ts && ts <= since) continue;
      out.push({ sakId: r[i.sak_id], ts: ts, from: r[i.from], text: r[i.text] });
    }
    out.sort((a,b) => (a.ts && b.ts) ? (a.ts.getTime() - b.ts.getTime()) : 0);
    return out;
  }

  // ==================== SLETTING (valgfritt) ====================
  function deleteAgendaItem(sakId, opts){
    const cascade = !opts || opts.cascade !== false;
    const saker = ensureSheetWithHeaders_(SAKER_SHEET, SAKER_HEADERS);
    const index = getIndex_(SAKER_SHEET, SAKER_HEADERS, 'sak_id');
    const row = index[String(sakId)];
    if (row) {
      saker.deleteRow(row);
      indexRebuild_(SAKER_SHEET, SAKER_HEADERS, 'sak_id'); // tryggest etter sletting
    }

    if (cascade){
      const inn = ensureSheetWithHeaders_(INNSPILL_SHEET, INNSPILL_HEADERS);
      const H = inn.getRange(1,1,1,inn.getLastColumn()).getValues()[0];
      const data = inn.getRange(2,1,Math.max(0,inn.getLastRow()-1), inn.getLastColumn()).getValues();
      const i_ref = H.indexOf('sak_id');
      const toDelete = [];
      for (let i=0;i<data.length;i++) if (String(data[i][i_ref]) === String(sakId)) toDelete.push(i+2);
      toDelete.sort((a,b)=>b-a).forEach(r=>inn.deleteRow(r));
    }
    return { ok:true, message:'Sak slettet' };
  }

  function deleteMeeting(moteId, opts){
    const cascade = !opts || opts.cascade !== false;
    const moter = ensureSheetWithHeaders_(MEETINGS_SHEET, MEETINGS_HEADERS);
    const idxM = getIndex_(MEETINGS_SHEET, MEETINGS_HEADERS, 'id');
    const rowM = idxM[String(moteId)];
    if (rowM) {
      moter.deleteRow(rowM);
      indexRebuild_(MEETINGS_SHEET, MEETINGS_HEADERS, 'id');
    }

    if (cascade){
      const saker = ensureSheetWithHeaders_(SAKER_SHEET, SAKER_HEADERS);
      if (saker.getLastRow()>1){
        const Hs = saker.getRange(1,1,1,saker.getLastColumn()).getValues()[0];
        const ds = saker.getRange(2,1,Math.max(0,saker.getLastRow()-1), saker.getLastColumn()).getValues();
        const i_mid = Hs.indexOf('mote_id');
        const i_sid = Hs.indexOf('sak_id');

        const sakRows = [], sakIds = [];
        for (let i=0;i<ds.length;i++){
          if (String(ds[i][i_mid]) === String(moteId)) { sakRows.push(i+2); sakIds.push(ds[i][i_sid]); }
        }
        sakRows.sort((a,b)=>b-a).forEach(r=>saker.deleteRow(r));
        indexRebuild_(SAKER_SHEET, SAKER_HEADERS, 'sak_id');

        if (sakIds.length){
          const inn = ensureSheetWithHeaders_(INNSPILL_SHEET, INNSPILL_HEADERS);
          if (inn.getLastRow()>1){
            const Hi = inn.getRange(1,1,1,inn.getLastColumn()).getValues()[0];
            const di = inn.getRange(2,1,Math.max(0,inn.getLastRow()-1), inn.getLastColumn()).getValues();
            const i_ref = Hi.indexOf('sak_id');
            const rows = [];
            for (let i=0;i<di.length;i++){
              if (sakIds.indexOf(di[i][i_ref]) !== -1) rows.push(i+2);
            }
            rows.sort((a,b)=>b-a).forEach(r=>inn.deleteRow(r));
          }
        }
      }
    }
    return { ok:true, message:'Møte slettet' };
  }

  // ==================== “NESTEN SANNTID” (for HTML-klienten) ====================
  function rtServerNow(){
    return { now: new Date().toISOString() };
  }

  /**
   * Returnerer endringer siden `sinceISO`:
   * {
   *   serverNow: ISO,
   *   meetingUpdated: {id, updated_ts}?,
   *   updatedSaker: [{sakId, ...}]?,
   *   newInnspill: [{sakId, ts, from, text}]?
   * }
   */
  function rtGetChanges(moteId, sinceISO){
    const serverNow = new Date().toISOString();
    const since = sinceISO ? new Date(sinceISO) : null;
    let meetingUpdated = null, updatedSaker = [], newInnspill = [];

    // 1) Møte endret?
    const moterSheet = ensureSheetWithHeaders_(MEETINGS_SHEET, MEETINGS_HEADERS);
    const moterH = moterSheet.getRange(1,1,1,moterSheet.getLastColumn()).getValues()[0];
    const iUpdated = moterH.indexOf('updated_ts');
    const moterIndex = getIndex_(MEETINGS_SHEET, MEETINGS_HEADERS, 'id');
    const moteRowNum = moterIndex[moteId];
    if (moteRowNum) {
      const uts = moterSheet.getRange(moteRowNum, iUpdated+1).getValue();
      if (uts instanceof Date && (!since || uts > since)) {
        meetingUpdated = { id: moteId, updated_ts: uts };
      }
    }

    // 2) Saker oppdatert?
    const alleSaker = listAgenda(moteId); // inkluderer updated_ts
    updatedSaker = alleSaker.filter(sak => sak.updated_ts && (!since || new Date(sak.updated_ts) > since));

    // 3) Nye innspill på noen av sakene?
    const sakIds = new Set(alleSaker.map(s => s.sakId));
    if (sakIds.size > 0) {
      const inn = ensureSheetWithHeaders_(INNSPILL_SHEET, INNSPILL_HEADERS);
      if (inn.getLastRow() > 1) {
        const data = inn.getDataRange().getValues();
        const H = data.shift();
        const i_sid = H.indexOf('sak_id');
        const i_ts  = H.indexOf('ts');
        const i_from= H.indexOf('from');
        const i_txt = H.indexOf('text');

        newInnspill = data.filter(r => {
          const ts = r[i_ts] instanceof Date ? r[i_ts] : null;
          return sakIds.has(r[i_sid]) && ts && (!since || ts > since);
        }).map(r => ({ sakId: r[i_sid], ts: r[i_ts], from: r[i_from], text: r[i_txt] }));
      }
    }

    return { serverNow, meetingUpdated, updatedSaker, newInnspill };
  }

  // -------------------- UI-ALIAS --------------------
  function moterSaveFromUI(payload){ return upsertMeeting(payload); }

  // -------------------- EKSPORT --------------------
  this.uiBootstrap      = uiBootstrap;
  this.openMeetingsUI   = openMeetingsUI;

  this.upsertMeeting    = upsertMeeting;
  this.listMeetings_    = listMeetings_;

  this.newAgendaItem    = newAgendaItem;
  this.saveAgenda       = saveAgenda;
  this.listAgenda       = listAgenda;

  this.appendInnspill   = appendInnspill;
  this.listInnspill     = listInnspill;

  this.deleteAgendaItem = deleteAgendaItem;
  this.deleteMeeting    = deleteMeeting;

  this.rtServerNow      = rtServerNow;
  this.rtGetChanges     = rtGetChanges;

  this.moterSaveFromUI  = moterSaveFromUI;
})();
