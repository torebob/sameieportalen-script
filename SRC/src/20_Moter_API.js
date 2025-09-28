/* ======================= Møter & Agenda (Komplett API) =======================
 * FILE: 20_Moter_API.gs | VERSION: 1.6.1 | UPDATED: 2025-09-14
 * FORMÅL: Komplett, høytytende backend for den avanserte møtemodulen.
 * Støtter sanntidspolling, indeksering, innspill og avstemming.
 * ============================================================================== */

(function (global) {
  // -------------------- KONFIG --------------------
  const PROPS = PropertiesService.getScriptProperties();

  const MEETINGS_SHEET = SHEETS.MOTER;
  const SAKER_SHEET    = SHEETS.MOTE_SAKER;
  const INNSPILL_SHEET = SHEETS.MOTE_KOMMENTARER;
  const STEMMER_SHEET  = SHEETS.MOTE_STEMMER;

  const MEETINGS_HEADERS = ['id','type','dato','start','slutt','sted','tittel','agenda','status','created_ts','updated_ts'];
  const SAKER_HEADERS = ['mote_id', 'sak_id', 'saksnr', 'tittel', 'forslagAv', 'gdprNote', 'bakgrunn', 'forslagVedtak', 'vedtak', 'status', 'ansvarlig', 'created_ts', 'updated_ts'];
  const INNSPILL_HEADERS = ['sak_id','ts','from','text'];
  const STEMMER_HEADERS = ['vote_id','sak_id','mote_id','email','name','vote','ts'];

  // -------------------- HELPERE --------------------
  function tz_() { return Session.getScriptTimeZone() || 'Europe/Oslo'; }
  function log_(topic, msg) { try { if (typeof _logEvent === 'function') _logEvent(topic, msg); } catch (_) {} }
  
  function ensureSheet_(name, headers) {
    if (typeof _ensureSheetWithHeaders_ === 'function') {
        return _ensureSheetWithHeaders_(name, headers);
    }
    const ss = SpreadsheetApp.getActive();
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }
  function getCurrentUser_() {
    const email = (Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail() || '').toLowerCase();
    let name = '';
    try {
      const boardSheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.STYRET);
      if (boardSheet && boardSheet.getLastRow() > 1) {
        const boardData = boardSheet.getRange(2, 1, boardSheet.getLastRow() - 1, 2).getValues();
        const match = boardData.find(row => String(row[1] || '').toLowerCase() === email);
        if (match) name = match[0];
      }
    } catch(e) {}
    return { email, name };
  }
  function getBoardMembers() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.STYRET);
    if (!sheet || sheet.getLastRow() < 2) return [];
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    return data.map(row => row[0]).filter(name => name && name.trim() !== '');
  } catch (e) {
    log_('Styremedlemmer_FEIL', e.message);
    return [];
  }
}
  function getMoteIdForSak_(sakId) {
    const sakerSheet = ensureSheet_(SAKER_SHEET, SAKER_HEADERS);
    const index = Indexer.get(SAKER_SHEET, SAKER_HEADERS, 'sak_id');
    const rowNum = index[sakId];
    if (!rowNum) return '';
    const cMoteId = SAKER_HEADERS.indexOf('mote_id');
    return sakerSheet.getRange(rowNum, cMoteId + 1).getValue();
  }

  // -------------------- RAD-INDEKSERING --------------------
  const Indexer = {
    getKey: (sheetName) => `IDX::${sheetName}`,
    get: function(sheetName, headers, idHeader) {
      const raw = PROPS.getProperty(this.getKey(sheetName));
      if (!raw) return this.rebuild(sheetName, headers, idHeader);
      try {
        const parsed = JSON.parse(raw);
        return (parsed && parsed.h === idHeader && typeof parsed.m === 'object') ? parsed.m : this.rebuild(sheetName, headers, idHeader);
      } catch (_) {
        return this.rebuild(sheetName, headers, idHeader);
      }
    },
    set: function(sheetName, idHeader, id, row) {
      const key = this.getKey(sheetName);
      const data = JSON.parse(PROPS.getProperty(key) || '{}');
      if (!data.h) data.h = idHeader;
      if (!data.m) data.m = {};
      data.m[id] = row;
      PROPS.setProperty(key, JSON.stringify(data));
    },
    rebuild: function(sheetName, headers, idHeader) {
      const sh = ensureSheet_(sheetName, headers);
      const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idCol = H.indexOf(idHeader);
      if (idCol < 0) throw new Error(`Fant ikke ID-kolonne '${idHeader}' i ${sheetName}`);
      
      const map = {};
      const last = sh.getLastRow();
      if (last > 1) {
        const ids = sh.getRange(2, idCol + 1, last - 1, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (ids[i][0]) map[ids[i][0]] = i + 2;
        }
      }
      PROPS.setProperty(this.getKey(sheetName), JSON.stringify({ h: idHeader, m: map }));
      log_('Indexer', `Indeks for ${sheetName} ble gjenoppbygd.`);
      return map;
    }
  };
  
  function getVoteIndex_(sakId){
    const key = `VOTEIDX::${sakId}`;
    const raw = PROPS.getProperty(key);
    if (raw) { try { return JSON.parse(raw) || {}; } catch(_){ return {}; } }
    const sh = ensureSheet_(STEMMER_SHEET, STEMMER_HEADERS);
    const H = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    const last = sh.getLastRow();
    const map = {};
    if (last > 1){
      const vals = sh.getRange(2,1,last-1,sh.getLastColumn()).getValues();
      const iS = H.indexOf('sak_id'), iE = H.indexOf('email');
      for (let i=0;i<vals.length;i++){
        if (String(vals[i][iS]) === String(sakId)) {
          const em = String(vals[i][iE]||'').toLowerCase();
          if (em) map[em] = i+2;
        }
      }
    }
    PROPS.setProperty(key, JSON.stringify(map));
    return map;
  }
  function setVoteIndexRow_(sakId, email, row){
    const key = `VOTEIDX::${sakId}`;
    const map = JSON.parse(PROPS.getProperty(key) || '{}');
    map[String(email).toLowerCase()] = row;
    PROPS.setProperty(key, JSON.stringify(map));
  }

  // -------------------- UI & API --------------------
  function openMeetingsUI() {
    const tpl = HtmlService.createTemplateFromFile('30_Moteoversikt.html');
    tpl.FILE = '30_Moteoversikt.html';
    tpl.APP_VERSION = APP.VERSION;
    tpl.TITLE = 'Møteoversikt & Protokoller';
    tpl.PURPOSE = 'UI for møter, agenda, innspill og vedtak.';
    
    const html = tpl.evaluate().setWidth(1100).setHeight(760);
    SpreadsheetApp.getUi().showModalDialog(html, 'Møteoversikt & Protokoller');
  }
  
  function uiBootstrap(){
    const { email, name } = getCurrentUser_();
    return { user: { email, name } };
  }

  // ==================== MØTER, SAKER, INNSPILL, STEMMING ====================
  function upsertMeeting(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (!payload?.tittel?.trim()) return { ok: false, message: 'Møtetittel er påkrevd' };
      
      const sh = ensureSheet_(MEETINGS_SHEET, MEETINGS_HEADERS);
      const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = Object.fromEntries(H.map((h, i) => [h, i]));
      const now = new Date();
      
      const index = Indexer.get(MEETINGS_SHEET, MEETINGS_HEADERS, 'id');
      const rowNum = payload.moteId ? index[payload.moteId] : null;

      let id = payload.moteId;
      if (!rowNum) {
        id = id || `M-${Utilities.formatDate(now, tz_(), 'yyyyMMdd-HHmmss')}`;
        const newRow = Array(H.length).fill('');
        newRow[idx.id] = id;
        newRow[idx.type] = payload.type || 'Styremøte';
        newRow[idx.dato] = new Date(payload.datoISO);
        newRow[idx.start] = payload.start || '';
        newRow[idx.slutt] = payload.slutt || '';
        newRow[idx.sted] = payload.sted || '';
        newRow[idx.tittel] = payload.tittel.trim();
        newRow[idx.agenda] = payload.agenda || '';
        newRow[idx.status] = 'Planlagt';
        newRow[idx.created_ts] = now;
        newRow[idx.updated_ts] = now;
        sh.appendRow(newRow);
        Indexer.set(MEETINGS_SHEET, 'id', id, sh.getLastRow());
      } else {
        const range = sh.getRange(rowNum, 1, 1, H.length);
        const cur = range.getValues()[0];
        cur[idx.type] = payload.type ?? cur[idx.type];
        cur[idx.dato] = new Date(payload.datoISO);
        cur[idx.start] = payload.start ?? cur[idx.start];
        cur[idx.slutt] = payload.slutt ?? cur[idx.slutt];
        cur[idx.sted] = payload.sted ?? cur[idx.sted];
        cur[idx.tittel] = payload.tittel.trim();
        cur[idx.agenda] = payload.agenda ?? cur[idx.agenda];
        cur[idx.updated_ts] = now;
        range.setValues([cur]);
        id = cur[idx.id];
      }
      log_('Møte', `Lagring OK (${id})`);
      return { ok: true, id, message: `Møte lagret (${id})` };
    } catch (e) {
      log_('Møte_FEIL', e.message);
      return { ok: false, message: e.message };
    } finally {
      lock.releaseLock();
    }
  }

  function listMeetings_(args){ 
    const scope = args?.scope || 'planned';
    const sh = ensureSheet_(MEETINGS_SHEET, MEETINGS_HEADERS);
    const last = sh.getLastRow();
    if (last < 2) return [];
    
    const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const i = Object.fromEntries(H.map((h, i) => [h, i]));

    const today = new Date(); today.setHours(0,0,0,0);

    return data.map(r => ({
      id: r[i.id], type: r[i.type] || 'Styremøte', dato: r[i.dato], start: r[i.start] || '',
      slutt: r[i.slutt] || '', sted: r[i.sted] || '', tittel: r[i.tittel] || '',
      agenda: r[i.agenda] || '', status: r[i.status] || 'Planlagt'
    }))
    .filter(m => m.status !== 'Slettet' && m.status !== 'Arkivert')
    .filter(m => {
      if (!m.dato) return scope === 'planned';
      const meetingDate = m.dato instanceof Date ? m.dato : new Date(m.dato);
      return scope === 'past' ? meetingDate < today : meetingDate >= today;
    })
    .sort((a,b) => (a.dato?.getTime() || 0) - (b.dato?.getTime() || 0));
  }

  function nextSaksnr_(moteId) {
    const year = new Date().getFullYear();
    const key = `SAKSSEQ::${moteId}::${year}`;
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      let seq = parseInt(PROPS.getProperty(key) || '0', 10);
      seq++;
      PROPS.setProperty(key, String(seq));
      return `S-${String(seq).padStart(3, '0')}${year}`;
    } finally {
      lock.releaseLock();
    }
  }

  function addAgendaItem(moteId, payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sakerSheet = ensureSheet_(SAKER_SHEET, SAKER_HEADERS);
      const H = SAKER_HEADERS;
      const idx = Object.fromEntries(H.map((h, i) => [h, i]));

      const sakId = `SAK-${Utilities.getUuid().slice(0, 8)}`;
      const saksnr = nextSaksnr_(moteId);
      const now = new Date();
      
      const newRow = Array(H.length).fill('');
      newRow[idx.mote_id] = moteId;
      newRow[idx.sak_id] = sakId;
      newRow[idx.saksnr] = saksnr;
      newRow[idx.tittel] = payload.tittel || '';
      newRow[idx.forslagAv] = payload.forslagAv || '';
      newRow[idx.gdprNote] = payload.gdprNote || '';
      newRow[idx.bakgrunn] = payload.bakgrunn || '';
      newRow[idx.forslagVedtak] = payload.forslagVedtak || '';
      newRow[idx.status] = 'Planlagt';
      newRow[idx.ansvarlig] = payload.ansvarlig || '';
      newRow[idx.created_ts] = now;
      newRow[idx.updated_ts] = now;

      sakerSheet.appendRow(newRow);
      Indexer.set(SAKER_SHEET, 'sak_id', sakId, sakerSheet.getLastRow());

      log_('Sak', `Ny sak ${sakId} (${saksnr}) for møte ${moteId}`);
      return { ok: true, sakId, saksnr };
    } catch(e) {
       log_('Sak_FEIL', `Add item: ${e.stack}`);
       return {ok: false, message: e.message};
    }
    finally {
      lock.releaseLock();
    }
  }

  function updateAgendaItem(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sakerSheet = ensureSheet_(SAKER_SHEET, SAKER_HEADERS);
      const H = SAKER_HEADERS;
      const idx = Object.fromEntries(H.map((h, i) => [h, i]));

      const index = Indexer.get(SAKER_SHEET, SAKER_HEADERS, 'sak_id');
      const rowNum = index[payload.sakId];
      if (!rowNum) return { ok: false, message: `Fant ikke sak ${payload.sakId}` };

      const range = sakerSheet.getRange(rowNum, 1, 1, H.length);
      const cur = range.getValues()[0];
      
      cur[idx.tittel] = payload.tittel ?? cur[idx.tittel];
      cur[idx.forslagAv] = payload.forslagAv ?? cur[idx.forslagAv];
      cur[idx.gdprNote] = payload.gdprNote ?? cur[idx.gdprNote];
      cur[idx.bakgrunn] = payload.bakgrunn ?? cur[idx.bakgrunn];
      cur[idx.forslagVedtak] = payload.forslagVedtak ?? cur[idx.forslagVedtak];
      cur[idx.ansvarlig] = payload.ansvarlig ?? cur[idx.ansvarlig];
      cur[idx.updated_ts] = new Date();
      
      range.setValues([cur]);
      return { ok: true, message: 'Sak lagret' };
    } catch(e) {
       log_('Sak_FEIL', `Update item: ${e.stack}`);
       return {ok: false, message: e.message};
    }
    finally {
      lock.releaseLock();
    }
  }

  function listAgenda(moteId) {
    const sakerSheet = ensureSheet_(SAKER_SHEET, SAKER_HEADERS);
    const last = sakerSheet.getLastRow();
    if (last < 2) return [];
    
    const data = sakerSheet.getRange(2, 1, last - 1, sakerSheet.getLastColumn()).getValues();
    const H = SAKER_HEADERS;
    const i = Object.fromEntries(H.map((h, i) => [h, i]));

    return data.filter(r => r[i.mote_id] === moteId)
      .map(r => ({
        moteId: r[i.mote_id],
        sakId: r[i.sak_id],
        saksnr: r[i.saksnr],
        tittel: r[i.tittel],
        forslagAv: r[i.forslagAv],
        gdprNote: r[i.gdprNote],
        bakgrunn: r[i.bakgrunn],
        forslagVedtak: r[i.forslagVedtak],
        vedtak: r[i.vedtak],
        status: r[i.status],
        ansvarlig: r[i.ansvarlig],
        updated_ts: r[i.updated_ts]
      }))
      .sort((a,b)=> String(a.saksnr).localeCompare(String(b.saksnr)));
  }
  
  function deleteAgendaItem(sakId, opts) {
    const cascade = opts?.cascade !== false;
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sakerSheet = ensureSheet_(SAKER_SHEET, SAKER_HEADERS);
      const index = Indexer.get(SAKER_SHEET, SAKER_HEADERS, 'sak_id');
      const rowNum = index[sakId];
      if(rowNum) {
        sakerSheet.deleteRow(rowNum);
        Indexer.rebuild(SAKER_SHEET, SAKER_HEADERS, 'sak_id');
      }
      
      if (cascade) {
        const innspillSheet = ensureSheet_(INNSPILL_SHEET, INNSPILL_HEADERS);
        if (innspillSheet.getLastRow() > 1) {
          let data = innspillSheet.getDataRange().getValues();
          const H = data.shift();
          const cSakId = H.indexOf('sak_id');
          const filteredData = data.filter(row => row[cSakId] !== sakId);
          innspillSheet.getRange(2, 1, innspillSheet.getLastRow() - 1, H.length).clearContent();
          if (filteredData.length > 0) {
            innspillSheet.getRange(2, 1, filteredData.length, H.length).setValues(filteredData);
          }
        }
      }
      return { ok: true, message: 'Sak slettet' };
    } finally {
      lock.releaseLock();
    }
  }

  function appendInnspill(sakId, text) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sh = ensureSheet_(INNSPILL_SHEET, INNSPILL_HEADERS);
      const email = getCurrentUser_().email;
      sh.appendRow([sakId, new Date(), email, text]);
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  }

  function listInnspill(sakId, sinceISO) {
    const sh = ensureSheet_(INNSPILL_SHEET, INNSPILL_HEADERS);
    if (sh.getLastRow() < 2) return [];
    
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const i = { sak_id:H.indexOf('sak_id'), ts:H.indexOf('ts'), from:H.indexOf('from'), text:H.indexOf('text') };
    const since = sinceISO ? new Date(sinceISO) : null;

    return data
      .filter(r => {
        const ts = r[i.ts] instanceof Date ? r[i.ts] : null;
        return r[i.sak_id] === sakId && (!since || (ts && ts > since));
      })
      .map(r => ({ sakId: r[i.sak_id], ts: r[i.ts], from: r[i.from], text: r[i.text] }))
      .sort((a,b) => (a.ts?.getTime() || 0) - (b.ts?.getTime() || 0));
  }

  function castVote(sakId, value){
    if (!sakId) throw new Error('Mangler sakId.');
    const v = String(value||'').toUpperCase();
    if (['JA','NEI','BLANK'].indexOf(v) === -1) throw new Error('Ugyldig stemme.');

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const { email, name } = getCurrentUser_();
      if (!email) throw new Error('Kunne ikke identifisere bruker.');
      const moteId = getMoteIdForSak_(sakId);
      if (!moteId) throw new Error(`Fant ikke tilhørende møte for sak ${sakId}`);
      
      const sh = ensureSheet_(STEMMER_SHEET, STEMMER_HEADERS);
      const voteIdx = getVoteIndex_(sakId);
      const rowNum = voteIdx[email.toLowerCase()];

      if (rowNum) {
        const cVote = STEMMER_HEADERS.indexOf('vote') + 1;
        const cTs = STEMMER_HEADERS.indexOf('ts') + 1;
        sh.getRange(rowNum, cVote).setValue(v);
        sh.getRange(rowNum, cTs).setValue(new Date());
      } else {
        const voteId = `V-${Utilities.getUuid().slice(0,8)}`;
        sh.appendRow([voteId, sakId, moteId, email, name, v, new Date()]);
        setVoteIndexRow_(sakId, email, sh.getLastRow());
      }
      log_('Stemme', `${email} -> ${sakId} = ${v}`);
      return { ok:true };
    } finally {
      lock.releaseLock();
    }
  }

  function getVoteSummary(sakId){
    if (!sakId) return {JA:0,NEI:0,BLANK:0};
    const sh = ensureSheet_(STEMMER_SHEET, STEMMER_HEADERS);
    if (sh.getLastRow() < 2) return {JA:0,NEI:0,BLANK:0};

    const data = sh.getDataRange().getValues();
    const H = data.shift();
    const iSak = H.indexOf('sak_id'), iVote = H.indexOf('vote');
    
    const summary = {JA:0,NEI:0,BLANK:0};
    data.forEach(row => {
      if (row[iSak] === sakId) {
        const vote = String(row[iVote] || '').toUpperCase();
        if (summary[vote] !== undefined) summary[vote]++;
      }
    });
    return summary;
  }

  function rtServerNow() { return { now: new Date().toISOString() }; }
  
  function rtGetChanges(moteId, sinceISO) {
    const serverNow = new Date().toISOString();
    const since = sinceISO ? new Date(sinceISO) : null;
    let meetingUpdated = null, updatedSaker = [], newInnspill = [];

    const moterSheet = ensureSheet_(MEETINGS_SHEET, MEETINGS_HEADERS);
    const moterH = moterSheet.getRange(1,1,1,moterSheet.getLastColumn()).getValues()[0];
    const iUpdated = moterH.indexOf('updated_ts');
    const moterIndex = Indexer.get(MEETINGS_SHEET, MEETINGS_HEADERS, 'id');
    const moteRowNum = moterIndex[moteId];
    if (moteRowNum) {
      const uts = moterSheet.getRange(moteRowNum, iUpdated + 1).getValue();
      if (uts instanceof Date && (!since || uts > since)) {
        meetingUpdated = listMeetings_({scope: 'all'}).find(m => m.id === moteId);
      }
    }

    const alleSaker = listAgenda(moteId);
    updatedSaker = alleSaker.filter(sak => sak.updated_ts && (!since || new Date(sak.updated_ts) > since));

    const sakIds = new Set(alleSaker.map(s => s.sakId));
    if (sakIds.size > 0) {
      newInnspill = [];
      for (const sakId of sakIds) {
        newInnspill.push(...listInnspill(sakId, sinceISO));
      }
    }
    return { serverNow, meetingUpdated, updatedSaker, newInnspill };
  }
  
  // -------------------- EKSPORT --------------------
  global.openMeetingsUI = openMeetingsUI;
  global.uiBootstrap = uiBootstrap;
  global.upsertMeeting = upsertMeeting;
  global.listMeetings_ = listMeetings_;
  global.addAgendaItem = addAgendaItem;
  global.updateAgendaItem = updateAgendaItem;
  global.listAgenda = listAgenda;
  global.getBoardMembers = getBoardMembers;
  global.deleteAgendaItem = deleteAgendaItem;
  global.appendInnspill = appendInnspill;
  global.listInnspill = listInnspill;
  global.castVote = castVote;
  global.getVoteSummary = getVoteSummary;
  global.rtServerNow = rtServerNow;
  global.rtGetChanges = rtGetChanges;

})(this);

