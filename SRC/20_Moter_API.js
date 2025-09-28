/* ======================= Møter & Agenda (Komplett API) =======================
 * FILE: 20_Moter_API.gs | VERSION: 2.0.0 | UPDATED: 2025-09-26
 * FORMÅL: Komplett, høytytende backend for den avanserte møtemodulen.
 * - Modernisert med let/const, arrow functions, og forbedret lesbarhet.
 * - Støtter sanntidspolling, indeksering, innspill og avstemming.
 * ============================================================================== */

((global) => {
  const PROPS = PropertiesService.getScriptProperties();
  const { MOTER, MOTE_SAKER, MOTE_KOMMENTARER, MOTE_STEMMER, BOARD } = SHEETS;

  const MEETINGS_HEADERS = ['id', 'type', 'dato', 'start', 'slutt', 'sted', 'tittel', 'agenda', 'status', 'created_ts', 'updated_ts', 'participants'];
  const SAKER_HEADERS = ['mote_id', 'sak_id', 'saksnr', 'tittel', 'forslag', 'vedtak', 'created_ts', 'updated_ts'];
  const INNSPILL_HEADERS = ['sak_id', 'ts', 'from', 'text'];
  const STEMMER_HEADERS = ['vote_id', 'sak_id', 'mote_id', 'email', 'name', 'vote', 'ts'];

  const _tz_ = () => Session.getScriptTimeZone() || 'Europe/Oslo';
  const _log_ = (topic, msg) => {
    try {
      if (typeof _logEvent === 'function') _logEvent(topic, msg);
    } catch (e) { /* ignore */ }
  };

  const _ensureSheet_ = (name, headers) => {
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
  };

  const _getCurrentUser_ = () => {
    const email = (Session.getActiveUser()?.getEmail() || Session.getEffectiveUser()?.getEmail() || '').toLowerCase();
    let name = '';
    try {
      const boardSheet = SpreadsheetApp.getActive().getSheetByName(BOARD);
      if (boardSheet && boardSheet.getLastRow() > 1) {
        const boardData = boardSheet.getRange(2, 1, boardSheet.getLastRow() - 1, 2).getValues();
        const match = boardData.find(row => String(row[1] || '').toLowerCase() === email);
        if (match) name = match[0];
      }
    } catch (e) { /* ignore */ }
    return { email, name };
  };

  const getBoardMembers = () => {
    try {
      const boardSheet = SpreadsheetApp.getActive().getSheetByName(BOARD);
      if (!boardSheet) return [];
      const lastRow = boardSheet.getLastRow();
      if (lastRow < 2) return [];
      // Assumes Name in col 1, Email in col 2
      const data = boardSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      return data.map(row => ({ name: row[0], email: row[1] })).filter(p => p.email);
    } catch (e) {
      _log_('Styremedlemmer_FEIL', e.message);
      return [];
    }
  };

  const getMoteIdForSak_ = (sakId) => {
    const sakerSheet = _ensureSheet_(MOTE_SAKER, SAKER_HEADERS);
    const index = Indexer.get(MOTE_SAKER, SAKER_HEADERS, 'sak_id');
    const rowNum = index[sakId];
    if (!rowNum) return '';
    const cMoteId = SAKER_HEADERS.indexOf('mote_id');
    return sakerSheet.getRange(rowNum, cMoteId + 1).getValue();
  };

  const Indexer = {
    getKey: (sheetName) => `IDX::${sheetName}`,
    get(sheetName, headers, idHeader) {
      const raw = PROPS.getProperty(this.getKey(sheetName));
      if (!raw) return this.rebuild(sheetName, headers, idHeader);
      try {
        const parsed = JSON.parse(raw);
        return (parsed?.h === idHeader && typeof parsed.m === 'object') ? parsed.m : this.rebuild(sheetName, headers, idHeader);
      } catch (e) {
        return this.rebuild(sheetName, headers, idHeader);
      }
    },
    set(sheetName, idHeader, id, row) {
      const key = this.getKey(sheetName);
      const data = JSON.parse(PROPS.getProperty(key) || '{}');
      data.h = data.h || idHeader;
      data.m = data.m || {};
      data.m[id] = row;
      PROPS.setProperty(key, JSON.stringify(data));
    },
    rebuild(sheetName, headers, idHeader) {
      const sh = _ensureSheet_(sheetName, headers);
      const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idCol = H.indexOf(idHeader);
      if (idCol < 0) throw new Error(`Fant ikke ID-kolonne '${idHeader}' i ${sheetName}`);

      const map = {};
      const last = sh.getLastRow();
      if (last > 1) {
        const ids = sh.getRange(2, idCol + 1, last - 1, 1).getValues();
        ids.forEach((id, i) => {
          if (id[0]) map[id[0]] = i + 2;
        });
      }
      PROPS.setProperty(this.getKey(sheetName), JSON.stringify({ h: idHeader, m: map }));
      _log_('Indexer', `Indeks for ${sheetName} ble gjenoppbygd.`);
      return map;
    }
  };

  const getVoteIndex_ = (sakId) => {
    const key = `VOTEIDX::${sakId}`;
    const raw = PROPS.getProperty(key);
    if (raw) {
      try { return JSON.parse(raw) || {}; } catch(e){ return {}; }
    }
    const sh = _ensureSheet_(MOTE_STEMMER, STEMMER_HEADERS);
    const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const last = sh.getLastRow();
    const map = {};
    if (last > 1){
      const vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
      const iS = H.indexOf('sak_id'), iE = H.indexOf('email');
      vals.forEach((row, i) => {
        if (String(row[iS]) === String(sakId)) {
          const em = String(row[iE] || '').toLowerCase();
          if (em) map[em] = i + 2;
        }
      });
    }
    PROPS.setProperty(key, JSON.stringify(map));
    return map;
  };

  const setVoteIndexRow_ = (sakId, email, row) => {
    const key = `VOTEIDX::${sakId}`;
    const map = JSON.parse(PROPS.getProperty(key) || '{}');
    map[String(email).toLowerCase()] = row;
    PROPS.setProperty(key, JSON.stringify(map));
  };

  const uiBootstrap = () => {
    const { email, name } = _getCurrentUser_();
    return { user: { email, name } };
  };

  function upsertMeeting(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      requirePermission('MANAGE_MEETINGS', 'Administrere møter');
      if (!payload?.tittel?.trim()) return { ok: false, message: 'Møtetittel er påkrevd' };
      
      const sh = _ensureSheet_(MOTER, MEETINGS_HEADERS);
      const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const idx = H.reduce((acc, h, i) => ({ ...acc, [h]: i }), {});
      const now = new Date();
      
      const index = Indexer.get(MOTER, MEETINGS_HEADERS, 'id');
      const rowNum = payload.moteId ? index[payload.moteId] : null;

      let id = payload.moteId;
      if (!rowNum) {
        id = id || `M-${Utilities.formatDate(now, _tz_(), 'yyyyMMdd-HHmmss')}`;
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
        newRow[idx.participants] = Array.isArray(payload.participants) ? payload.participants.join(',') : '';
        sh.appendRow(newRow);
        Indexer.set(MOTER, 'id', id, sh.getLastRow());
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
        if (payload.participants !== undefined) {
          cur[idx.participants] = Array.isArray(payload.participants) ? payload.participants.join(',') : '';
        }
        range.setValues([cur]);
        id = cur[idx.id];
      }
      _log_('Møte', `Lagring OK (${id})`);
      return { ok: true, id, message: `Møte lagret (${id})` };
    } catch (e) {
      _log_('Møte_FEIL', e.message);
      return { ok: false, message: e.message };
    } finally {
      lock.releaseLock();
    }
  }

  const listMeetings_ = (args) => {
    const scope = args?.scope || 'planned';
    const sh = _ensureSheet_(MOTER, MEETINGS_HEADERS);
    const last = sh.getLastRow();
    if (last < 2) return [];
    
    const data = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
    const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const i = H.reduce((acc, h, i) => ({ ...acc, [h]: i }), {});

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    return data
      .map(r => ({
        id: r[i.id], type: r[i.type] || 'Styremøte', dato: r[i.dato], start: r[i.start] || '',
        slutt: r[i.slutt] || '', sted: r[i.sted] || '', tittel: r[i.tittel] || '',
        agenda: r[i.agenda] || '', status: r[i.status] || 'Planlagt',
        participants: r[i.participants] ? String(r[i.participants]).split(',') : []
      }))
      .filter(m => m.status !== 'Slettet' && m.status !== 'Arkivert')
      .filter(m => {
        if (!m.dato) return scope === 'planned';
        const meetingDate = m.dato instanceof Date ? m.dato : new Date(m.dato);
        return scope === 'past' ? meetingDate < today : meetingDate >= today;
      })
      .sort((a, b) => (a.dato?.getTime() || 0) - (b.dato?.getTime() || 0));
  };

  const nextSaksnr_ = (moteId) => {
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
  };

  function newAgendaItem(moteId) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      requirePermission('MANAGE_MEETINGS', 'Administrere møter');
      const sakerSheet = _ensureSheet_(MOTE_SAKER, SAKER_HEADERS);
      const H = sakerSheet.getRange(1, 1, 1, sakerSheet.getLastColumn()).getValues()[0];
      const idx = H.reduce((acc, h, i) => ({ ...acc, [h]: i }), {});
      const sakId = `SAK-${Utilities.getUuid().slice(0, 8)}`;
      const saksnr = nextSaksnr_(moteId);
      const now = new Date();
      
      const newRow = Array(H.length).fill('');
      newRow[idx.mote_id] = moteId;
      newRow[idx.sak_id] = sakId;
      newRow[idx.saksnr] = saksnr;
      newRow[idx.created_ts] = now;
      newRow[idx.updated_ts] = now;

      sakerSheet.appendRow(newRow);
      Indexer.set(MOTE_SAKER, 'sak_id', sakId, sakerSheet.getLastRow());

      _log_('Sak', `Ny sak ${sakId} (${saksnr}) for møte ${moteId}`);
      return { ok: true, sakId, saksnr };
    } finally {
      lock.releaseLock();
    }
  }

  function saveAgenda(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      requirePermission('MANAGE_MEETINGS', 'Administrere møter');
      const sakerSheet = _ensureSheet_(MOTE_SAKER, SAKER_HEADERS);
      const H = sakerSheet.getRange(1, 1, 1, sakerSheet.getLastColumn()).getValues()[0];
      const idx = { tittel: H.indexOf('tittel'), forslag: H.indexOf('forslag'), vedtak: H.indexOf('vedtak'), updated_ts: H.indexOf('updated_ts') };

      const index = Indexer.get(MOTE_SAKER, SAKER_HEADERS, 'sak_id');
      const rowNum = index[payload.sakId];
      if (!rowNum) return { ok: false, message: `Fant ikke sak ${payload.sakId}` };

      const range = sakerSheet.getRange(rowNum, 1, 1, H.length);
      const cur = range.getValues()[0];
      
      cur[idx.tittel] = payload.tittel ?? cur[idx.tittel];
      cur[idx.forslag] = payload.forslag ?? cur[idx.forslag];
      cur[idx.vedtak] = payload.vedtak ?? cur[idx.vedtak];
      cur[idx.updated_ts] = new Date();
      
      range.setValues([cur]);
      return { ok: true, message: 'Sak lagret' };
    } finally {
      lock.releaseLock();
    }
  }

  const listAgenda = (moteId) => {
    const sakerSheet = _ensureSheet_(MOTE_SAKER, SAKER_HEADERS);
    const last = sakerSheet.getLastRow();
    if (last < 2) return [];
    
    const data = sakerSheet.getRange(2, 1, last - 1, sakerSheet.getLastColumn()).getValues();
    const H = sakerSheet.getRange(1, 1, 1, sakerSheet.getLastColumn()).getValues()[0];
    const i = H.reduce((acc, h, i) => ({ ...acc, [h]: i }), {});

    return data.filter(r => r[i.mote_id] === moteId)
      .map(r => ({
        moteId: r[i.mote_id], sakId: r[i.sak_id], saksnr: r[i.saksnr],
        tittel: r[i.tittel], forslag: r[i.forslag], vedtak: r[i.vedtak],
        updated_ts: r[i.updated_ts]
      }))
      .sort((a, b) => String(a.saksnr).localeCompare(String(b.saksnr)));
  };
  
  function deleteAgendaItem(sakId, opts) {
    const cascade = opts?.cascade !== false;
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      requirePermission('MANAGE_MEETINGS', 'Administrere møter');
      const sakerSheet = _ensureSheet_(MOTE_SAKER, SAKER_HEADERS);
      const index = Indexer.get(MOTE_SAKER, SAKER_HEADERS, 'sak_id');
      const rowNum = index[sakId];
      if (rowNum) {
        sakerSheet.deleteRow(rowNum);
        Indexer.rebuild(MOTE_SAKER, SAKER_HEADERS, 'sak_id');
      }
      
      if (cascade) {
        const innspillSheet = _ensureSheet_(MOTE_KOMMENTARER, INNSPILL_HEADERS);
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
      const sh = _ensureSheet_(MOTE_KOMMENTARER, INNSPILL_HEADERS);
      const { email } = _getCurrentUser_();
      sh.appendRow([sakId, new Date(), email, text]);
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  }

  const listInnspill = (sakId, sinceISO) => {
    const sh = _ensureSheet_(MOTE_KOMMENTARER, INNSPILL_HEADERS);
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
      .sort((a, b) => (a.ts?.getTime() || 0) - (b.ts?.getTime() || 0));
  };

  function castVote(sakId, value) {
    if (!sakId) throw new Error('Mangler sakId.');
    const v = String(value || '').toUpperCase();
    if (!['JA', 'NEI', 'BLANK'].includes(v)) throw new Error('Ugyldig stemme.');

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const { email, name } = _getCurrentUser_();
      if (!email) throw new Error('Kunne ikke identifisere bruker.');
      const moteId = getMoteIdForSak_(sakId);
      if (!moteId) throw new Error(`Fant ikke tilhørende møte for sak ${sakId}`);
      
      const sh = _ensureSheet_(MOTE_STEMMER, STEMMER_HEADERS);
      const voteIdx = getVoteIndex_(sakId);
      const rowNum = voteIdx[email.toLowerCase()];

      if (rowNum) {
        const cVote = STEMMER_HEADERS.indexOf('vote') + 1;
        const cTs = STEMMER_HEADERS.indexOf('ts') + 1;
        sh.getRange(rowNum, cVote).setValue(v);
        sh.getRange(rowNum, cTs).setValue(new Date());
      } else {
        const voteId = `V-${Utilities.getUuid().slice(0, 8)}`;
        sh.appendRow([voteId, sakId, moteId, email, name, v, new Date()]);
        setVoteIndexRow_(sakId, email, sh.getLastRow());
      }
      _log_('Stemme', `${email} -> ${sakId} = ${v}`);
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  }

  const getVoteSummary = (sakId) => {
    if (!sakId) return { JA: 0, NEI: 0, BLANK: 0 };
    const sh = _ensureSheet_(MOTE_STEMMER, STEMMER_HEADERS);
    if (sh.getLastRow() < 2) return { JA: 0, NEI: 0, BLANK: 0 };

    const data = sh.getDataRange().getValues();
    const H = data.shift();
    const iSak = H.indexOf('sak_id'), iVote = H.indexOf('vote');
    
    return data.reduce((summary, row) => {
      if (row[iSak] === sakId) {
        const vote = String(row[iVote] || '').toUpperCase();
        if (summary[vote] !== undefined) summary[vote]++;
      }
      return summary;
    }, { JA: 0, NEI: 0, BLANK: 0 });
  };

  const rtServerNow = () => ({ now: new Date().toISOString() });
  
  function rtGetChanges(moteId, sinceISO) {
    const serverNow = new Date().toISOString();
    const since = sinceISO ? new Date(sinceISO) : null;
    let meetingUpdated = null;

    const moterSheet = _ensureSheet_(MOTER, MEETINGS_HEADERS);
    const moterH = moterSheet.getRange(1, 1, 1, moterSheet.getLastColumn()).getValues()[0];
    const iUpdated = moterH.indexOf('updated_ts');
    const moterIndex = Indexer.get(MOTER, MEETINGS_HEADERS, 'id');
    const moteRowNum = moterIndex[moteId];

    if (moteRowNum) {
      const uts = moterSheet.getRange(moteRowNum, iUpdated + 1).getValue();
      if (uts instanceof Date && (!since || uts > since)) {
        meetingUpdated = listMeetings_({ scope: 'all' }).find(m => m.id === moteId);
      }
    }

    const alleSaker = listAgenda(moteId);
    const updatedSaker = alleSaker.filter(sak => sak.updated_ts && (!since || new Date(sak.updated_ts) > since));

    const sakIds = new Set(alleSaker.map(s => s.sakId));
    let newInnspill = [];
    if (sakIds.size > 0) {
      newInnspill = Array.from(sakIds).flatMap(sakId => listInnspill(sakId, sinceISO));
    }
    return { serverNow, meetingUpdated, updatedSaker, newInnspill };
  }

  function getAiAssistance(text, mode) {
    try {
      const API_KEY = PROPS.getProperty('AI_API_KEY');
      if (!API_KEY) {
        return 'AI API-nøkkel er ikke konfigurert. Vennligst kontakt en administrator.';
      }

      const API_URL = 'https://api.openai.com/v1/completions'; // Placeholder URL

      let prompt = '';
      if (mode === 'summarize') {
        prompt = `Oppsummer følgende møtesak på en konsis måte (maks 3-4 setninger):\n\n${text}`;
      } else if (mode === 'tasks') {
        prompt = `Basert på følgende møtesak, lag en punktliste med konkrete oppgaver som må gjøres. Inkluder hvem som kan være ansvarlig hvis det er nevnt. Hvis ingen oppgaver virker nødvendige, svar "Ingen åpenbare oppgaver".\n\n${text}`;
      } else {
        return 'Ugyldig AI-modus.';
      }

      const payload = {
        model: 'text-davinci-003', // Placeholder model
        prompt: prompt,
        max_tokens: 150,
        temperature: 0.5,
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
          'Authorization': 'Bearer ' + API_KEY,
        },
        payload: JSON.stringify(payload),
      };

      const response = UrlFetchApp.fetch(API_URL, options);
      const jsonResponse = JSON.parse(response.getContentText());
      const aiText = jsonResponse.choices && jsonResponse.choices[0] && jsonResponse.choices[0].text;

      return aiText ? aiText.trim() : 'Fikk ikke noe svar fra AI-tjenesten.';

    } catch (e) {
      _log_('AI_FEIL', e.message);
      return `En feil oppstod under kall til AI-tjenesten: ${e.message}`;
    }
  }

  global.getBoardMembers = getBoardMembers;
  global.uiBootstrap = uiBootstrap;
  global.upsertMeeting = upsertMeeting;
  global.listMeetings_ = listMeetings_;
  global.newAgendaItem = newAgendaItem;
  global.saveAgenda = saveAgenda;
  global.listAgenda = listAgenda;
  global.deleteAgendaItem = deleteAgendaItem;
  global.appendInnspill = appendInnspill;
  global.listInnspill = listInnspill;
  global.castVote = castVote;
  global.getVoteSummary = getVoteSummary;
  global.rtServerNow = rtServerNow;
  global.rtGetChanges = rtGetChanges;
  global.getAiAssistance = getAiAssistance;
})(this);