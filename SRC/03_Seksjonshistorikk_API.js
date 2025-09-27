/* ====================== Seksjonshistorikk – Enhanced API & UI ======================
 * FILE: 03_Seksjonshistorikk_Enhanced.gs | VERSION: 2.1.0 | UPDATED: 2025-09-15
 * FORMÅL: Optimized and secure section history with caching, filtering, and improved performance
 * Avhenger av: SHEETS, _tz_(), _logEvent()
 * ================================================================================ */

////////////////////////////////////////////////////////////////////////////////
// CONFIGURATION AND CONSTANTS
////////////////////////////////////////////////////////////////////////////////

const CACHE_DURATION = 5 * 60 * 1000; // 5 minutes
const MAX_RESULTS_DEFAULT = 1000;
const cache = new Map();

/** @enum {string} */
const EVENT_TYPES = Object.freeze({
  EIERSKAP: 'Eierskap',
  LEIE: 'Leie',
  OPPGAVE: 'Oppgave',
  INNSPILL: 'Innspill',
  VEDLEGG: 'Vedlegg'
});

const EVENT_PROCESSORS = Object.freeze({
  eierskap: {
    sheet: 'EIERSKAP',
    type: EVENT_TYPES.EIERSKAP,
    columns: {
      seksjonsnr: ['seksjonsnr'],
      fra: ['fra_dato', 'start', 'fra'],
      til: ['til_dato', 'slutt', 'til'],
      personId: ['eier_person_id', 'person_id', 'eier']
    }
  },
  leie: {
    sheet: 'LEIE',
    type: EVENT_TYPES.LEIE,
    columns: {
      seksjonsnr: ['seksjonsnr', 'seksjon'],
      fra: ['fra_dato', 'start'],
      til: ['til_dato', 'slutt'],
      personId: ['leietaker_person_id', 'person_id', 'leietaker'],
      kontrakt: ['kontrakt_url', 'lenke', 'url']
    }
  },
  tasks: {
    sheet: 'TASKS',
    type: EVENT_TYPES.OPPGAVE,
    columns: {
      tittel: ['Tittel', 'Sak', 'Emne', 'Title'],
      kategori: ['Kategori'],
      status: ['Status'],
      frist: ['Frist', 'Due', 'Forfallsdato'],
      opprettet: ['Opprettet', 'Opprettet dato', 'Created'],
      ansvarlig: ['Ansvarlig', 'Owner'],
      seksjonsnr: ['Seksjonsnr', 'seksjonsnr', 'Seksjon', 'seksjon']
    }
  },
  support: {
    sheet: 'SUPPORT',
    type: EVENT_TYPES.INNSPILL,
    columns: {
      seksjonsnr: ['Seksjonsnr', 'seksjonsnr', 'Seksjon', 'seksjon', 'Leil', 'leil'],
      tittel: ['Tittel', 'Emne', 'Subject', 'Sak'],
      status: ['Status'],
      opprettet: ['Opprettet', 'Mottatt', 'Dato', 'ts', 'timestamp'],
      link: ['Lenke', 'URL', 'Link']
    }
  }
});

////////////////////////////////////////////////////////////////////////////////
// LOGGING HELPERS
////////////////////////////////////////////////////////////////////////////////

function logInfo(msg) {
  Logger.log(`INFO: ${msg}`);
  if (typeof _logEvent === 'function') _logEvent('History', msg);
}

function logError(msg) {
  Logger.log(`ERROR: ${msg}`);
  if (typeof _logEvent === 'function') _logEvent('HistoryError', msg);
}

////////////////////////////////////////////////////////////////////////////////
// PUBLIC API FUNCTIONS
////////////////////////////////////////////////////////////////////////////////

/*
 * MERK: openSectionHistory() er fjernet fra denne filen for å unngå konflikter.
 * Funksjonen kalles nå fra 00_App_Core.js, som bruker den sentrale UI_FILES-mappingen.
 */

/**
 * Enhanced section history with filtering options
 * @param {string} seksjonsnr
 * @param {Object} options
 * @returns {{ok:boolean, seksjonsnr?:string, count?:number, events?:Array, error?:string, hasMore?:boolean}}
 */
function getCompleteSectionHistory(seksjonsnr, options = {}) {
  const operationId = Utilities.getUuid().slice(0, 8);
  try {
    const sx = _validateSectionNumber_(seksjonsnr);
    const config = _parseHistoryOptions_(options);
    logInfo(`[${operationId}] Henter historikk for seksjon ${sx}`);
    const events = _collectAllSectionEvents_(sx, config);
    events.sort((a, b) => (new Date(b.ts) - new Date(a.ts)));
    const limitedEvents = config.maxResults > 0 ? events.slice(0, config.maxResults) : events;
    logInfo(`[${operationId}] Fant ${limitedEvents.length} hendelser for seksjon ${sx}`);
    return { ok: true, seksjonsnr: sx, count: limitedEvents.length, events: limitedEvents, hasMore: events.length > limitedEvents.length };
  } catch (error) {
    logError(`[${operationId}] ${error.message}`);
    const userMessage = error.message.startsWith('VALIDERING:') ? error.message : `Feil ved henting av historikk: ${error.message}`;
    return { ok: false, error: userMessage };
  }
}

/** Eksporter seksjonshistorikk til nytt ark */
function exportSectionHistoryToSheet(seksjonsnr, options = {}) {
  try {
    const result = getCompleteSectionHistory(seksjonsnr, { ...options, maxResults: 0 });
    if (!result.ok) return result;
    const ss = SpreadsheetApp.getActive();
    const sheetName = `Historikk_${String(seksjonsnr).trim()}`;
    let sheet = ss.getSheetByName(sheetName);
    if (sheet) sheet.clear(); else sheet = ss.insertSheet(sheetName);
    const headers = ['Tidspunkt', 'Type', 'Tittel', 'Beskrivelse', 'Kilde', 'Lenke', 'Vedlegg'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
    if (result.events?.length) {
      const rows = result.events.map(ev => [ev.ts ? new Date(ev.ts) : '', ev.type, ev.title, ev.desc, ev.source, ev.link, (ev.attachments || []).join('\n')]);
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows).setVerticalAlignment('top');
      sheet.getRange(2, 1, rows.length, 1).setNumberFormat('dd.MM.yyyy hh:mm');
      try { sheet.autoResizeColumns(1, headers.length); } catch (e) {}
    }
    return { ok: true, sheet: sheetName, rows: (result.events || []).length, hasMore: result.hasMore };
  } catch (error) {
    return { ok: false, error: error.message };
  }
}

////////////////////////////////////////////////////////////////////////////////
// CORE DATA COLLECTION
////////////////////////////////////////////////////////////////////////////////

function _collectAllSectionEvents_(seksjonsnr, config) {
  const allEvents = [];
  for (const [name, cfg] of Object.entries(EVENT_PROCESSORS)) {
    if (config.eventTypes && !config.eventTypes.includes(cfg.type)) continue;
    try {
      const sheetName = SHEETS[cfg.sheet];
      if (!sheetName) continue;
      const data = _getCachedSheetData_(sheetName);
      if (!data) continue;
      const events = _processEventsByType_(seksjonsnr, data, cfg, config);
      allEvents.push(...events);
    } catch (err) {
      logError(`Feil i prosessering av ${name}: ${err.message}`);
    }
  }
  if (config.includeAttachments) {
    try { allEvents.push(..._collectAttachmentHistory_(seksjonsnr, config)); }
    catch (err) { logError(`Feil i vedlegg: ${err.message}`); }
  }
  return allEvents;
}

function _processEventsByType_(seksjonsnr, data, cfg, config) {
  const events = [];
  const col = _getColumnIndices_(data.header, cfg.columns);
  if (!col.seksjonsnr) return events;
  for (const row of data.rows) {
    if (String(row[col.seksjonsnr - 1] || '').trim() !== seksjonsnr) continue;
    const evs = _createEventsFromRow_(row, col, cfg, config);
    events.push(...evs.filter(e => _isEventInDateRange_(e, config)));
  }
  return events;
}

function _makeEvent({ ts, type, title = '', desc = '', source = '', link = '', attachments = [] }) {
  return { ts, type, title, desc, source, link, attachments };
}

function _createEventsFromRow_(row, col, cfg, config) {
  const tz = _getTimezone_();
  const evs = [];
  switch (cfg.type) {
    case EVENT_TYPES.EIERSKAP:
      const fraDato = col.fra ? _parseDate_(row[col.fra - 1]) : null;
      const tilDato = col.til ? _parseDate_(row[col.til - 1]) : null;
      const personId = col.personId ? String(row[col.personId - 1] || '').trim() : '';
      if (fraDato) evs.push(_makeEvent({ ts: fraDato, type: 'Eierskap start', title: personId ? `Ny eier (${personId})` : 'Ny eier', desc: `Eierskap registrert fra ${_formatDate_(fraDato, tz)}`, source: EVENT_TYPES.EIERSKAP }));
      if (tilDato) evs.push(_makeEvent({ ts: tilDato, type: 'Eierskap slutt', title: personId ? `Eier sluttet (${personId})` : 'Eier sluttet', desc: `Eierskap avsluttet ${_formatDate_(tilDato, tz)}`, source: EVENT_TYPES.EIERSKAP }));
      break;
    case EVENT_TYPES.LEIE:
      const leieFra = col.fra ? _parseDate_(row[col.fra - 1]) : null;
      const leieTil = col.til ? _parseDate_(row[col.til - 1]) : null;
      const leietaker = col.personId ? String(row[col.personId - 1] || '').trim() : '';
      const kontrakt = col.kontrakt ? String(row[col.kontrakt - 1] || '') : '';
      if (leieFra) evs.push(_makeEvent({ ts: leieFra, type: 'Leie start', title: leietaker ? `Leietaker (${leietaker})` : 'Leietaker', desc: `Leieforhold fra ${_formatDate_(leieFra, tz)}`, source: EVENT_TYPES.LEIE, link: kontrakt }));
      if (leieTil) evs.push(_makeEvent({ ts: leieTil, type: 'Leie slutt', title: leietaker ? `Leietaker sluttet (${leietaker})` : 'Leie avsluttet', desc: `Leieforhold avsluttet ${_formatDate_(leieTil, tz)}`, source: EVENT_TYPES.LEIE, link: kontrakt }));
      break;
    case EVENT_TYPES.OPPGAVE:
      const opprettet = col.opprettet ? _parseDate_(row[col.opprettet - 1]) : null;
      const frist = col.frist ? _parseDate_(row[col.frist - 1]) : null;
      const kategori = col.kategori ? String(row[col.kategori - 1] || '') : '';
      const tittel = col.tittel ? String(row[col.tittel - 1] || '(uten tittel)') : '(uten tittel)';
      const status = col.status ? String(row[col.status - 1] || '') : '';
      const ansvarlig = col.ansvarlig ? String(row[col.ansvarlig - 1] || '') : '';
      const when = opprettet || frist;
      if (when) {
        const descParts = [];
        if (kategori && kategori.toLowerCase() !== 'hms') descParts.push(`Kategori ${kategori}`);
        if (frist) descParts.push(`Frist ${_formatDate_(frist, tz)}`);
        if (ansvarlig) descParts.push(`Ansv. ${ansvarlig}`);
        if (status) descParts.push(status);
        const type = kategori.toLowerCase() === 'hms' ? 'HMS' : 'Oppgave';
        evs.push(_makeEvent({ ts: when, type, title: tittel, desc: descParts.join(' • '), source: EVENT_TYPES.OPPGAVE }));
      }
      break;
    case EVENT_TYPES.INNSPILL:
      const dato = col.opprettet ? _parseDate_(row[col.opprettet - 1]) : null;
      const innspillTittel = col.tittel ? String(row[col.tittel - 1] || '(uten tittel)') : '(uten tittel)';
      const innspillStatus = col.status ? String(row[col.status - 1] || '') : '';
      const innspillLink = col.link ? String(row[col.link - 1] || '') : '';
      evs.push(_makeEvent({ ts: dato, type: EVENT_TYPES.INNSPILL, title: innspillTittel, desc: innspillStatus, source: 'Support', link: innspillLink }));
      break;
  }
  return evs;
}

////////////////////////////////////////////////////////////////////////////////
// UTILITY FUNCTIONS
////////////////////////////////////////////////////////////////////////////////

function _getCachedSheetData_(sheetName) {
  const cacheKey = `sheet_${sheetName}`;
  const now = Date.now();
  // In-memory
  const cached = cache.get(cacheKey);
  if (cached && now - cached.timestamp < CACHE_DURATION) return cached.data;
  // CacheService
  const scriptCache = CacheService.getScriptCache();
  const cachedStr = scriptCache.get(cacheKey);
  if (cachedStr) {
    const data = JSON.parse(cachedStr);
    cache.set(cacheKey, { data, timestamp: now });
    return data;
  }
  // Fallback lesing
  const data = _readSheetSafe_(sheetName);
  if (data) {
    scriptCache.put(cacheKey, JSON.stringify(data), CACHE_DURATION / 1000);
    cache.set(cacheKey, { data, timestamp: now });
  }
  return data;
}

function _readSheetSafe_(sheetName) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 1) return null;
    const values = sheet.getDataRange().getValues();
    if (!values?.length) return null;
    const header = values.shift().map(h => String(h));
    return { header, rows: values };
  } catch (err) {
    logError(`_readSheetSafe_(${sheetName}): ${err.message}`);
    return null;
  }
}

function _getColumnIndices_(header, columnDefs) {
  const headerLower = header.map(h => String(h || '').trim().toLowerCase());
  const indices = {};
  for (const [key, aliases] of Object.entries(columnDefs)) {
    let index = 0;
    for (const alias of aliases) {
      const pos = headerLower.indexOf(String(alias || '').trim().toLowerCase());
      if (pos >= 0) { index = pos + 1; break; }
    }
    indices[key] = index;
  }
  return indices;
}

function _parseDate_(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  const str = String(value || '').trim();
  if (!str) return null;
  const ddMmYyyy = str.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (ddMmYyyy) return new Date(Number(ddMmYyyy[3]), Number(ddMmYyyy[2]) - 1, Number(ddMmYyyy[1]));
  const yyyyMmDd = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (yyyyMmDd) return new Date(Number(yyyyMmDd[1]), Number(yyyyMmDd[2]) - 1, Number(yyyyMmDd[3]));
  const parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function _formatDate_(date, tz) {
  try { return Utilities.formatDate(date, tz || 'Europe/Oslo', 'dd.MM.yyyy'); }
  catch { return date instanceof Date ? date.toDateString() : String(date || ''); }
}

function _getTimezone_() {
  return (typeof _tz_ === 'function') ? _tz_() : (Session.getScriptTimeZone() || 'Europe/Oslo');
}

function _validateSectionNumber_(seksjonsnr) {
  const sx = String(seksjonsnr || '').trim();
  if (!sx) throw new Error('VALIDERING: Mangler seksjonsnummer.');
  return sx;
}

function _parseHistoryOptions_(options) {
  return {
    startDate: options.startDate ? _parseDate_(options.startDate) : null,
    endDate: options.endDate ? _parseDate_(options.endDate) : null,
    eventTypes: Array.isArray(options.eventTypes) ? options.eventTypes : null,
    includeAttachments: options.includeAttachments !== false,
    maxResults: Math.max(0, Number(options.maxResults) || MAX_RESULTS_DEFAULT)
  };
}

function _isEventInDateRange_(event, config) {
  if (!event.ts) return true;
  const eventDate = new Date(event.ts);
  if (config.startDate && eventDate < config.startDate) return false;
  if (config.endDate && eventDate > config.endDate) return false;
  return true;
}
