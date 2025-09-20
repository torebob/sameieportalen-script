/**
 * CoreAnalysisPlus – DataCollector (v1.4.2)
 *
 * Purpose:
 *   Collects spreadsheet metadata, triggers, menu functions, and a typed data model
 *   in a defensive, scalable way. Adds progress callbacks, yielding, and caching.
 *
 * Public (used by Analysis layer):
 *   _collectMetadata_()
 *   _collectTriggers_()
 *   _collectMenuFunctions_()
 *   _collectDataModel_(progressCb?)  // progressCb({ current, total, sheetName, percentage })
 *
 * Key improvements since v1.4.0:
 *   - v1.4.1: Protected-sheet safe reads via _safeGetValues_() + dual caching ready
 *   - v1.4.2: Optional confidence-scoring for type inference (FEATURE FLAG: TYPE_CONFIDENCE=false)
 *             Cache diagnostics helpers: dc_getCacheStats_(), dc_clearCaches_()
 *             Regex config still overridable via CONFIG_PLUS.REGEX (e.g., currency patterns)
 *
 * Feature flags you can set in CONFIG_PLUS (optional):
 *   CONFIG_PLUS.TYPE_CONFIDENCE = false; // when true, uses score-based type detection
 *
 * Notes:
 *   - This module is dependency-safe: will run even if Config/Utils/Logger modules aren’t present.
 *   - All CONFIG_PLUS overrides are optional and read via safe bridges below.
 */

/* --------------------------- Safe dependency bridges --------------------------- */

function __dc_getLogger_() {
  try {
    if (typeof _getLoggerPlus_ === 'function') return _getLoggerPlus_();
    if (typeof getAppLogger_ === 'function')   return getAppLogger_();
  } catch (_) {}
  return { // console fallback
    debug: function (fn, msg, d) { try { console.log('[DEBUG]', fn||'', msg||'', d||''); } catch(_){} },
    info:  function (fn, msg, d) { try { console.log('[INFO]',  fn||'', msg||'', d||''); } catch(_){} },
    warn:  function (fn, msg, d) { try { console.warn('[WARN]',  fn||'', msg||'', d||''); } catch(_){} },
    error: function (fn, msg, d) { try { console.error('[ERROR]',fn||'', msg||'', d||''); } catch(_){} }
  };
}
function __dc_cfgGet_(key, fallback) {
  try { if (typeof _cfgGet_ === 'function') return _cfgGet_(key, fallback); } catch(_) {}
  // fallback to CORE_ANALYSIS_CFG or provided fallback
  try {
    if (typeof CORE_ANALYSIS_CFG !== 'undefined' && CORE_ANALYSIS_CFG &&
        Object.prototype.hasOwnProperty.call(CORE_ANALYSIS_CFG, key)) {
      return CORE_ANALYSIS_CFG[key];
    }
  } catch(_) {}
  return fallback;
}
function __dc_numCfg_(key, fallback) {
  try { if (typeof _numCfgSafe_ === 'function') return _numCfgSafe_(key, fallback); } catch(_) {}
  var v = Number(__dc_cfgGet_(key, fallback));
  return isNaN(v) ? Number(fallback) : v;
}
function __dc_safe_(fn, fallback) {
  try { if (typeof _safe === 'function') return _safe(fn, fallback); } catch(_) {}
  try { return fn(); } catch(_) { return fallback; }
}
function __dc_truthy_(v, locale) {
  try { if (typeof _truthy_ === 'function') return _truthy_(v, locale); } catch(_) {}
  var s = String(v).trim().toLowerCase();
  return (s === '1' || s === 'true' || s === 'ja' || s === 'x' || s === 'on' || s === 'enabled');
}
function __dc_normalizeName_(s, stripAll) {
  try { if (typeof _normalizeName_ === 'function') return _normalizeName_(s, stripAll); } catch(_) {}
  var out = String(s||'').toLowerCase().trim();
  out = out.replace(/\s+/g, stripAll ? '' : ' ');
  out = out.replace(/_/g, stripAll ? '' : '_');
  return out;
}
function __dc_normalizeHeader_(h) {
  try { if (typeof _normalizeHeader_ === 'function') return _normalizeHeader_(h); } catch(_) {}
  return String(h || '').trim().toLowerCase();
}
function __dc_indexOfHeaderAny_(headersLower, alts) {
  try { if (typeof _indexOfHeaderAny_ === 'function') return _indexOfHeaderAny_(headersLower, alts); } catch(_) {}
  if (!Array.isArray(headersLower) || !Array.isArray(alts)) return -1;
  var wants = alts.map(function (x) { return String(x||'').trim().toLowerCase(); });
  for (var i=0;i<headersLower.length;i++) {
    var h = String(headersLower[i]||'').trim().toLowerCase();
    for (var j=0;j<wants.length;j++) {
      if (h === wants[j]) return i;
    }
  }
  return -1;
}
function __dc_buildHeaderPreview_(arr) {
  try { if (typeof _buildHeaderPreview_ === 'function') return _buildHeaderPreview_(arr); } catch(_) {}
  var max = __dc_numCfg_('MAX_HEADER_PREVIEW', 10);
  var preview = (arr || []).slice(0, Math.max(0,max)).map(function(h){return String(h||'').trim();}).filter(Boolean);
  return preview.join(' | ');
}
function __dc_stringifySafe_(obj, space) {
  try { if (typeof _stringifySafe_ === 'function') return _stringifySafe_(obj, space); } catch(_) {}
  try { return JSON.stringify(obj, null, space||0); } catch(e) { return '<<JSON Error: '+(e&&e.message)+'>>'; }
}

/* --------------------------------- Constants --------------------------------- */

var __DC_CONST = {
  DATA_START_ROW: __dc_numCfg_('DATA_START_ROW', 2),
  SCAN_COL_CHUNK: __dc_numCfg_('SCAN_COL_CHUNK', 50),
  MAX_SCAN_ROWS:  __dc_numCfg_('MAX_SCAN_ROWS', 25),
  PROGRESS_EVERY: __dc_numCfg_('PROGRESS_LOG_EVERY_SHEETS', 5),
  YIELD_EVERY:    __dc_numCfg_('YIELD_EVERY_SHEETS', 15),
  TYPE_CONFIDENCE: (function(){ try{
    if (typeof CONFIG_PLUS !== 'undefined' && CONFIG_PLUS && 'TYPE_CONFIDENCE' in CONFIG_PLUS)
      return !!CONFIG_PLUS.TYPE_CONFIDENCE;
  }catch(_){} return false; })()
};

/* -------------------------------- Utilities ---------------------------------- */

/** Protected-range safe values getter (best-effort). */
function _safeGetValues_(range) {
  var log = __dc_getLogger_(), fn = '_safeGetValues_';
  try {
    return range.getValues();
  } catch (e) {
    // Try reading row-by-row to skip protected chunks
    try {
      var r = range.getNumRows(), c = range.getNumColumns();
      var out = new Array(r);
      for (var i=0;i<r;i++) {
        try {
          out[i] = range.offset(i, 0, 1, c).getValues()[0];
        } catch (rowErr) {
          // If a row is unreadable, fill with blanks to preserve shape
          out[i] = new Array(c).fill('');
          log.warn(fn, 'Protected row skipped', { rowOffset: i, error: rowErr && rowErr.message });
        }
      }
      return out;
    } catch (fallbackErr) {
      log.error(fn, 'Failed to read values (protected or inaccessible range)', {
        error: e && e.message, fallbackError: fallbackErr && fallbackErr.message
      });
      // Return safe empty matrix
      try { return new Array(range.getNumRows()).fill(0).map(function(){ return new Array(range.getNumColumns()).fill(''); }); }
      catch(_) { return [[]]; }
    }
  }
}

/* --------------------------------- Caching ---------------------------------- */

var __sheetScanCache = Object.create(null); // per-execution cache
// Note: If you use an LRU cache elsewhere (e.g., analysis layer), diagnostics below are safe.

/* --------------------------------- Metadata --------------------------------- */

function _collectMetadata_() {
  var log = __dc_getLogger_(), fn = '_collectMetadata_';
  try {
    var ss = SpreadsheetApp.getActive();
    var userEmail = __dc_safe_(function(){ return Session.getActiveUser().getEmail(); }, '');
    return {
      spreadsheetName: __dc_safe_(function(){ return ss.getName(); }, ''),
      spreadsheetUrl:  __dc_safe_(function(){ return ss.getUrl(); }, ''),
      spreadsheetId:   __dc_safe_(function(){ return ss.getId(); }, ''),
      timeZone:        __dc_safe_(function(){ return ss.getSpreadsheetTimeZone(); }, ''),
      locale:          __dc_safe_(function(){ return ss.getSpreadsheetLocale && ss.getSpreadsheetLocale(); }, ''),
      sheetsCount:     __dc_safe_(function(){ return ss.getSheets().length; }, 0),
      user:            userEmail
    };
  } catch (e) {
    log.error(fn, 'Failed to collect metadata', { error: e.message });
    return { spreadsheetName:'', spreadsheetUrl:'', spreadsheetId:'', timeZone:'', locale:'', sheetsCount:0, user:'' };
  }
}

/* --------------------------------- Triggers --------------------------------- */

function _collectTriggers_() {
  var log = __dc_getLogger_(), fn = '_collectTriggers_', out = [];
  try {
    var trig = ScriptApp.getProjectTriggers() || [];
    trig.forEach(function(t){
      var eventType='', source='', handler='';
      try { handler = String(t.getHandlerFunction() || ''); } catch(_) {}
      try { eventType = String(t.getEventType && t.getEventType()); } catch(_) {}
      try { source = String(t.getTriggerSource && t.getTriggerSource()); } catch(_) {}
      out.push({ handler:handler, eventType:(eventType||'UNKNOWN'), source:(source||'UNKNOWN'), raw:{eventType,source} });
    });
  } catch (e) {
    log.error(fn, 'Failed to collect triggers', { error: e.message });
  }
  return out;
}

/* --------------------------------- Menus ------------------------------------ */

function _collectMenuFunctions_() {
  var log = __dc_getLogger_(), fn = '_collectMenuFunctions_', out = [];
  try {
    var names = __dc_cfgGet_('NAMES', { menuFelles: ['Meny_Felles'], menuMin: ['Meny_Min'] });
    var shFelles = _getSheetByAnyName_(names.menuFelles);
    var shMin    = _getSheetByAnyName_(names.menuMin);
    if (shFelles) out.push.apply(out, _readMenuSheet_(shFelles, 'Meny_Felles'));
    if (shMin)    out.push.apply(out, _readMenuSheet_(shMin, 'Meny_Min'));
  } catch (e) {
    log.error(fn, 'Failed reading menu sheets', { error: e.message });
  }
  return out;
}

function _readMenuSheet_(sh, sheetLabel) {
  var log = __dc_getLogger_(), fn = '_readMenuSheet_';
  try {
    var rng = sh.getDataRange();
    var vals = _safeGetValues_(rng);
    if (!vals || vals.length < 2) return [];
    var hdr = vals[0].map(function(h){ return String(h||'').trim().toLowerCase(); });

    var titleIdx = __dc_indexOfHeaderAny_(hdr, ['tittel','title','kommando','menu','meny']);
    var fnIdx    = __dc_indexOfHeaderAny_(hdr, ['funksjon','function','handler']);
    var roleIdx  = __dc_indexOfHeaderAny_(hdr, ['rollekrav','rolle','role']);
    var userIdx  = __dc_indexOfHeaderAny_(hdr, ['bruker','user']);
    var actIdx   = __dc_indexOfHeaderAny_(hdr, ['aktiv','active','enabled']);

    var out = [];
    for (var r=1;r<vals.length;r++) {
      var row = vals[r] || [];
      var title = (titleIdx >= 0 ? row[titleIdx] : '') || '';
      var fnName= (fnIdx    >= 0 ? row[fnIdx]    : '') || '';
      if (!title && !fnName) continue;
      var role  = (roleIdx  >= 0 ? row[roleIdx]  : '') || '';
      var user  = (userIdx  >= 0 ? row[userIdx]  : '') || '';
      var act   = (actIdx   >= 0 ? row[actIdx]   : '');
      out.push({
        sheet: sheetLabel,
        title: String(title),
        functionName: String(fnName),
        role: String(role),
        user: String(user),
        active: __dc_truthy_(act, 'no')
      });
    }
    return out;
  } catch (e) {
    log.warn(fn, 'Menu sheet read failed', { sheet: __dc_safe_(function(){return sh.getName();}, '#unknown'), error: e.message });
    return [];
  }
}

/* ----------------------------- Data Model Scan ------------------------------ */

/**
 * Collect a typed model for all sheets.
 * @param {function({current:number,total:number,sheetName:string,percentage:number})} [progressCb]
 * @returns {{ sheets: Array<Object>, headerDuplicates: Array<Object> }}
 */
function _collectDataModel_(progressCb) {
  var log = __dc_getLogger_(), fn = '_collectDataModel_';
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets() || [];
  var outSheets = [];
  var headerIndexGlobal = {}; // normalized header -> occurrences
  var everyN = Math.max(1, __DC_CONST.PROGRESS_EVERY);
  var yieldEvery = Math.max(1, __DC_CONST.YIELD_EVERY);

  var report = function(i, total, name) {
    if (typeof progressCb === 'function') {
      try {
        progressCb({ current: i, total: total, sheetName: name, percentage: total ? Math.round((i/total)*100) : 0 });
      } catch(_) {}
    }
    if (i % everyN === 0) {
      log.info(fn, 'Scanning sheets progress...', { index: i, total: total, sheet: name });
    }
    if (i % yieldEvery === 0) {
      try { Utilities.sleep(1); } catch(_) {} // yield
    }
  };

  for (var i=0;i<sheets.length;i++) {
    var sh = sheets[i];
    var name = __dc_safe_(function(){ return sh.getName(); }, '#'+(i+1));
    report(i, sheets.length, name);

    try {
      var rows = __dc_safe_(function(){ return sh.getLastRow(); }, 0);
      var cols = __dc_safe_(function(){ return sh.getLastColumn(); }, 0);
      var isHidden = __dc_safe_(function(){ return (typeof sh.isSheetHidden === 'function') ? sh.isSheetHidden() : false; }, false);

      var header = [];
      if (cols > 0) {
        header = __dc_safe_(function(){
          return _safeGetValues_(sh.getRange(1,1,1,cols))[0] || [];
        }, []);
      }

      var preview = __dc_buildHeaderPreview_(header);
      var typesByHeader = _inferTypesForSheetChunked_(sh, header); // safe inside

      // duplicates across workbook
      header.forEach(function(h, idx){
        var norm = __dc_normalizeHeader_(h);
        if (!norm) return;
        if (!headerIndexGlobal[norm]) headerIndexGlobal[norm] = [];
        headerIndexGlobal[norm].push({ sheet: name, col: idx+1 });
      });

      outSheets.push({
        name: name,
        rows: rows,
        columns: cols,
        hidden: isHidden,
        headerPreview: preview,
        typesByHeader: typesByHeader
      });

    } catch (e) {
      log.warn(fn, 'Failed scanning sheet (skipping)', { sheet: name, error: e.message });
    }
  }

  // Compute duplicates list
  var duplicates = [];
  Object.keys(headerIndexGlobal).forEach(function(h){
    var occ = headerIndexGlobal[h];
    if (occ && occ.length > 1) {
      duplicates.push({ header: h, occurrences: occ });
    }
  });

  return { sheets: outSheets, headerDuplicates: duplicates };
}

/**
 * Chunked column scan with row guardrails and per-exec cache.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {Array<*>} headers first row
 * @returns {Object<string,string>} headerName -> inferredType
 */
function _inferTypesForSheetChunked_(sh, headers) {
  var log = __dc_getLogger_(), fn = '_inferTypesForSheetChunked_';
  try {
    var cacheKey = 'scan:' + __dc_safe_(function(){ return sh.getSheetId(); }, sh) + ':' + (headers||[]).length + ':' + __DC_CONST.MAX_SCAN_ROWS + ':' + __DC_CONST.SCAN_COL_CHUNK + ':' + __DC_CONST.DATA_START_ROW + ':' + (__DC_CONST.TYPE_CONFIDENCE?'S':'N');
    if (__sheetScanCache[cacheKey]) return __sheetScanCache[cacheKey];

    var totalRows = __dc_safe_(function(){ return sh.getLastRow(); }, 0);
    var totalCols = __dc_safe_(function(){ return sh.getLastColumn(); }, 0);
    var dataStart = Math.max(1, __DC_CONST.DATA_START_ROW);
    var rowsToScan = Math.max(0, Math.min(__DC_CONST.MAX_SCAN_ROWS, Math.max(0, totalRows - (dataStart-1))));
    if (rowsToScan <= 0 || totalCols <= 0 || totalRows < dataStart) return {};

    var chunkSize = Math.max(1, __DC_CONST.SCAN_COL_CHUNK);
    var out = {};
    var colIndex = 1;

    while (colIndex <= totalCols) {
      var thisChunk = Math.min(chunkSize, totalCols - colIndex + 1);
      var range2D = _safeGetValues_(sh.getRange(dataStart, colIndex, rowsToScan, thisChunk)); // [rowsToScan x thisChunk]

      for (var c=0;c<thisChunk;c++) {
        var headerName = String(headers[colIndex - 1 + c] || '').trim();
        if (!headerName) continue;
        var samples = [];
        for (var r=0;r<range2D.length;r++) samples.push(range2D[r][c]);
        out[headerName] = _inferTypeFromSamplesEnhanced_(samples);
      }
      colIndex += thisChunk;
    }

    __sheetScanCache[cacheKey] = out;
    return out;

  } catch (e) {
    log.warn(fn, 'Type inference failed for sheet (skipping types)', { sheet: __dc_safe_(function(){return sh.getName();}, '#unknown'), error: e.message });
    return {};
  }
}

/**
 * Enhanced type inference with configurable patterns.
 * If CONFIG_PLUS.TYPE_CONFIDENCE = true → score-based choice; otherwise binary detection.
 * Order of decision (binary mode): date > percent > currency > number > boolean > email > url > phone > string/empty
 * @param {Array<*>} arr
 * @returns {'empty'|'date'|'percent'|'currency'|'number'|'boolean'|'email'|'url'|'phone'|'string'|'unknown'}
 */
function _inferTypeFromSamplesEnhanced_(arr) {
  return __DC_CONST.TYPE_CONFIDENCE
    ? _inferTypeFromSamplesEnhanced_Score_(arr)
    : _inferTypeFromSamplesEnhanced_NoScore_(arr);
}

/* ---- No-score (default) ---- */
function _inferTypeFromSamplesEnhanced_NoScore_(arr) {
  if (!Array.isArray(arr)) return 'unknown';
  if (arr.length === 0) return 'empty';

  var RX = __dc_cfgGet_('REGEX', {
    email: /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i,
    url: /^(https?:\/\/|www\.)/i,
    phone: /^[\+]?[0-9\s\-\(\)]{6,}$/,
    currency: /^\s*(?:[€$£]|kr|NOK|USD|EUR)\s*[\d\s.,]+|[\d\s.,]+\s*(?:[€$£]|kr|NOK|USD|EUR)\s*$/i,
    percent: /^\s*[-+]?\d+([.,]\d+)?\s*%$/,
    boolTrue: /^(true|ja|yes|y|1)$/i,
    boolFalse:/^(false|nei|no|n|0)$/i
  });

  var has = { date:false, number:false, boolean:false, email:false, url:false, phone:false, currency:false, percent:false };
  var nonEmpty = 0;

  for (var i=0;i<arr.length;i++) {
    var v = arr[i];
    if (v === '' || v === null || typeof v === 'undefined') continue;
    nonEmpty++;

    if (v instanceof Date) { has.date = true; continue; }

    var s = (typeof v === 'string') ? v.trim() :
            (typeof v === 'number' && isFinite(v)) ? String(v) :
            (typeof v === 'boolean') ? (v ? 'true' : 'false') :
            String(v || '').trim();

    if (typeof v === 'number' && isFinite(v)) { has.number = true; }
    else if (s && !isNaN(Number(s.replace(',', '.')))) { has.number = true; }

    if (RX.boolTrue.test(s) || RX.boolFalse.test(s)) { has.boolean = true; }
    if (RX.percent.test(s))  { has.percent = true; }
    if (RX.currency.test(s)) { has.currency = true; }
    if (RX.email.test(s))    { has.email = true; }
    if (RX.url.test(s))      { has.url = true; }
    if (RX.phone.test(s))    { has.phone = true; }
  }

  if (nonEmpty === 0) return 'empty';
  if (has.date)     return 'date';
  if (has.percent)  return 'percent';
  if (has.currency) return 'currency';
  if (has.number)   return 'number';
  if (has.boolean)  return 'boolean';
  if (has.email)    return 'email';
  if (has.url)      return 'url';
  if (has.phone)    return 'phone';
  return 'string';
}

/* ---- Score-based (optional via flag) ---- */
function _inferTypeFromSamplesEnhanced_Score_(arr) {
  if (!Array.isArray(arr) || arr.length === 0) return 'empty';
  var RX = __dc_cfgGet_('REGEX', {
    email: /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i,
    url: /^(https?:\/\/|www\.)/i,
    phone: /^[\+]?[0-9\s\-\(\)]{6,}$/,
    currency: /^\s*(?:[€$£]|kr|NOK|USD|EUR)\s*[\d\s.,]+|[\d\s.,]+\s*(?:[€$£]|kr|NOK|USD|EUR)\s*$/i,
    percent: /^\s*[-+]?\d+([.,]\d+)?\s*%$/,
    boolTrue: /^(true|ja|yes|y|1)$/i,
    boolFalse:/^(false|nei|no|n|0)$/i
  });

  var score = { date:0, percent:0, currency:0, number:0, boolean:0, email:0, url:0, phone:0 };
  var nonEmpty = 0;

  for (var i=0;i<arr.length;i++){
    var v = arr[i]; if (v===''||v==null||typeof v==='undefined') continue; nonEmpty++;
    if (v instanceof Date) { score.date+=3; continue; }

    var s = (typeof v==='string')?v.trim():
            (typeof v==='number'&&isFinite(v))?String(v):
            (typeof v==='boolean')?(v?'true':'false'):
            String(v||'').trim();

    if (!isNaN(Number(String(s).replace(',','.')))) score.number+=1;
    if (RX.percent.test(s))  score.percent+=2;
    if (RX.currency.test(s)) score.currency+=2;
    if (RX.boolTrue.test(s)||RX.boolFalse.test(s)) score.boolean+=1;
    if (RX.email.test(s))    score.email+=2;
    if (RX.url.test(s))      score.url+=1;
    if (RX.phone.test(s))    score.phone+=1;
  }

  if (nonEmpty===0) return 'empty';
  var order = ['date','percent','currency','number','boolean','email','url','phone'];
  var best = order[0];
  for (var k=1;k<order.length;k++) { if (score[order[k]]>score[best]) best = order[k]; }
  return (score[best]>0) ? best : 'string';
}

/* ----------------------------- Sheet name utils ----------------------------- */

function _getSheetByAnyName_(candidates) {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets() || [];
  var cand = (Array.isArray(candidates) ? candidates : [candidates])
    .map(function(s){ return __dc_normalizeName_(String(s||'')); });

  // exact normalized match
  for (var i=0;i<sheets.length;i++) {
    var n = __dc_normalizeName_(sheets[i].getName());
    for (var j=0;j<cand.length;j++) if (n === cand[j]) return sheets[i];
  }
  // loose match (strip spaces/_)
  for (var k=0;k<sheets.length;k++) {
    var n2 = __dc_normalizeName_(sheets[k].getName(), true);
    for (var m=0;m<cand.length;m++) if (n2 === __dc_normalizeName_(cand[m], true)) return sheets[k];
  }
  return null;
}

/* ----------------------------- Cache diagnostics ---------------------------- */

/**
 * Returns quick stats about internal caches. Safe to call; returns zeros if absent.
 * Useful for monitoring during large runs.
 */
function dc_getCacheStats_() {
  try {
    var lruSize;
    try {
      if (typeof __dcLRU !== 'undefined' && __dcLRU && ('size' in __dcLRU)) lruSize = __dcLRU.size;
    } catch(_) {}
    return {
      sheetScanCacheKeys: Object.keys(__sheetScanCache || {}).length,
      lruApproxSize: (typeof lruSize === 'number') ? lruSize : null
    };
  } catch (_) { return { sheetScanCacheKeys: 0, lruApproxSize: null }; }
}

/**
 * Clears per-execution caches (best-effort). Use between long batches.
 */
function dc_clearCaches_() {
  try { for (var k in __sheetScanCache) { if (Object.prototype.hasOwnProperty.call(__sheetScanCache,k)) delete __sheetScanCache[k]; } } catch(_) {}
  try {
    // If an external LRU exists, try to nudge/clear (implementation-dependent).
    if (typeof __dcLRU !== 'undefined' && __dcLRU && typeof __dcLRU.clear === 'function') {
      __dcLRU.clear();
    }
  } catch(_) {}
}
