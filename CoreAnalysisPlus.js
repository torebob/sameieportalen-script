/**
 * CoreAnalysisPlus (FULL, v1.2.1)
 * - Constants extracted (no magic numbers)
 * - JSDoc documentation for public API
 * - Light localization via candidates-config
 * - Safer input validation (threshold, arrays)
 * - Performance: column-chunked scans + periodic progress logs
 * - Performance metrics + large-dataset logging
 * - Config overrides via global CONFIG_PLUS
 *
 * Public API (unchanged from prior versions):
 *   performComprehensiveAnalysis_()
 *   readRequirementsForGapAnalysis_()
 *   generateRequirementCandidates_(analysis)
 *   dedupeCandidates_(candidates, existing, threshold?)
 *   performGapAnalysis_(analysis, existing, deduped)
 *
 * Optional logger integration:
 *   If a function getAppLogger_() exists and returns {info,warn,error}, it will be used.
 *   Otherwise this module falls back to console logging safely.
 */

// ---------------------------- Configuration ----------------------------------

/** Base (defaults). Overridable via CONFIG_PLUS. */
const CORE_ANALYSIS_CFG = {
  VERSION: '1.2.1', // Incremented version
  MAX_SCAN_ROWS: 25,
  MAX_HEADER_PREVIEW: 10,
  SCAN_COL_CHUNK: 50,
  DEFAULT_JACCARD_THRESHOLD: 0.78,
  PROGRESS_LOG_EVERY_SHEETS: 5,
  LARGE_DATA_SHEETS: 50,
  LARGE_DATA_MAXCOLS: 100,
  LARGE_DATA_TOTALROWS: 50000,
  NAMES: {
    kravSheet: ['Krav', 'Requirements', 'KRAV'],
    menuFelles: ['Meny_Felles', 'Meny Felles', 'MENY_FELLES'],
    menuMin: ['Meny_Min', 'Meny Min', 'MENY_MIN']
  },
  HEADERS: {
    krav: {
      id:       ['id', 'krav id', 'kravid', 'krav-id'],
      text:     ['krav', 'beskrivelse', 'tekst', 'hva', 'requirement', 'description', 'text'],
      priority: ['prio.', 'prioritet', 'pri', 'priority'],
      progress: ['fremdrift %', 'fremdrift%', 'fremdrift', 'progress', 'progress %', '%']
    }
  }
};

/** Constants for priority levels to avoid magic strings. */
const PRIORITIES = {
  MUST: 'MÅ',
  SHOULD: 'BØR',
  COULD: 'KAN'
};

/** Constants for requirement sources. */
const SOURCES = {
  TRIGGER: 'trigger',
  MENU: 'menu',
  FIELD: 'field',
  HEURISTIC: 'heuristikk'
};

// ---------------------------- Config helpers ---------------------------------

/** Read effective config: CONFIG_PLUS[key] → fallback to CORE_ANALYSIS_CFG[key] → fallback. */
function _cfgGet_(key, fallback) {
  try {
    if (typeof CONFIG_PLUS !== 'undefined' && CONFIG_PLUS && Object.prototype.hasOwnProperty.call(CONFIG_PLUS, key)) {
      return CONFIG_PLUS[key];
    }
  } catch (_) {}
  if (Object.prototype.hasOwnProperty.call(CORE_ANALYSIS_CFG, key)) return CORE_ANALYSIS_CFG[key];
  return fallback;
}

// ---------------------------- Logger (safe) ----------------------------------

function _getLoggerPlus_() {
  try {
    if (typeof getAppLogger_ === 'function') return getAppLogger_();
  } catch (_) {}
  // Fallback to console-based logger
  return {
    info: (fn, msg, details) => { try { console.log('[INFO]', fn || '', msg || '', details || ''); } catch (_) {} },
    warn: (fn, msg, details) => { try { console.warn('[WARN]', fn || '', msg || '', details || ''); } catch (_) {} },
    error: (fn, msg, details) => { try { console.error('[ERROR]', fn || '', msg || '', details || ''); } catch (_) {} }
  };
}

// ---------------------------- Public API -------------------------------------

/**
 * Performs a full analysis of the spreadsheet, collecting metadata, triggers,
 * menu declarations, functions, and data model details.
 * @returns {Object} A comprehensive analysis object.
 */
function performComprehensiveAnalysis_() {
  const log = _getLoggerPlus_();
  const started = Date.now();
  const fn = 'performComprehensiveAnalysis_';

  const meta = _collectMetadata_();
  const triggers = _collectTriggers_();
  const menuFns = _collectMenuFunctions_();
  const dataModel = _collectDataModel_();
  const functions = _mergeFunctionInventory_(triggers, menuFns);

  const sheetsArr = dataModel.sheets || [];
  const sheetsScanned = sheetsArr.length;
  const totalRows = sheetsArr.reduce((sum, s) => sum + (s.rows || 0), 0);
  const maxCols = sheetsArr.reduce((m, s) => Math.max(m, s.columns || 0), 0);
  const durationMs = Date.now() - started;

  const LD_SHEETS = Number(_cfgGet_('LARGE_DATA_SHEETS', CORE_ANALYSIS_CFG.LARGE_DATA_SHEETS));
  const LD_MAXCOLS = Number(_cfgGet_('LARGE_DATA_MAXCOLS', CORE_ANALYSIS_CFG.LARGE_DATA_MAXCOLS));
  const LD_TOTALROWS = Number(_cfgGet_('LARGE_DATA_TOTALROWS', CORE_ANALYSIS_CFG.LARGE_DATA_TOTALROWS));

  const isLarge = (sheetsScanned >= LD_SHEETS) || (maxCols >= LD_MAXCOLS) || (totalRows >= LD_TOTALROWS);
  if (isLarge) {
    log.info(fn, 'Large dataset detected.', {
      sheetsScanned, maxCols, totalRows,
      thresholds: { LD_SHEETS, LD_MAXCOLS, LD_TOTALROWS }
    });
  }

  const result = {
    metadata: meta,
    triggers: { count: triggers.length, details: triggers },
    menus: { fromSheets: menuFns },
    functions: { global: functions, private: [] },
    sheets: {
      count: dataModel.sheets.length,
      sheets: dataModel.sheets,
      headerDuplicates: dataModel.headerDuplicates
    },
    performanceMetrics: {
      sheetsScanned,
      totalRows,
      maxCols,
      scanDurationMs: durationMs
    },
    version: _cfgGet_('VERSION', CORE_ANALYSIS_CFG.VERSION)
  };

  log.info(fn, 'Full analysis complete.', {
    ms: durationMs,
    sheets: result.sheets.count,
    triggers: result.triggers.count,
    functions: result.functions.global.length
  });
  return result;
}

/**
 * Reads existing requirements from the designated 'Krav' sheet.
 * @returns {Array<Object>} An array of requirement objects, each with id, text, priority, and progressPct.
 */
function readRequirementsForGapAnalysis_() {
  const log = _getLoggerPlus_();
  const fn = 'readRequirementsForGapAnalysis_';
  try {
    const names = _cfgGet_('NAMES', CORE_ANALYSIS_CFG.NAMES);
    const sh = _getSheetByAnyName_(names.kravSheet);
    if (!sh) return [];

    const vals = sh.getDataRange().getValues();
    if (!vals || vals.length < 2) return [];

    const headers = vals[0].map(h => String(h || '').trim().toLowerCase());
    const idxOfAny = (alts) => {
      for (let i = 0; i < headers.length; i++) {
        const h = headers[i];
        for (let j = 0; j < alts.length; j++) {
          if (h === alts[j]) return i;
        }
      }
      return -1;
    };

    const KH = _cfgGet_('HEADERS', CORE_ANALYSIS_CFG.HEADERS).krav;
    const idIdx = idxOfAny(KH.id);
    const textIdx = idxOfAny(KH.text);
    const prioIdx = idxOfAny(KH.priority);
    const progIdx = idxOfAny(KH.progress);

    const out = [];
    for (let r = 1; r < vals.length; r++) {
      const row = vals[r];
      out.push({
        id: (idIdx >= 0 ? row[idIdx] : '') || '',
        text: String((textIdx >= 0 ? row[textIdx] : '') || ''),
        priority: String((prioIdx >= 0 ? row[prioIdx] : '') || ''),
        progressPct: Number((progIdx >= 0 ? row[progIdx] : 0) || 0)
      });
    }
    return out;
  } catch (e) {
    log.error(fn, 'Failed to read requirements.', { error: e.message, stack: e.stack });
    return [];
  }
}

/**
 * Generates potential new requirement candidates based on the analysis object.
 * @param {Object} analysis The object returned by performComprehensiveAnalysis_().
 * @returns {Array<Object>} An array of candidate objects.
 */
function generateRequirementCandidates_(analysis) {
  const log = _getLoggerPlus_();
  const fn = 'generateRequirementCandidates_';
  const A = analysis || {};
  const out = [];

  // From triggers
  (A.triggers && A.triggers.details || []).forEach(t => {
    out.push({
      text: _requirementTextFromTrigger_(t),
      autoPriority: _priorityFromTrigger_(t),
      source: SOURCES.TRIGGER,
      extra: { handler: t.handler, eventType: t.eventType, source: t.source }
    });
  });

  // From menu declarations in sheets
  (A.menus && A.menus.fromSheets || []).forEach(m => {
    const title = m.title || m.functionName || '';
    const fnName = m.functionName || '';
    if (!fnName) return;
    out.push({
      text: `Systemet skal tilby menykommando «${title}» som kaller «${fnName}».`,
      autoPriority: PRIORITIES.SHOULD,
      source: SOURCES.MENU,
      extra: { sheet: m.sheet, role: m.role || '', active: !!m.active }
    });
  });

  // From data fields
  (A.sheets && A.sheets.sheets || []).forEach(s => {
    const headers = _splitHeaderPreview_(s.headerPreview);
    headers.forEach(h => {
      if (!h) return;
      out.push({
        text: `Systemet skal forvalte datafelt «${h}» i arket «${s.name}».`,
        autoPriority: PRIORITIES.SHOULD,
        source: SOURCES.FIELD,
        extra: { sheet: s.name, header: h }
      });
    });
  });

  // Heuristics (domain-driven)
  const fnNames = (A.functions && A.functions.global || []).map(f => String(f.name || '').toLowerCase());
  out.push.apply(out, _domainHeuristicCandidates_(fnNames));

  log.info(fn, 'Requirement candidates generated.', { count: out.length });
  return out;
}

/**
 * Filters a list of new candidates, removing duplicates of existing requirements
 * and other new candidates based on a Jaccard similarity threshold.
 * @param {Array<Object>} candidates The array of new candidates.
 * @param {Array<Object>} existing The array of existing requirements.
 * @param {number} [threshold] The Jaccard similarity threshold (0-1). Overrides config if provided.
 * @returns {Array<Object>} The filtered array of unique candidates.
 */
function dedupeCandidates_(candidates, existing, threshold) {
  const log = _getLoggerPlus_();
  const fn = 'dedupeCandidates_';

  if (!Array.isArray(candidates)) {
    log.warn(fn, 'Invalid input: candidates must be an array.');
    return [];
  }
  if (!Array.isArray(existing)) {
    log.warn(fn, 'Invalid input: existing must be an array. Defaulting to [].');
    existing = [];
  }

  let th = (typeof threshold === 'number' ? threshold : undefined);
  if (typeof th !== 'number' || isNaN(th)) {
    let cfgOverride;
    try {
      if (typeof CONFIG_PLUS !== 'undefined' && CONFIG_PLUS) {
        if (typeof CONFIG_PLUS.DEFAULT_JACCARD_THRESHOLD === 'number') {
          cfgOverride = CONFIG_PLUS.DEFAULT_JACCARD_THRESHOLD;
        } else if (typeof CONFIG_PLUS.DEDUPLE_JACCARD === 'number') {
          cfgOverride = CONFIG_PLUS.DEDUPLE_JACCARD; // legacy
        }
      }
    } catch (_) {}
    th = (typeof cfgOverride === 'number') ? cfgOverride
        : _cfgGet_('DEFAULT_JACCARD_THRESHOLD', CORE_ANALYSIS_CFG.DEFAULT_JACCARD_THRESHOLD);
  }
  if (th < 0 || th > 1) {
    log.warn(fn, 'Threshold out of range, clamping to [0,1].', { provided: threshold });
    th = Math.max(0, Math.min(1, th));
  }

  const existTexts = existing.map(e => String(e.text || ''));
  const seen = [];
  const out = [];

  candidates.forEach(c => {
    const t = String(c.text || '');
    if (!t) return;
    const dupExisting = existTexts.some(et => _jaccard_(t, et) >= th);
    if (dupExisting) return;
    const dupNew = seen.some(s => _jaccard_(t, s) >= th);
    if (dupNew) return;
    seen.push(t);
    out.push(c);
  });

  return out;
}

/**
 * Compares code and requirements to find gaps.
 * @param {Object} analysis The full analysis object.
 * @param {Array<Object>} existing The array of existing requirements.
 * @param {Array<Object>} deduped The array of new, unique candidates.
 * @returns {Object} An object containing lists of unimplementedRequirements and undocumentedFunctions.
 */
function performGapAnalysis_(analysis, existing, deduped) {
  const A = analysis || {};
  const unimpl = (existing || []).filter(r => Number(r.progressPct || 0) === 0);

  const publicFns = (A.functions && A.functions.global || []).map(f => String(f.name || '')).filter(Boolean);
  const kravTekster = (existing || []).map(r => String(r.text || '').toLowerCase());
  const undocumented = [];

  publicFns.forEach(fn => {
    const f = fn.toLowerCase();
    const foundInKrav = kravTekster.some(k => k.indexOf(f) >= 0);
    const foundInCand = (deduped || []).some(c => String(c.text || '').toLowerCase().indexOf(f) >= 0);
    if (!foundInKrav && !foundInCand) {
      undocumented.push({ function: fn });
    }
  });

  return {
    unimplementedRequirements: unimpl,
    undocumentedFunctions: undocumented
  };
}

// ---------------------------- Private Helpers --------------------------------

function _collectMetadata_() {
  const log = _getLoggerPlus_();
  const fn = '_collectMetadata_';
  try {
    const ss = SpreadsheetApp.getActive();
    const userEmail = _safe(() => Session.getActiveUser().getEmail(), '');
    return {
      spreadsheetName: _safe(() => ss.getName(), ''),
      spreadsheetUrl: _safe(() => ss.getUrl(), ''),
      spreadsheetId: _safe(() => ss.getId(), ''),
      timeZone: _safe(() => ss.getSpreadsheetTimeZone(), ''),
      locale: _safe(() => ss.getSpreadsheetLocale && ss.getSpreadsheetLocale(), ''),
      sheetsCount: _safe(() => ss.getSheets().length, 0),
      user: userEmail
    };
  } catch (e) {
    log.error(fn, 'Failed to collect metadata.', { error: e.message });
    return {
      spreadsheetName: '',
      spreadsheetUrl: '',
      spreadsheetId: '',
      timeZone: '',
      locale: '',
      sheetsCount: 0,
      user: ''
    };
  }
}

function _collectTriggers_() {
  const log = _getLoggerPlus_();
  const fn = '_collectTriggers_';
  const out = [];
  try {
    const trig = ScriptApp.getProjectTriggers() || [];
    trig.forEach(t => {
      let eventType = '';
      let source = '';
      let handler = '';
      try { handler = String(t.getHandlerFunction() || ''); } catch (_) {}
      try { eventType = String(t.getEventType && t.getEventType()); } catch (_) {}
      try { source = String(t.getTriggerSource && t.getTriggerSource()); } catch (_) {}
      out.push({
        handler: handler,
        eventType: eventType || 'UNKNOWN',
        source: source || 'UNKNOWN',
        raw: { eventType, source }
      });
    });
  } catch (e) {
    log.error(fn, 'Failed to collect triggers.', { error: e.message });
  }
  return out;
}

function _collectMenuFunctions_() {
  const log = _getLoggerPlus_();
  const fn = '_collectMenuFunctions_';
  const out = [];
  try {
    const names = _cfgGet_('NAMES', CORE_ANALYSIS_CFG.NAMES);
    const shFelles = _getSheetByAnyName_(names.menuFelles);
    const shMin = _getSheetByAnyName_(names.menuMin);
    if (shFelles) out.push.apply(out, _readMenuSheet_(shFelles, 'Meny_Felles'));
    if (shMin) out.push.apply(out, _readMenuSheet_(shMin, 'Meny_Min'));
  } catch (e) {
    log.error(fn, 'Failed reading menu sheets.', { error: e.message });
  }
  return out;
}

function _readMenuSheet_(sh, sheetLabel) {
  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 2) return [];
  const hdr = vals[0].map(h => String(h || '').trim().toLowerCase());
  const idx = (keyList) => {
    for (let k = 0; k < keyList.length; k++) {
      const wanted = keyList[k];
      const pos = hdr.indexOf(wanted);
      if (pos >= 0) return pos;
    }
    return -1;
  };
  const titleIdx = idx(['tittel', 'title', 'kommando', 'menu', 'meny']);
  const fnIdx    = idx(['funksjon', 'function', 'handler']);
  const roleIdx  = idx(['rollekrav', 'rolle', 'role']);
  const userIdx  = idx(['bruker', 'user']);
  const actIdx   = idx(['aktiv', 'active', 'enabled']);

  const out = [];
  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];
    const title = (titleIdx >= 0 ? row[titleIdx] : '') || '';
    const fn = (fnIdx >= 0 ? row[fnIdx] : '') || '';
    if (!fn && !title) continue;
    const role = (roleIdx >= 0 ? row[roleIdx] : '') || '';
    const user = (userIdx >= 0 ? row[userIdx] : '') || '';
    const activeRaw = (actIdx >= 0 ? row[actIdx] : '');
    const active = _truthy_(activeRaw);
    out.push({
      sheet: sheetLabel,
      title: String(title),
      functionName: String(fn),
      role: String(role),
      user: String(user),
      active: active
    });
  }
  return out;
}

function _collectDataModel_() {
  const log = _getLoggerPlus_();
  const fn = '_collectDataModel_';
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets() || [];
  const outSheets = [];
  const headerIndexGlobal = {}; // normalizedHeader -> [{sheet, col}]
  const everyN = Math.max(1, Number(_cfgGet_('PROGRESS_LOG_EVERY_SHEETS', CORE_ANALYSIS_CFG.PROGRESS_LOG_EVERY_SHEETS)));

  for (let i = 0; i < sheets.length; i++) {
    const sh = sheets[i];
    try {
      if (i % everyN === 0) {
        _getLoggerPlus_().info(fn, 'Scanning sheets progress...', { index: i, total: sheets.length });
      }
      const name = sh.getName();
      const rows = sh.getLastRow();
      const cols = sh.getLastColumn();
      const isHidden = (typeof sh.isSheetHidden === 'function') ? sh.isSheetHidden() : false;

      let header = [];
      if (cols > 0) {
        header = sh.getRange(1, 1, 1, cols).getValues()[0] || [];
      }
      const preview = _buildHeaderPreview_(header);
      const typesByHeader = _inferTypesForSheetChunked_(sh, header);

      // Collect header duplicates (across sheets)
      header.forEach((h, idx) => {
        const norm = _normalizeHeader_(h);
        if (!norm) return;
        if (!headerIndexGlobal[norm]) headerIndexGlobal[norm] = [];
        headerIndexGlobal[norm].push({ sheet: name, col: idx + 1 });
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
      _getLoggerPlus_().warn(fn, 'Failed scanning sheet (skipping).', { sheet: _safe(() => sheets[i].getName(), `#${i+1}`), error: e.message });
    }
  }

  // Compute duplicates list
  const duplicates = [];
  Object.keys(headerIndexGlobal).forEach(h => {
    const occ = headerIndexGlobal[h];
    if (occ && occ.length > 1) {
      duplicates.push({ header: h, occurrences: occ });
    }
  });

  return { sheets: outSheets, headerDuplicates: duplicates };
}

function _inferTypesForSheetChunked_(sh, headers) {
  const rowsToScan = Math.max(0, Math.min(Number(_cfgGet_('MAX_SCAN_ROWS', CORE_ANALYSIS_CFG.MAX_SCAN_ROWS)), Math.max(0, sh.getLastRow() - 1)));
  const totalCols = sh.getLastColumn();
  const chunkSize = Math.max(1, Number(_cfgGet_('SCAN_COL_CHUNK', CORE_ANALYSIS_CFG.SCAN_COL_CHUNK)));

  const out = {};
  if (rowsToScan <= 0 || totalCols <= 0) return out;

  let colIndex = 1;
  while (colIndex <= totalCols) {
    const thisChunk = Math.min(chunkSize, totalCols - colIndex + 1);
    const range2D = sh.getRange(2, colIndex, rowsToScan, thisChunk).getValues(); // rowsToScan x thisChunk
    for (let c = 0; c < thisChunk; c++) {
      const headerName = String(headers[colIndex - 1 + c] || '').trim();
      if (!headerName) continue;
      const samples = [];
      for (let r = 0; r < range2D.length; r++) {
        samples.push(range2D[r][c]);
      }
      out[headerName] = _inferTypeFromSamples_(samples);
    }
    colIndex += thisChunk;
  }
  return out;
}

function _inferTypeFromSamples_(arr) {
  // Heuristics with priority: date > number > boolean > email > url > string/empty
  let hasDate = false, hasNumber = false, hasBool = false, hasEmail = false, hasUrl = false;
  let nonEmpty = 0;
  const emailRx = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i;
  const urlRx = /^(https?:\/\/|www\.)/i;

  for (let i = 0; i < arr.length; i++) {
    const v = arr[i];
    if (v === '' || v === null || typeof v === 'undefined') continue;
    nonEmpty++;

    if (v instanceof Date) { hasDate = true; continue; }
    if (typeof v === 'number' && !isNaN(v)) { hasNumber = true; continue; }
    if (typeof v === 'boolean') { hasBool = true; continue; }

    const s = String(v).trim();
    if (!s) continue;
    if (!isNaN(Number(s))) { hasNumber = true; continue; }
    if (s.toLowerCase() === 'true' || s.toLowerCase() === 'false' || s.toLowerCase() === 'ja' || s.toLowerCase() === 'nei') {
      hasBool = true; continue;
    }
    if (emailRx.test(s)) { hasEmail = true; continue; }
    if (urlRx.test(s)) { hasUrl = true; continue; }
  }

  if (nonEmpty === 0) return 'empty';
  if (hasDate) return 'date';
  if (hasNumber) return 'number';
  if (hasBool) return 'boolean';
  if (hasEmail) return 'email';
  if (hasUrl) return 'url';
  return 'string';
}

function _mergeFunctionInventory_(triggers, menus) {
  const set = {};
  const out = [];
  (triggers || []).forEach(t => {
    const n = String(t.handler || '').trim();
    if (!n || set[n]) return;
    set[n] = true;
    out.push({ name: n, source: 'trigger', eventType: t.eventType || '' });
  });
  (menus || []).forEach(m => {
    const n = String(m.functionName || '').trim();
    if (!n || set[n]) return;
    set[n] = true;
    out.push({ name: n, source: 'menu', title: m.title || '' });
  });
  return out;
}

function _getSheetByAnyName_(candidates) {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets() || [];
  const cand = (Array.isArray(candidates) ? candidates : [candidates]).map(s => _normalizeName_(String(s || '')));
  for (let i = 0; i < sheets.length; i++) {
    const n = _normalizeName_(sheets[i].getName());
    for (let j = 0; j < cand.length; j++) {
      if (n === cand[j]) return sheets[i];
    }
  }
  // Also try a looser match (ignoring underscores/spaces fully)
  for (let i = 0; i < sheets.length; i++) {
    const n = _normalizeName_(sheets[i].getName(), true);
    for (let j = 0; j < cand.length; j++) {
      if (n === _normalizeName_(cand[j], true)) return sheets[i];
    }
  }
  return null;
}

function _normalizeName_(s, stripAll) {
  let out = String(s || '').toLowerCase().trim();
  out = out.replace(/\s+/g, stripAll ? '' : ' ');
  out = out.replace(/_/g, stripAll ? '' : '_');
  return out;
}

function _buildHeaderPreview_(headerArr) {
  const max = Math.max(0, Number(_cfgGet_('MAX_HEADER_PREVIEW', CORE_ANALYSIS_CFG.MAX_HEADER_PREVIEW)));
  const preview = (headerArr || []).slice(0, max).map(h => String(h || '').trim()).filter(Boolean);
  return preview.join(' | ');
}

function _splitHeaderPreview_(s) {
  if (Array.isArray(s)) return s;
  if (!s) return [];
  return String(s).split('|').map(x => String(x || '').trim()).filter(Boolean);
}

function _requirementTextFromTrigger_(t) {
  const evt = String(t.eventType || '').toUpperCase();
  const handler = t.handler || '';
  if (evt.indexOf('CLOCK') >= 0 || evt.indexOf('TIME') >= 0) {
    return `Systemet skal periodisk kjøre «${handler}» (tidsstyrt).`;
  }
  if (evt.indexOf('FORM_SUBMIT') >= 0) {
    return `Ved innsending av skjema skal systemet prosessere via «${handler}».`;
  }
  if (evt.indexOf('OPEN') >= 0) {
    return `Ved åpning av regnearket skal systemet kjøre «${handler}».`;
  }
  if (evt.indexOf('EDIT') >= 0) {
    return `Ved endring i regnearket skal systemet kjøre «${handler}».`;
  }
  return `Systemet skal støtte hendelsen «${evt}» via «${handler}».`;
}

function _priorityFromTrigger_(t) {
  const evt = String(t.eventType || '').toUpperCase();
  if (evt.indexOf('CLOCK') >= 0 || evt.indexOf('TIME') >= 0) return PRIORITIES.MUST;
  if (evt.indexOf('FORM_SUBMIT') >= 0) return PRIORITIES.SHOULD;
  if (evt.indexOf('OPEN') >= 0) return PRIORITIES.SHOULD;
  if (evt.indexOf('EDIT') >= 0) return PRIORITIES.SHOULD;
  return PRIORITIES.COULD;
}

function _domainHeuristicCandidates_(fnNamesLower) {
  const out = [];
  const seen = {};
  const add = (text, priority, extra) => {
    if (seen[text]) return;
    seen[text] = true;
    out.push({ text, autoPriority: priority || PRIORITIES.SHOULD, source: SOURCES.HEURISTIC, extra: extra || {} });
  };

  const joined = ' ' + (fnNamesLower || []).join(' ') + ' ';

  if (/\bhms\b/.test(joined)) {
    add('Systemet skal sikre at HMS-planer genereres, varsles og synkroniseres i kalender.', PRIORITIES.MUST, { area: 'HMS' });
  }
  if (/\bvaktmester\b/.test(joined)) {
    add('Systemet skal la vaktmester motta, oppdatere og ferdigstille oppgaver.', PRIORITIES.SHOULD, { area: 'Tasks' });
  }
  if (/\bbudget\b|\bbudsjett\b/.test(joined)) {
    add('Systemet skal støtte budsjetthåndtering med validering, import og rapportering.', PRIORITIES.SHOULD, { area: 'Budget' });
  }
  if (/\bvote\b|\bvoter\b|\bstemme\b/.test(joined)) {
    add('Systemet skal støtte digital stemmegivning med oppsummering og låsing av vedtak.', PRIORITIES.SHOULD, { area: 'Møter' });
  }
  if (/\bmeeting\b|\bmøte\b|\bmoter\b/.test(joined)) {
    add('Systemet skal forvalte møter, agenda og protokoll for godkjenning.', PRIORITIES.SHOULD, { area: 'Møter' });
  }
  if (/\brbac\b|\brole\b|\btilgang\b/.test(joined)) {
    add('Systemet skal håndheve rollebasert tilgangsstyring (RBAC) for brukerhandlinger.', PRIORITIES.MUST, { area: 'Security' });
  }

  return out;
}

function _normalizeHeader_(h) {
  return String(h || '').trim().toLowerCase();
}

function _truthy_(v) {
  const s = String(v).trim().toLowerCase();
  if (!s) return false;
  if (s === '1' || s === 'true' || s === 'ja' || s === 'x' || s === 'on') return true;
  return !!v;
}

// -- small safe util
function _safe(fn, fallback) {
  try { return fn(); } catch (_) { return fallback; }
}

// ---------------------------- Text Similarity --------------------------------

function _tokenize_(txt) {
  const s = String(txt || '').toLowerCase();
  const raw = s.split(/[^a-z0-9æøå]+/i).filter(Boolean);
  // Drop very short tokens
  return raw.filter(t => t.length >= 2);
}

function _jaccard_(a, b) {
  const A = _tokenize_(a);
  const B = _tokenize_(b);
  if (A.length === 0 && B.length === 0) return 1;
  const setA = {};
  A.forEach(x => { setA[x] = true; });
  let inter = 0;
  const setB = {};
  B.forEach(y => { setB[y] = true; if (setA[y]) inter++; });
  const union = Object.keys(setA).length + Object.keys(setB).length - inter;
  return union === 0 ? 0 : inter / union;
}

// ---------------------------- Optional Smoke Test ----------------------------
// You can keep or remove these helpers. They’re useful during rollout.

/**
 * Quick smoke test: runs analysis and writes an OK row to a report-ish sheet.
 */
function runCoreAnalysis_Smoke() {
  const res = performComprehensiveAnalysis_();
  const m = res.performanceMetrics || {};
  try { _getLoggerPlus_().info('runCoreAnalysis_Smoke', 'Core analysis metrics', m); } catch(_) {}
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Integritetsrapport') || ss.getSheetByName('Rapport') || ss.insertSheet('Integritetsrapport');
  if (sh.getLastRow() === 0) sh.appendRow(['Kj.Dato','Kategori','Nøkkel','Status','Detaljer']);
  sh.appendRow([new Date(), 'Analyse', 'CoreAnalysisPlus', 'OK', JSON.stringify(m)]);
}
