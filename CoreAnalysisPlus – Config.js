/**
 * CoreAnalysisPlus – Config (v1.3.3)
 * Base configuration + safe getters (overridable via global CONFIG_PLUS) + schema validation.
 *
 * Highlights:
 * - Clean override path: CONFIG_PLUS -> CORE_ANALYSIS_CFG -> fallback
 * - i18n-ready requirement templates (no, en)
 * - Extended regex patterns (email, url, phone, percent, isoDate, currency-like)
 * - JSDoc + lightweight schema/type validation at load
 */

// ---------------------------- Configuration ----------------------------------
const CORE_ANALYSIS_CFG = {
  VERSION: '1.3.3',

  // Scanning & performance
  MAX_SCAN_ROWS: 25,
  MAX_HEADER_PREVIEW: 10,
  SCAN_COL_CHUNK: 50,
  DEFAULT_JACCARD_THRESHOLD: 0.78,
  PROGRESS_LOG_EVERY_SHEETS: 5,
  LARGE_DATA_SHEETS: 50,
  LARGE_DATA_MAXCOLS: 100,
  LARGE_DATA_TOTALROWS: 50000,

  // Tokenization / data reading conventions
  TOKEN_MIN_LEN: 2,
  DATA_START_ROW: 2,

  // Requirement templates (basic NO + EN)
  REQUIREMENT_TEMPLATES: {
    no: {
      trigger_clock: function (handler) { return 'Systemet skal periodisk kjøre «' + handler + '» (tidsstyrt).'; },
      trigger_form_submit: function (handler) { return 'Ved innsending av skjema skal systemet prosessere via «' + handler + '».'; },
      trigger_open: function (handler) { return 'Ved åpning av regnearket skal systemet kjøre «' + handler + '».'; },
      trigger_edit: function (handler) { return 'Ved endring i regnearket skal systemet kjøre «' + handler + '».'; },
      trigger_generic: function (evt, handler) { return 'Systemet skal støtte hendelsen «' + evt + '» via «' + handler + '».'; },
      menu_item: function (title, fnName) { return 'Systemet skal tilby menykommando «' + title + '» som kaller «' + fnName + '».'; },
      field_item: function (field, sheet) { return 'Systemet skal forvalte datafelt «' + field + '» i arket «' + sheet + '».'; }
    },
    en: {
      trigger_clock: function (handler) { return 'System shall periodically run “' + handler + '” (time-driven).'; },
      trigger_form_submit: function (handler) { return 'On form submission, system shall process via “' + handler + '”.'; },
      trigger_open: function (handler) { return 'On spreadsheet open, system shall run “' + handler + '”.'; },
      trigger_edit: function (handler) { return 'On spreadsheet edit, system shall run “' + handler + '”.'; },
      trigger_generic: function (evt, handler) { return 'System shall support event “' + evt + '” via “' + handler + '”.'; },
      menu_item: function (title, fnName) { return 'System shall expose menu command “' + title + '” calling “' + fnName + '”.'; },
      field_item: function (field, sheet) { return 'System shall manage data field “' + field + '” in sheet “' + sheet + '”.'; }
    }
  },

  // Common patterns (kept simple; tune to your domain)
  REGEX: {
    email: /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/i,
    url: /^(https?:\/\/|www\.)/i,
    phone: /^[+]?[\d\s().-]{6,20}$/,
    percent: /^\s*\d{1,3}(\.\d+)?\s*%?\s*$/,
    isoDate: /^\d{4}-\d{2}-\d{2}$/,
    currencyLike: /^\s*[-+]?\s*\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{1,2})?\s*(?:kr|nok|€|\$)?\s*$/i
  },

  // Friendly names (NO/EN variants) for sheet/header matching
  NAMES: {
    kravSheet: ['Krav', 'Requirements', 'KRAV'],
    menuFelles: ['Meny_Felles', 'Meny Felles', 'MENY_FELLES'],
    menuMin: ['Meny_Min', 'Meny Min', 'MENY_MIN']
  },

  HEADERS: {
    krav: {
      id:       ['id', 'krav id', 'kravid', 'krav-id', 'requirement id'],
      text:     ['krav', 'beskrivelse', 'tekst', 'hva', 'requirement', 'description', 'text'],
      priority: ['prio.', 'prioritet', 'pri', 'priority'],
      progress: ['fremdrift %', 'fremdrift%', 'fremdrift', 'progress', 'progress %', '%']
    }
  }
};

// Common literals used by other modules (kept here for one import point)
const PRIORITIES = { MUST: 'MÅ', SHOULD: 'BØR', COULD: 'KAN' };
const SOURCES =   { TRIGGER: 'trigger', MENU: 'menu', FIELD: 'field', HEURISTIC: 'heuristikk' };

// ---------------------------- Schema / Validation ----------------------------
/**
 * Minimal schema ensuring numbers are numbers and within sane bounds.
 * Extend as needed if you want stricter validation.
 */
const _CORE_CFG_SCHEMA = {
  MAX_SCAN_ROWS:               { type: 'number', min: 0, max: 10000 },
  MAX_HEADER_PREVIEW:          { type: 'number', min: 0, max: 2000 },
  SCAN_COL_CHUNK:              { type: 'number', min: 1, max: 1000 },
  DEFAULT_JACCARD_THRESHOLD:   { type: 'number', min: 0, max: 1 },
  PROGRESS_LOG_EVERY_SHEETS:   { type: 'number', min: 1, max: 10000 },
  LARGE_DATA_SHEETS:           { type: 'number', min: 0, max: 10000 },
  LARGE_DATA_MAXCOLS:          { type: 'number', min: 0, max: 10000 },
  LARGE_DATA_TOTALROWS:        { type: 'number', min: 0, max: 5000000 },
  TOKEN_MIN_LEN:               { type: 'number', min: 1, max: 10 },
  DATA_START_ROW:              { type: 'number', min: 1, max: 1000 }
};

var _CORE_CFG_VALIDATED = false;

/**
 * Run once on load to self-validate numeric fields and clamp into bounds.
 * This protects downstream callers even if CONFIG_PLUS overrides are off-range.
 */
function _validateCoreCfgOnce_() {
  if (_CORE_CFG_VALIDATED) return;
  _CORE_CFG_VALIDATED = true;

  // Validate base config
  _clampCfgObject_(CORE_ANALYSIS_CFG, _CORE_CFG_SCHEMA);

  // Validate overrides if present
  try {
    if (typeof CONFIG_PLUS !== 'undefined' && CONFIG_PLUS && typeof CONFIG_PLUS === 'object') {
      _clampCfgObject_(CONFIG_PLUS, _CORE_CFG_SCHEMA);
    }
  } catch (e) {
    // ignore
  }
}

function _clampCfgObject_(obj, schema) {
  var log = _getLoggerPlus_ ? _getLoggerPlus_() : { warn: function(){} };
  Object.keys(schema).forEach(function (key) {
    var rule = schema[key];
    var hasOwn = Object.prototype.hasOwnProperty.call(obj, key);
    if (!hasOwn) return;

    var raw = obj[key];
    if (rule.type === 'number') {
      var n = Number(raw);
      if (isNaN(n)) {
        try { log.warn('_clampCfgObject_', 'Non-numeric config, using default if available', { key: key, value: raw }); } catch (_e) {}
        return;
      }
      if (typeof rule.min === 'number' && n < rule.min) n = rule.min;
      if (typeof rule.max === 'number' && n > rule.max) n = rule.max;
      obj[key] = n;
    }
  });
}

// Make sure validation runs when file loads
_validateCoreCfgOnce_();

// ---------------------------- Config helpers ---------------------------------
/**
 * Safe config getter (shallow). Looks in CONFIG_PLUS first, then falls back to CORE_ANALYSIS_CFG,
 * then to provided fallback.
 * @param {string} key
 * @param {*} fallback
 * @returns {*}
 */
function _cfgGet_(key, fallback) {
  if (typeof key !== 'string' || key.length === 0) {
    throw new Error('Config key must be a non-empty string');
  }
  try {
    if (typeof CONFIG_PLUS !== 'undefined' &&
        CONFIG_PLUS &&
        Object.prototype.hasOwnProperty.call(CONFIG_PLUS, key)) {
      return CONFIG_PLUS[key];
    }
  } catch (e) {
    // ignore
  }
  return Object.prototype.hasOwnProperty.call(CORE_ANALYSIS_CFG, key)
    ? CORE_ANALYSIS_CFG[key]
    : fallback;
}

/**
 * Safe deep config getter using dot-notation (e.g., "REQUIREMENT_TEMPLATES.no").
 * Checks CONFIG_PLUS first, then CORE_ANALYSIS_CFG. Returns fallback if not found.
 * @param {string} path
 * @param {*} fallback
 * @returns {*}
 */
function _cfgDeep_(path, fallback) {
  if (typeof path !== 'string' || path.length === 0) {
    throw new Error('Config path must be a non-empty string');
  }

  var segs = path.split('.');
  var traverse = function (root) {
    var cur = root;
    for (var i = 0; i < segs.length; i++) {
      var seg = segs[i];
      if (!cur || typeof cur !== 'object' || !Object.prototype.hasOwnProperty.call(cur, seg)) {
        return undefined;
      }
      cur = cur[seg];
    }
    return cur;
  };

  try {
    var fromPlus = (typeof CONFIG_PLUS !== 'undefined' && CONFIG_PLUS) ? traverse(CONFIG_PLUS) : undefined;
    if (fromPlus !== undefined) return fromPlus;
  } catch (e) {
    // ignore
  }

  var fromCore = traverse(CORE_ANALYSIS_CFG);
  return (fromCore !== undefined) ? fromCore : fallback;
}

/**
 * Numeric config helper with NaN safety and fallbacks.
 * @param {string} key
 * @param {number} fallback
 * @returns {number}
 */
function _numCfg_(key, fallback) {
  var v = Number(_cfgGet_(key, fallback));
  return isNaN(v) ? Number(fallback) : v;
}

// ---------------------------- Minimal Logger Shim ----------------------------
/**
 * Small shim so this file can log warnings during validation even if LoggerPlus
 * isn’t loaded yet. If LoggerPlus exists later, it’ll be used by other modules.
 */
function _getLoggerPlus_() {
  try {
    if (typeof getAppLogger_ === 'function') return getAppLogger_();
  } catch (e) {
    // ignore
  }
  return { warn: function(){}, info: function(){}, error: function(){} };
}
