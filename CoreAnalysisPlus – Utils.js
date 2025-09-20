/**
 * CoreAnalysisPlus – Utilities (v1.3.4)
 *
 * Purpose:
 *  - Small, defensive helper functions shared across CoreAnalysisPlus modules.
 *  - Safe fallbacks if Config/Logger modules are not loaded yet.
 *
 * Highlights:
 *  - Safe config bridging to _numCfg_/CORE_ANALYSIS_CFG with local fallback
 *  - Tokenization tuned for Norwegian (æøå) without the unnecessary /i flag
 *  - Set-based Jaccard similarity (faster & cleaner)
 *  - Locale-aware truthy parsing (no/en) with legacy compatibility
 *  - Robust JSON stringify that handles circular structures
 *  - Bounded memoization cache for tokenization
 */

// ---------------------------- Local Safe Bridges -----------------------------

/**
 * Safe logger shim with a debug level. If LoggerPlus is present, uses it.
 * @returns {{debug:function,info:function,warn:function,error:function}}
 */
function _getLoggerPlus_() {
  try {
    if (typeof getAppLogger_ === 'function') return getAppLogger_();
  } catch (_) {}
  // Console fallback
  return {
    debug: function (fn, msg, details) { try { console.log('[DEBUG]', fn || '', msg || '', details || ''); } catch (_) {} },
    info:  function (fn, msg, details) { try { console.log('[INFO]',  fn || '', msg || '', details || ''); } catch (_) {} },
    warn:  function (fn, msg, details) { try { console.warn('[WARN]',  fn || '', msg || '', details || ''); } catch (_) {} },
    error: function (fn, msg, details) { try { console.error('[ERROR]', fn || '', msg || '', details || ''); } catch (_) {} }
  };
}

/**
 * Safe numeric config fetcher.
 * Prefers _numCfg_ from Config module; falls back to CORE_ANALYSIS_CFG; then to provided fallback.
 * @param {string} key
 * @param {number} fallback
 * @returns {number}
 */
function _numCfgSafe_(key, fallback) {
  try {
    if (typeof _numCfg_ === 'function') {
      var v = _numCfg_(key, fallback);
      return (typeof v === 'number' && !isNaN(v)) ? v : Number(fallback);
    }
  } catch (_) {}
  try {
    if (typeof CORE_ANALYSIS_CFG !== 'undefined' && CORE_ANALYSIS_CFG && Object.prototype.hasOwnProperty.call(CORE_ANALYSIS_CFG, key)) {
      var n = Number(CORE_ANALYSIS_CFG[key]);
      return isNaN(n) ? Number(fallback) : n;
    }
  } catch (_) {}
  return Number(fallback);
}

// ------------------------------- Safe Utils ---------------------------------

/**
 * Run a function and return fallback on any error.
 * @template T
 * @param {function():T} fn
 * @param {*} fallback
 * @returns {T|*}
 */
function _safe(fn, fallback) {
  try { return fn(); } catch (_) { return fallback; }
}

/**
 * Normalize a display name to a comparison-friendly key.
 * @param {string} s
 * @param {boolean} [stripAll=false] If true, removes spaces and underscores entirely.
 * @returns {string}
 */
function _normalizeName_(s, stripAll) {
  var out = String(s || '').toLowerCase().trim();
  out = out.replace(/\s+/g, stripAll ? '' : ' ');
  out = out.replace(/_/g, stripAll ? '' : '_');
  return out;
}

/**
 * Normalize header text to a canonical form.
 * @param {string} h
 * @returns {string}
 */
function _normalizeHeader_(h) {
  return String(h || '').trim().toLowerCase();
}

/**
 * Find the first index where any of the alternative header names match.
 * @param {string[]} headersLower Lowercased header array.
 * @param {string[]|null|undefined} alts Alternative names (lowercased or raw).
 * @returns {number} index or -1 if not found
 */
function _indexOfHeaderAny_(headersLower, alts) {
  if (!Array.isArray(headersLower) || !Array.isArray(alts)) return -1;
  // Normalize alts once
  var wants = alts.map(function (x) { return String(x || '').trim().toLowerCase(); });
  for (var i = 0; i < headersLower.length; i++) {
    var h = String(headersLower[i] || '').trim().toLowerCase();
    for (var j = 0; j < wants.length; j++) {
      if (h === wants[j]) return i;
    }
  }
  return -1;
}

/**
 * Build a preview string from the header row with a configurable cap.
 * @param {Array<*>} headerArr
 * @returns {string}
 */
function _buildHeaderPreview_(headerArr) {
  var max = _numCfgSafe_('MAX_HEADER_PREVIEW', 10);
  var preview = (headerArr || [])
    .slice(0, Math.max(0, max))
    .map(function (h) { return String(h || '').trim(); })
    .filter(Boolean);
  return preview.join(' | ');
}

/**
 * Split a previously built header preview back into an array.
 * @param {string|string[]} s
 * @returns {string[]}
 */
function _splitHeaderPreview_(s) {
  if (Array.isArray(s)) return s;
  if (!s) return [];
  return String(s).split('|').map(function (x) { return String(x || '').trim(); }).filter(Boolean);
}

// ------------------------- Tokenizing & Similarity ---------------------------

// small bounded memoization for tokenization to speed up repeated comparisons
var __tokenCache = Object.create(null);
var __tokenCacheKeys = [];
var __tokenCacheMax = 500;

/**
 * Tokenize a string into comparable terms.
 * - Lowercases
 * - Splits on non [a-z0-9æøå]
 * - Filters tokens shorter than TOKEN_MIN_LEN
 * @param {string} txt
 * @returns {string[]}
 */
function _tokenize_(txt) {
  var key = 'k:' + String(txt || '');
  if (__tokenCache[key]) return __tokenCache[key].slice(); // return copy

  var minLen = _numCfgSafe_('TOKEN_MIN_LEN', 2);
  var s = String(txt || '').toLowerCase();
  // Note: ‘i’ flag not needed since we lowercase first.
  var raw = s.split(/[^a-z0-9æøå]+/).filter(Boolean);
  var out = raw.filter(function (t) { return t.length >= minLen; });

  // store in bounded cache
  __tokenCache[key] = out;
  __tokenCacheKeys.push(key);
  if (__tokenCacheKeys.length > __tokenCacheMax) {
    var old = __tokenCacheKeys.shift();
    try { delete __tokenCache[old]; } catch (_) {}
  }

  return out.slice();
}

/**
 * Calculate Jaccard similarity between two texts (0..1).
 * Uses Set-based intersection/union for performance.
 * @param {string} a
 * @param {string} b
 * @returns {number}
 */
function _jaccard_(a, b) {
  if (a == null || b == null) return 0;
  var A = _tokenize_(a);
  var B = _tokenize_(b);
  if (A.length === 0 && B.length === 0) return 1;
  if (A.length === 0 || B.length === 0) return 0;

  var setA = new Set(A);
  var setB = new Set(B);

  var inter = 0;
  setA.forEach(function (x) { if (setB.has(x)) inter++; });

  var union = new Set();
  setA.forEach(function (x) { union.add(x); });
  setB.forEach(function (x) { union.add(x); });

  return union.size === 0 ? 0 : inter / union.size;
}

// -------------------------- Locale Truthy Parsing ----------------------------

/**
 * Locale-aware truthy parsing with legacy compatibility.
 * @param {*} v
 * @param {'no'|'en'} [locale='no']
 * @returns {boolean}
 */
function _truthy_(v, locale) {
  var s = String(v).trim().toLowerCase();
  if (!s) return false;
  var loc = (locale || 'no').toLowerCase();

  var truthyWords = {
    en: ['1', 'true', 'yes', 'y', 'on', 'enabled', 'x'],
    no: ['1', 'true', 'ja', 'j', 'på', 'x']
  };
  var words = truthyWords[loc] || truthyWords.en;
  return words.indexOf(s) !== -1;
}

// ---------------------------- JSON stringify safe ----------------------------

/**
 * Safe JSON stringify with circular reference protection.
 * @param {*} obj
 * @param {number} [space=0]
 * @returns {string}
 */
function _stringifySafe_(obj, space) {
  if (obj === null || typeof obj === 'undefined') return 'null';
  try {
    var seen = (typeof WeakSet !== 'undefined') ? new WeakSet() : [];
    return JSON.stringify(obj, function (k, v) {
      if (v && typeof v === 'object') {
        if (seen.add) {
          if (seen.has(v)) return '[Circular]';
          seen.add(v);
        } else {
          if (seen.indexOf(v) !== -1) return '[Circular]';
          seen.push(v);
        }
      }
      return v;
    }, space || 0);
  } catch (e) {
    return '<<JSON Error: ' + (e && e.message) + '>>';
  }
}

// ------------------------------- Exports note --------------------------------
/**
 * This file intentionally does not use module systems. All functions are defined
 * on the global scope (Apps Script pattern) and are intended to be used by
 * other CoreAnalysisPlus files:
 *
 * - _safe
 * - _normalizeName_
 * - _normalizeHeader_
 * - _indexOfHeaderAny_
 * - _buildHeaderPreview_
 * - _splitHeaderPreview_
 * - _tokenize_
 * - _jaccard_
 * - _truthy_
 * - _stringifySafe_
 * - _getLoggerPlus_
 * - _numCfgSafe_
 *
 * It expects (optionally) that a Config file provides:
 *   CORE_ANALYSIS_CFG and _numCfg_
 * If not, _numCfgSafe_ uses local fallbacks so consumers don’t crash.
 */
