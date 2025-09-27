/**
 * LoggerPlus (v1.3.0)
 * Production-grade logging for Google Apps Script.
 *
 * Features
 * - Batch writes to a "Logg" sheet (configurable) to reduce API calls
 * - Log rotation (keeps sheet size under control)
 * - Log levels with runtime filtering (ERROR, WARN, INFO, DEBUG)
 * - Daily quota guard using ScriptProperties
 * - Sensitive data sanitization
 * - JSON stringify with circular reference protection
 * - Cached sheet handle with TTL (auto re-validate)
 * - Optional external endpoint integration (POST)
 * - Simple stats API for observability
 *
 * Configuration (optional)
 *   const CONFIG_PLUS = {
 *     LOGGER: {
 *       SHEET_NAME: 'Logg',
 *       BATCH_SIZE: 10,
 *       ROTATE_MAX_ROWS: 10000,
 *       ROTATE_KEEP_RATIO: 0.7,
 *       MAX_DAILY_LOGS: 5000,
 *       SENSITIVE_KEYS: ['password','token','key','secret','apiKey','authorization','email'],
 *       CACHE_TTL_MS: 5 * 60 * 1000,
 *       LEVEL: 'INFO', // 'ERROR' | 'WARN' | 'INFO' | 'DEBUG'
 *       EXTERNAL_ENABLED: false,
 *       EXTERNAL_ENDPOINT: '', // URL
 *       EXTERNAL_TIMEOUT_MS: 8000,
 *       EXTERNAL_MAX_RETRIES: 2
 *     }
 *   };
 *
 * Usage
 *   const log = getAppLogger_();
 *   log.setLevel('INFO');
 *   log.info('myFunction', 'Started processing', { user: Session.getActiveUser().getEmail() });
 *   try { ... } catch (e) {
 *     log.error('myFunction', 'Unexpected error', { error: e.message, stack: e.stack });
 *   } finally {
 *     log.flush();
 *   }
 */

function getAppLogger_() {
  if (getAppLogger_._singleton) return getAppLogger_._singleton;

  // --------------------- Defaults + Config Overrides ---------------------
  var defaults = {
    SHEET_NAME: 'Logg',
    BATCH_SIZE: 10,
    ROTATE_MAX_ROWS: 10000,
    ROTATE_KEEP_RATIO: 0.7,
    MAX_DAILY_LOGS: 5000,
    SENSITIVE_KEYS: ['password','token','key','secret','apiKey','authorization','email'],
    CACHE_TTL_MS: 5 * 60 * 1000, // 5 min
    LEVELS: { ERROR: 0, WARN: 1, INFO: 2, DEBUG: 3 },
    LEVEL: 'INFO',
    EXTERNAL_ENABLED: false,
    EXTERNAL_ENDPOINT: '',
    EXTERNAL_TIMEOUT_MS: 8000,
    EXTERNAL_MAX_RETRIES: 2,
    MAX_BUFFER_SIZE: 100 // hard cap to avoid runaway buffering
  };

  var cfg = (function resolveCfg() {
    try {
      if (typeof CONFIG_PLUS !== 'undefined' && CONFIG_PLUS && CONFIG_PLUS.LOGGER) {
        var out = {};
        var keys = Object.keys(defaults);
        for (var i = 0; i < keys.length; i++) {
          var k = keys[i];
          out[k] = (CONFIG_PLUS.LOGGER.hasOwnProperty(k)) ? CONFIG_PLUS.LOGGER[k] : defaults[k];
        }
        // Keep LEVELS map intact
        out.LEVELS = defaults.LEVELS;
        return out;
      }
    } catch (_) {}
    return defaults;
  })();

  // Validate basic cfg
  if (cfg.BATCH_SIZE < 1 || cfg.BATCH_SIZE > 100) cfg.BATCH_SIZE = defaults.BATCH_SIZE;
  if (cfg.ROTATE_KEEP_RATIO <= 0 || cfg.ROTATE_KEEP_RATIO >= 1) cfg.ROTATE_KEEP_RATIO = defaults.ROTATE_KEEP_RATIO;

  // ----------------------------- State -----------------------------------
  var _buffer = [];
  var _sheet = null;
  var _spreadsheet = null;
  var _lastCheck = 0;
  var _currentLevel = (cfg.LEVELS[cfg.LEVEL] != null) ? cfg.LEVELS[cfg.LEVEL] : cfg.LEVELS.INFO;
  var _stats = {
    written: 0,
    buffered: 0,
    rotated: 0,
    droppedByQuota: 0,
    flushCount: 0,
    avgFlushTime: 0,
    lastFlushTime: null
  };

  // ------------------------- Internal Helpers ----------------------------
  function _nowStr_() {
    try {
      var tz = Session.getScriptTimeZone ? Session.getScriptTimeZone() : 'Etc/UTC';
      return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
    } catch (_) {
      return new Date().toISOString();
    }
  }

  function _shouldLog_(levelStr) {
    var lvl = cfg.LEVELS[levelStr];
    return lvl != null && lvl <= _currentLevel;
  }

  function _sanitizeDetails_(details) {
    if (!details || typeof details !== 'object') return details;
    var out = {};
    var keys = Object.keys(details);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      var v = details[k];
      var lower = k.toLowerCase();
      var sensitive = false;
      for (var j = 0; j < cfg.SENSITIVE_KEYS.length; j++) {
        if (lower.indexOf(String(cfg.SENSITIVE_KEYS[j]).toLowerCase()) !== -1) {
          sensitive = true;
          break;
        }
      }
      out[k] = sensitive ? '[REDACTED]' : v;
    }
    return out;
  }

  /*
   * MERK: _stringifySafe_() er fjernet fra denne filen. Den globale
   * versjonen fra 000_Utils.js brukes i stedet.
   */

  function _props_() {
    try { return PropertiesService.getScriptProperties(); } catch (_) {}
    return null;
  }

  function _quotaKey_() {
    var d = new Date();
    var day = d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
    return 'LOGGER_DAILY_' + day;
  }

  function _incDailyCountAndCheckQuota_() {
    if (!_props_()) return true; // if properties unavailable, don't block
    var p = _props_();
    var key = _quotaKey_();
    var val = Number(p.getProperty(key) || 0);
    if (val >= cfg.MAX_DAILY_LOGS) {
      _stats.droppedByQuota++;
      return false;
    }
    p.setProperty(key, String(val + 1));
    return true;
  }

  function _ensureSheet_() {
    var now = Date.now();
    if (_sheet && (now - _lastCheck) < cfg.CACHE_TTL_MS) return _sheet;

    _lastCheck = now;

    try {
      if (!_spreadsheet) _spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var sh = _spreadsheet.getSheetByName(cfg.SHEET_NAME);
      if (!sh) {
        sh = _spreadsheet.insertSheet(cfg.SHEET_NAME);
        sh.getRange('A1:E1').setValues([[
          'Tidsstempel', 'Nivå', 'Funksjon', 'Melding', 'Detaljer (JSON)'
        ]]);
        sh.getRange('A1:E1').setFontWeight('bold');
        sh.setFrozenRows(1);
      } else {
        // Ensure headers exist (idempotent)
        var rng = sh.getRange(1, 1, 1, 5);
        var vals = rng.getValues();
        if (!vals || !vals[0] || String(vals[0][0]).toLowerCase() !== 'tidsstempel') {
          rng.setValues([[
            'Tidsstempel', 'Nivå', 'Funksjon', 'Melding', 'Detaljer (JSON)'
          ]]);
          sh.setFrozenRows(1);
        }
      }
      _sheet = sh;
      return _sheet;
    } catch (e) {
      // Fallback to no-sheet mode (console only)
      _sheet = null;
      return null;
    }
  }

  function _rotateIfNeeded_() {
    var sh = _ensureSheet_();
    if (!sh) return;
    try {
      var rows = sh.getLastRow();
      if (rows <= cfg.ROTATE_MAX_ROWS) return;
      // Keep header + the newest N rows
      var keep = Math.max(1, Math.floor(cfg.ROTATE_MAX_ROWS * cfg.ROTATE_KEEP_RATIO));
      var deleteCount = Math.max(0, rows - keep - 1); // minus 1 for header
      if (deleteCount > 0) {
        sh.deleteRows(2, deleteCount);
        _stats.rotated += deleteCount;
      }
    } catch (_) {}
  }

  function _sendExternal_(entry) {
    try {
      if (!cfg.EXTERNAL_ENABLED || !cfg.EXTERNAL_ENDPOINT) return;
      var payload = {
        timestamp: entry[0],
        level: entry[1],
        function: entry[2],
        message: entry[3],
        details: (function () {
          try { return JSON.parse(entry[4] || '{}'); } catch (_) { return { raw: entry[4] }; }
        })(),
        source: 'google-apps-script',
        scriptId: (function () {
          try { return ScriptApp.getScriptId(); } catch (_) { return null; }
        })()
      };
      var maxRetries = Math.max(0, Number(cfg.EXTERNAL_MAX_RETRIES || 0));
      for (var attempt = 0; attempt <= maxRetries; attempt++) {
        try {
          var resp = UrlFetchApp.fetch(cfg.EXTERNAL_ENDPOINT, {
            method: 'post',
            contentType: 'application/json',
            muteHttpExceptions: true,
            payload: JSON.stringify(payload)
          });
          var rc = resp && resp.getResponseCode ? resp.getResponseCode() : 0;
          if (rc >= 200 && rc < 300) break;
        } catch (e) {
          if (attempt === maxRetries) break;
          try { Utilities.sleep(Math.pow(2, attempt) * 250); } catch (_) {}
        }
      }
    } catch (_) {}
  }

  function _flushBuffer_() {
    if (!_buffer.length) return;
    var local = _buffer.slice(0);
    _buffer.length = 0;
    _stats.buffered = 0;

    // external first (best effort)
    try {
      for (var i = 0; i < local.length; i++) {
        _sendExternal_(local[i]);
      }
    } catch (_) {}

    var sh = _ensureSheet_();
    if (!sh) {
      // fallback to console
      try {
        for (var j = 0; j < local.length; j++) {
          // Log minimal structure to console
          console.log('[LOG]', local[j][0], local[j][1], local[j][2], local[j][3], local[j][4]);
        }
      } catch (_) {}
      return;
    }

    try {
      var startRow = sh.getLastRow() + 1;
      var flushStart = Date.now();
      sh.getRange(startRow, 1, local.length, 5).setValues(local);
      var flushTime = Date.now() - flushStart;
      _stats.avgFlushTime = (_stats.avgFlushTime * _stats.flushCount + flushTime) / (_stats.flushCount + 1);
      _stats.flushCount++;
      _stats.lastFlushTime = new Date();
      _stats.written += local.length;
      _rotateIfNeeded_();
    } catch (e) {
      // Fallback to console on failure
      try {
        for (var k = 0; k < local.length; k++) {
          console.error('[LOGGER WRITE FAIL]', e && e.message, local[k]);
        }
      } catch (_) {}
    }
  }

  function _write_(level, fnName, message, details) {
    if (!_shouldLog_(level)) return;
    if (!_incDailyCountAndCheckQuota_()) return;

    var ts = _nowStr_();
    var det = _sanitizeDetails_(details);
    var json = _stringifySafe_(det);
    var entry = [ts, level, fnName || '', message || '', json];

    _buffer.push(entry);
    _stats.buffered = _buffer.length;

    // Hard cap protection
    if (_buffer.length >= cfg.MAX_BUFFER_SIZE) {
      _flushBuffer_();
      return;
    }

    if (_buffer.length >= cfg.BATCH_SIZE) {
      _flushBuffer_();
    }
  }

  // ------------------------------- API -----------------------------------
  var api = {
    info: function (fn, msg, details) { _write_('INFO', fn, msg, details); },
    warn: function (fn, msg, details) { _write_('WARN', fn, msg, details); },
    error: function (fn, msg, details) { _write_('ERROR', fn, msg, details); },
    debug: function (fn, msg, details) { _write_('DEBUG', fn, msg, details); },
    flush: function () { _flushBuffer_(); },
    stats: function () { return JSON.parse(JSON.stringify(_stats)); },
    setLevel: function (levelName) {
      if (cfg.LEVELS[levelName] == null) return;
      _currentLevel = cfg.LEVELS[levelName];
    },
    setExternalEnabled: function (enabled) {
      cfg.EXTERNAL_ENABLED = !!enabled;
    },
    setExternalEndpoint: function (url) {
      cfg.EXTERNAL_ENDPOINT = String(url || '');
    },
    setBatchSize: function (n) {
      var v = Number(n);
      if (v >= 1 && v <= 100) cfg.BATCH_SIZE = v;
    }
  };

  getAppLogger_._singleton = api;
  return api;
}
