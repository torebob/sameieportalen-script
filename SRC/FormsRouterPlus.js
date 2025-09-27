/**
 * FormsRouterPlus + Ownership Handler (v2.2-lite, single file)
 *
 * Changes vs 2.1-lite:
 *  - Race-condition fix in rate limiter (strict critical section under LockService)
 *  - Consistent locking (waitLock) for rate-limit and write
 *  - Strong input validation for event shape
 *  - Sanitize → then validate (no gap)
 *  - Lightweight i18n (NO/EN messages)
 *  - Per-field sanitize via validation rules; safe default for others
 *  - Backward compatible writes; flags for optional features
 */

// ============================= VERSION & FLAGS ===============================

const VERSION_INFO = { version: '2.2.0-lite' };

const FEATURE_FLAGS = {
  METRICS_ENABLED: true,        // best-effort telemetry to "Metrics"
  NOTIFY_ON_FAILURE: true,      // admin e-mail on errors
  SUCCESS_UI_ALERTS: true,      // UI alerts on success (Sheets)
  RECEIPT_TO_USER: true,        // e-mail receipt to user if email provided
};

// ================================ I18N ======================================

const I18N = {
  lang: 'no', // 'no' | 'en'
  t(k, vars) {
    const msg = (I18N.dict[I18N.lang] && I18N.dict[I18N.lang][k]) || (I18N.dict.en[k]) || k;
    if (!vars) return msg;
    return msg.replace(/\{(\w+)\}/g, (_, key) => (vars[key] != null ? String(vars[key]) : ''));
  },
  dict: {
    no: {
      ROUTER_MISSING_HEADER: 'Skjema-register mangler forventede kolonner',
      ROUTER_EMPTY: 'Register er tomt (kun header).',
      ROUTER_CREATED: 'Register manglet – opprettet tomt.',
      ROUTER_NO_MATCH: 'Ingen match i Skjema-registeret.',
      ROUTER_ERROR: 'Feil i ruter',
      SUBMIT_DUPLICATE: 'Innsending allerede registrert (duplikat).',
      SUBMIT_OK: 'Takk! Skjemaet er registrert.',
      SUBMIT_RATE_LIMIT: 'Det er mange innsendinger på kort tid. Vent litt og prøv igjen.',
      SUBMIT_VALIDATE_FAIL: 'Vennligst kontroller feltene i skjemaet og prøv igjen. {errors}',
      SUBMIT_GENERIC_FAIL: 'Det oppstod en feil under innsending. Administrator er varslet.',
      ADMIN_SUBJECT: '🚨 Feil i skjemainnsending: {form} [{env}]',
      ADMIN_BODY_INTRO: 'Det oppstod en feil ved behandling av skjemainnsending.',
      RECEIPT_SUBJECT: 'Kvittering: {form}',
      RECEIPT_BODY: 'Hei {name},\n\nVi har registrert din innsending for leilighet {apt}.\nTakk!\n\n— Automatisk kvittering.'
    },
    en: {
      ROUTER_MISSING_HEADER: 'Route register is missing required columns',
      ROUTER_EMPTY: 'Register is empty (header only).',
      ROUTER_CREATED: 'Register missing – created empty.',
      ROUTER_NO_MATCH: 'No matching route.',
      ROUTER_ERROR: 'Router error',
      SUBMIT_DUPLICATE: 'Submission already recorded (duplicate).',
      SUBMIT_OK: 'Thank you! Your form was recorded.',
      SUBMIT_RATE_LIMIT: 'Too many submissions. Please wait and try again.',
      SUBMIT_VALIDATE_FAIL: 'Please check your form fields and try again. {errors}',
      SUBMIT_GENERIC_FAIL: 'An error occurred. The administrator has been notified.',
      ADMIN_SUBJECT: '🚨 Form submission error: {form} [{env}]',
      ADMIN_BODY_INTRO: 'An error occurred while processing a form submission.',
      RECEIPT_SUBJECT: 'Receipt: {form}',
      RECEIPT_BODY: 'Hi {name},\n\nWe recorded your submission for apartment {apt}.\nThanks!\n\n— Automatic receipt.'
    }
  }
};

// ============================= ENV & SECRETS ================================

const ENV_CFG = {
  defaultEnv: 'prod',
  propKeys: {
    ENV: 'ENV',
    ADMIN_EMAIL: 'ADMIN_EMAIL',
    SPREADSHEET_ID: 'SPREADSHEET_ID',
  }
};

function getEnv_() {
  try {
    const p = PropertiesService.getScriptProperties();
    return (p.getProperty(ENV_CFG.propKeys.ENV) || ENV_CFG.defaultEnv).toLowerCase();
  } catch (_) { return ENV_CFG.defaultEnv; }
}

function getSecret_(key, fallback) {
  try {
    const p = PropertiesService.getScriptProperties();
    const v = p.getProperty(key);
    return (v == null || v === '') ? fallback : v;
  } catch (_) { return fallback; }
}

function _getSs_() {
  try {
    const id = getSecret_(ENV_CFG.propKeys.SPREADSHEET_ID, null);
    if (id) return SpreadsheetApp.openById(id);
  } catch (err) {
    _getLogger_().warn('_getSs_', 'openById failed; falling back to active', { error: err.message });
  }
  return SpreadsheetApp.getActive();
}

function _safeUiAlert_(title, message) {
  try { SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK); } catch (_) {}
}

// ================================ LOGGER ====================================

function _getLogger_() {
  if (typeof getAppLogger_ === 'function') {
    try { return getAppLogger_(); } catch (_) {}
  }
  return {
    info: (fn, msg, data)  => { try { console.log('[INFO]', fn, msg, data || ''); } catch (_) {} },
    warn: (fn, msg, data)  => { try { console.warn('[WARN]', fn, msg, data || ''); } catch (_) {} },
    error:(fn, msg, data)  => { try { console.error('[ERROR]', fn, msg, data || ''); } catch (_) {} },
    setLevel: function(){}, flush: function(){}, stats: function(){ return {}; }
  };
}

// ============================== ROUTER CONFIG ===============================

const ROUTER_CFG = {
  registerSheetName: 'Skjema-register',
  header: ['Aktiv', 'MatchFelt', 'MatchVerdi', 'Handler', 'Beskrivelse'],
  truthy: v => /^(true|ja|1|x|on|enabled)$/i.test(String(v || '').trim()),
  maxRoutes: 100
};

/**
 * Route incoming form submit to a handler listed in "Skjema-register".
 */
function routeFormSubmit(e) {
  const log = _getLogger_();
  const fn = 'routeFormSubmit';
  const nv = (e && e.namedValues) || {};

  try {
    const ss = _getSs_();
    let sh = ss.getSheetByName(ROUTER_CFG.registerSheetName);
    if (!sh) {
      sh = ss.insertSheet(ROUTER_CFG.registerSheetName);
      sh.appendRow(ROUTER_CFG.header);
      sh.setFrozenRows(1);
      log.warn(fn, I18N.t('ROUTER_CREATED'), {});
      return;
    }

    const vals = sh.getDataRange().getValues();
    if (!vals || vals.length < 2) {
      log.warn(fn, I18N.t('ROUTER_EMPTY'), {});
      return;
    }

    const header = vals[0].map(x => String(x || '').trim().toLowerCase());
    const idx = {
      aktiv: header.indexOf('aktiv'),
      felt: header.indexOf('matchfelt'),
      verdi: header.indexOf('matchverdi'),
      handler: header.indexOf('handler'),
    };
    if (idx.aktiv < 0 || idx.felt < 0 || idx.verdi < 0 || idx.handler < 0) {
      throw new Error(I18N.t('ROUTER_MISSING_HEADER') + ': ' + JSON.stringify(ROUTER_CFG.header));
    }

    let dispatched = false;

    for (let r = 1; r < vals.length && r <= ROUTER_CFG.maxRoutes; r++) {
      const row = vals[r];
      if (!ROUTER_CFG.truthy(row[idx.aktiv])) continue;

      const matchField = String(row[idx.felt] || '').trim();
      const matchValue = String(row[idx.verdi] || '').trim();
      const handlerName = String(row[idx.handler] || '').trim();

      const matches = (matchField === '*') ||
        (Object.prototype.hasOwnProperty.call(nv, matchField) && _first_(nv[matchField]) === matchValue);

      if (!matches) continue;

      const root = (typeof globalThis !== 'undefined') ? globalThis : this;
      if (typeof root[handlerName] === 'function') {
        log.info(fn, 'Routing to handler', { handler: handlerName, matchField, matchValue });
        root[handlerName](e);
        _metric_('router_success', fn, { handler: handlerName });
        dispatched = true;
        break;
      } else {
        log.warn(fn, 'Handler not found', { handler: handlerName });
      }
    }

    if (!dispatched) {
      log.warn(fn, I18N.t('ROUTER_NO_MATCH'), { sampleKeys: Object.keys(nv).slice(0, 10) });
      _metric_('router_no_match', fn, { sampleKeys: Object.keys(nv).slice(0, 10) });
    }
  } catch (err) {
    log.error(fn, I18N.t('ROUTER_ERROR'), { error: err.message, stack: err.stack });
    if (FEATURE_FLAGS.NOTIFY_ON_FAILURE) _notifyAdminSafe_(I18N.t('ROUTER_ERROR'), err, { fn });
    _safeUiAlert_(I18N.t('ROUTER_ERROR'), String(err && err.message || 'Unknown'));
    _metric_('router_error', fn, { error: err.message });
  } finally {
    log.info(fn, 'Done', { version: VERSION_INFO.version });
  }
}

// ============================= TELEMETRY (lite) =============================

const METRICS_CFG = {
  enabled: FEATURE_FLAGS.METRICS_ENABLED,
  sheetName: 'Metrics',
  header: ['Tid', 'Event', 'Handler', 'Miljø', 'Versjon', 'Detaljer (JSON)'],
};

function _metric_(eventName, handlerName, details) {
  if (!METRICS_CFG.enabled) return;
  try {
    const ss = _getSs_();
    let sh = ss.getSheetByName(METRICS_CFG.sheetName);
    if (!sh) {
      sh = ss.insertSheet(METRICS_CFG.sheetName);
      sh.appendRow(METRICS_CFG.header);
      sh.setFrozenRows(1);
    }
    const row = [new Date(), eventName, handlerName || '', getEnv_(), VERSION_INFO.version, _stringifySafe_(details || {})];
    sh.appendRow(row);
  } catch (_) { /* best-effort */ }
}

// =========================== OWNERSHIP HANDLER ==============================

const OWNERSHIP_CFG = {
  sheetName: 'Skjema-mottak',
  formIdentifier: 'Eierskap',
  formFields: {
    name: 'Navn',
    apartment: 'Leilighetsnummer',
    email: 'E-post'
  },
  requiredFields: ['name', 'apartment'],
  outputColumns: ['Tidspunkt', 'Skjema', 'Navn', 'Leilighet', 'E-post'],

  validation: {
    name:      { minLength: 2,  maxLength: 100, pattern: /^[a-zA-ZæøåÆØÅ\s\-\.\']+$/, sanitize: v => v.replace(/[^a-zA-ZæøåÆØÅ\s\-\.\']/g, '').trim() },
    apartment: {                 maxLength: 10,  pattern: /^[A-Z]\d{3,4}$|^\d{3,4}$/,    sanitize: v => v.toUpperCase().replace(/[^A-Z0-9]/g, '') },
    email:     {                 maxLength: 254, pattern: /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/, sanitize: v => v.toLowerCase().trim() }
  },

  lockTimeoutMs: 15000,
  maxRetries: 3,
  retryDelayMs: 1000,

  rateLimitWindowMs: 60000,
  maxSubmissionsPerWindow: 10,
  cacheKeyPrefix: 'ownership_submit',
  enableDuplicateWindowMs: 120000,
  duplicateKeyFields: ['name', 'apartment', 'email'],

  warnOnHeaderMismatch: true
};

/**
 * Ownership form handler (production-safe, lean)
 */
function onOwnershipSubmit(e) {
  const log = _getLogger_();
  const fn  = 'onOwnershipSubmit';

  try {
    _validateOwnershipConfig_();

    // Strong input validation for event shape
    if (!e || typeof e !== 'object' || !e.namedValues || typeof e.namedValues !== 'object') {
      throw new Error('Invalid form submission event (missing namedValues)');
    }

    _ownershipRateLimit_();

    const formData = _ownershipExtract_(e);
    const { isValid, errors, data } = _ownershipSanitizeThenValidate_(formData);
    if (!isValid) {
      throw new Error('Validering feilet: ' + errors.join(', '));
    }

    const dup = _ownershipCheckAndMarkDuplicate_(data);
    if (dup.isDuplicate) {
      log.warn(fn, 'Duplicate within idempotency window – ignoring', { key: dup.key });
      _metric_('ownership_duplicate', fn, { key: dup.key });
      if (FEATURE_FLAGS.SUCCESS_UI_ALERTS) _safeUiAlert_('OK', I18N.t('SUBMIT_DUPLICATE'));
      return;
    }

    const res = _ownershipWriteWithRetry_(data);
    log.info(fn, 'Saved', { retries: res.retries });
    _metric_('ownership_saved', fn, { retries: res.retries });

    if (FEATURE_FLAGS.SUCCESS_UI_ALERTS) _safeUiAlert_('OK', I18N.t('SUBMIT_OK'));
    if (FEATURE_FLAGS.RECEIPT_TO_USER) _maybeSendReceipt_(data);

  } catch (err) {
    _ownershipHandleError_(err, e, { fn, env: getEnv_() });
  }
}

// -------- OWNERSHIP helpers --------

function _validateOwnershipConfig_() {
  const req = ['sheetName', 'formIdentifier', 'formFields', 'requiredFields', 'outputColumns'];
  const missing = req.filter(k => !OWNERSHIP_CFG[k]);
  if (missing.length) throw new Error('OWNERSHIP_CFG missing: ' + missing.join(', '));

  if (FEATURE_FLAGS.NOTIFY_ON_FAILURE) {
    const adminEmail = getSecret_(ENV_CFG.propKeys.ADMIN_EMAIL, null);
    if (!adminEmail) throw new Error('ADMIN_EMAIL missing in Script Properties (notifications enabled)');
  }
}

function _ownershipExtract_(e) {
  const nv = e.namedValues || {};
  return {
    name: _first_(nv[OWNERSHIP_CFG.formFields.name]),
    apartment: _first_(nv[OWNERSHIP_CFG.formFields.apartment]),
    email: _first_(nv[OWNERSHIP_CFG.formFields.email]),
    timestamp: new Date(),
    rawData: nv
  };
}

/** Sanitize THEN validate so patterns check stored values */
function _ownershipSanitizeThenValidate_(data) {
  const errors = [];
  const sanitized = { ...data };

  // Apply sanitize rule if present, else a safe default
  Object.keys(OWNERSHIP_CFG.validation).forEach(field => {
    const rules = OWNERSHIP_CFG.validation[field] || {};
    const raw = String(data[field] || '');
    const clean = (typeof rules.sanitize === 'function')
      ? rules.sanitize(raw)
      : raw.replace(/[<>]/g, '').trim(); // very light default
    sanitized[field] = clean;
  });

  // Required fields after sanitization
  OWNERSHIP_CFG.requiredFields.forEach(field => {
    const v = sanitized[field];
    if (!v || String(v).trim().length === 0) {
      errors.push(`${OWNERSHIP_CFG.formFields[field] || field} mangler`);
    }
  });

  // Length & pattern validation after sanitization
  Object.keys(OWNERSHIP_CFG.validation).forEach(field => {
    const rules = OWNERSHIP_CFG.validation[field] || {};
    const v = String(sanitized[field] || '');

    if (!v && !OWNERSHIP_CFG.requiredFields.includes(field)) return;

    if (rules.minLength && v.length < rules.minLength) errors.push(`${OWNERSHIP_CFG.formFields[field]} for kort (min ${rules.minLength})`);
    if (rules.maxLength && v.length > rules.maxLength) errors.push(`${OWNERSHIP_CFG.formFields[field]} for lang (maks ${rules.maxLength})`);
    if (rules.pattern && v && !rules.pattern.test(v)) errors.push(`${OWNERSHIP_CFG.formFields[field]} har ugyldig format`);
  });

  return { isValid: errors.length === 0, errors, data: sanitized };
}

/** Strict critical section under LockService to avoid race in counter */
function _ownershipRateLimit_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(OWNERSHIP_CFG.lockTimeoutMs);
  try {
    const cache = CacheService.getScriptCache();
    const now = Date.now();
    const key = `${OWNERSHIP_CFG.cacheKeyPrefix}:window:${Math.floor(now / OWNERSHIP_CFG.rateLimitWindowMs)}`;

    // Single read/modify/write under the same lock
    let count = 0;
    try { count = parseInt(cache.get(key) || '0', 10) || 0; } catch (_) { count = 0; }

    if (count >= OWNERSHIP_CFG.maxSubmissionsPerWindow) {
      throw new Error('Rate limit exceeded. For mange innsendinger på kort tid.');
    }

    const msIntoWindow = now % OWNERSHIP_CFG.rateLimitWindowMs;
    const ttlSec = Math.max(1, Math.ceil((OWNERSHIP_CFG.rateLimitWindowMs - msIntoWindow) / 1000));
    try { cache.put(key, String(count + 1), ttlSec); } catch (_) {}
  } finally {
    lock.releaseLock();
  }
}

function _ownershipCheckAndMarkDuplicate_(data) {
  const windowMs = Number(OWNERSHIP_CFG.enableDuplicateWindowMs || 0);
  if (!windowMs) return { isDuplicate: false, key: '' };

  const norm = {};
  OWNERSHIP_CFG.duplicateKeyFields.forEach(k => { norm[k] = (data[k] || '').toString().trim().toLowerCase(); });
  const json = JSON.stringify(norm);
  const bytes = Utilities.newBlob(json).getBytes();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, bytes);
  const hex = digest.map(b => ((b & 0xFF).toString(16).padStart(2,'0'))).join('');

  const key = `${OWNERSHIP_CFG.cacheKeyPrefix}:dup:${hex}`;
  const cache = CacheService.getScriptCache();
  const existing = (function(){ try { return cache.get(key); } catch(_){ return null; }})();
  const isDup = !!existing;
  try { cache.put(key, '1', Math.max(1, Math.ceil(windowMs / 1000))); } catch (_){}
  return { isDuplicate: isDup, key };
}

function _ownershipWriteWithRetry_(data) {
  let lastErr = null;
  for (let attempt = 0; attempt <= OWNERSHIP_CFG.maxRetries; attempt++) {
    try {
      _ownershipWriteOnce_(data);
      return { data, retries: attempt };
    } catch (err) {
      lastErr = err;
      if (attempt < OWNERSHIP_CFG.maxRetries) Utilities.sleep(OWNERSHIP_CFG.retryDelayMs * Math.pow(2, attempt));
    }
  }
  throw new Error(`Kunne ikke skrive etter ${OWNERSHIP_CFG.maxRetries + 1} forsøk: ${lastErr && lastErr.message}`);
}

function _ownershipWriteOnce_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(OWNERSHIP_CFG.lockTimeoutMs); // consistent with rate limiter
  try {
    const ss = _getSs_();
    let sh = ss.getSheetByName(OWNERSHIP_CFG.sheetName);
    if (!sh) {
      sh = ss.insertSheet(OWNERSHIP_CFG.sheetName);
      _setupSheet_(sh, OWNERSHIP_CFG.outputColumns);
    }
    _ensureHeader_(sh, OWNERSHIP_CFG.outputColumns, OWNERSHIP_CFG.warnOnHeaderMismatch);

    const row = [
      new Date(),
      OWNERSHIP_CFG.formIdentifier,
      data.name,
      data.apartment,
      data.email || ''
    ];

    const cols = Math.min(row.length, sh.getLastColumn() || OWNERSHIP_CFG.outputColumns.length);
    const next = sh.getLastRow() + 1;
    sh.getRange(next, 1, 1, cols).setValues([row.slice(0, cols)]);
    if (next === 2) sh.autoResizeColumns(1, cols);
  } finally {
    lock.releaseLock();
  }
}

function _ownershipHandleError_(error, originalEvent, context) {
  const log = _getLogger_();
  const payload = {
    message: error && error.message,
    stack: error && error.stack,
    context: context || {},
    env: getEnv_(),
    ts: new Date().toISOString(),
    formKeys: Object.keys((originalEvent && originalEvent.namedValues) || {})
  };
  log.error(context && context.fn || 'onOwnershipSubmit', 'Feil ved innsending', payload);
  _metric_('ownership_error', context && context.fn || 'onOwnershipSubmit', { message: payload.message });

  if (FEATURE_FLAGS.NOTIFY_ON_FAILURE) {
    const subject = I18N.t('ADMIN_SUBJECT', { form: OWNERSHIP_CFG.formIdentifier, env: payload.env });
    const body =
`${I18N.t('ADMIN_BODY_INTRO')}

Feil: ${payload.message}
Miljø: ${payload.env}
Tid: ${payload.ts}

Kontekst:
${_stringifySafe_(payload.context, 2)}

Skjema-felter:
${payload.formKeys.join(', ')}

Stack:
${payload.stack}

— Automatisk melding.`;
    _notifyAdminSafe_(subject, error, body);
  }

  // User-facing UI message
  let msg = I18N.t('SUBMIT_GENERIC_FAIL');
  const m = String(payload.message || '');
  if (m.indexOf('Rate limit exceeded') !== -1) {
    msg = I18N.t('SUBMIT_RATE_LIMIT');
  } else if (m.indexOf('Validering feilet') !== -1) {
    msg = I18N.t('SUBMIT_VALIDATE_FAIL', { errors: m.replace('Validering feilet: ', '') });
  }
  _safeUiAlert_('Feil', msg);
}

function _maybeSendReceipt_(data) {
  const to = (data.email || '').trim();
  if (!to) return;
  try {
    const subject = I18N.t('RECEIPT_SUBJECT', { form: OWNERSHIP_CFG.formIdentifier });
    const body = I18N.t('RECEIPT_BODY', { name: data.name, apt: data.apartment });
    MailApp.sendEmail({ to, subject, body });
  } catch (_) { /* non-critical */ }
}

// ============================= SHARED HELPERS ===============================

function _setupSheet_(sheet, header) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#E8F0FE');
}

function _ensureHeader_(sheet, expectedHeader, warn) {
  const last = sheet.getLastRow();
  if (last === 0) {
    _setupSheet_(sheet, expectedHeader);
    return;
  }
  if (warn) {
    try {
      const actual = sheet.getRange(1, 1, 1, expectedHeader.length).getValues()[0] || [];
      const mismatch = actual.length !== expectedHeader.length ||
        actual.some((v, i) => String(v || '').trim() !== String(expectedHeader[i] || '').trim());
      if (mismatch) {
        _getLogger_().warn('_ensureHeader_', 'Header differs (no auto-change)', { expected: expectedHeader, actual });
      }
    } catch (e) {
      _getLogger_().warn('_ensureHeader_', 'Header check failed', { error: e && e.message });
    }
  }
}

function _first_(v) {
  if (v == null) return '';
  if (Array.isArray(v)) return (v[0] || '').toString().trim();
  return (v || '').toString().trim();
}

/*
 * MERK: _stringifySafe_() er fjernet fra denne filen. Den globale
 * versjonen fra 000_Utils.js brukes i stedet.
 */

/** Best-effort admin mail (context may be string body or object) */
function _notifyAdminSafe_(subject, error, context) {
  try {
    const adminEmail = getSecret_(ENV_CFG.propKeys.ADMIN_EMAIL, null);
    if (!adminEmail) return;
    const body = typeof context === 'string'
      ? context
      : `${subject}\n\nError: ${error && error.message}\nEnv: ${getEnv_()}\nTime: ${new Date().toISOString()}\n\nContext:\n${_stringifySafe_(context, 2)}\n\nStack:\n${error && error.stack}\n\n— Automated`;
    MailApp.sendEmail({ to: adminEmail, subject, body, htmlBody: body.replace(/\n/g, '<br>') });
  } catch (_) { /* best-effort */ }
}

// =============================== DEV / SEED =================================

function dev_seed_register() {
  const ss = _getSs_();
  let sh = ss.getSheetByName(ROUTER_CFG.registerSheetName);
  if (!sh) {
    sh = ss.insertSheet(ROUTER_CFG.registerSheetName);
    sh.appendRow(ROUTER_CFG.header);
    sh.setFrozenRows(1);
  }
  sh.appendRow(['JA', OWNERSHIP_CFG.formFields.name, 'Ola Nordmann', 'onOwnershipSubmit', 'Eksempel: Navn == Ola Nordmann']);
  SpreadsheetApp.flush();
}

function dev_test_router() {
  const e = {
    namedValues: {
      [OWNERSHIP_CFG.formFields.name]: ['Ola Nordmann'],
      [OWNERSHIP_CFG.formFields.apartment]: ['H0302'],
      [OWNERSHIP_CFG.formFields.email]: ['ola@example.com']
    }
  };
  routeFormSubmit(e);
}

function dev_test_onOwnershipSubmit() {
  const cases = [
    {namedValues: { [OWNERSHIP_CFG.formFields.name]: ['Ola Nordmann'], [OWNERSHIP_CFG.formFields.apartment]: ['H0302'], [OWNERSHIP_CFG.formFields.email]: ['ola@example.com'] }},
    {namedValues: { [OWNERSHIP_CFG.formFields.name]: ['Kari Hansen'],   [OWNERSHIP_CFG.formFields.email]:     ['kari@example.com'] }},
    {namedValues: { [OWNERSHIP_CFG.formFields.name]: ['Per Olsen'],     [OWNERSHIP_CFG.formFields.apartment]: ['A123'], [OWNERSHIP_CFG.formFields.email]: ['feil'] }}
  ];
  cases.forEach((tc, i) => {
    try { onOwnershipSubmit(tc); console.log(`Case ${i+1}: OK`); }
    catch (e) { console.log(`Case ${i+1}: FAIL -> ${e.message}`); }
  });
}

// --- NEW: tiny wrappers to invoke the migration helper (in the other file) ---

/** Dry-run: see what the migration would do (no changes). */
function dev_migration_dryRun() {
  return ownershipMigrationPlan_();
}

/** Apply the migration safely with defaults (backup + ensure columns + reorder). */
function dev_migration_apply() {
  return applyOwnershipMigration_({
    dryRun: false,
    createBackupCopy: true,
    backupSuffix: '_backup',
    ensureColumns: true,
    reorderToDefault: true,
    newSheetSuffix: '_v2'
  });
}

/** Fallback mini-logger if LoggerPlus is not present */
if (typeof getAppLogger_ !== 'function') {
  function getAppLogger_() {
    return {
      info: (fn, msg, data) => { try { console.log('[INFO]', fn || '', msg || '', data || ''); } catch (_) {} },
      warn: (fn, msg, data) => { try { console.warn('[WARN]', fn || '', msg || '', data || ''); } catch (_) {} },
      error:(fn, msg, data) => { try { console.error('[ERROR]', fn || '', msg || '', data || ''); } catch (_) {} },
      setLevel: function(){}, flush: function(){}, stats: function(){ return {}; }
    };
  }
}
