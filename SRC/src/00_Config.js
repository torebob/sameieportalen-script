/**
 * Sameieportalen – Global Konfigurasjon og Konstanter
 * FIL: 00_Config.gs
 * VERSJON: 2.1.0
 * SIST OPPDATERT: 2025-09-23
 *
 * FORMÅL
 * - Én sentral og selvvaliderende "sannhetskilde" for hele løsningen.
 * - Standardiserte navn på ark, kolonner, UI-ressurser og systemgrenser.
 * - Miljøstyring (development/production), feature-flags og tilgangsregler.
 * - Trygge hjelpefunksjoner for å lese/validere konfig i andre moduler.
 *
 * NØKKELFUNKSJONER
 * - validateConfiguration(): Kjører helsesjekk av konfig og returnerer funn.
 * - getUIConfig(key): Henter UI-oppsett (med miljøtilpasning).
 * - getSheetMetadata(name): Returnerer metadata/validering for et ark.
 * - isFeatureEnabled(flag): Sjekker om en funksjon er aktivert.
 * - setupDevelopmentTools(): Aktiverer utviklerverktøy (kun i development).
 * - initializeConfiguration() [IIFE]: Kjører oppstartsvalidering trygt.
 */

/* ============================= Miljødeteksjon ============================== */

var ENVIRONMENT = (function () {
  try {
    var sp = PropertiesService.getScriptProperties();
    var env = sp.getProperty('ENVIRONMENT');
    if (env === 'development') return 'development';
    if (env === 'production') return 'production';
  } catch (_) {}
  return 'production'; // safe default
})();

/* ============================ App-metadata (APP) =========================== */

var APP = Object.freeze({
  NAME: 'Sameieportalen',
  VERSION: '2.1.0',
  BUILD_DATE: '2025-09-23',
  ENVIRONMENT: ENVIRONMENT,
  DESCRIPTION: 'Portal for administrasjon av boligsameie',
  AUTHOR: 'System Administrator',

  FEATURES: Object.freeze({
    MEETINGS_MODULE: true,
    VAKTMESTER_MODULE: true,
    HMS_MODULE: true,
    ADVANCED_REPORTING: true,
    AUDIT_LOGGING: true,
    PERFORMANCE_MONITORING: (ENVIRONMENT === 'development'),
    DEVELOPMENT_TOOLS: (ENVIRONMENT === 'development'),
    AUTO_BACKUP: (ENVIRONMENT === 'production')
  }),

  REQUIRED_SCOPES: [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/script.external_request',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/gmail.send'
  ]
});

/* ========================= Ark-navn (kanoniske) =========================== */

var CONFIG_SHEETS = Object.freeze({
  SEKSJONER: 'Seksjoner',
  PERSONER: 'Personer',
  EIERSKAP: 'Eierskap',
  LEIE: 'Leieforhold',
  OPPGAVER: 'Oppgaver',
  STYRET: 'Styret',
  RAPPORT: 'Rapport',
  KONFIG: 'Konfig',
  HMS: 'HMS_Egenkontroll',
  MENY_FELLES: 'Meny_Felles',
  MENY_MIN: 'Meny_Min',
  KRAV: 'Requirements',
  SYNC_LOG: 'Sync_Log',
  VERSION_LOG: 'Version_Log',
  ANALYSIS_LOG: 'Analysis_Log',
  AUDIT_LOG: 'System_AuditLog',
  ERROR_LOG: 'System_ErrorLog',
  PERFORMANCE: 'System_Performance'
});

// Bakoverkompabilitet (valgfritt alias)
var SHEETS = CONFIG_SHEETS;

/* ============================ Ark-metadata (rich) ========================== */

var SHEET_METADATA = Object.freeze({
  'Seksjoner': {
    displayName: 'Seksjoner',
    description: 'Register over alle seksjoner i sameiet',
    protected: false,
    requiredColumns: ['SeksjonID', 'Adresse', 'Størrelse', 'Type'],
    dataValidation: {
      'SeksjonID': 'UNIQUE_INTEGER',
      'Størrelse': 'POSITIVE_NUMBER',
      'Type': ['Leilighet', 'Rekkehus', 'Enebolig']
    },
    accessRole: 'user'
  },
  'Personer': {
    displayName: 'Personer',
    description: 'Kontaktinformasjon for eiere og beboere',
    protected: true,
    requiredColumns: ['PersonID', 'Navn', 'Epost', 'Telefon'],
    dataValidation: {
      'PersonID': 'UNIQUE_INTEGER',
      'Epost': 'EMAIL',
      'Telefon': 'PHONE_NO'
    },
    accessRole: 'admin'
  },
  'Requirements': {
    displayName: 'Krav',
    description: 'Kildedata for kravdokument, synk mot Google Docs',
    protected: false,
    requiredColumns: ['KravID', 'Krav', 'Prioritet', 'Fremdrift %', 'Kapittel'],
    dataValidation: {
      'KravID': 'KRAV_ID',      // f.eks. K1-01
      'Prioritet': ['MÅ', 'BØR', 'KAN', 'MUST', 'SHOULD', 'COULD'],
      'Fremdrift %': 'PERCENT_0_100',
      'Kapittel': 'CHAPTER_1_39'
    },
    accessRole: 'admin'
  }
});

/* ============================== UI-register =============================== */

var CONFIG_UI = Object.freeze({
  DASHBOARD: { file: '37_Dashboard.html', title: 'Sameieportalen – Dashbord', width: 1280, height: 840, resizable: true, modal: true, requiredRole: 'user' },
  MOTE_OVERSIKT: { file: '30_Moteoversikt.html', title: 'Møteoversikt', width: 1100, height: 760, resizable: true, modal: true, requiredRole: 'user' },
  MOTE_SAK_EDITOR: { file: '31_MoteSakEditor.html', title: 'Møtesaker – Editor', width: 1100, height: 760, resizable: true, modal: true, requiredRole: 'admin' },
  SYNC_WIZARD: { file: '20_SyncWizard.html', title: 'Krav Sync – Veiviser', width: 980, height: 720, resizable: true, modal: true, requiredRole: 'admin' },
  ANALYSIS_DASH: { file: '38_AnalysisDashboard.html', title: 'Analyse – Oversikt', width: 1280, height: 860, resizable: true, modal: true, requiredRole: 'admin' }
});

// Bakoverkompabilitets-alias (trygt i V8)
try { globalThis.UI_FILES = CONFIG_UI; } catch (_) {}

/* ============================ Systemgrenser =============================== */

var CONFIG_LIMITS = Object.freeze({
  SHEET_MAX_ROWS_SOFT: 50000,
  SHEET_MAX_COLS_SOFT: 100,
  BATCH_SIZE_DEFAULT: 500,
  BATCH_SIZE_SMALL: 100,
  BATCH_SIZE_LARGE: 1000,
  DOC_SEED_BATCH_SIZE: 12,
  DOC_SEED_YIELD_EVERY: 2,
  DOC_SEED_YIELD_MS: 50,
  LOCK_TIMEOUT_MS: 30000,
  TOAST_SECS: 3
});

/* ========================= Valideringsregler (regex) ====================== */

var CONFIG_VALIDATION = Object.freeze({
  EMAIL: /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/,
  PHONE_NO: /^[+]?[\d\s\-()]{6,}$/,
  UNIQUE_INTEGER: /^[1-9]\d*$/,
  POSITIVE_NUMBER: /^(?:\d+|\d+\.\d+|\d+,\d+)$/,
  PERCENT_0_100: /^(?:100|[1-9]?\d)(?:[.,]\d+)?$/,
  KRAV_ID: /^K(?:[1-9]|[12]\d|3[0-9]|X)-\d{2,3}$/,
  CHAPTER_1_39: /^(?:[1-9]|[12]\d|3[0-9])$/
});

/* ======================= Script Properties nøkkelnavn ===================== */

var CONFIG_PROP_KEYS = Object.freeze({
  ENVIRONMENT: 'ENVIRONMENT',
  LOG_LEVEL: 'LOG_LEVEL',
  RSP_DOC_ID: 'RSP_DOC_ID',
  DEV_TOOLS_ENABLED: 'DEV_TOOLS_ENABLED',
  LAST_CONFIG_VALIDATION: 'LAST_CONFIG_VALIDATION'
});

var PROP_KEYS = CONFIG_PROP_KEYS; // alias

/* ============================== Hjelpefunksjoner ========================== */

/**
 * Hent UI-konfig trygt. Legger på miljøindikator i tittelen i development.
 * @param {string} uiKey
 * @returns {object|null}
 */
function getUIConfig(uiKey) {
  try {
    if (!uiKey || !CONFIG_UI[uiKey]) return null;
    var cfg = JSON.parse(JSON.stringify(CONFIG_UI[uiKey]));
    if (APP.ENVIRONMENT === 'development') {
      cfg.title = '[DEV] ' + cfg.title;
    }
    return cfg;
  } catch (e) {
    _configLog_('WARN', 'getUIConfig', { error: e && e.message, key: uiKey });
    return null;
  }
}

/**
 * Hent metadata for ark (med safe defaults).
 * @param {string} sheetName
 * @returns {object}
 */
function getSheetMetadata(sheetName) {
  var name = String(sheetName || '');
  if (SHEET_METADATA[name]) return SHEET_METADATA[name];
  return {
    displayName: name,
    description: '',
    protected: false,
    requiredColumns: [],
    dataValidation: {},
    accessRole: 'user'
  };
}

/**
 * Er en funksjon aktivert via feature-flags?
 * @param {string} featureName
 * @returns {boolean}
 */
function isFeatureEnabled(featureName) {
  try {
    return APP.FEATURES[featureName] === true;
  } catch (_) {
    return false;
  }
}

/**
 * Liten intern logger til Sync_Log (best effort).
 */
function _configLog_(level, event, data) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(CONFIG_SHEETS.SYNC_LOG) || ss.insertSheet(CONFIG_SHEETS.SYNC_LOG);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['Tid', 'Nivå', 'Hendelse', 'Detaljer']);
    }
    sh.appendRow([new Date(), String(level || 'INFO'), String(event || ''), JSON.stringify(data || {})]);
  } catch (_) {
    // swallow
  }
}

/**
 * Safe get/set av Script Properties.
 */
function _getProp_(key, defVal) {
  try {
    var sp = PropertiesService.getScriptProperties();
    var v = sp.getProperty(key);
    return (v === null || v === undefined) ? defVal : v;
  } catch (_) {
    return defVal;
  }
}
function _setProp_(key, val) {
  try {
    PropertiesService.getScriptProperties().setProperty(key, String(val));
    return true;
  } catch (_) {
    return false;
  }
}

/* =========================== Konfig-validering ============================ */

/**
 * Validerer konfigurasjon og struktur. Skriver til Sync_Log.
 * @returns {{isValid:boolean, errors:string[], warnings:string[], info:Object}}
 */
function validateConfiguration() {
  var res = { isValid: true, errors: [], warnings: [], info: {} };

  // Miljø
  if (APP.ENVIRONMENT !== 'development' && APP.ENVIRONMENT !== 'production') {
    res.isValid = false;
    res.errors.push('Ugyldig ENVIRONMENT: ' + APP.ENVIRONMENT);
  }

  // Påkrevde ark
  var ss = SpreadsheetApp.getActive();
  var requiredSheets = [CONFIG_SHEETS.SYNC_LOG, CONFIG_SHEETS.VERSION_LOG, CONFIG_SHEETS.KRAV];
  requiredSheets.forEach(function (name) {
    if (!ss.getSheetByName(name)) {
      res.warnings.push('Mangler anbefalt ark: ' + name);
    }
  });

  // Valider ark-metadata
  Object.keys(SHEET_METADATA).forEach(function (name) {
    var meta = SHEET_METADATA[name];
    if (meta.requiredColumns && meta.requiredColumns.length) {
      var sh = ss.getSheetByName(name);
      if (sh && sh.getLastRow() > 0) {
        var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
        meta.requiredColumns.forEach(function (col) {
          if (hdr.indexOf(col) < 0) {
            res.warnings.push('Ark "' + name + '" mangler kolonne "' + col + '".');
          }
        });
      }
    }
  });

  // DOC ID format (valgfritt)
  var docId = _getProp_(PROP_KEYS.RSP_DOC_ID, null);
  if (docId && !/^[A-Za-z0-9\-_]{20,}$/.test(docId)) {
    res.warnings.push('RSP_DOC_ID ser ugyldig ut i Script Properties.');
  }

  res.info.environment = APP.ENVIRONMENT;
  res.info.version = APP.VERSION;
  res.info.requiredScopes = APP.REQUIRED_SCOPES.length;

  _configLog_('INFO', 'validateConfiguration', res);
  _setProp_(PROP_KEYS.LAST_CONFIG_VALIDATION, new Date().toISOString());

  return res;
}

/* =============================== Init-sekvens ============================= */

/**
 * Init-rutine som kjører ved lasting av prosjektet (IIFE).
 * Utfører en lett validering og logger resultatet. Feiler aldri hardt.
 */
(function initializeConfiguration() {
  try {
    var r = validateConfiguration();
    var msg = r.isValid ? 'Konfig OK' : 'Konfig med avvik';
    _configLog_('INFO', 'initializeConfiguration', { message: msg, warnings: r.warnings.length, errors: r.errors.length });
  } catch (e) {
    _configLog_('ERROR', 'initializeConfiguration', { error: e && e.message });
  }
})();

/* ============================ Dev-verktøy (opsjon) ======================== */

/**
 * Aktiverer enkle utviklerverktøy i development-miljø.
 */
function setupDevelopmentTools() {
  if (!isFeatureEnabled('DEVELOPMENT_TOOLS')) return;
  try {
    globalThis.DEV_CONFIG_TOOLS = {
      validate: validateConfiguration,
      getUIConfig: getUIConfig,
      getSheetMetadata: getSheetMetadata,
      isFeatureEnabled: isFeatureEnabled,
      getProp: _getProp_,
      setProp: _setProp_
    };
    _configLog_('INFO', 'setupDevelopmentTools', { enabled: true });
  } catch (e) {
    _configLog_('WARN', 'setupDevelopmentTools', { error: e && e.message });
  }
}
