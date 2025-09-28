/**
 * Sameieportalen – Dynamisk Menybygger
 * FILE: MenuBuilder.gs
 * VERSION: 2.7.0
 * UPDATED: 2025-09-24
 *
 * FORMÅL:
 * - Bygger en dynamisk, sikker og robust meny i Google Sheets.
 * - Tilpasser menyen basert på brukerrolle (admin, vaktmester, bruker) og systemhelse.
 * - Skjuler valg når avhengigheter mangler (RSP, analyse) og tilbyr guiding.
 * - Viser helsetilstand (emoji) i menytittel + interaktiv helseside (sidebar).
 * - Inneholder “Hjelp / Kom i gang” med sjekkliste og hurtighandlinger.
 *
 * NYTT I 2.7.0:
 * - getConfig(): sentral konfigurator (DocProps → ScriptProps → default).
 * - Raskere _computeSystemHealth_() med samlet sheet-oppslag (Set).
 * - Forbedret feilbehandling: MenuBuilderError + handleMenuError().
 * - rateLimitCheck() brukt på sensitive operasjoner (eks. setRSPDocId).
 * - Testhooks: MENU_BUILDER (TEST_MODE, eksponerte helpers, mock-providers).
 *
 * HOVEDFUNKSJONER:
 * - onOpen(e): Låser, validerer og bygger meny (med helse-emoji + ev. toast).
 * - openHealthSidebar(): Helse-sidebar med status/warnings/snarveier + inline DOC_ID-lagring.
 * - openHelpDialog(): Kom-i-gang med sjekkliste og hurtighandlinger (lagres per bruker).
 * - openAboutDialog(): Om / Versjonsinfo (miljø, rolle, ID-er).
 * - openSetRSPDocIdDialog(), setRSPDocId(docId): Admin-dialog for DOC_ID (RSP) + lagring (rate limited).
 * - openKravDokument(): Viser lenke til kravdokumentet (hvis DOC_ID satt).
 * - clearMenuCache(), forceShowMenu(): Cache-rydding / tvangsoppbygging.
 * - showSystemInfo(): Admindiagnostikk.
 * - updateAdminEmails(), updateVaktmesterEmails(): Vedlikehold roller.
 */

/* ============================= Konfig & Avhengigheter ============================= */

var __MB_CFG__ = (function () {
  var out = {
    VERSION: '2.7.0',
    UPDATED: '2025-09-24',
    // onOpen-beskyttelse
    LOCK_TIMEOUT_MS: 15000,
    OPEN_DEBOUNCE_SECS: 3,
    // Ark vi sjekker i helse
    SHEETS: {
      SYNC_LOG: 'Sync_Log',
      VERSION_LOG: 'Version_Log',
      KRAV: 'Requirements'
    },
    PROP_KEYS: {
      RSP_DOC_ID: 'RSP_DOC_ID',
      ADMIN_EMAILS: 'ADMIN_EMAILS',
      VAKTMESTER_EMAILS: 'VAKTMESTER_EMAILS',
      ENVIRONMENT: 'ENVIRONMENT'
    },
    USER_PROP_KEYS: {
      ONBOARDING_STATE: 'MB_ONBOARDING_JSON' // per bruker sjekkliste
    },
    SIDEBAR: {
      TITLE: 'Systemhelse – Sameieportalen',
      WIDTH: 360
    },
    RATE_LIMIT: {
      WINDOW_SEC: 3600, // 1 time
      MAX_OPS: 5        // maks 5 forsøk per time
    }
  };
  try {
    if (typeof CONFIG_SHEETS === 'object' && CONFIG_SHEETS) {
      out.SHEETS.SYNC_LOG    = CONFIG_SHEETS.SYNC_LOG    || out.SHEETS.SYNC_LOG;
      out.SHEETS.VERSION_LOG = CONFIG_SHEETS.VERSION_LOG || out.SHEETS.VERSION_LOG;
      out.SHEETS.KRAV        = CONFIG_SHEETS.KRAV        || out.SHEETS.KRAV;
    }
  } catch (_) {}
  try {
    if (typeof CONFIG_PROP_KEYS === 'object' && CONFIG_PROP_KEYS) {
      out.PROP_KEYS.RSP_DOC_ID        = CONFIG_PROP_KEYS.RSP_DOC_ID        || out.PROP_KEYS.RSP_DOC_ID;
      out.PROP_KEYS.ENVIRONMENT       = CONFIG_PROP_KEYS.ENVIRONMENT       || out.PROP_KEYS.ENVIRONMENT;
      out.PROP_KEYS.ADMIN_EMAILS      = CONFIG_PROP_KEYS.ADMIN_EMAILS      || out.PROP_KEYS.ADMIN_EMAILS;
      out.PROP_KEYS.VAKTMESTER_EMAILS = CONFIG_PROP_KEYS.VAKTMESTER_EMAILS || out.PROP_KEYS.VAKTMESTER_EMAILS;
    }
  } catch (_) {}
  return out;
})();

/* ================================== Helpers ================================== */

/** Sentral config-henter: DocProps → ScriptProps → default. */
function getConfig(key, defaultValue) {
  try {
    var dp = PropertiesService.getDocumentProperties();
    var sp = PropertiesService.getScriptProperties();
    var v  = dp.getProperty(key);
    if (v == null) v = sp.getProperty(key);
    return (v != null ? v : defaultValue);
  } catch (_) {
    return defaultValue;
  }
}

/* ================================== Logger ================================== */

function _getMenuLogger_() {
  var consoleLike = {
    debug: function (fn, msg, data) { try { Logger.log('[DEBUG] %s: %s %s', fn||'', msg||'', data?JSON.stringify(data):''); } catch(_) {} },
    info:  function (fn, msg, data) { try { Logger.log('[INFO]  %s: %s %s', fn||'', msg||'', data?JSON.stringify(data):''); } catch(_) {} },
    warn:  function (fn, msg, data) { try { Logger.log('[WARN]  %s: %s %s', fn||'', msg||'', data?JSON.stringify(data):''); } catch(_) {} },
    error: function (fn, msg, data) { try { Logger.log('[ERROR] %s: %s %s', fn||'', msg||'', data?JSON.stringify(data):''); } catch(_) {} }
  };
  try {
    if (typeof _getLoggerPlus_ === 'function') return _getLoggerPlus_();
  } catch (_) {}
  return consoleLike;
}

/* =============================== Meny-konfig =============================== */

var MENU_CONFIG = {
  version: __MB_CFG__.VERSION,
  lastUpdated: __MB_CFG__.UPDATED,
  mainMenu: {
    name: 'Sameieportalen',
    items: [
      { id: 'dashboard',     text: 'Åpne Dashboard',             function: 'openDashboard',             required: true,  permission: 'user',       description: 'Oversikt og nøkkeltall' },
      { id: 'meetings',      text: 'Møteoversikt & Protokoller',  function: 'openMeetingsUI',            required: false, permission: 'user',       description: 'Administrer møter og protokoller' },
      { id: 'vaktmester',    text: 'Mine Oppgaver (Vaktmester)',  function: 'openVaktmesterUI',          required: false, permission: 'vaktmester', separator: 'after', description: 'Vaktmester-oppgaver' },
      { id: 'qualityCheck',  text: 'Kjør kvalitetssjekk',         function: 'runAllChecks',              required: true,  permission: 'user',       description: 'Valider integritet og helse' },
      { id: 'createSheets',  text: 'Opprett basisfaner',          function: 'createBaseSheets',          required: true,  permission: 'admin',      description: 'Grunnleggende datastruktur' },
      { id: 'healthSidebar', text: 'Vis Systemhelse (sidebar)',   function: 'openHealthSidebar',         required: false, permission: 'user',       description: 'Se status, advarsler og snarveier' },
      { id: 'help',          text: 'Hjelp / Kom i gang',          function: 'openHelpDialog',            required: false, permission: 'user',       description: 'Guidet sjekkliste og hurtighandlinger', separator: 'after' },
      { id: 'about',         text: 'Om / Versjonsinfo',           function: 'openAboutDialog',           required: false, permission: 'user',       description: 'Versjon, miljø og status' }
    ]
  },
  adminMenu: {
    name: 'Admin',
    items: [
      { id: 'docVersioning', text: 'Dokumentversjonering',       function: 'showDocVersioning',         required: false, permission: 'admin', description: 'Versjonskontroll for styringsdokumenter', separator: 'after' },
      // RSP (skjules automatisk hvis ikke tilgjengelig)
      { id: 'rspWizard',   text: 'Krav Sync – Førstegangsoppsett', function: 'rsp_menu_firstRunWizard', required: false, permission: 'admin', description: 'Wizard for kravdokument og synk (auto-seed/struktur)' },
      { id: 'rspValidate', text: 'Krav Sync – Valider',            function: 'rsp_menu_validate',       required: false, permission: 'admin', description: 'Sjekk systemtilstand før synk' },
      { id: 'rspPush',     text: 'Krav Sync – Push (Sheet → Doc)', function: 'rsp_menu_pushRun',        required: false, permission: 'admin', description: 'Oppdater DOC fra kravarket' },
      { id: 'rspPull',     text: 'Krav Sync – Pull (Doc → Sheet)', function: 'rsp_menu_pullRun',        required: false, permission: 'admin', description: 'Oppdater kravarket fra DOC' },

      // DOC ID / Kravdokument-hurtigtilgang
      { id: 'setDocId',    text: 'Sett RSP_DOC_ID…',               function: 'openSetRSPDocIdDialog',   required: false, permission: 'admin', description: 'Konfigurer DOC-ID for kravdokument' },
      { id: 'openDoc',     text: 'Åpne kravdokument',              function: 'openKravDokument',        required: false, permission: 'admin', description: 'Åpne kravdokumentet i ny fane (lenke)', separator: 'after' },

      // Analyse (skjules automatisk hvis ikke tilgjengelig)
      { id: 'analysisDash', text: 'Analyse – Oversikt',            function: 'openAnalysisDashboard',   required: false, permission: 'admin', description: 'Interaktiv analysevisning' },

      // System
      { id: 'clearMenuCache', text: 'Tøm meny-cache',              function: 'clearMenuCache',          required: false, permission: 'admin', description: 'Rydd functionExists-cache' },
      { id: 'clearCache',  text: 'Tøm dashboard-cache',            function: 'clearDashboardCache',     required: false, permission: 'admin', description: 'Fjern mellomlagret data' },
      { id: 'forceMenu',   text: 'Bygg meny på nytt',              function: 'forceShowMenu',           required: true,  permission: 'admin', description: 'Tving menyoppbygging' },
      { id: 'systemInfo',  text: 'Systemdiagnostikk',              function: 'showSystemInfo',          required: true,  permission: 'admin', description: 'Vis status og avhengigheter' },
      { id: 'configureUsers', text: 'Administrer brukerroller',    function: 'configureUserRoles',      required: false, permission: 'admin', description: 'Vedlikehold roller og e-postlister' }
    ]
  }
};

// Cache for function existence checks
var __menuFnExistCache = new Map();

/* =============================== Error handling =============================== */

function MenuBuilderError(message, code, recoverable) {
  this.name = 'MenuBuilderError';
  this.message = String(message || 'MenuBuilder error');
  this.code = code || 'GENERIC';
  this.recoverable = !!recoverable;
  this.stack = (new Error()).stack;
}
MenuBuilderError.prototype = Object.create(Error.prototype);
MenuBuilderError.prototype.constructor = MenuBuilderError;

function handleMenuError(error, context) {
  var log = _getMenuLogger_();
  try {
    if (error && error.recoverable) {
      log.warn('handleMenuError', 'Recoverable error – attempting recovery', { code: error.code, ctx: context });
      // Enkle recovery-strategier (kan utvides):
      if (error.code === 'LOCK_TIMEOUT') {
        _toast_('Et annet script bygger menyen. Prøv igjen om litt.', 'Info', 4);
        return;
      }
    }
    // Ikke recoverable → fallback
    buildBasicFallbackMenu();
  } catch (e) {
    // Siste skanse: minimal fallback-alert
    try { SpreadsheetApp.getUi().alert('Menyfeil', 'Kunne ikke bygge meny. Prøv "Bygg meny på nytt".', SpreadsheetApp.getUi().ButtonSet.OK); } catch(_) {}
    _getMenuLogger_().error('handleMenuError', 'Fallback also failed', { err: e && e.message });
  }
}

/* ================================ Rate limit ================================ */

function rateLimitCheck(operation) {
  try {
    var user = (Session.getActiveUser() && Session.getActiveUser().getEmail()) || 'anon';
    var cache = CacheService.getUserCache();
    var key = 'RL_' + operation + '_' + user;
    var count = parseInt(cache.get(key) || '0', 10) || 0;
    if (count >= __MB_CFG__.RATE_LIMIT.MAX_OPS) {
      throw new MenuBuilderError('Rate limit exceeded for '+operation, 'RATE_LIMIT', false);
    }
    cache.put(key, String(count + 1), __MB_CFG__.RATE_LIMIT.WINDOW_SEC);
  } catch (e) {
    if (e instanceof MenuBuilderError) throw e;
    // Fail-closed om cache ikke funker
  }
}

/* ================================= onOpen ================================= */

function onOpen(e) {
  var log = _getMenuLogger_();
  var start = Date.now();
  var fn = 'onOpen';

  // Debounce multiple rapid onOpen events
  try {
    var cache = CacheService.getScriptCache();
    var key = 'MENU_ONOPEN_DEBOUNCE';
    if (cache.get(key)) {
      log.warn(fn, 'Debounced duplicate onOpen within window.');
      return;
    }
    cache.put(key, '1', __MB_CFG__.OPEN_DEBOUNCE_SECS);
  } catch (_) {}

  // Global lock to avoid concurrent builds
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(__MB_CFG__.LOCK_TIMEOUT_MS);
  } catch (err) {
    handleMenuError(new MenuBuilderError(err && err.message, 'LOCK_TIMEOUT', true), { phase: 'acquire_lock' });
    return;
  }

  try {
    var ui = SpreadsheetApp.getUi();
    var role = getCurrentUserRole();
    var health = _computeSystemHealth_(); // {emoji, warnings, docIdOk}

    // Build main menu with health badge
    var mainName = MENU_CONFIG.mainMenu.name + ' ' + health.emoji;
    var mainMenu = ui.createMenu(mainName);
    var addedMain = _appendConfiguredItems_(mainMenu, MENU_CONFIG.mainMenu.items, role);

    // Admin submenu if applicable
    if (hasPermission(role, 'admin')) {
      var adminMenu = ui.createMenu(MENU_CONFIG.adminMenu.name);
      var addedAdmin = _appendConfiguredItems_(adminMenu, MENU_CONFIG.adminMenu.items, role, true);
      if (addedAdmin > 0) {
        mainMenu.addSeparator();
        mainMenu.addSubMenu(adminMenu);
      }
    }

    mainMenu.addToUi();
    log.info(fn, 'Menu built', { role: role, addedMain: addedMain, ms: Date.now() - start, healthWarnings: health.warnings });

    if (health.emoji === '❌') {
      _toast_('Systemhelse: kritisk – åpne «Vis Systemhelse (sidebar)» for detaljer.', 'Kritisk', 6);
    }

  } catch (error) {
    log.error(fn, 'Menu build failed; using fallback', { error: error && error.message, stack: error && error.stack });
    handleMenuError(error, { phase: 'build_menu' });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ============================== Builder helpers ============================== */

function _appendConfiguredItems_(menu, items, role) {
  var log = _getMenuLogger_();
  var count = 0;

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    try {
      if (!hasPermission(role, it.permission)) continue;

      // Dependency gating: if function is absent, hide (log if required)
      if (!functionExists(it.function)) {
        if (it.required) {
          log.warn('_appendConfiguredItems_', 'Required function missing; item hidden', { id: it.id, fn: it.function });
        }
        continue;
      }

      menu.addItem(it.text, it.function);
      count++;

      if (it.separator === 'after') menu.addSeparator();

    } catch (e) {
      log.error('_appendConfiguredItems_', 'Failed to add menu item', { id: it && it.id, error: e && e.message });
    }
  }
  return count;
}

/* =============================== Role handling =============================== */

function getCurrentUserRole() {
  var log = _getMenuLogger_();
  try {
    var email = (Session.getActiveUser().getEmail() || '').toLowerCase();
    if (!email) return 'user';

    var admins = _readEmailListProp_(__MB_CFG__.PROP_KEYS.ADMIN_EMAILS);
    if (admins.indexOf(email) >= 0) return 'admin';

    var vaktm = _readEmailListProp_(__MB_CFG__.PROP_KEYS.VAKTMESTER_EMAILS);
    if (vaktm.indexOf(email) >= 0) return 'vaktmester';

    return 'user';
  } catch (e) {
    log.warn('getCurrentUserRole', 'Falling back to user', { error: e && e.message });
    return 'user';
  }
}

function _readEmailListProp_(key) {
  var out = [];
  try {
    var raw = getConfig(key, null);
    if (raw) {
      var arr = JSON.parse(raw);
      if (Array.isArray(arr)) {
        out = arr.map(function (s) { return String(s || '').trim().toLowerCase(); }).filter(Boolean);
      }
    }
  } catch (_) {}
  return out;
}

function hasPermission(userRole, requiredPermission) {
  var map = {
    'user': ['user'],
    'vaktmester': ['user', 'vaktmester'],
    'admin': ['user', 'vaktmester', 'admin']
  };
  return (map[userRole] || []).indexOf(requiredPermission) >= 0;
}

/* ============================= Function existence ============================ */

function functionExists(functionName) {
  if (!functionName) return false;
  if (__menuFnExistCache.has(functionName)) return __menuFnExistCache.get(functionName);
  var ok = false;
  try { ok = (typeof globalThis[functionName] === 'function'); } catch (_) { ok = false; }
  __menuFnExistCache.set(functionName, ok);
  return ok;
}

/* ================================ Health probe =============================== */

function _computeSystemHealth_() {
  var log = _getMenuLogger_();
  var warnings = [];
  var ss = SpreadsheetApp.getActive();

  // Batch sheet name discovery
  var namesSet = (function() {
    try {
      var arr = ss.getSheets().map(function(s){ return s.getName(); });
      return new Set(arr);
    } catch (e) {
      // Fallback: probe individually if needed (unlikely to be faster)
      return new Set([]);
    }
  })();

  // Required sheets
  [__MB_CFG__.SHEETS.KRAV, __MB_CFG__.SHEETS.SYNC_LOG, __MB_CFG__.SHEETS.VERSION_LOG].forEach(function (name) {
    try {
      if (!namesSet.size) {
        if (!ss.getSheetByName(name)) warnings.push('Mangler ark: ' + name);
      } else if (!namesSet.has(name)) {
        warnings.push('Mangler ark: ' + name);
      }
    } catch (e) {
      warnings.push('Feil ved sjekk av ark "' + name + '": ' + (e && e.message));
    }
  });

  // RSP Document ID presence (optional but recommended)
  var docIdOk = false;
  var docId = '';
  try {
    docId = getConfig(__MB_CFG__.PROP_KEYS.RSP_DOC_ID, '') || '';
    docIdOk = !!docId && /^[A-Za-z0-9\-_]{20,}$/.test(docId);
    if (!docIdOk) warnings.push('RSP_DOC_ID ikke konfigurert eller ugyldig.');
  } catch (_) {
    warnings.push('Klarte ikke å lese RSP_DOC_ID.');
  }

  var critical = (!namesSet.has(__MB_CFG__.SHEETS.KRAV) || !docIdOk);
  var emoji = warnings.length === 0 ? '✅' : (critical ? '❌' : '⚠️');

  log.debug('_computeSystemHealth_', 'health', { emoji: emoji, warnings: warnings, docIdPreview: docId ? (docId.substring(0,6)+'…') : '' });
  return { emoji: emoji, warnings: warnings, docId: docId, docIdOk: docIdOk };
}

/* ================================== Sidebar ================================= */

function openHealthSidebar() {
  var health = _computeSystemHealth_();
  var ss = SpreadsheetApp.getActive();
  var hasWizard   = functionExists('rsp_menu_firstRunWizard');
  var hasValidate = functionExists('rsp_menu_validate');
  var hasPush     = functionExists('rsp_menu_pushRun');
  var hasPull     = functionExists('rsp_menu_pullRun');

  var warningsHtml = (health.warnings.length === 0)
    ? '<li>Ingen advarsler 🎉</li>'
    : health.warnings.map(function (w) { return '<li>' + _escHtml_(w) + '</li>'; }).join('');

  var actions = [];
  if (hasWizard)   actions.push('<button onclick="google.script.run.rsp_menu_firstRunWizard()">Kjør Førstegangsoppsett</button>');
  if (hasValidate) actions.push('<button onclick="google.script.run.rsp_menu_validate()">Valider konfig</button>');
  if (hasPush)     actions.push('<button onclick="google.script.run.rsp_menu_pushRun()">Push Sheet → Doc</button>');
  if (hasPull)     actions.push('<button onclick="google.script.run.rsp_menu_pullRun()">Pull Doc → Sheet</button>');

  var docIdConfigurator = '';
  if (!health.docIdOk) {
    docIdConfigurator =
      '<div style="margin-top:10px;padding:8px;border:1px solid #e3e3e3;border-radius:6px;background:#fafafa;">' +
        '<div style="font-weight:600;margin-bottom:6px;">Konfigurer RSP_DOC_ID</div>' +
        '<div style="font-size:12px;color:#555;margin-bottom:6px;">Lim inn dokument-ID til kravdokumentet (Google Docs).</div>' +
        '<input id="docid" type="text" style="width:100%;box-sizing:border-box;padding:6px;border:1px solid #ccc;border-radius:4px;" placeholder="1a2B3c_dEfG-…">' +
        '<div style="display:flex;gap:6px;margin-top:8px;">' +
          '<button onclick="saveDocId()" style="flex:0 0 auto;">Lagre</button>' +
          '<span id="docidStatus" style="font-size:12px;color:#777;align-self:center;"></span>' +
        '</div>' +
      '</div>' +
      '<script>' +
      'function saveDocId(){' +
        'var el=document.getElementById("docid"); var v=(el&&el.value||"").trim();' +
        'var status=document.getElementById("docidStatus");' +
        'status.textContent="Lagrer…";' +
        'google.script.run.withSuccessHandler(function(res){' +
          'status.style.color = res.ok ? "#126300" : "#b00020";' +
          'status.textContent = res.message;' +
          'if(res.ok){ setTimeout(function(){ google.script.host.close(); }, 900); }' +
        '}).withFailureHandler(function(err){' +
          'status.style.color="#b00020"; status.textContent="Feil: "+(err && err.message || err);' +
        '}).setRSPDocId(v);' +
      '}' +
      '</script>';
  }

  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;padding:12px 10px 20px 10px;max-width:540px;">' +
      '<h3 style="margin:0 0 8px 0;">Systemhelse ' + _escHtml_(health.emoji) + '</h3>' +
      '<div style="font-size:12px;color:#555;margin-bottom:8px;">' +
        'Regneark-ID: ' + _escHtml_(ss.getId()) + '<br>' +
        'RSP_DOC_ID: ' + (health.docId ? _escHtml_(health.docId.substring(0,12) + '…') : '<i>ikke satt</i>') +
      '</div>' +
      docIdConfigurator +
      '<h4 style="margin:12px 0 6px 0;">Advarsler</h4>' +
      '<ul style="margin:0 0 10px 16px;">' + warningsHtml + '</ul>' +
      '<h4 style="margin:12px 0 6px 0;">Handlinger</h4>' +
      (actions.length ? actions.join(' ') : '<div style="color:#777;">Ingen tilgjengelige handlinger (moduler ikke lastet).</div>') +
      '<div style="margin-top:12px;border-top:1px solid #eee;padding-top:10px;color:#777;font-size:12px;">' +
        'Versjon: ' + _escHtml_(__MB_CFG__.VERSION) + ' · Oppdatert: ' + _escHtml_(__MB_CFG__.UPDATED) +
      '</div>' +
    '</div>'
  );
  html.setTitle(__MB_CFG__.SIDEBAR.TITLE);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ============================= Hjelp / Kom i gang ============================ */

function openHelpDialog() {
  var up = PropertiesService.getUserProperties();
  var state = _readOnboardingState_(up);
  var steps = _defaultOnboardingSteps_();

  // Render sjekkliste
  var itemsHtml = steps.map(function (st) {
    var done = !!state[st.key];
    var disabled = !functionExists(st.actionFn) && st.actionFn !== 'SET_RSP_DOC_ID';
    var label = _escHtml_(st.label);
    var hint  = st.hint ? ('<div style="font-size:11px;color:#777;margin-top:2px;">' + _escHtml_(st.hint) + '</div>') : '';
    var btn   = '';

    if (st.actionFn === 'SET_RSP_DOC_ID') {
      btn = '<button onclick="setDocIdPrompt()" ' + (done?'disabled':'') + '>Sett RSP_DOC_ID…</button>';
    } else if (!disabled) {
      btn = '<button onclick="runAction(\'' + _escHtmlAttr_(st.actionFn) + '\')" ' + (done?'disabled':'') + '>' + _escHtml_(st.button||'Kjør') + '</button>';
    } else {
      btn = '<button disabled>Ikke tilgjengelig</button>';
    }

    return (
      '<li style="margin-bottom:10px;">' +
        '<label style="display:flex;align-items:center;gap:8px;">' +
        '<input type="checkbox" ' + (done ? 'checked' : '') + ' onchange="toggleStep(\''+ _escHtmlAttr_(st.key) +'\', this.checked)">' +
        '<span>' + label + '</span>' +
        '</label>' +
        hint +
        '<div style="margin-top:6px;">' + btn + '</div>' +
      '</li>'
    );
  }).join('');

  var role = getCurrentUserRole();

  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;padding:12px;max-width:600px;">' +
      '<h3 style="margin:0 0 8px 0;">Hjelp / Kom i gang</h3>' +
      '<div style="font-size:12px;color:#555;margin-bottom:10px;">' +
        'Din rolle: <b>' + _escHtml_(role) + '</b>. Denne sjekklisten lagres for deg og kan endres når som helst.' +
      '</div>' +
      '<ol style="margin:0 0 12px 18px;padding:0;">' + itemsHtml + '</ol>' +
      '<div style="display:flex;gap:8px;margin-top:8px;">' +
        '<button onclick="google.script.host.close()">Lukk</button>' +
        '<button onclick="resetChecklist()" style="margin-left:auto;">Nullstill sjekkliste</button>' +
      '</div>' +
      '<script>' +
      'function toggleStep(key, checked){' +
        'google.script.run.withSuccessHandler(function(){})' +
          '.setOnboardingStep(key, !!checked);' +
      '}' +
      'function resetChecklist(){' +
        'google.script.run.withSuccessHandler(function(){ location.reload(); })' +
          '.resetOnboardingChecklist();' +
      '}' +
      'function runAction(fn){' +
        'if(!fn) return;' +
        'var api = google.script.run.withFailureHandler(function(err){alert("Feil: "+(err&&err.message||err));});' +
        'api[fn]();' +
      '}' +
      'function setDocIdPrompt(){' +
        'google.script.run.withFailureHandler(function(err){alert("Feil: "+(err&&err.message||err));}).openSetRSPDocIdDialog();' +
      '}' +
      '</script>' +
    '</div>'
  ).setWidth(520).setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, 'Hjelp / Kom i gang');
}

function _defaultOnboardingSteps_() {
  return [
    { key: 'docid', label: 'Sett RSP_DOC_ID (kravdokument-ID)', button: 'Sett…', hint: 'Binder RSP-synk til riktig dokument', actionFn: 'SET_RSP_DOC_ID' },
    { key: 'wizard', label: 'Kjør Krav Sync – Førstegangsoppsett', button: 'Kjør wizard', hint: 'Oppretter basestruktur og initierer DOC', actionFn: 'rsp_menu_firstRunWizard' },
    { key: 'validate', label: 'Valider konfigurasjon', button: 'Valider', hint: 'Sjekker at systemet er klart for synk', actionFn: 'rsp_menu_validate' },
    { key: 'push', label: 'Push (Sheet → Doc)', button: 'Push nå', hint: 'Generer/oppdater kravseksjoner i dokumentet', actionFn: 'rsp_menu_pushRun' },
    { key: 'pull', label: 'Pull (Doc → Sheet)', button: 'Pull nå', hint: 'Les endringer fra dokumentet tilbake til arket', actionFn: 'rsp_menu_pullRun' }
  ];
}

function _readOnboardingState_(userProps) {
  try {
    var raw = userProps.getProperty(__MB_CFG__.USER_PROP_KEYS.ONBOARDING_STATE);
    if (raw) {
      var obj = JSON.parse(raw);
      if (obj && typeof obj === 'object') return obj;
    }
  } catch (_) {}
  return {}; // tomt state
}

function setOnboardingStep(key, value) {
  var up = PropertiesService.getUserProperties();
  var state = _readOnboardingState_(up);
  state[String(key||'').trim()] = !!value;
  up.setProperty(__MB_CFG__.USER_PROP_KEYS.ONBOARDING_STATE, JSON.stringify(state));
}

function resetOnboardingChecklist() {
  var up = PropertiesService.getUserProperties();
  up.deleteProperty(__MB_CFG__.USER_PROP_KEYS.ONBOARDING_STATE);
}

/* =============================== Om / Versjon =============================== */

function openAboutDialog() {
  var env = (typeof APP === 'object' && APP && APP.ENVIRONMENT) ?
             APP.ENVIRONMENT : (getConfig(__MB_CFG__.PROP_KEYS.ENVIRONMENT, 'production'));
  var role = getCurrentUserRole();
  var ssId = SpreadsheetApp.getActive().getId();

  var parts = [
    '<div style="font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;padding:12px;max-width:520px;">',
    '<h3 style="margin:0 0 8px 0;">Om / Versjonsinfo</h3>',
    '<div style="font-size:13px;margin-bottom:8px;"><b>Sameieportalen – Menybygger</b></div>',
    '<div style="font-size:12px;color:#555;">Versjon: <b>'+ _escHtml_(__MB_CFG__.VERSION) +'</b> · Oppdatert: '+ _escHtml_(__MB_CFG__.UPDATED) +'</div>',
    '<div style="font-size:12px;color:#555;">Miljø: <b>'+ _escHtml_(env) +'</b> · Din rolle: <b>'+ _escHtml_(role) +'</b></div>',
    '<div style="font-size:12px;color:#555;margin-top:6px;">Regneark-ID: '+ _escHtml_(ssId) +'</div>',
    '<div style="border-top:1px solid #eee;margin:10px 0;"></div>',
    '<div style="font-size:12px;color:#444;">',
    '<ul style="margin:0 0 0 16px;padding:0;">',
    '<li>Helse-sidebar: <code>openHealthSidebar()</code></li>',
    '<li>Kom i gang: <code>openHelpDialog()</code></li>',
    '<li>Systemdiagnostikk: <code>showSystemInfo()</code></li>',
    '</ul>',
    '</div>',
    '<div style="margin-top:10px;display:flex;gap:8px;">',
    '<button onclick="google.script.host.close()">Lukk</button>',
    '</div>',
    '</div>'
  ].join('');

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(parts).setWidth(420).setHeight(220), 'Om / Versjonsinfo');
}

/* ======================== RSP_DOC_ID dialog og lagring ======================= */

function openSetRSPDocIdDialog() {
  var current = getRSPDocId_() || '';
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;padding:12px;max-width:520px;">' +
      '<h3 style="margin:0 0 10px 0;">Sett RSP_DOC_ID</h3>' +
      '<div style="font-size:12px;color:#555;margin-bottom:6px;">Angi dokument-ID til kravdokumentet (Google Docs). Eksempel: <code>1a2B3c_dEfG-…</code></div>' +
      '<input id="docid" type="text" value="' + _escHtml_(current) + '" ' +
            'style="width:100%;box-sizing:border-box;padding:8px;border:1px solid #ccc;border-radius:4px;" ' +
            'placeholder="1a2B3c_dEfG-…">' +
      '<div style="display:flex;gap:8px;margin-top:10px;">' +
        '<button onclick="saveDocId()">Lagre</button>' +
        '<button onclick="google.script.host.close()">Avbryt</button>' +
        '<span id="docidStatus" style="font-size:12px;color:#777;align-self:center;"></span>' +
      '</div>' +
      '<script>' +
      'function saveDocId(){' +
        'var el=document.getElementById("docid"); var v=(el&&el.value||"").trim();' +
        'var status=document.getElementById("docidStatus");' +
        'if(!/^[A-Za-z0-9\\-_]{20,}$/.test(v)){ status.style.color="#b00020"; status.textContent="Ugyldig ID-format."; return; }' +
        'status.textContent="Lagrer…";' +
        'google.script.run.withSuccessHandler(function(res){' +
          'status.style.color = res.ok ? "#126300" : "#b00020";' +
          'status.textContent = res.message;' +
          'if(res.ok){ setTimeout(function(){ google.script.host.close(); }, 900); }' +
        '}).withFailureHandler(function(err){' +
          'status.style.color="#b00020"; status.textContent="Feil: "+(err && err.message || err);' +
        '}).setRSPDocId(v);' +
      '}' +
      '</script>' +
    '</div>'
  );
  html.setWidth(420).setHeight(170);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sett RSP_DOC_ID');
}

function setRSPDocId(docId) {
  rateLimitCheck('setRSPDocId'); // rate limiting
  var log = _getMenuLogger_();
  try {
    docId = String(docId || '').trim();
    if (!/^[A-Za-z0-9\-_]{20,}$/.test(docId)) {
      return { ok: false, message: 'Ugyldig dokument-ID.' };
    }
    try {
      var doc = DocumentApp.openById(docId);
      var name = doc.getName();
      doc.saveAndClose();
      log.info('setRSPDocId', 'Doc OK', { name: name });
    } catch (e) {
      log.warn('setRSPDocId', 'Kunne ikke bekrefte dokumenttilgang nå – lagrer likevel.', { error: e && e.message });
    }
    PropertiesService.getScriptProperties().setProperty(__MB_CFG__.PROP_KEYS.RSP_DOC_ID, docId);
    _toast_('RSP_DOC_ID lagret. Menyen oppdateres…', 'Konfig', 4);
    try { Utilities.sleep(400); } catch(_) {}
    try { forceShowMenu(); } catch(_) {}
    return { ok: true, message: 'Lagret ✅' };
  } catch (err) {
    log.error('setRSPDocId', 'Feil ved lagring', { error: err && err.message });
    return { ok: false, message: 'Kunne ikke lagre: ' + (err && err.message) };
  }
}

function openKravDokument() {
  var docId = getRSPDocId_();
  if (!docId) {
    SpreadsheetApp.getUi().alert('Kravdokument', 'RSP_DOC_ID er ikke satt. Bruk «Sett RSP_DOC_ID…».', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var url = 'https://docs.google.com/document/d/' + docId + '/edit';
  SpreadsheetApp.getUi().alert('Kravdokument', 'Åpne dokumentet:\n\n' + url, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ================================= Helpers ================================= */

function getRSPDocId_() {
  try {
    return getConfig(__MB_CFG__.PROP_KEYS.RSP_DOC_ID, '') || '';
  } catch (_) { return ''; }
}

function _escHtml_(s) {
  s = String(s || '');
  return s.replace(/[&<>"']/g, function (m) {
    return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]);
  });
}

function _escHtmlAttr_(s) {
  return _escHtml_(s).replace(/"/g, '&quot;');
}

/* ================================== UI bits ================================= */

function buildBasicFallbackMenu() {
  var log = _getMenuLogger_();
  try {
    var ui = SpreadsheetApp.getUi();
    var m = ui.createMenu('Sameieportalen (Feilmodus)');
    if (functionExists('forceShowMenu')) m.addItem('Bygg meny på nytt', 'forceShowMenu');
    m.addToUi();
    log.warn('buildBasicFallbackMenu', 'Fallback menu exposed.');
  } catch (e) {
    log.error('buildBasicFallbackMenu', 'Failed building fallback menu', { error: e && e.message });
  }
}

function forceShowMenu() {
  rateLimitCheck('forceShowMenu'); // rate limiting
  var log = _getMenuLogger_();
  try {
    __menuFnExistCache.clear();
    onOpen();
    SpreadsheetApp.getUi().alert('Meny oppdatert', 'Menyen er bygget på nytt.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    log.error('forceShowMenu', 'Failed to rebuild menu', { error: e && e.message, stack: e && e.stack });
    SpreadsheetApp.getUi().alert('Feil', 'Kunne ikke bygge menyen: ' + (e && e.message), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function clearMenuCache() {
  __menuFnExistCache.clear();
  _toast_('Meny-cache tømt', 'Cache', 3);
}

/* ================================ Diagnose ================================ */

function showSystemInfo() {
  var log = _getMenuLogger_();
  try {
    var allFns = []
      .concat((MENU_CONFIG.mainMenu.items || []).map(function (x) { return x.function; }))
      .concat((MENU_CONFIG.adminMenu.items || []).map(function (x) { return x.function; }));
    var uniq = Array.from(new Set(allFns));

    var info = {
      timestamp: new Date().toISOString(),
      user: (Session.getActiveUser() && Session.getActiveUser().getEmail()) || '',
      role: getCurrentUserRole(),
      spreadsheetId: SpreadsheetApp.getActive().getId(),
      environment: (typeof APP === 'object' && APP && APP.ENVIRONMENT) ? APP.ENVIRONMENT :
                   (getConfig(__MB_CFG__.PROP_KEYS.ENVIRONMENT, 'production')),
      functionCacheSize: __menuFnExistCache.size,
      keyFunctionStatus: uniq.map(function (f) { return { name: f, available: functionExists(f) }; })
    };

    SpreadsheetApp.getUi().alert('Systeminfo', JSON.stringify(info, null, 2), SpreadsheetApp.getUi().ButtonSet.OK);
    log.info('showSystemInfo', 'Presented system info.', info);

  } catch (e) {
    log.error('showSystemInfo', 'Failed system info', { error: e && e.message });
    SpreadsheetApp.getUi().alert('Feil', 'Kunne ikke hente systeminfo: ' + (e && e.message), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/* ============================== Role updaters =============================== */

function updateAdminEmails(emails) {
  if (!Array.isArray(emails)) throw new Error('Input må være en liste (array) med e-postadresser.');
  PropertiesService.getDocumentProperties().setProperty(__MB_CFG__.PROP_KEYS.ADMIN_EMAILS, JSON.stringify(emails));
  _getMenuLogger_().info('updateAdminEmails', 'Oppdatert admin-e-poster.', { count: emails.length });
}

function updateVaktmesterEmails(emails) {
  if (!Array.isArray(emails)) throw new Error('Input må være en liste (array) med e-postadresser.');
  PropertiesService.getDocumentProperties().setProperty(__MB_CFG__.PROP_KEYS.VAKTMESTER_EMAILS, JSON.stringify(emails));
  _getMenuLogger_().info('updateVaktmesterEmails', 'Oppdatert vaktmester-e-poster.', { count: emails.length });
}

/* ================================ Utilities ================================= */

function _toast_(msg, title, seconds) {
  try {
    SpreadsheetApp.getActive().toast(String(msg || ''), String(title || ''), Math.max(1, Number(seconds) || 3));
  } catch (_) {}
}

/* ============================== Testing hooks =============================== */

var MENU_BUILDER = {
  TEST_MODE: false,
  // Expose internals for unit tests
  _computeSystemHealth_: _computeSystemHealth_,
  _readOnboardingState_: _readOnboardingState_,
  _appendConfiguredItems_: _appendConfiguredItems_,
  // Mocks (simple placeholders; expand as needed)
  getMockUI: function () {
    return {
      _items: [],
      createMenu: function (name) { return this; },
      addItem: function () { this._items.push('item'); return this; },
      addSeparator: function () { this._items.push('sep'); return this; },
      addSubMenu: function () { this._items.push('submenu'); return this; },
      addToUi: function () { return this; }
    };
  }
};
