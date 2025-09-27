/* ====================== Roller & Tilgangskontroll (RBAC) ======================
 * FILE: 08_Roller_og_Tilgang_RBAC.gs | VERSION: 1.7.3 | UPDATED: 2025-09-14
 * FORMÅL: Sentralisert logikk for å bestemme brukerroller og validere tillatelser.
 * NYTT v1.7.3: Lagt til den formelle rollen "Leietaker".
 * =========================================================================== */

/* ---------- Konfigurasjon ---------- */

/** Kanoniske rollenavn som brukes i systemet. */
const RBAC_CANONICAL_ROLES = Object.freeze([
  'Gjest', 'Beboer', 'Seksjonseier', 'Leietaker', 'Styremedlem', 'Vaktmester', 'Kjernebruker', 'Admin'
]);

/** Alias/normalisering av roller (case-insensitive). */
const RBAC_ROLE_ALIASES = Object.freeze({
  'gjest': 'Gjest',
  'beboer': 'Beboer',
  'eier': 'Seksjonseier',
  'seksjonseier': 'Seksjonseier',
  // NYE ALIASER for Leietaker
  'leietaker': 'Leietaker',
  'leieboer': 'Leietaker',
  'styre': 'Styremedlem',
  'styremedlem': 'Styremedlem',
  'vaktmester': 'Vaktmester',
  'kjernebruker': 'Kjernebruker',
  'admin': 'Admin',
  'administrator': 'Admin',
});

/**
 * Access Control List (ACL) – hvilke roller gir en tillatelse.
 * NB: Admin har alltid tilgang (implisitt).
 */
const PERMISSIONS = Object.freeze({
  'VIEW_ADMIN_MENU':       ['Styremedlem', 'Kjernebruker', 'Admin'],
  'EDIT_CONFIG':           ['Kjernebruker', 'Admin'],
  'VIEW_ALL_TASKS':        ['Styremedlem', 'Kjernebruker', 'Admin'],
  'EDIT_ALL_TASKS':        ['Styremedlem', 'Kjernebruker', 'Admin'],
  'VIEW_PERSON_REGISTER':  ['Styremedlem', 'Kjernebruker', 'Admin'],
  'EXPORT_DATA':           ['Styremedlem', 'Kjernebruker', 'Admin'],
  'VIEW_VAKTMESTER_UI':    ['Vaktmester', 'Admin'],
  'OPEN_MEETINGS_UI':     ['Styremedlem', 'Kjernebruker', 'Admin'],
  'GENERATE_REPORTS':     ['Styremedlem', 'Kjernebruker', 'Admin'],
  'VIEW_BUDGET_MENU':     ['Styremedlem', 'Kjernebruker', 'Admin']
});

/* ---------- Lettvektscache for roller ---------- */
var __RBAC_CACHE = {
  email: null,
  roles: null,
  ts: 0
};
const __RBAC_CACHE_MS = 2 * 60 * 1000; // 2 minutter

/* ---------- Hjelpere ---------- */

function _rbacLog_(type, msg){
  try {
    if (typeof _logEvent === 'function') _logEvent(type, msg);
    else Logger.log(type + '> ' + msg);
  } catch(_) {}
}

function _getEmailLower_(){
  try {
    const act = Session.getActiveUser() && Session.getActiveUser().getEmail();
    const eff = Session.getEffectiveUser() && Session.getEffectiveUser().getEmail();
    return String((act || eff || '')).trim().toLowerCase();
  } catch(_) {
    return '';
  }
}

function _normalizeRole_(raw){
  if (!raw) return null;
  const s = String(raw).trim().toLowerCase();
  const canon = RBAC_ROLE_ALIASES[s] || (s.charAt(0).toUpperCase() + s.slice(1));
  return RBAC_CANONICAL_ROLES.indexOf(canon) >= 0 ? canon : null;
}

function _splitRolesCell_(val){
  if (!val) return [];
  return String(val)
    .split(/[;,|\r\n]+/)
    .map(x => _normalizeRole_(x))
    .filter(Boolean);
}

function _getAdminWhitelist_(){
  try {
    if (typeof getConfigMap === 'function') {
      const cfg = getConfigMap();
      const raw = cfg && cfg['ADMIN_WHITELIST'];
      if (!raw) return [];
      return String(raw).split(/[,;\s\r\n]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
    }
  } catch(_) {}
  return [];
}

/* ---------- Kjernelogikk ---------- */

/**
 * Finner rollene til den innloggede brukeren.
 * Kilder:
 *  - Styret-arket (kolonner: Navn, E-post, Rolle) – støtte for flere roller pr celle.
 *  - Konfig ADMIN_WHITELIST – personer her får rollen "Admin".
 * @returns {string[]} Kanoniske roller (minst 'Beboer' hvis identifisert, ellers 'Gjest').
 */
function resolveCurrentUserRoles(){
  try {
    const now = Date.now();
    const me = _getEmailLower_();
    if (!me) return ['Gjest'];

    // Cache-treff?
    if (__RBAC_CACHE.email === me && (now - __RBAC_CACHE.ts) < __RBAC_CACHE_MS && Array.isArray(__RBAC_CACHE.roles)) {
      return __RBAC_CACHE.roles.slice();
    }

    const roles = new Set(['Beboer']); // baseline for identifiserte brukere

    // Admin via whitelist
    const admins = _getAdminWhitelist_();
    if (admins.indexOf(me) >= 0) roles.add('Admin');

    // Roller fra Styret-arket
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEETS.BOARD);
    if (sh && sh.getLastRow() > 1){
      // Forventer headers: ['Navn','E-post','Rolle']
      const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
      for (var i=0; i<vals.length; i++){
        var email = String(vals[i][1] || '').trim().toLowerCase();
        if (email === me){
          const cellRoles = _splitRolesCell_(vals[i][2] || 'Styremedlem');
          cellRoles.forEach(r => roles.add(r));
        }
      }
    }

    // Fallback hvis vi av en eller annen grunn mangler noe
    if (roles.size === 0) roles.add('Gjest');

    // Lagre i cache
    const out = Array.from(roles);
    __RBAC_CACHE = { email: me, roles: out.slice(), ts: now };
    return out;

  } catch (e) {
    _rbacLog_('RBAC_Error', 'resolveCurrentUserRoles feilet: ' + e.message);
    return ['Gjest'];
  }
}

/**
 * Sjekker om innlogget bruker har en gitt tillatelse.
 * @param {string} permission - Navn på tillatelse (fra PERMISSIONS).
 * @returns {boolean}
 */
function hasPermission(permission) {
  try {
    const required = PERMISSIONS[permission];
    if (!required) {
      _rbacLog_('RBAC_Warning', 'Ukjent tillatelse sjekket: ' + permission);
      return false;
    }
    const userRoles = resolveCurrentUserRoles();

    // Admin har alltid tilgang
    if (userRoles.indexOf('Admin') >= 0) return true;

    // Vanlig sjekk: minst én av rollene
    for (var i=0; i<userRoles.length; i++){
      if (required.indexOf(userRoles[i]) >= 0) return true;
    }
    return false;

  } catch (e) {
    _rbacLog_('RBAC_Error', "hasPermission('" + permission + "') feilet: " + e.message);
    return false;
  }
}

/**
 * Kaster en vennlig feil hvis brukeren ikke har tillatelse.
 * @param {string} permission
 * @param {string} [actionName] - Hva forsøkte brukeren å gjøre (for feilmelding).
 */
function requirePermission(permission, actionName){
  if (!hasPermission(permission)) {
    const me = _getEmailLower_();
    const msg = 'Du har ikke tilgang til denne handlingen'
      + (actionName ? (': ' + actionName) : '')
      + '. Kontakt administrator hvis du mener dette er feil.';
    _rbacLog_('RBAC_Deny', `Permission='${permission}' | User='${me}'`);
    throw new Error(msg);
  }
}

/**
 * Debug-hjelper: returnerer e-post, roller og en liste over tillatelser (true/false).
 * Nyttig for feilsøking i konsollen (View → Logs).
 */
function rbacDebug(){
  const me = _getEmailLower_();
  const roles = resolveCurrentUserRoles();
  const perms = {};
  Object.keys(PERMISSIONS).forEach(p => perms[p] = hasPermission(p));
  const info = { email: me, roles: roles, permissions: perms, ts: new Date().toISOString() };
  _rbacLog_('RBAC_Debug', JSON.stringify(info));
  return info;
}

