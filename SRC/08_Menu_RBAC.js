/* ====================== Roller & Tilgangskontroll (RBAC) ======================
 * FILE: 08_Menu_RBAC.gs | VERSION: 2.0.0 | UPDATED: 2025-09-26
 * FORMÅL: Sentralisert logikk for å bestemme brukerroller og validere tillatelser.
 * ENDRINGER v2.0.0:
 *  - Modernisert til let/const og arrow functions.
 *  - Innkapslet i en IIFE for å unngå globale variabler.
 *  - Forenklet og forbedret logikk for rollehåndtering.
 * =========================================================================== */

(() => {
  const RBAC_CANONICAL_ROLES = Object.freeze([
    'Gjest', 'Beboer', 'Seksjonseier', 'Leietaker', 'Styremedlem', 'Vaktmester', 'Kjernebruker', 'Admin'
  ]);

  const RBAC_ROLE_ALIASES = Object.freeze({
    'gjest': 'Gjest',
    'beboer': 'Beboer',
    'eier': 'Seksjonseier',
    'seksjonseier': 'Seksjonseier',
    'leietaker': 'Leietaker',
    'leieboer': 'Leietaker',
    'styre': 'Styremedlem',
    'styremedlem': 'Styremedlem',
    'vaktmester': 'Vaktmester',
    'kjernebruker': 'Kjernebruker',
    'admin': 'Admin',
    'administrator': 'Admin',
  });

  const PERMISSIONS = Object.freeze({
    'VIEW_ADMIN_MENU': ['Styremedlem', 'Kjernebruker', 'Admin'],
    'EDIT_CONFIG': ['Kjernebruker', 'Admin'],
    'VIEW_ALL_TASKS': ['Styremedlem', 'Kjernebruker', 'Admin'],
    'EDIT_ALL_TASKS': ['Styremedlem', 'Kjernebruker', 'Admin'],
    'VIEW_PERSON_REGISTER': ['Styremedlem', 'Kjernebruker', 'Admin'],
    'EXPORT_DATA': ['Styremedlem', 'Kjernebruker', 'Admin'],
    'VIEW_VAKTMESTER_UI': ['Vaktmester', 'Admin'],
    'OPEN_MEETINGS_UI': ['Styremedlem', 'Kjernebruker', 'Admin'],
    'GENERATE_REPORTS': ['Styremedlem', 'Kjernebruker', 'Admin'],
    'VIEW_BUDGET_MENU': ['Styremedlem', 'Kjernebruker', 'Admin']
  });

  let __RBAC_CACHE = {
    email: null,
    roles: null,
    ts: 0
  };
  const __RBAC_CACHE_MS = 2 * 60 * 1000;

  const _rbacLog_ = (type, msg) => {
    try {
      if (typeof _logEvent === 'function') _logEvent(type, msg);
      else Logger.log(`${type}> ${msg}`);
    } catch (e) { /* ignore */ }
  };

  const _getEmailLower_ = () => {
    try {
      const activeUser = Session.getActiveUser()?.getEmail();
      const effectiveUser = Session.getEffectiveUser()?.getEmail();
      return (activeUser || effectiveUser || '').trim().toLowerCase();
    } catch (e) {
      return '';
    }
  };

  const _normalizeRole_ = (raw) => {
    if (!raw) return null;
    const s = String(raw).trim().toLowerCase();
    const canon = RBAC_ROLE_ALIASES[s] || (s.charAt(0).toUpperCase() + s.slice(1));
    return RBAC_CANONICAL_ROLES.includes(canon) ? canon : null;
  };

  const _splitRolesCell_ = (val) => {
    if (!val) return [];
    return String(val).split(/[;,|\r\n]+/).map(_normalizeRole_).filter(Boolean);
  };

  const _getAdminWhitelist_ = () => {
    try {
      if (typeof getConfigMap === 'function') {
        const cfg = getConfigMap();
        const raw = cfg?.ADMIN_WHITELIST;
        if (!raw) return [];
        return String(raw).split(/[,;\s\r\n]+/).map(s => s.trim().toLowerCase()).filter(Boolean);
      }
    } catch (e) { /* ignore */ }
    return [];
  };

  globalThis.resolveCurrentUserRoles = () => {
    try {
      const now = Date.now();
      const me = _getEmailLower_();
      if (!me) return ['Gjest'];

      if (__RBAC_CACHE.email === me && (now - __RBAC_CACHE.ts) < __RBAC_CACHE_MS && Array.isArray(__RBAC_CACHE.roles)) {
        return [...__RBAC_CACHE.roles];
      }

      const roles = new Set(['Beboer']);
      const admins = _getAdminWhitelist_();
      if (admins.includes(me)) roles.add('Admin');

      const ss = SpreadsheetApp.getActive();
      const sh = ss.getSheetByName(SHEETS.BOARD);
      if (sh && sh.getLastRow() > 1) {
        const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
        for (const row of vals) {
          const email = String(row[1] || '').trim().toLowerCase();
          if (email === me) {
            const cellRoles = _splitRolesCell_(row[2] || 'Styremedlem');
            cellRoles.forEach(r => roles.add(r));
          }
        }
      }

      if (roles.size === 0) roles.add('Gjest');

      const out = [...roles];
      __RBAC_CACHE = { email: me, roles: [...out], ts: now };
      return out;

    } catch (e) {
      _rbacLog_('RBAC_Error', `resolveCurrentUserRoles feilet: ${e.message}`);
      return ['Gjest'];
    }
  };

  globalThis.hasPermission = (permission) => {
    try {
      const required = PERMISSIONS[permission];
      if (!required) {
        _rbacLog_('RBAC_Warning', `Ukjent tillatelse sjekket: ${permission}`);
        return false;
      }
      const userRoles = resolveCurrentUserRoles();
      if (userRoles.includes('Admin')) return true;
      return userRoles.some(role => required.includes(role));
    } catch (e) {
      _rbacLog_('RBAC_Error', `hasPermission('${permission}') feilet: ${e.message}`);
      return false;
    }
  };

  globalThis.requirePermission = (permission, actionName) => {
    if (!hasPermission(permission)) {
      const me = _getEmailLower_();
      const msg = `Du har ikke tilgang til denne handlingen${actionName ? `: ${actionName}` : ''}. Kontakt administrator hvis du mener dette er feil.`;
      _rbacLog_('RBAC_Deny', `Permission='${permission}' | User='${me}'`);
      throw new Error(msg);
    }
  };

  globalThis.rbacDebug = () => {
    const me = _getEmailLower_();
    const roles = resolveCurrentUserRoles();
    const perms = {};
    Object.keys(PERMISSIONS).forEach(p => perms[p] = hasPermission(p));
    const info = { email: me, roles, permissions: perms, ts: new Date().toISOString() };
    _rbacLog_('RBAC_Debug', JSON.stringify(info));
    return info;
  };
})();