/**
 * FIL: 22_WebApp_Server.gs
 * Leser fra spesifikt regneark via SPREADSHEET_ID.
 * Ark som forventes:
 *  - "Brukere":  Epost | Navn | Roller
 *  - "Meny":     Rolle | Meny | Funksjon | Sort (valgfri)
 *  - "Logg":     (opprettes automatisk)
 */

const DATA = Object.freeze({
  SPREADSHEET_ID: '1v91oJ7F5qiZHbp6trHnRiC5kKyWYXPklvIsVvC2lVao',
  SHEET_USERS: 'Brukere',
  SHEET_MENU:  'Meny',
  SHEET_LOG:   'Logg',
});

// ------------------------------------------------------------------
// Hoved-API for frontend
// ------------------------------------------------------------------
function uiBootstrap() {
  try {
    const email = _getUserEmail_();
    Logger.log('[uiBootstrap] session email: %s', email);

    const user  = _getUserRecord_(email);          // { name, email, roles[] }
    const menu  = _getMenuForRoles_(user.roles);   // [{ name, action }]
    logUserVisit_(user.email, user.name, user.roles);

    return { user, menu };
  } catch (err) {
    Logger.log('uiBootstrap error: ' + (err && err.message ? err.message : err));
    const guest = { name: 'Gjest', email: 'ukjent', roles: ['Gjest'] };
    return { user: guest, menu: _getMenuForRoles_(guest.roles) };
  }
}

// ------------------------------------------------------------------
// Datatilgang / hjelpere
// ------------------------------------------------------------------
function _ss_() {
  return SpreadsheetApp.openById(DATA.SPREADSHEET_ID);
}

/** Trygg uthenting av e-post. */
function _getUserEmail_() {
  try {
    const a = Session.getActiveUser() && Session.getActiveUser().getEmail();
    const b = Session.getEffectiveUser && Session.getEffectiveUser().getEmail();
    return String(a || b || '').trim();
  } catch (e) {
    return '';
  }
}

/** Leser "Brukere" (Epost | Navn | Roller) og returnerer {name,email,roles[]}. */
function _getUserRecord_(email) {
  const ss = _ss_();
  const sh = ss.getSheetByName(DATA.SHEET_USERS);
  if (!sh || sh.getLastRow() < 2) {
    Logger.log('[Brukere] ark mangler eller tomt.');
    return { name: email || 'Gjest', email: email || '', roles: ['Gjest'] };
  }

  // Normaliser for match
  const norm = s => String(s || '').trim().toLowerCase().replace(/\s+/g, '');
  const emailNorm = norm(email);

  const values = sh.getDataRange().getValues();
  const header = values.shift().map(h => String(h).trim());
  const idx = _indexByNames_(header, {
    Epost:  ['Epost','Email','E-mail','E post','E-post'],
    Navn:   ['Navn','Name'],
    Roller: ['Roller','Roles']
  });

  const row = values.find(r => norm(r[idx.Epost]) === emailNorm);
  if (!row) return { name: email || 'Gjest', email: email || '', roles: ['Gjest'] };

  const name  = row[idx.Navn] || email || 'Ukjent';
  const roles = _parseRoles_(row[idx.Roller]);
  return { name: String(name), email: String(email || ''), roles };
}

/**
 * Leser "Meny" (Rolle | Meny | Funksjon | Sort) og filtrerer på roller.
 * Implementerer arving:
 *  - Admin       -> + Styremedlem, Beboer, Vaktmester
 *  - Styremedlem -> + Beboer
 */
function _getMenuForRoles_(roles) {
  roles = Array.isArray(roles) ? roles.slice() : ['Gjest'];

  // Arving av roller (unik liste)
  if (roles.includes('Admin')) {
    roles = [...new Set([...roles, 'Styremedlem', 'Beboer', 'Vaktmester'])];
  }
  if (roles.includes('Styremedlem')) {
    roles = [...new Set([...roles, 'Beboer'])];
  }

  const ss = _ss_();
  const sh = ss.getSheetByName(DATA.SHEET_MENU);
  if (!sh || sh.getLastRow() < 2) {
    Logger.log('[Meny] ark mangler eller tomt.');
    return [];
  }

  const values = sh.getDataRange().getValues();
  const header = values.shift().map(String);
  const idx = _indexByNames_(header, {
    Rolle:    ['Rolle','Roller','Role'], // tåler "Roller"
    Meny:     ['Meny','Menu','Navn','Name'],
    Funksjon: ['Funksjon','Function','Action'],
    Sort:     ['Sort','Order','Idx']
  });

  let rows = values.filter(r => {
    const role = String(r[idx.Rolle] || '').trim().toLowerCase();
    return roles.some(userRole => role === String(userRole).trim().toLowerCase());
  });

  if (idx.Sort !== -1) {
    rows = rows.sort((a,b) => (Number(a[idx.Sort]||0)) - (Number(b[idx.Sort]||0)));
  }

  // Fjern duplikater basert på (name, action)
  const seen = new Set();
  const out = [];
  for (const r of rows) {
    const name   = String(r[idx.Meny]||'').trim();
    const action = String(r[idx.Funksjon]||'').trim();
    if (!name || !action) continue;
    const key = name + '|' + action;
    if (!seen.has(key)) { seen.add(key); out.push({ name, action }); }
  }

  return out;
}

/** Logger besøk til "Logg". */
function logUserVisit_(email, name, roles) {
  try {
    const ss = _ss_();
    const sh = ss.getSheetByName(DATA.SHEET_LOG) || ss.insertSheet(DATA.SHEET_LOG);
    sh.appendRow([new Date(), String(email||''), String(name||''), (roles||[]).join(', ')]);
  } catch (e) {
    // stille
  }
}

// ------------------------------------------------------------------
// Små hjelpere
// ------------------------------------------------------------------
function _parseRoles_(cell) {
  if (!cell) return ['Gjest'];
  if (Array.isArray(cell)) cell = cell.join(',');
  return String(cell).split(/[,\;]/).map(s => s.trim()).filter(Boolean);
}

/** Case-insensitive headeroppslag med støtte for flere navn. */
function _indexByNames_(header, wanted) {
  const lower = header.map(h => String(h).toLowerCase().trim());
  const out = {};
  Object.keys(wanted).forEach(key => {
    const candidates = wanted[key].map(x => String(x).toLowerCase().trim());
    let idx = -1;
    for (let i = 0; i < lower.length; i++) {
      if (candidates.indexOf(lower[i]) !== -1) { idx = i; break; }
    }
    out[key] = idx;
  });
  return out;
}

// ------------------------------------------------------------------
// Debug (kjør fra editor → se i Logger)
// ------------------------------------------------------------------
function debugBootstrap() {
  const result = uiBootstrap();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}
