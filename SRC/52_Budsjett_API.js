// =============================================================================
// Budsjett – API (les/skriv + oppsummering + roller)
// FILE: 52_Budsjett_API.gs
// VERSION: 1.1.1
// UPDATED: 2025-09-15
// REQUIRES: Ark 'BUDSJETT' (normalisert), 'TILGANG' (Email|Rolle)
// ROLES: LEDER/KASSERER = rediger; STYRE/LESER = lese
// DESIGN: Idempotent namespace (globalThis.BUDGET) for å unngå global-kollisjoner
// =============================================================================

// Namespace og konfig (idempotent)
(function (glob) {
  var S = glob.SHEETS || {};
  glob.BUDGET = Object.assign(glob.BUDGET || {}, {
    SHEET: S.BUDSJETT || 'BUDSJETT',
    ACCESS_SHEET: S.TILGANG || 'TILGANG',
    EDIT_ROLES: new Set(['LEDER', 'KASSERER']),
    VIEW_ROLES: new Set(['LEDER','KASSERER','STYRE','LESER']),
    VERSION: '1.1.1',
    UPDATED: '2025-09-15'
  });
})(globalThis);

// Skjema/metadata
function getBudgetSchema() {
  return {
    months: ['Jan','Feb','Mar','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Des'],
    required: ['year','account','month','amount'],
    optional: ['version','name','costCenter','project','vat','type','comment'],
    defaults: { version: 'main' }
  };
}

// Brukerprofil (for UI)
function getCurrentUserProfile() {
  var email = budgetGetUserEmail_();
  var role  = budgetGetRoleForEmail_(email);
  return { email: email, role: role, canEdit: globalThis.BUDGET.EDIT_ROLES.has(role), displayName: email };
}

// Hent budsjettlinjer (år/versjon)
function getBudget(year, version) {
  var sh = SpreadsheetApp.getActive().getSheetByName(globalThis.BUDGET.SHEET);
  if (!sh) return { ok:false, error:'Mangler ark: ' + globalThis.BUDGET.SHEET };
  var values = sh.getDataRange().getValues();
  if (!values.length) return { ok:true, rows:0, items:[] };

  var header = values.shift().map(String);
  var idx = function(name){ return header.indexOf(name); };
  if (idx('År') < 0) return { ok:false, error:'BUDSJETT-ark mangler forventet header.' };

  var ver = String(version || 'main');
  var items = [];
  for (var i=0;i<values.length;i++) {
    var r = values[i];
    if (String(r[idx('År')]) !== String(year)) continue;
    if (String(r[idx('Versjon')] || 'main') !== ver) continue;
    items.push({
      year: Number(r[idx('År')]),
      version: String(r[idx('Versjon')] || 'main'),
      account: String(r[idx('Konto')] || ''),
      name: String(r[idx('Navn')] || ''),
      costCenter: String(r[idx('Kostnadssted')] || ''),
      project: String(r[idx('Prosjekt')] || ''),
      vat: String(r[idx('MVA')] || ''),
      type: String(r[idx('Type')] || ''),
      month: String(r[idx('Måned')] || ''),
      amount: Number(r[idx('Beløp')] || 0),
      comment: String(r[idx('Kommentar')] || '')
    });
  }
  return { ok:true, rows: items.length, items: items };
}

// Append linjer
function saveBudgetLines(lines) {
  budgetEnsureCanEdit_();
  if (!Array.isArray(lines) || !lines.length) return { ok:false, error:'Tomt datasett' };

  var sh = SpreadsheetApp.getActive().getSheetByName(globalThis.BUDGET.SHEET) || budgetInitBudgetSheet_();
  var header = ['År','Versjon','Konto','Navn','Kostnadssted','Prosjekt','MVA','Type','Måned','Beløp','Kommentar'];
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');

  var out = lines.map(function(x){
    return [
      Number(x.year),
      String(x.version||'main'),
      String(x.account||''),
      String(x.name||''),
      String(x.costCenter||''),
      String(x.project||''),
      String(x.vat||''),
      String(x.type||''),
      String(x.month||''),
      Number(x.amount||0),
      String(x.comment||'')
    ];
  });

  sh.getRange(sh.getLastRow()+1, 1, out.length, header.length).setValues(out);
  return { ok:true, appended: out.length };
}

// Erstatt hele år+versjon
function replaceBudget(year, version, lines) {
  budgetEnsureCanEdit_();
  if (!Number.isInteger(year)) return { ok:false, error:'Ugyldig år' };
  var ver = String(version||'main');

  var sh = SpreadsheetApp.getActive().getSheetByName(globalThis.BUDGET.SHEET) || budgetInitBudgetSheet_();
  var values = sh.getDataRange().getValues();
  if (!values.length) return saveBudgetLines(lines);

  var header = values.shift().map(String);
  var idx = function(name){ return header.indexOf(name); };
  if (idx('År') < 0) return { ok:false, error:'BUDSJETT-ark mangler forventet header.' };

  var keep = [header];
  for (var i=0;i<values.length;i++) {
    var r = values[i];
    var y = String(r[idx('År')]);
    var v = String(r[idx('Versjon')]||'main');
    if (y === String(year) && v === ver) continue;
    keep.push(r);
  }

  sh.clear();
  sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
  if (keep.length > 1) sh.getRange(2,1,keep.length-1,header.length).setValues(keep.slice(1));

  return saveBudgetLines(lines);
}

// Sammendrag per konto (inkl. enkel MVA)
function calculateBudgetSummary(year, version) {
  var res = getBudget(year, version);
  if (!res.ok) return res;

  var items = res.items || [];
  var months = getBudgetSchema().months;
  var byAccount = {};

  for (var i=0;i<items.length;i++) {
    var it = items[i];
    var key = it.account || '(mangler konto)';
    if (!byAccount[key]) byAccount[key] = { account: key, name: it.name || '', totals: {}, sum: 0 };
    var base = Number(it.amount) || 0;
    var vatPct = String(it.vat||'').trim() === '25' ? 0.25 : 0;
    var withVat = base * (1 + vatPct);
    byAccount[key].totals[it.month] = (byAccount[key].totals[it.month] || 0) + withVat;
    byAccount[key].sum += withVat;
  }

  var rows = Object.values(byAccount).map(function(r){
    var obj = { account: r.account, name: r.name, sum: r.sum };
    for (var m=0;m<months.length;m++) obj[months[m]] = r.totals[months[m]] || 0;
    return obj;
  }).sort(function(a,b){ return String(a.account).localeCompare(String(b.account)); });

  return { ok:true, rows: rows };
}

// ----------------------------- Helpers (namespacet) --------------------------

function budgetGetUserEmail_() {
  return String(Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').trim();
}

function budgetGetRoleForEmail_(email) {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(globalThis.BUDGET.ACCESS_SHEET);
    if (!sh) return 'LESER';
    var values = sh.getDataRange().getValues(); values.shift();
    var row = values.find(function(r){
      return String(r[0]||'').trim().toLowerCase() === String(email||'').trim().toLowerCase();
    });
    var role = row ? String(row[1]||'LESER').toUpperCase().trim() : 'LESER';
    return globalThis.BUDGET.VIEW_ROLES.has(role) ? role : 'LESER';
  } catch (e) {
    return 'LESER';
  }
}

function budgetEnsureCanEdit_() {
  var role = budgetGetRoleForEmail_(budgetGetUserEmail_());
  if (!globalThis.BUDGET.EDIT_ROLES.has(role)) throw new Error('Tilgang nektet: Du har ikke redigeringstilgang.');
}

function budgetInitBudgetSheet_() {
  var sh = SpreadsheetApp.getActive().insertSheet(globalThis.BUDGET.SHEET);
  var header = ['År','Versjon','Konto','Navn','Kostnadssted','Prosjekt','MVA','Type','Måned','Beløp','Kommentar'];
  sh.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
  return sh;
}
