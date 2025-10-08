// =============================================================================
// Budsjett – Tilgang/roller (TILGANG-ark)
// FILE: 53_Budsjett_Access.gs
// VERSION: 1.0.1
// UPDATED: 2025-09-15
// NOTE: Idempotent – bruker globalThis.BUDGET.* (ingen globale const-er)
// =============================================================================

// Namespace-konfig (flett inn uten redeklarasjon)
(function (glob) {
  var prev = glob.BUDGET || {};
  var S = glob.SHEETS || {};
  glob.BUDGET = Object.assign({}, prev, {
    ACCESS_SHEET: prev.ACCESS_SHEET || S.TILGANG || 'TILGANG',
    VALID_ROLES:  prev.VALID_ROLES  || ['LEDER','KASSERER','STYRE','LESER']
  });
})(globalThis);

// Interne hjelpere (navn-prefiks for å unngå kollisjoner)
function budgetAccessSheetName_() {
  return (globalThis.BUDGET && globalThis.BUDGET.ACCESS_SHEET) || 'TILGANG';
}
function budgetValidRoles_() {
  return (globalThis.BUDGET && globalThis.BUDGET.VALID_ROLES) || ['LEDER','KASSERER','STYRE','LESER'];
}

// Opprett TILGANG-ark hvis mangler
function ensureAccessSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheetName = budgetAccessSheetName_();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1,1,1,2).setValues([['Email','Rolle']]).setFontWeight('bold');
    sh.freezeRows(1);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,2).setValues([['Email','Rolle']]).setFontWeight('bold');
  }
  return { ok:true, sheet: sheetName };
}

// Hent alle roller
function getRoles() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(budgetAccessSheetName_());
  if (!sh || sh.getLastRow() < 2) return { ok:true, items:[] };
  var values = sh.getDataRange().getValues(); values.shift();
  var items = values.map(function(r){
    return { email: String(r[0]||'').trim(), role: String(r[1]||'').toUpperCase().trim() };
  }).filter(function(x){ return x.email; });
  return { ok:true, items: items };
}

// Sett/oppdater rolle for en e-post
function setRole(email, role) {
  var r = String(role||'').toUpperCase().trim();
  if (budgetValidRoles_().indexOf(r) < 0) return { ok:false, error: 'Ugyldig rolle: ' + role };

  var ss = SpreadsheetApp.getActive();
  var sheetName = budgetAccessSheetName_();
  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,2).setValues([['Email','Rolle']]).setFontWeight('bold');

  var values = sh.getDataRange().getValues();
  values.shift(); // header

  var target = String(email||'').trim().toLowerCase();
  var rowIndex = -1;
  for (var i=0;i<values.length;i++) {
    if (String(values[i][0]||'').trim().toLowerCase() === target) { rowIndex = i+2; break; }
  }

  if (rowIndex < 0) {
    sh.appendRow([email, r]);
  } else {
    sh.getRange(rowIndex, 1, 1, 2).setValues([[email, r]]);
  }
  return { ok:true, email: email, role: r };
}

// Fjern rolle for en e-post
function removeRole(email) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(budgetAccessSheetName_());
  if (!sh || sh.getLastRow() < 2) return { ok:true, removed:null };

  var values = sh.getDataRange().getValues(); values.shift();
  var target = String(email||'').trim().toLowerCase();

  for (var i=0;i<values.length;i++) {
    if (String(values[i][0]||'').trim().toLowerCase() === target) {
      sh.deleteRow(i+2);
      return { ok:true, removed: email };
    }
  }
  return { ok:true, removed: null };
}
