// =============================================================================
// Økonomisk Oversikt – API (leser regnskap, fordringer, budsjett)
// FILE: 53_Okonomi_API.gs
// VERSION: 1.0.0
// UPDATED: 2025-09-28
// REQUIRES:
//   - Ark 'REGNSKAP' (Dato, Beskrivelse, Konto, Beløp)
//   - Ark 'FORDRINGER' (Beboer, Leilighet, Beløp, Forfallsdato, Status)
//   - Ark 'BUDSJETT' (fra 52_Budsjett_API.gs)
//   - Ark 'TILGANG' (Email|Rolle)
// ROLES: LEDER/KASSERER = full tilgang; STYRE/LESER = lesetilgang
// =============================================================================

// Namespace og konfig (idempotent)
(function (glob) {
  var S = glob.SHEETS || {};
  glob.FINANCE = Object.assign(glob.FINANCE || {}, {
    ACTUALS_SHEET: S.REGNSKAP || 'REGNSKAP',
    RECEIVABLES_SHEET: S.FORDRINGER || 'FORDRINGER',
    ACCESS_SHEET: S.TILGANG || 'TILGANG',
    EDIT_ROLES: new Set(['LEDER', 'KASSERER']),
    VIEW_ROLES: new Set(['LEDER','KASSERER','STYRE']),
    VERSION: '1.0.0',
    UPDATED: '2025-09-28'
  });
})(globalThis);

// ----------------------------- Public API Functions --------------------------

/**
 * Fetches actual transactions from the 'REGNSKAP' sheet for a given year.
 * Assumes columns: Dato, Beskrivelse, Konto, Beløp
 * @param {number} year The year to fetch data for.
 * @returns {object} A response object with the filtered data.
 */
function getActuals(year) {
  financeEnsureCanView_();
  try {
    if (!year || !Number.isInteger(year)) {
      return { ok: false, error: "Et gyldig årstall må oppgis." };
    }
    const sheet = SpreadsheetApp.getActive().getSheetByName(globalThis.FINANCE.ACTUALS_SHEET);
    if (!sheet) return { ok: false, error: `Mangler ark: ${globalThis.FINANCE.ACTUALS_SHEET}` };

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { ok: true, data: [] }; // No data is not an error

    const headers = values.shift();
    const dateIdx = headers.indexOf('Dato');
    const accountIdx = headers.indexOf('Konto');
    const amountIdx = headers.indexOf('Beløp');

    if (dateIdx === -1 || accountIdx === -1 || amountIdx === -1) {
      return { ok: false, error: "REGNSKAP-arket mangler påkrevde kolonner: Dato, Konto, Beløp." };
    }

    const data = values.map(row => {
      const transactionDate = new Date(row[dateIdx]);
      if (transactionDate.getFullYear() === year) {
        const item = {};
        headers.forEach((header, i) => {
          item[header] = row[i];
        });
        return item;
      }
      return null;
    }).filter(Boolean); // Filter out nulls (rows from other years)

    return { ok: true, data: data };
  } catch (e) {
    return { ok: false, error: `En feil oppstod ved henting av regnskapsdata: ${e.message}` };
  }
}

/**
 * Fetches outstanding receivables from the 'FORDRINGER' sheet.
 * Assumes columns: Beboer, Leilighet, Beløp, Forfallsdato, Status
 * @returns {object} A response object with the list of receivables.
 */
function getReceivables() {
  financeEnsureCanView_();
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(globalThis.FINANCE.RECEIVABLES_SHEET);
    if (!sheet) return { ok: false, error: `Mangler ark: ${globalThis.FINANCE.RECEIVABLES_SHEET}` };

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { ok: true, data: [] }; // No data is valid

    const headers = values.shift();
    const data = values.map(row => {
      const receivable = {};
      headers.forEach((header, i) => {
        receivable[header] = row[i];
      });
      return receivable;
    });

    return { ok: true, data: data };
  } catch (e) {
    return { ok: false, error: `En feil oppstod ved henting av fordringer: ${e.message}` };
  }
}

/**
 * Calculates a simplified financial summary (income vs. expenses) for a given year.
 * Income is defined as transactions with account numbers starting with '3'.
 * @param {number} year The year to calculate the summary for.
 * @returns {object} A response object with the summary data.
 */
function getFinancialSummary(year) {
  financeEnsureCanView_();
  const actualsResponse = getActuals(year);
  if (!actualsResponse.ok) return actualsResponse;

  const transactions = actualsResponse.data;
  let totalIncome = 0;
  let totalExpenses = 0;

  const accountHeader = 'Konto';
  const amountHeader = 'Beløp';

  transactions.forEach(t => {
    const account = String(t[accountHeader] || '');
    const amount = Number(t[amountHeader] || 0);

    if (account.startsWith('3')) {
      totalIncome += amount;
    } else {
      totalExpenses += amount;
    }
  });

  return {
    ok: true,
    data: {
      year: year,
      totalIncome: totalIncome,
      totalExpenses: totalExpenses,
      netResult: totalIncome - totalExpenses
    }
  };
}


// ----------------------------- Helpers (namespacet) --------------------------

/**
 * Gets the current user's email address.
 * @private
 * @returns {string} The user's email.
 */
function financeGetUserEmail_() {
  return String(Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').trim();
}

/**
 * Gets the role for a given email from the access sheet.
 * @private
 * @param {string} email The email to look up.
 * @returns {string} The user's role (e.g., 'LESER', 'STYRE'). Defaults to 'LESER'.
 */
function financeGetRoleForEmail_(email) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(globalThis.FINANCE.ACCESS_SHEET);
    if (!sh) return 'LESER';
    const values = sh.getDataRange().getValues();
    values.shift(); // Remove header
    const row = values.find(function(r){
      return String(r[0] || '').trim().toLowerCase() === String(email || '').trim().toLowerCase();
    });
    const role = row ? String(row[1] || 'LESER').toUpperCase().trim() : 'LESER';
    // Ensure the role is valid before returning
    return globalThis.FINANCE.VIEW_ROLES.has(role) ? role : 'LESER';
  } catch (e) {
    // Log the error for debugging, but return a safe default.
    console.error("Error in financeGetRoleForEmail_: " + e.message);
    return 'LESER';
  }
}

/**
 * Throws an error if the current user does not have view permissions.
 * @private
 */
function financeEnsureCanView_() {
  const email = financeGetUserEmail_();
  const role = financeGetRoleForEmail_(email);
  if (!globalThis.FINANCE.VIEW_ROLES.has(role)) {
    throw new Error('Tilgang nektet: Du har ikke tilgang til å se økonomiske data.');
  }
}

/**
 * Throws an error if the current user does not have edit permissions.
 * @private
 */
function financeEnsureCanEdit_() {
  const email = financeGetUserEmail_();
  const role = financeGetRoleForEmail_(email);
  if (!globalThis.FINANCE.EDIT_ROLES.has(role)) {
    throw new Error('Tilgang nektet: Du har ikke redigeringstilgang.');
  }
}