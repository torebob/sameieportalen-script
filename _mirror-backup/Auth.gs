/**
 * Auth.gs - Autorisasjonsmodul for Sameieportalen
 *
 * VIKTIG: Denne filen håndterer all autorisasjon.
 * Bruk getCurrentUser() i ALLE backend-funksjoner.
 */

// VIKTIG: ERSTATT "ERSTATT_MED_ID" MED ID-EN TIL GOOGLE SHEET-DATABASEN DIN
const DB_SHEET_ID = "ERSTATT_MED_ID";

/**
 * Henter innlogget bruker fra Google Session
 * @returns {object} Brukerobjekt med email, navn og rolle
 * @throws {Error} Hvis ikke autentisert
 */
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();

  if (!email) {
    throw new Error("Ikke autentisert. Vennligst logg inn.");
  }

  // Hent brukerinfo fra Users-sheet
  const userInfo = getUserInfo(email);

  if (!userInfo) {
    throw new Error("Bruker ikke funnet i systemet: " + email);
  }

  return {
    email: email,
    name: userInfo.name || email.split('@')[0],
    role: userInfo.role || 'beboer',
    apartmentId: userInfo.apartmentId || null
  };
}

/**
 * Krever autentisering og eventuelt spesifikk rolle
 * @param {Array<string>} allowedRoles - Liste over tillatte roller (valgfritt)
 * @returns {object} Brukerobjekt
 * @throws {Error} Hvis ikke autentisert eller ikke autorisert
 */
function requireAuth(allowedRoles = []) {
  const user = getCurrentUser();

  if (allowedRoles.length > 0 && !allowedRoles.includes(user.role)) {
    throw new Error(
      `Ikke autorisert. Krever rolle: ${allowedRoles.join(' eller ')}. ` +
      `Din rolle: ${user.role}`
    );
  }

  return user;
}

/**
 * Henter brukerinfo fra Users-sheet
 * @param {string} email - Brukerens e-post
 * @returns {object|null} Brukerinfo eller null hvis ikke funnet
 */
function getUserInfo(email) {
  try {
    const headers = ['email', 'name', 'role', 'apartmentId', 'phone'];
    const sheet = _getOrCreateSheet('Users', headers);
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) { // Ark er tomt eller har bare overskrifter
      return null;
    }

    const headerRow = data.shift();
    const emailIndex = headerRow.indexOf('email');

    if (emailIndex === -1) {
      console.error("Kolonnen 'email' ble ikke funnet i Users-arket.");
      return null;
    }

    const userRow = data.find(row => row[emailIndex] === email);

    if (!userRow) return null;

    const userInfo = {};
    headerRow.forEach((header, i) => {
      userInfo[header] = userRow[i];
    });

    return userInfo;
  } catch (e) {
    console.error("Error in getUserInfo: " + e.message);
    return null;
  }
}

/**
 * Sjekker om bruker har tilgang til en spesifikk leilighet
 * @param {string} apartmentId - Leilighets-ID
 * @param {object} user - Brukerobjekt (valgfritt, hentes hvis ikke oppgitt)
 * @returns {boolean} True hvis bruker har tilgang
 */
function hasAccessToApartment(apartmentId, user = null) {
  if (!user) {
    user = getCurrentUser();
  }

  // Admin og styremedlemmer har tilgang til alt
  if (['admin', 'board_member', 'board_leader'].includes(user.role)) {
    return true;
  }

  // Beboere har kun tilgang til egen leilighet
  return user.apartmentId === apartmentId;
}

/**
 * Hjelpefunksjon for å hente eller opprette et ark
 * @param {string} sheetName - Navnet på arket
 * @param {Array<string>} headers - Liste med overskrifter for nytt ark
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Ark-objektet
 */
function _getOrCreateSheet(sheetName, headers = []) {
  try {
    const ss = SpreadsheetApp.openById(DB_SHEET_ID);
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      if (headers.length > 0) {
        sheet.appendRow(headers);
      }
    }
    return sheet;
  } catch (e) {
    console.error(`Error accessing or creating sheet ${sheetName}: ${e.message}`);
    // Hvis vi ikke kan åpne regnearket, er det en kritisk feil.
    throw new Error(`Could not open spreadsheet with ID ${DB_SHEET_ID}. `+
      `Please ensure the ID is correct and you have access.`);
  }
}

/**
 * Logger revisjonshendelse (for GDPR-compliance)
 * @param {string} action - Handling (f.eks. 'CREATE_BOOKING')
 * @param {string} resource - Ressurs (f.eks. 'Bookings')
 * @param {object} details - Detaljer om hendelsen
 */
function logAuditEvent(action, resource, details = {}) {
  try {
    const user = getCurrentUser();
    const sheet = _getOrCreateSheet('AuditLog',
      ['timestamp', 'userEmail', 'action', 'resource', 'details']
    );

    sheet.appendRow([
      new Date().toISOString(),
      user.email,
      action,
      resource,
      JSON.stringify(details)
    ]);
  } catch (e) {
    // Ikke la logging-feil stoppe hovedoperasjonen
    console.error("Failed to log audit event: " + e.message);
  }
}