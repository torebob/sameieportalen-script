/* ====================== Min Side API ======================
 * FILE: 26_MinSide_API.js | VERSION: 1.0.0 | UPDATED: 2025-09-28
 * FORMÅL: Håndterer servertjenester for "Min Side"-funksjonaliteten.
 *
 * FUNKSJONER:
 *  - getUserInfo(): Henter innlogget brukers kontaktinformasjon.
 *  - updateUserInfo(updates): Mottar endringsforespørsel fra bruker
 *    og logger den for styrets godkjenning.
 * ================================================================== */

/**
 * Henter kontaktinformasjonen for den innloggede brukeren.
 * @returns {object} Et objekt med status og data (navn, e-post, telefon) eller en feilmelding.
 */
function getUserInfo() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { ok: false, message: 'Kunne ikke identifisere brukeren.' };
    }

    const personerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PERSONER);
    if (!personerSheet) {
      return { ok: false, message: `Finner ikke arket: ${SHEETS.PERSONER}` };
    }

    const data = personerSheet.getDataRange().getValues();
    const headers = data.shift();
    const emailCol = headers.indexOf('E-post');
    const nameCol = headers.indexOf('Navn');
    const phoneCol = headers.indexOf('Telefon');

    if ([emailCol, nameCol, phoneCol].includes(-1)) {
      return { ok: false, message: 'Nødvendige kolonner (Navn, E-post, Telefon) finnes ikke i Personer-arket.' };
    }

    for (const row of data) {
      if (row[emailCol] && row[emailCol].toString().toLowerCase() === userEmail.toLowerCase()) {
        return {
          ok: true,
          data: {
            name: row[nameCol],
            email: row[emailCol],
            phone: row[phoneCol],
          },
        };
      }
    }

    return { ok: false, message: 'Brukeren ble ikke funnet i registeret.' };

  } catch (e) {
    Logger.log(`Feil i getUserInfo: ${e.message}`);
    return { ok: false, message: 'En teknisk feil oppstod ved henting av brukerdata.' };
  }
}

/**
 * Mottar og logger en forespørsel om å oppdatere brukerinformasjon.
 * @param {object} updates Et objekt som inneholder feltene som skal oppdateres (f.eks. { email: '...', phone: '...' }).
 * @returns {object} Et objekt som indikerer om operasjonen var vellykket.
 */
function updateUserInfo(updates) {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { ok: false, message: 'Kunne ikke identifisere brukeren for oppdatering.' };
    }

    // Enkel validering
    if (!updates || (typeof updates.email === 'undefined' && typeof updates.phone === 'undefined')) {
      return { ok: false, message: 'Ingen gyldige endringer ble sendt med.' };
    }

    const newEmail = updates.email;
    const newPhone = updates.phone;

    // TODO: Implementer en mer robust varslingsmekanisme.
    // For nå, logger vi forespørselen til Hendelsesloggen.
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOGG);
    if (logSheet) {
      const timestamp = new Date();
      const logMessage = `Bruker ${userEmail} ba om å oppdatere kontaktinfo. Ny e-post: "${newEmail}", Ny telefon: "${newPhone}".`;
      logSheet.appendRow([timestamp, 'BRUKER-OPPDATERING', userEmail, logMessage]);

      // Send e-post til styret/admin
      const appConfig = getAppConfig();
      const adminEmail = appConfig.ADMIN_EMAIL;
      const subject = `Endringsforespørsel for kontaktinfo: ${userEmail}`;
      const body = `Brukeren ${userEmail} har bedt om å oppdatere sin kontaktinformasjon:\n\n` +
                   `Ny e-post: ${newEmail}\n` +
                   `Ny telefon: ${newPhone}\n\n` +
                   `En administrator må godkjenne denne endringen i systemet.`;

      if (adminEmail && adminEmail !== 'styret@example.com') {
        MailApp.sendEmail(adminEmail, subject, body);
      } else {
        Logger.log(`Admin-epost er ikke konfigurert. Kan ikke sende varsel for endringsforespørsel fra ${userEmail}.`);
      }

    } else {
      Logger.log(`VIKTIG: Kunne ikke finne logg-arket (${SHEETS.LOGG}). Endringsforespørsel for ${userEmail} er ikke logget sentralt.`);
    }

    Logger.log(`Endringsforespørsel fra ${userEmail}: E-post -> ${newEmail}, Telefon -> ${newPhone}`);

    return { ok: true, message: 'Forespørsel om endring er mottatt og vil bli behandlet av styret.' };

  } catch (e) {
    Logger.log(`Feil i updateUserInfo: ${e.message}`);
    return { ok: false, message: 'En teknisk feil oppstod under lagring av endringsforespørselen.' };
  }
}

/**
 * Henter en oversikt over felleskostnader for brukeren.
 * MERK: Returnerer mock-data i denne versjonen.
 * @returns {object} Et objekt med status og mock-data for felleskostnader.
 */
function getFinancials() {
  try {
    // I en reell applikasjon ville denne funksjonen integrert med et økonomisystem
    // via API for å hente sanntidsdata for den innloggede brukeren.
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { ok: false, message: 'Kunne ikke identifisere brukeren.' };
    }

    // Mock-data for demonstrasjon
    const mockData = {
      monthly_cost: '4.250,- NOK',
      next_due_date: '01.10.2025',
      status: 'Betalt',
      outstanding_amount: '0,- NOK'
    };

    return { ok: true, data: mockData };

  } catch (e) {
    Logger.log(`Feil i getFinancials: ${e.message}`);
    return { ok: false, message: 'En teknisk feil oppstod ved henting av økonomisk data.' };
  }
}

/**
 * Henter eierskapshistorikken for den innloggede brukerens seksjon.
 * @returns {object} Et objekt med status og en liste over eierskapshistorikk.
 */
function getOwnershipHistory() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { ok: false, message: 'Kunne ikke identifisere brukeren.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const personerSheet = ss.getSheetByName(SHEETS.PERSONER);
    const eierskapSheet = ss.getSheetByName(SHEETS.EIERSKAP);

    if (!personerSheet || !eierskapSheet) {
      return { ok: false, message: 'Nødvendige data-ark (Personer, Eierskap) ble ikke funnet.' };
    }

    // Bruk hjelpefunksjon for å få tak i sheet-data og headers
    const personerData = personerSheet.getDataRange().getValues();
    const personerHeaders = personerData.shift();
    const eierskapData = eierskapSheet.getDataRange().getValues();
    const eierskapHeaders = eierskapData.shift();

    const personIdCol = personerHeaders.indexOf('Person-ID');
    const personNavnCol = personerHeaders.indexOf('Navn');
    const personEmailCol = personerHeaders.indexOf('E-post');

    const eierPersonIdCol = eierskapHeaders.indexOf('Person-ID');
    const seksjonIdCol = eierskapHeaders.indexOf('Seksjons-ID');
    const startDatoCol = eierskapHeaders.indexOf('Start-dato');
    const sluttDatoCol = eierskapHeaders.indexOf('Slutt-dato');

    // Finn brukerens Person-ID
    const userPersonRow = personerData.find(row => row[personEmailCol].toLowerCase() === userEmail.toLowerCase());
    if (!userPersonRow) {
      return { ok: false, message: 'Brukeren ble ikke funnet i personregisteret.' };
    }
    const userPersonId = userPersonRow[personIdCol];

    // Finn brukerens nåværende seksjon
    const currentEierskap = eierskapData.find(row => row[eierPersonIdCol] == userPersonId && !row[sluttDatoCol]);
    if (!currentEierskap) {
      return { ok: false, message: 'Fant ikke nåværende eierskap for brukeren.' };
    }
    const userSeksjonId = currentEierskap[seksjonIdCol];

    // Finn all historikk for den seksjonen
    const historyForSeksjon = eierskapData.filter(row => row[seksjonIdCol] == userSeksjonId);

    // Map Person-ID til Navn for raskt oppslag
    const personIdToNameMap = new Map(personerData.map(row => [row[personIdCol], row[personNavnCol]]));

    const history = historyForSeksjon.map(row => {
      const ownerName = personIdToNameMap.get(row[eierPersonIdCol]) || 'Ukjent Eier';
      return {
        from_date: row[startDatoCol] ? new Date(row[startDatoCol]).toLocaleDateString('nb-NO') : 'Ukjent dato',
        to_date: row[sluttDatoCol] ? new Date(row[sluttDatoCol]).toLocaleDateString('nb-NO') : 'Nåværende',
        owner_name: ownerName
      };
    }).sort((a, b) => new Date(b.from_date.split('.').reverse().join('-')) - new Date(a.from_date.split('.').reverse().join('-'))); // Sorter med nyeste først

    return { ok: true, data: history };

  } catch (e) {
    Logger.log(`Feil i getOwnershipHistory: ${e.message}`);
    return { ok: false, message: 'En teknisk feil oppstod ved henting av eierskapshistorikk.' };
  }
}

/**
 * Henter en liste over seksjonsspesifikke dokumenter.
 * MERK: Returnerer mock-data i denne versjonen.
 * @returns {object} Et objekt med status og en mock-liste over dokumenter.
 */
function getSectionDocuments() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { ok: false, message: 'Kunne ikke identifisere brukeren.' };
    }

    // I en reell applikasjon ville denne funksjonen slått opp i et dokumentarkiv
    // eller en Drive-mappe-struktur basert på brukerens seksjons-ID.
    const mockData = [
      { name: 'Salgsoppgave 2022', url: '#' },
      { name: 'Takstrapport 2022', url: '#' },
      { name: 'Vedlikeholdslogg for bad', url: '#' }
    ];

    return { ok: true, data: mockData };

  } catch (e) {
    Logger.log(`Feil i getSectionDocuments: ${e.message}`);
    return { ok: false, message: 'En teknisk feil oppstod ved henting av dokumenter.' };
  }
}