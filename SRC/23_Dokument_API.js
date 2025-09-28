/* ====================== Dokumentdeling - API ======================
 * FILE: 23_Dokument_API.js | VERSION: 1.0.0 | UPDATED: 2025-09-28
 * FORMÅL: Håndtere logikk for deling av dokumenter fra Vedlegg-arket.
 * ================================================================== */

/**
 * Henter en liste over delbare dokumenter fra 'Vedlegg'-arket.
 *
 * Antar at 'Vedlegg'-arket har følgende kolonner:
 * A: Dokumentnavn
 * B: Kategori
 * C: Drive-URL
 *
 * @returns {Array<Object>} En liste med dokumentobjekter, hver med 'navn', 'kategori', og 'url'.
 */
function getShareableDocuments() {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.VEDLEGG);
    if (!sheet) {
      // Oppretter arket hvis det ikke finnes, for å unngå feil ved førstegangsbruk.
      const newSheet = SpreadsheetApp.getActive().insertSheet(SHEETS.VEDLEGG);
      newSheet.getRange('A1:C1').setValues([['Dokumentnavn', 'Kategori', 'Drive-URL']]).setFontWeight('bold');
      newSheet.setFrozenRows(1);
      // Returnerer en tom liste siden arket nettopp ble opprettet.
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return []; // Tomt ark, kun header
    }

    const headers = data.shift().map(h => String(h).toLowerCase());
    const cNavn = headers.indexOf('dokumentnavn');
    const cKategori = headers.indexOf('kategori');
    const cUrl = headers.indexOf('drive-url');

    if (cNavn === -1 || cKategori === -1 || cUrl === -1) {
      throw new Error("Mangler forventede kolonner i 'Vedlegg'-arket: Dokumentnavn, Kategori, Drive-URL.");
    }

    return data.map(row => ({
      navn: row[cNavn],
      kategori: row[cKategori],
      url: row[cUrl]
    })).filter(doc => doc.navn && doc.url); // Filtrer ut tomme rader

  } catch (e) {
    safeLog('getShareableDocuments_Feil', e.message);
    throw new Error(`Kunne ikke hente dokumenter: ${e.message}`);
  }
}

/**
 * Deler et valgt dokument med alle beboere ved å sende et oppslag.
 *
 * @param {Object} document - Objektet som inneholder dokumentinformasjon.
 * @param {string} document.navn - Navnet på dokumentet.
 * @param {string} document.url - URL-en til dokumentet.
 * @returns {Object} Et resultatobjekt fra sendOppslag-funksjonen.
 */
function shareDocumentWithResidents(document) {
  try {
    requirePermission('SEND_OPPSLAG');

    if (!document || !document.navn || !document.url) {
      throw new Error("Ugyldig dokumentinformasjon oppgitt.");
    }

    const tittel = `Nytt dokument publisert: ${document.navn}`;
    const innhold = `
Hei,

Et nytt dokument har blitt delt og er nå tilgjengelig for deg:
"${document.navn}"

Du kan se dokumentet her:
${document.url}

Med vennlig hilsen,
Styret
    `.trim();

    const payload = {
      tittel: tittel,
      innhold: innhold,
      maalgruppe: 'Alle beboere'
    };

    // Gjenbruker den eksisterende funksjonen for å sende oppslag
    return sendOppslag(payload);

  } catch (e) {
    safeLog('shareDocument_Feil', e.message);
    throw e; // Kaster feilen videre til klienten
  }
}