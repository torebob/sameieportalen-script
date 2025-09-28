/* ==================================================================
 *  SEARCH API
 * ================================================================== */

/**
 * Searches for content across specified sheets and returns a structured response.
 * This function is exposed to the client-side UI.
 *
 * @param {string} query The search query from the user.
 * @returns {object} An object with `ok` status, and either `results` or `message`.
 *                   Example: { ok: true, results: [...] } or { ok: false, message: "Error details" }
 */
function searchContent(query) {
  if (!query || typeof query !== 'string' || query.trim().length < 3) {
    return { ok: false, message: 'Søketeksten må være minst 3 tegn.', results: [] };
  }

  // Use a global helper if available, otherwise fall back to a basic logger.
  const _log = typeof safeLog === 'function' ? safeLog : (topic, msg) => console.log(`${topic}: ${msg}`);

  try {
    // Ensure SHEETS global is available.
    if (typeof SHEETS === 'undefined') {
      throw new Error("Kritisk feil: 'SHEETS' er ikke definert.");
    }

    const normalizedQuery = query.toLowerCase();
    let results = [];

    // Centralized configuration for all searchable content.
    // This makes it easy to add new searchable sheets/columns in the future.
    const searchConfig = [
      {
        sheetName: SHEETS.MOTER,
        columns: ['tittel', 'agenda', 'sted'],
        type: 'Møte',
        idColumn: 'id',
        titleColumn: 'tittel'
      },
      {
        sheetName: SHEETS.MOTE_SAKER,
        columns: ['tittel', 'forslag', 'vedtak'],
        type: 'Møtesak',
        idColumn: 'sak_id',
        titleColumn: 'tittel'
      },
      {
        sheetName: SHEETS.PROTOKOLL_GODKJENNING,
        columns: ['Møte-ID', 'Kommentar'],
        type: 'Protokoll',
        idColumn: 'Godkjenning-ID',
        titleColumn: 'Møte-ID',
        linkColumn: 'Protokoll-URL'
      }
    ];

    searchConfig.forEach(config => {
      const sheetResults = searchInSheet_(config, normalizedQuery);
      if (sheetResults.length > 0) {
        results = results.concat(sheetResults);
      }
    });

    _log('SearchAPI', `Søk på "${query}" ga ${results.length} treff.`);
    return { ok: true, results: results };

  } catch (e) {
    _log('SearchAPI_Error', `Søk feilet: ${e.message}`);
    return { ok: false, message: `Søket feilet: ${e.message}`, results: [] };
  }
}

/**
 * Searches for a query within a single Google Sheet based on a configuration object.
 * This is a private helper function and not intended to be called directly from the client.
 *
 * @param {object} config The configuration for the sheet to search.
 * @param {string} normalizedQuery The lower-cased search query.
 * @returns {Array<object>} An array of result objects found in the sheet.
 * @private
 */
function searchInSheet_(config, normalizedQuery) {
  const { sheetName, columns, type, idColumn, titleColumn, linkColumn } = config;
  const results = [];

  let sheet;
  try {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
      return []; // Sheet is missing, empty, or only has headers.
    }
  } catch (e) {
    return []; // Could not access the sheet, return empty results.
  }

  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Get and remove header row.

  // Map column names to their indices for efficient lookup.
  const colIndices = columns.map(colName => headers.indexOf(colName)).filter(index => index !== -1);
  const idColIndex = headers.indexOf(idColumn);
  const titleColIndex = headers.indexOf(titleColumn);
  const linkColIndex = linkColumn ? headers.indexOf(linkColumn) : -1;

  // If essential columns are missing, we cannot process this sheet.
  if (idColIndex === -1 || titleColIndex === -1 || colIndices.length === 0) {
    return [];
  }

  data.forEach((row, rowIndex) => {
    // Check if any of the specified columns in the current row contain the query.
    const isMatch = colIndices.some(colIndex => {
      const cellValue = row[colIndex];
      return cellValue && typeof cellValue === 'string' && cellValue.toLowerCase().includes(normalizedQuery);
    });

    if (isMatch) {
      const id = row[idColIndex];
      let title = row[titleColIndex];

      // Custom title for protocols to make them more descriptive.
      if (type === 'Protokoll') {
        title = `Protokoll for "${title}"`;
      }

      results.push({
        id: id,
        title: title || `Uten tittel (${id})`,
        type: type,
        // Provide a reference for context, e.g., for debugging or linking.
        reference: `${sheetName} (Rad ${rowIndex + 2})`,
        link: (linkColIndex !== -1 && row[linkColIndex]) ? row[linkColIndex] : null
      });
    }
  });

  return results;
}

// Expose the main search function to be callable from the client-side UI.
globalThis.searchContent = searchContent;