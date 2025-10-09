/**
 * ════════════════════════════════════════════════════════════════════════
 * HTML INCLUDE HELPER
 * FIL: 00_HTML_Include.gs
 * VERSJON: 1.0.0
 * DATO: 2025-10-09
 * 
 * FORMÅL:
 * - Inkludere HTML-filer i andre HTML-filer
 * - Brukes av styling-systemet (_tokens, _base, _utilities)
 * - Brukes for gjenbrukbare HTML-komponenter
 * 
 * BRUK:
 * I HTML-filer:
 *   <?!= include('_tokens') ?>
 *   <?!= include('_base') ?>
 *   <?!= include('_utilities') ?>
 * ════════════════════════════════════════════════════════════════════════
 */

/**
 * Inkluderer en HTML-fil i en annen HTML-fil
 * 
 * @param {string} filename - Navn på filen (uten .html)
 * @returns {string} HTML-innholdet
 * 
 * @example
 * // I en HTML-fil:
 * <?!= include('_tokens') ?>
 * <?!= include('header') ?>
 * <?!= include('footer') ?>
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger.log('FEIL ved include(' + filename + '): ' + error.message);
    
    // Returner en synlig feilmelding i HTML
    return '<!-- FEIL: Kunne ikke inkludere "' + filename + '.html" -->' +
           '<div style="background: #fee; border: 2px solid #c00; padding: 10px; margin: 10px;">' +
           '<strong>Include-feil:</strong> Filen "' + filename + '.html" ble ikke funnet.<br>' +
           'Sjekk at filen eksisterer i Apps Script-prosjektet.' +
           '</div>';
  }
}

/**
 * Inkluderer flere HTML-filer samtidig
 * Nyttig for å inkludere alle styling-filer på én gang
 * 
 * @param {string[]} filenames - Array med filnavn
 * @returns {string} Samlet HTML-innhold
 * 
 * @example
 * // I en HTML-fil:
 * <?!= includeMultiple(['_tokens', '_base', '_utilities']) ?>
 */
function includeMultiple(filenames) {
  var content = '';
  for (var i = 0; i < filenames.length; i++) {
    content += include(filenames[i]);
  }
  return content;
}

/**
 * Inkluderer styling-systemet (shortcut)
 * 
 * @returns {string} All styling (tokens + base + utilities)
 * 
 * @example
 * // I en HTML-fil:
 * <?!= includeStyling() ?>
 */
function includeStyling() {
  return includeMultiple(['_tokens', '_base', '_utilities']);
}
