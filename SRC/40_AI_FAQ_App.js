/* ====================== AI-basert FAQ App ======================
 * FILE: 40_AI_FAQ_App.js | VERSION: 1.0.0 | UPDATED: 2025-09-28
 * FORMÅL: Håndtere søk i FAQ-databasen.
 * ================================================================== */

/**
 * Loads FAQ data from the 'Konfig' spreadsheet.
 * Assumes 'Konfig' sheet has 'Nøkkel' (key/question) and 'Verdi' (value/answer) columns.
 * @returns {Array<Object>} An array of question/answer objects.
 */
function getFaqData() {
  try {
    const sheetName = SHEETS.KONFIG;
    const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sh) {
      Logger.log(`Sheet not found: ${sheetName}`);
      return [];
    }
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return []; // Not enough data

    const headers = data.shift().map(h => String(h || '').toLowerCase());
    let keyIndex = headers.indexOf('nøkkel');
    let valueIndex = headers.indexOf('verdi');

    // Fallback to other potential column names if primary are not found
    if (keyIndex === -1 || valueIndex === -1) {
      const ruleIndex = headers.indexOf('regel');
      const descIndex = headers.indexOf('beskrivelse');
      if (ruleIndex !== -1 && descIndex !== -1) {
        keyIndex = ruleIndex;
        valueIndex = descIndex;
      } else {
        Logger.log(`Required columns ('Nøkkel'/'Verdi' or 'Regel'/'Beskrivelse') not found in ${sheetName}`);
        return [];
      }
    }

    return data.map(row => ({
      question: row[keyIndex],
      answer: row[valueIndex]
    })).filter(item => item.question && item.answer); // Filter out empty rows

  } catch (e) {
    Logger.log(`Error reading from sheet '${SHEETS.KONFIG}': ${e.message}`);
    return [];
  }
}

// --- Start of testable search logic ---

const stopWords = new Set(['og', 'i', 'jeg', 'det', 'at', 'en', 'et', 'den', 'til', 'er', 'som', 'på', 'de', 'med', 'han', 'av', 'ikke', 'der', 'så', 'var', 'meg', 'seg', 'hun', 'men', 'ett', 'har', 'om', 'vi', 'min', 'mitt', 'ha', 'hadde', 'hun', 'inn', 'ut', 'opp', 'ned', 'kan', 'kunne', 'skal', 'skulle', 'vil', 'ville', 'må', 'måtte', 'bli', 'blir', 'ble', 'blev', 'vær', 'være', 'vært', 'fra', 'du', 'deg', 'dere', 'deres', 'hva', 'hvem', 'hvor', 'hvorfor', 'hvordan']);

function tokenize(text) {
  return text.toLowerCase().replace(/[^\wæøå\s]/g, '').split(/\s+/).filter(Boolean);
}

function removeStopWords(tokens) {
  return tokens.filter(token => !stopWords.has(token));
}

function calculateIDF(term, docs, idfCache = {}) {
  if (idfCache[term]) return idfCache[term];
  const docsWithTerm = docs.filter(doc => doc.tokens.includes(term)).length;
  const idf = Math.log((docs.length) / (1 + docsWithTerm));
  idfCache[term] = idf;
  return idf;
}

function calculateTF(term, docTokens) {
  const termCount = docTokens.filter(t => t === term).length;
  return termCount / docTokens.length;
}

function cosineSimilarity(vecA, vecB) {
    let dotProduct = 0.0;
    let normA = 0.0;
    let normB = 0.0;
    for (const key in vecA) {
        if (vecB.hasOwnProperty(key)) {
            dotProduct += vecA[key] * vecB[key];
        }
        normA += Math.pow(vecA[key], 2);
    }
    for (const key in vecB) {
        normB += Math.pow(vecB[key], 2);
    }
    if (normA === 0 || normB === 0) return 0;
    return dotProduct / (Math.sqrt(normA) * Math.sqrt(normB));
}

/**
 * Pre-processes a database of documents for searching.
 * @param {Array<Object>} db - Array of {question, answer} objects.
 * @returns {Array<Object>} The processed documents with tokens and TF-IDF vectors.
 */
function processFaqData(db) {
  const idfCache = {};
  const docsWithTokens = db.map(doc => {
    const questionTokens = removeStopWords(tokenize(doc.question));
    const answerTokens = removeStopWords(tokenize(doc.answer));
    const allTokens = [...new Set([...questionTokens, ...answerTokens])];
    return { ...doc, tokens: allTokens };
  });

  docsWithTokens.forEach(doc => {
    const tfidfVector = {};
    const allTokensInDoc = removeStopWords(tokenize(doc.question + " " + doc.answer));
    const uniqueTokens = [...new Set(doc.tokens)];
    uniqueTokens.forEach(token => {
      const tf = calculateTF(token, allTokensInDoc);
      const idf = calculateIDF(token, docsWithTokens, idfCache);
      tfidfVector[token] = tf * idf;
    });
    doc.vector = tfidfVector;
  });

  return docsWithTokens;
}

/**
 * Performs the search logic on pre-processed documents.
 * @param {string} query - The user's search query.
 * @param {Array<Object>} processedDocs - The documents, processed by processFaqData.
 * @returns {Array<Object>} Sorted list of matching documents.
 */
function searchFaqLogic(query, processedDocs) {
  if (!query || typeof query !== 'string' || !processedDocs) {
    return [];
  }

  const queryTokens = removeStopWords(tokenize(query));
  if (queryTokens.length === 0) return [];

  const idfCache = {}; // Use a fresh cache for the query's context
  const queryVector = {};
  const uniqueQueryTokens = [...new Set(queryTokens)];
  uniqueQueryTokens.forEach(token => {
    const tf = calculateTF(token, queryTokens);
    const idf = calculateIDF(token, processedDocs, idfCache);
    queryVector[token] = tf * idf;
  });

  const rankedDocs = processedDocs.map(doc => ({
    ...doc,
    score: cosineSimilarity(queryVector, doc.vector)
  }));

  return rankedDocs.filter(doc => doc.score > 0.01).sort((a, b) => b.score - a.score);
}

// --- End of testable search logic ---

// Global state for the live app, initialized once.
let _processedFaqData;

function getProcessedFaqData_() {
    if (!_processedFaqData) {
        const faqDatabase = getFaqData();
        _processedFaqData = processFaqData(faqDatabase);
    }
    return _processedFaqData;
}

/**
 * The main search function called by the UI.
 * @param {string} query - The user's question.
 * @returns {Array<Object>} A list of relevant Q&A pairs, sorted by relevance.
 */
function searchFAQ(query) {
  const processedDocs = getProcessedFaqData_();
  const results = searchFaqLogic(query, processedDocs);

  Utilities.sleep(500); // Simulate network delay

  // Return the original document object without the extra properties
  return results.map(({question, answer}) => ({question, answer}));
}

/**
 * Funksjon for å rendre FAQ-siden.
 * @returns {HtmlOutput} HTML-siden for FAQ.
 */
function handleFaqRequest() {
  return HtmlService.createTemplateFromFile('SRC/40_AI_FAQ')
    .evaluate()
    .setTitle('AI-basert FAQ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}