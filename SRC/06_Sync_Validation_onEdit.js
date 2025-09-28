/* ====================== Sync & Validation (onEdit) ======================
 * FILE: 06_Sync_Validation_onEdit.gs | VERSION: 1.7.1 | UPDATED: 2025-09-13
 * FORMÅL: Sanntidsvalidering, synk og autoutfyllinger med ytelsesoptimalisering.
 * Endringer v1.7.1: Korrekt håndtering av tomme datoer vs. ugyldig format.
 * ===================================================================== */

const VALIDATION_STYLES = {
  INVALID_ROW_BG: '#fde68a',
  VALID_ROW_BG: null
};

// Cache av kolonneindekser pr. ark (sheetId_key)
const COLUMN_CACHE = new Map();

function onEdit(e){
  try{
    if (!e || !e.range || !e.source) return;
    const sh = e.range.getSheet();
    const name = sh.getName();

    // Trim tekst (enkel anti-feil), men unngå å loope (neste onEdit har allerede trimmed verdi)
    if (typeof e.value === 'string') {
      const trimmed = e.value.trim();
      if (trimmed !== e.value) {
        e.range.setValue(trimmed);
        return; // la neste onEdit ta resten, så vi ikke dobbeltkjører
      }
    }

    switch(name){
      case SHEETS.EIERSKAP:
      case SHEETS.LEIE:
        _validateDateRangeRow_(sh, e.range);
        break;

      case SHEETS.TASKS:
        _ensureTaskIdOnRow_(sh, e.range);
        break;

      default:
        // no-op
        break;
    }
  } catch(err){
    Logger.log('onEdit error: ' + err.message);
  }
}

/** Henter 0-basert kolonneindeks for gitt header-navn (case-insensitive), med cache. */
function _getCachedColIndex_(sheet, key){
  if (!sheet || !key) return -1;
  const cacheKey = `${sheet.getSheetId()}__${String(key).toLowerCase()}`;
  if (COLUMN_CACHE.has(cacheKey)) return COLUMN_CACHE.get(cacheKey);

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn())
                      .getValues()[0]
                      .map(h => String(h||'').trim().toLowerCase());
  const idx = header.indexOf(String(key).toLowerCase());
  COLUMN_CACHE.set(cacheKey, idx);
  return idx;
}

/** Marker én rad som (u)gyldig og vedlikehold Kommentar-feltet om det finnes. */
function _markRow_(sheet, row, isInvalid, note){
  if (!sheet || !row) return;

  // Marker bakgrunn for hele raden
  const rng = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  rng.setBackground(isInvalid ? VALIDATION_STYLES.INVALID_ROW_BG : VALIDATION_STYLES.VALID_ROW_BG);

  // Oppdater "kommentar"-kolonne hvis den finnes (eksisterer i EIERSKAP/LEIE)
  const cK = _getCachedColIndex_(sheet, 'kommentar');
  if (cK > -1) {
    const cell = sheet.getRange(row, cK + 1);
    const existing = String(cell.getValue()||'').trim();
    // Fjern tidligere VALIDATION-del, behold brukers egen kommentar
    const base = existing.replace(/\s*\|\s*VALIDATION:.*$/i, '').trim();
    const text = note ? String(note) : '';
    if (isInvalid) {
      const msg = base ? `${base} | VALIDATION: ${text}` : `VALIDATION: ${text}`;
      cell.setValue(msg);
    } else {
      if (base) cell.setValue(base);
      else cell.clearContent();
    }
  }
}

/** Trygg parsing av valgfri dato: tom => null, ugyldig => {invalid:true}. */
function _parseOptionalDate_(val){
  if (val == null || val === '') return { date: null, invalid: false };
  if (val instanceof Date) return isNaN(val.getTime()) ? { date:null, invalid:true } : { date:val, invalid:false };

  const s = String(val).trim();
  if (!s) return { date:null, invalid:false };

  // dd.MM.yyyy
  const m1 = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m1){
    const d = new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));
    return isNaN(d.getTime()) ? { date:null, invalid:true } : { date:d, invalid:false };
  }
  // yyyy-MM-dd
  const m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m2){
    const d = new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));
    return isNaN(d.getTime()) ? { date:null, invalid:true } : { date:d, invalid:false };
  }
  // Fallback: Date-parser
  const d = new Date(s);
  return isNaN(d.getTime()) ? { date:null, invalid:true } : { date:d, invalid:false };
}

/** Validerer datoperioder (fra_dato <= til_dato). Tomme datoer er tillatt. */
function _validateDateRangeRow_(sheet, range){
  if (!sheet || !range) return;

  const row = range.getRow();
  if (row === 1) return; // hopp header

  const cFrom = _getCachedColIndex_(sheet, 'fra_dato');
  const cTo   = _getCachedColIndex_(sheet, 'til_dato');
  if (cFrom < 0 || cTo < 0) return;

  // Kjør kun når en av dato-kolonnene er redigert
  const editedCol0 = range.getColumn() - 1;
  if (editedCol0 !== cFrom && editedCol0 !== cTo) return;

  // Les verdier i raden (batch)
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const { date: dFrom, invalid: invFrom } = _parseOptionalDate_(rowData[cFrom]);
  const { date: dTo,   invalid: invTo   } = _parseOptionalDate_(rowData[cTo]);

  if (invFrom) return _markRow_(sheet, row, true, 'Ugyldig fra_dato format');
  if (invTo)   return _markRow_(sheet, row, true, 'Ugyldig til_dato format');

  // Begge kan være tomme, eller kun én av dem – det er OK
  if (dFrom && dTo && dTo < dFrom){
    _markRow_(sheet, row, true, 'til_dato kan ikke være før fra_dato');
  } else {
    _markRow_(sheet, row, false);
  }
}

/** Auto-utfyll OppgaveID/Status/Opprettet i Oppgaver-fanen i én batch. */
function _ensureTaskIdOnRow_(sheet, range){
  if (!sheet || !range) return;
  const row = range.getRow();
  if (row === 1) return; // header

  const cId   = _getCachedColIndex_(sheet, 'oppgaveid');
  const cTit  = _getCachedColIndex_(sheet, 'tittel');
  const cStat = _getCachedColIndex_(sheet, 'status');
  const cOpp  = _getCachedColIndex_(sheet, 'opprettet');

  if (cId < 0 || cTit < 0) return;

  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const titleV = rowData[cTit];
  const idV = rowData[cId];

  if (titleV && !String(idV || '').trim()){
    try {
      const id = (typeof _nextTaskId_ === 'function') ? _nextTaskId_() : `TASK-${Date.now()}`;
      rowData[cId] = id;

      if (cOpp > -1 && !rowData[cOpp]) rowData[cOpp] = new Date();
      if (cStat > -1 && !rowData[cStat]) rowData[cStat] = 'Ny';

      sheet.getRange(row, 1, 1, rowData.length).setValues([rowData]); // batch write
      if (typeof _logEvent === 'function') _logEvent('Oppgaver', `Auto-ID tildelt: ${id} (batch update)`);
    } catch(e){
      Logger.log(`Auto-ID feilet: ${e.message}`);
    }
  }
}
