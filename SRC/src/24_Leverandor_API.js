/* ================== Leverandør API (Backend) ==================
 * FILE: 24_Leverandor_API.js | VERSION: 1.0.0 | UPDATED: 2025-09-28
 * FORMÅL: Backend for administrasjon av leverandørprofiler.
 * SIKKERHET: Tilgang styres via RBAC (hasPermission).
 * ============================================================== */

(function () {
  const SHEET_NAME = 'Leverandører';
  const HDR = ['ID', 'Navn', 'Kontaktperson', 'Telefon', 'E-post', 'Adresse'];

  function _getSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, HDR.length).setValues([HDR]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  function _getAccess() {
    if (typeof hasPermission === 'function' && !hasPermission('MANAGE_SUPPLIERS')) {
      throw new Error('Tilgang nektet. Krever MANAGE_SUPPLIERS rettighet.');
    }
  }

  function _headersMap(headers) {
    const map = {};
    headers.forEach((h, i) => map[h] = i);
    return map;
  }

  function _generateId() {
    return 'LEV-' + new Date().getTime().toString(36) + Math.random().toString(36).substr(2, 5).toUpperCase();
  }

  function getSuppliers() {
    _getAccess();
    try {
      const sheet = _getSheet();
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return { ok: true, suppliers: [] };

      const headers = data.shift();
      const c = _headersMap(headers);

      const suppliers = data.map(row => ({
        id: row[c.ID],
        navn: row[c.Navn],
        kontaktperson: row[c.Kontaktperson],
        telefon: row[c.Telefon],
        epost: row[c['E-post']],
        adresse: row[c.Adresse]
      }));

      return { ok: true, suppliers: suppliers };
    } catch (e) {
      return { ok: false, error: 'Kunne ikke hente leverandører: ' + e.message };
    }
  }

  function getSupplier(id) {
    _getAccess();
    if (!id) return { ok: false, error: 'Mangler ID' };
    try {
      const sheet = _getSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data.shift();
      const c = _headersMap(headers);
      const idCol = c.ID;

      const row = data.find(r => r[idCol] === id);
      if (!row) return { ok: false, error: 'Leverandør ikke funnet' };

      const supplier = {
        id: row[c.ID],
        navn: row[c.Navn],
        kontaktperson: row[c.Kontaktperson],
        telefon: row[c.Telefon],
        epost: row[c['E-post']],
        adresse: row[c.Adresse]
      };
      return { ok: true, supplier: supplier };
    } catch (e) {
      return { ok: false, error: 'Feil ved henting av leverandør: ' + e.message };
    }
  }

  function saveSupplier(supplierData) {
    _getAccess();
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      if (!supplierData || !supplierData.navn) {
        return { ok: false, error: 'Navn er påkrevd.' };
      }

      const sheet = _getSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data.shift();
      const c = _headersMap(headers);

      let rowIndex = -1;
      if (supplierData.id) {
        rowIndex = data.findIndex(r => r[c.ID] === supplierData.id) + 2;
      }

      const rowData = [
        supplierData.id || _generateId(),
        supplierData.navn,
        supplierData.kontaktperson || '',
        supplierData.telefon || '',
        supplierData.epost || '',
        supplierData.adresse || ''
      ];

      if (rowIndex > 1) { // Update existing
        sheet.getRange(rowIndex, 1, 1, HDR.length).setValues([rowData]);
      } else { // Create new
        sheet.appendRow(rowData);
      }

      return { ok: true, id: rowData[0] };
    } catch (e) {
      return { ok: false, error: 'Kunne ikke lagre leverandør: ' + e.message };
    } finally {
      lock.releaseLock();
    }
  }

  function deleteSupplier(id) {
    _getAccess();
    if (!id) return { ok: false, error: 'Mangler ID' };

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sheet = _getSheet();
      const data = sheet.getDataRange().getValues();
      const headers = data.shift();
      const c = _headersMap(headers);
      const idCol = c.ID;

      const rowIndex = data.findIndex(r => r[idCol] === id);
      if (rowIndex === -1) {
        return { ok: false, error: 'Leverandør ikke funnet' };
      }

      sheet.deleteRow(rowIndex + 2); // +2 because of header and 0-based index

      return { ok: true };
    } catch (e) {
      return { ok: false, error: 'Kunne ikke slette leverandør: ' + e.message };
    } finally {
      lock.releaseLock();
    }
  }

  globalThis.getSuppliers = getSuppliers;
  globalThis.getSupplier = getSupplier;
  globalThis.saveSupplier = saveSupplier;
  globalThis.deleteSupplier = deleteSupplier;

})();