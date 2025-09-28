/* ====================== Tasks API (Backend) ======================
 * FILE: 19_Tasks_API.js | VERSION: 1.0.0 | UPDATED: 2025-09-28
 *
 * FORMÅL:
 * - Tilbyr et API for å hente, filtrere og sortere oppgaver for administratorer.
 * - Støtter filtrering på status og ansvarlig, og sortering på frist.
 *
 * API:
 *  - getTasks(options): Henter en liste over oppgaver basert på filter/sortering.
 *  - getTaskFilterData(): Henter unike statuser og ansvarlige for filter-dropdowns.
 * ====================================================================== */

(() => {
  const SH_NAME = globalThis.SHEETS?.TASKS || 'Oppgaver';
  const TZ = Session.getScriptTimeZone() || 'Europe/Oslo';

  // Hjelpefunksjon for å hente overskrifter og mappe dem til kolonneindekser
  const _getHeaders = (sh) => {
    if (sh.getLastRow() < 1) return {};
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    return headers.reduce((acc, header, i) => {
      if (header) acc[header.trim()] = i;
      return acc;
    }, {});
  };

  /**
   * Henter data for å populere filter-dropdowns i UI.
   * @returns {object} Et objekt med lister over unike statuser og ansvarlige.
   */
  function getTaskFilterData() {
    try {
      const sh = SpreadsheetApp.getActive().getSheetByName(SH_NAME);
      if (!sh || sh.getLastRow() < 2) {
        return { ok: true, statuses: [], assignees: [] };
      }

      const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      const headers = _getHeaders(sh);
      const statusIdx = headers['Status'];
      const assigneeIdx = headers['Ansvarlig'];

      const statuses = new Set();
      const assignees = new Set();

      data.forEach(row => {
        if (statusIdx !== undefined && row[statusIdx]) {
          statuses.add(row[statusIdx].toString().trim());
        }
        if (assigneeIdx !== undefined && row[assigneeIdx]) {
          assignees.add(row[assigneeIdx].toString().trim());
        }
      });

      return {
        ok: true,
        statuses: [...statuses].sort(),
        assignees: [...assignees].sort()
      };
    } catch (e) {
      return { ok: false, error: `Kunne ikke hente filterdata: ${e.message}` };
    }
  }

  /**
   * Henter, filtrerer og sorterer oppgaver.
   * @param {object} options - Filter- og sorteringsvalg.
   * @param {string} [options.status] - Filtrer på denne statusen.
   * @param {string} [options.assignee] - Filtrer på denne ansvarlige.
   * @param {string} [options.sortBy='Frist'] - Kolonnen som skal sorteres.
   * @param {string} [options.sortOrder='asc'] - Sorteringsrekkefølge ('asc' eller 'desc').
   * @returns {object} Et objekt med en liste av oppgaver.
   */
  function getTasks(options = {}) {
    try {
      const { status, assignee, sortBy = 'Frist', sortOrder = 'asc' } = options;

      const sh = SpreadsheetApp.getActive().getSheetByName(SH_NAME);
      if (!sh || sh.getLastRow() < 2) {
        return { ok: true, tasks: [] };
      }

      const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      const headers = _getHeaders(sh);

      const statusIdx = headers['Status'];
      const assigneeIdx = headers['Ansvarlig'];
      const fristIdx = headers['Frist'];

      let tasks = data.map(row => {
        const task = {};
        for (const key in headers) {
          task[key] = row[headers[key]];
        }
        return task;
      });

      // Filtrering
      if (status) {
        tasks = tasks.filter(task => task.Status && task.Status.toString().trim() === status);
      }
      if (assignee) {
        tasks = tasks.filter(task => task.Ansvarlig && task.Ansvarlig.toString().trim() === assignee);
      }

      // Sortering
      const sortByIdx = headers[sortBy];
      if (sortByIdx !== undefined) {
        tasks.sort((a, b) => {
          let valA = a[sortBy];
          let valB = b[sortBy];

          // Behandle datoer korrekt
          if (valA instanceof Date && valB instanceof Date) {
            return sortOrder === 'asc' ? valA.getTime() - valB.getTime() : valB.getTime() - valA.getTime();
          }

          // Fallback for andre typer
          if (valA < valB) return sortOrder === 'asc' ? -1 : 1;
          if (valA > valB) return sortOrder === 'asc' ? 1 : -1;
          return 0;
        });
      }

      // Formater frist for visning
      tasks.forEach(task => {
        if (task.Frist instanceof Date) {
          task.Frist = Utilities.formatDate(task.Frist, TZ, 'yyyy-MM-dd');
        }
      });

      return { ok: true, tasks };

    } catch (e) {
      // Logger.log(`Feil i getTasks: ${e.stack}`);
      return { ok: false, error: `Kunne ikke hente oppgaver: ${e.message}` };
    }
  }

  globalThis.getTasks = getTasks;
  globalThis.getTaskFilterData = getTaskFilterData;

})();