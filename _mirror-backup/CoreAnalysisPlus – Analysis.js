/**
 * CoreAnalysisPlus – Analysis (v1.4.1)
 *
 * What’s new vs v1.4.0:
 *  - Phase-aware progress reporting (phase + overall)
 *  - Failure phase tracking surfaced in fallback error object
 *  - Set-based function dedupe
 *  - Config validation warning on missing keys
 *  - Richer performance metrics (avg per sheet, detailed marks, isLargeDataset)
 *  - Keeps defensive dep bridges + metrics hook
 *
 * Public:
 *   performComprehensiveAnalysis_(options?)
 *     @param {Object} [options]
 *     @param {(p:{current:number,total:number,sheetName?:string,percentage:number,phase?:string,phaseProgress?:number,overallProgress?:number,currentPhase?:number,totalPhases?:number})=>void} [options.progressCb]
 *     @param {string} [options.metricEventName='core_analysis']
 *     @param {boolean} [options.includeHidden] (reserved)
 */

(function (global) {
  'use strict';

  /* -------------------------- Safe dependency bridges -------------------------- */

  const _log = (() => {
    try {
      if (typeof _getLoggerPlus_ === 'function') return _getLoggerPlus_();
      if (typeof getAppLogger_ === 'function')   return getAppLogger_();
    } catch (_) {}
    return {
      debug: (fn, msg, d) => { try { console.log('[DEBUG]', fn||'', msg||'', d||''); } catch(_){} },
      info:  (fn, msg, d) => { try { console.log('[INFO]',  fn||'', msg||'', d||''); } catch(_){} },
      warn:  (fn, msg, d) => { try { console.warn('[WARN]', fn||'', msg||'', d||''); } catch(_){} },
      error: (fn, msg, d) => { try { console.error('[ERROR]',fn||'', msg||'', d||''); } catch(_){} }
    };
  })();

  const _cfgGet = (key, fallback) => {
    try { if (typeof _cfgGet_ === 'function') return _cfgGet_(key, fallback); } catch(_) {}
    try {
      if (typeof CORE_ANALYSIS_CFG !== 'undefined' &&
          CORE_ANALYSIS_CFG &&
          Object.prototype.hasOwnProperty.call(CORE_ANALYSIS_CFG, key)) {
        return CORE_ANALYSIS_CFG[key];
      }
    } catch(_) {}
    return fallback;
  };

  const _numCfg = (key, fallback) => {
    try { if (typeof _numCfgSafe_ === 'function') return _numCfgSafe_(key, fallback); } catch(_) {}
    const n = Number(_cfgGet(key, fallback));
    return Number.isFinite(n) ? n : Number(fallback);
  };

  const _safeCall = (fn, fb) => {
    try { if (typeof _safe === 'function') return _safe(fn, fb); } catch(_) {}
    try { return fn(); } catch(_) { return fb; }
  };

  const _metric = (eventName, handler, details) => {
    try {
      if (typeof _metric_ === 'function') _metric_(eventName, handler, details);
    } catch (_) {}
  };

  /* ------------------------------ Perf mini-helper ---------------------------- */

  class Perf {
    constructor(op) { this.op = op; this.t0 = Date.now(); this.marks = []; }
    mark(label) { this.marks.push({ label, ms: Date.now() - this.t0 }); return this; }
    done() { return { operation: this.op, totalMs: Date.now() - this.t0, marks: this.marks }; }
  }

  /* ------------------------------ Config validation --------------------------- */

  function _validateConfigKeys_(fn) {
    const required = ['LARGE_DATA_SHEETS','LARGE_DATA_MAXCOLS','LARGE_DATA_TOTALROWS','VERSION'];
    const missing = [];
    for (let i = 0; i < required.length; i++) {
      const k = required[i];
      const v = _cfgGet(k, null);
      if (v === null || typeof v === 'undefined') missing.push(k);
    }
    if (missing.length) {
      _log.warn(fn, 'Missing configuration values', { missing });
    }
  }

  /* ------------------------------- Main function ------------------------------ */

  function performComprehensiveAnalysis_(options) {
    const fn = 'performComprehensiveAnalysis_';
    const perf = new Perf(fn);
    const opt = options && typeof options === 'object' ? options : {};
    const progressCb = typeof opt.progressCb === 'function' ? opt.progressCb : null;
    const metricName = String(opt.metricEventName || 'core_analysis');

    const PHASES = ['metadata','triggers','menus','dataModel','merge'];
    let currentPhaseIndex = 0;
    let failedPhase = null;

    const phaseReport = (phase, phaseProgress) => {
      if (!progressCb) return;
      const overall = ((currentPhaseIndex + (phaseProgress || 0)) / PHASES.length) * 100;
      progressCb({
        phase,
        phaseProgress: phaseProgress || 0,
        overallProgress: Math.max(0, Math.min(100, overall)),
        currentPhase: currentPhaseIndex + 1,
        totalPhases: PHASES.length
      });
    };

    try {
      _validateConfigKeys_(fn);

      // 1) Metadata
      phaseReport(PHASES[currentPhaseIndex], 0);
      failedPhase = 'metadata';
      const meta = _collectMetadata_();
      perf.mark('metadata');
      phaseReport(PHASES[currentPhaseIndex], 1);
      currentPhaseIndex++;

      // 2) Triggers
      phaseReport(PHASES[currentPhaseIndex], 0);
      failedPhase = 'triggers';
      const triggers = _collectTriggers_() || [];
      perf.mark('triggers');
      phaseReport(PHASES[currentPhaseIndex], 1);
      currentPhaseIndex++;

      // 3) Menus
      phaseReport(PHASES[currentPhaseIndex], 0);
      failedPhase = 'menus';
      const menuFns = _collectMenuFunctions_() || [];
      perf.mark('menus');
      phaseReport(PHASES[currentPhaseIndex], 1);
      currentPhaseIndex++;

      // 4) Data model (+ per-sheet progress passthrough)
      phaseReport(PHASES[currentPhaseIndex], 0);
      failedPhase = 'dataModel';
      const dataModel = (typeof _collectDataModel_ === 'function')
        ? _collectDataModel_(p => {
            // p = {current,total,sheetName,percentage}
            phaseReport('dataModel', Math.max(0, Math.min(1, (p && p.percentage ? p.percentage : 0) / 100)));
            // Also stream the raw per-sheet progress if consumer wants it
            if (progressCb) progressCb(p);
          })
        : { sheets: [], headerDuplicates: [] };
      perf.mark('dataModel');
      phaseReport(PHASES[currentPhaseIndex], 1);
      currentPhaseIndex++;

      // 5) Merge function inventory
      phaseReport(PHASES[currentPhaseIndex], 0);
      failedPhase = 'merge';
      const fnSet = new Set();
      const functions = [];

      (triggers || []).forEach(t => {
        const name = String(_safeCall(() => t.handler, '') || '').trim();
        if (name && !fnSet.has(name)) {
          fnSet.add(name);
          functions.push({ name, source: 'trigger', eventType: _safeCall(() => t.eventType, '') || '' });
        }
      });

      (menuFns || []).forEach(m => {
        const name = String(_safeCall(() => m.functionName, '') || '').trim();
        if (name && !fnSet.has(name)) {
          fnSet.add(name);
          functions.push({ name, source: 'menu', title: _safeCall(() => m.title, '') || '' });
        }
      });
      perf.mark('merge');
      phaseReport(PHASES[currentPhaseIndex], 1);
      currentPhaseIndex++;

      // 6) Perf + large dataset detection
      const sheetsArr = dataModel.sheets || [];
      const sheetsScanned = sheetsArr.length;
      const totalRows = sheetsArr.reduce((sum, s) => sum + (s.rows || 0), 0);
      const maxCols = sheetsArr.reduce((m, s) => Math.max(m, s.columns || 0), 0);

      const LD_SHEETS    = _numCfg('LARGE_DATA_SHEETS', 50);
      const LD_MAXCOLS   = _numCfg('LARGE_DATA_MAXCOLS', 100);
      const LD_TOTALROWS = _numCfg('LARGE_DATA_TOTALROWS', 50000);

      const isLarge = (sheetsScanned >= LD_SHEETS) || (maxCols >= LD_MAXCOLS) || (totalRows >= LD_TOTALROWS);
      if (isLarge) {
        _log.info(fn, 'Large dataset detected', {
          sheetsScanned, maxCols, totalRows,
          thresholds: { LD_SHEETS, LD_MAXCOLS, LD_TOTALROWS }
        });
      }

      // 7) Build result
      const perfResult = perf.done();
      const avgRows = sheetsScanned ? (totalRows / sheetsScanned) : 0;
      const avgTimePerSheet = sheetsScanned ? (perfResult.totalMs / sheetsScanned) : 0;

      const result = {
        metadata: meta,
        triggers: { count: triggers.length, details: triggers },
        menus: { fromSheets: menuFns },
        functions: { global: functions, private: [] },
        sheets: {
          count: sheetsArr.length,
          sheets: sheetsArr,
          headerDuplicates: dataModel.headerDuplicates || []
        },
        performanceMetrics: {
          sheetsScanned,
          totalRows,
          maxCols,
          scanDurationMs: perfResult.totalMs,
          averageRowsPerSheet: avgRows,
          averageTimePerSheetMs: avgTimePerSheet,
          isLargeDataset: isLarge,
          detailedTimings: perfResult.marks
        },
        version: _cfgGet('VERSION', _safeCall(() => CORE_ANALYSIS_CFG.VERSION, '1.x'))
      };

      _log.info(fn, 'Full analysis complete', {
        ms: perfResult.totalMs,
        sheets: result.sheets.count,
        triggers: result.triggers.count,
        functions: result.functions.global.length
      });
      _metric(metricName, fn, {
        ms: perfResult.totalMs,
        sheets: result.sheets.count,
        triggers: result.triggers.count,
        functions: result.functions.global.length
      });

      return result;

    } catch (err) {
      const perfResult = perf.done();
      const errorObj = {
        message: err && err.message,
        stack: err && err.stack,
        phase: failedPhase || 'unknown'
      };

      _log.error(fn, `Analysis failed in phase: ${errorObj.phase}`, {
        error: errorObj.message, stack: errorObj.stack
      });
      _metric(metricName + '_error', fn, { error: errorObj.message, phase: errorObj.phase });

      // Safe fallback result so callers don’t crash
      return {
        metadata: {},
        triggers: { count: 0, details: [] },
        menus: { fromSheets: [] },
        functions: { global: [], private: [] },
        sheets: { count: 0, sheets: [], headerDuplicates: [] },
        performanceMetrics: {
          sheetsScanned: 0,
          totalRows: 0,
          maxCols: 0,
          scanDurationMs: perfResult.totalMs || 0,
          averageRowsPerSheet: 0,
          averageTimePerSheetMs: perfResult.totalMs || 0,
          isLargeDataset: false,
          detailedTimings: perfResult.marks || []
        },
        version: _cfgGet('VERSION', _safeCall(() => CORE_ANALYSIS_CFG.VERSION, '1.x')),
        error: errorObj
      };
    }
  }

  /* -------------------------- Attach to global scope -------------------------- */

  global.performComprehensiveAnalysis_ = performComprehensiveAnalysis_;

})(typeof globalThis !== 'undefined' ? globalThis : this);
