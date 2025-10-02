/**
 * CoreAnalysisPlus – LoggerShim (v1.3.1)
 * If LoggerPlus exists (getAppLogger_), we’ll use it. Otherwise no-op console.
 * You can omit this file if you already have LoggerPlus in the project.
 */
function getAppLogger_() {
  // If you have a real LoggerPlus, keep that one instead.
  return {
    info: (fn, msg, data) => { try { console.log('[INFO]', fn || '', msg || '', data || ''); } catch (_) {} },
    warn: (fn, msg, data) => { try { console.warn('[WARN]', fn || '', msg || '', data || ''); } catch (_) {} },
    error: (fn, msg, data) => { try { console.error('[ERROR]', fn || '', msg || '', data || ''); } catch (_) {} },
    setLevel: function(){}, flush: function(){}, stats: function(){ return {}; }
  };
}
