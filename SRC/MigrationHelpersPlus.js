/**
 * MigrationHelpersPlus — v1.0.0
 * -----------------------------------------------------------------------------
 * - Tracks schema version (semver) AND a deterministic headerSignature (SHA-1)
 * - Dry-run planning of header changes for OWNERSHIP_CFG.sheetName
 * - Safe apply with optional backup and reorder-by-rebuild
 * - Version/signature stored in ScriptProperties
 *
 * Typical use:
 *   ownershipMigrationPlan_();                 // inspect (shows versions & signatures)
 *   applyOwnershipMigration_({
 *     dryRun: false,
 *     createBackupCopy: true,
 *     ensureColumns: true,
 *     reorderToDefault: true,
 *     targetVersion: '2.1.0'
 *   });
 *
 * Dev helpers:
 *   dev_showSchemaVersion_();
 *   dev_setSchemaVersion_('2.1.0');
 *   dev_showSchemaSignature_();
 *   dev_recomputeAndStoreSchemaSignature_();  // recompute from current header
 */

// ----------------------------- Version / Signature --------------------------

const OWNERSHIP_SCHEMA_META = {
  MODULE_VERSION: '1.0.0',                      // this helper file version
  PROP_KEY_VERSION: 'OWNERSHIP_SCHEMA_VERSION', // semver string
  PROP_KEY_SIG:     'OWNERSHIP_SCHEMA_SIGNATURE', // header signature (hash)
  DEFAULT_CURRENT:  '1.0.0',
  DEFAULT_TARGET:   '2.0.0',
};

/** Read current schema version (semver) from ScriptProperties. */
function getOwnershipSchemaVersion_() {
  try {
    const p = PropertiesService.getScriptProperties();
    return p.getProperty(OWNERSHIP_SCHEMA_META.PROP_KEY_VERSION) || OWNERSHIP_SCHEMA_META.DEFAULT_CURRENT;
  } catch (_) {
    return OWNERSHIP_SCHEMA_META.DEFAULT_CURRENT;
  }
}

/** Persist schema version (semver) to ScriptProperties. */
function setOwnershipSchemaVersion_(versionStr) {
  try {
    const p = PropertiesService.getScriptProperties();
    p.setProperty(OWNERSHIP_SCHEMA_META.PROP_KEY_VERSION, String(versionStr || '').trim() || OWNERSHIP_SCHEMA_META.DEFAULT_CURRENT);
    return true;
  } catch (e) {
    _getLogger_().warn('setOwnershipSchemaVersion_', 'Failed to set schema version', { error: e.message });
    return false;
  }
}

/** Read current header signature from ScriptProperties. */
function getOwnershipSchemaSignature_() {
  try {
    const p = PropertiesService.getScriptProperties();
    return p.getProperty(OWNERSHIP_SCHEMA_META.PROP_KEY_SIG) || '';
  } catch (_) {
    return '';
  }
}

/** Persist header signature to ScriptProperties. */
function setOwnershipSchemaSignature_(sig) {
  try {
    const p = PropertiesService.getScriptProperties();
    p.setProperty(OWNERSHIP_SCHEMA_META.PROP_KEY_SIG, String(sig || ''));
    return true;
  } catch (e) {
    _getLogger_().warn('setOwnershipSchemaSignature_', 'Failed to set schema signature', { error: e.message });
    return false;
  }
}

/** Compute deterministic signature of a header array (case/space-insensitive). */
function computeHeaderSignature_(headerArr) {
  try {
    const norm = (headerArr || []).map(h => String(h || '').trim().toLowerCase()).join('|');
    const bytes = Utilities.newBlob(norm).getBytes();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, bytes);
    return digest.map(b => ((b & 0xff).toString(16).padStart(2, '0'))).join('');
  } catch (e) {
    _getLogger_().warn('computeHeaderSignature_', 'Failed to compute signature', { error: e.message });
    return '';
  }
}

/** Compare two semver-like strings. Returns -1, 0, 1. */
function _semverCmp_(a, b) {
  const pa = String(a || '0').split('.').map(x => parseInt(x, 10) || 0);
  const pb = String(b || '0').split('.').map(x => parseInt(x, 10) || 0);
  const len = Math.max(pa.length, pb.length);
  for (let i = 0; i < len; i++) {
    const da = pa[i] || 0;
    const db = pb[i] || 0;
    if (da < db) return -1;
    if (da > db) return 1;
  }
  return 0;
}

// ------------------------------- Dev Helpers --------------------------------

function dev_showSchemaVersion_() {
  const v = getOwnershipSchemaVersion_();
  _getLogger_().info('dev_showSchemaVersion_', 'Current ownership schema version', { version: v });
  return v;
}

function dev_setSchemaVersion_(v) {
  setOwnershipSchemaVersion_(v);
  _getLogger_().info('dev_setSchemaVersion_', 'Schema version set manually', { version: v });
  return getOwnershipSchemaVersion_();
}

function dev_showSchemaSignature_() {
  const sig = getOwnershipSchemaSignature_();
  _getLogger_().info('dev_showSchemaSignature_', 'Current schema signature', { signature: sig });
  return sig;
}

/** Recompute signature from the sheet’s current header and store it. */
function dev_recomputeAndStoreSchemaSignature_() {
  const ss = _getSs_();
  const sh = ss.getSheetByName(OWNERSHIP_CFG.sheetName);
  if (!sh) {
    _getLogger_().warn('dev_recomputeAndStoreSchemaSignature_', 'Sheet missing', { sheet: OWNERSHIP_CFG.sheetName });
    return '';
    }
  const lastCol = sh.getLastColumn();
  const current = (lastCol > 0) ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const sig = computeHeaderSignature_(current);
  setOwnershipSchemaSignature_(sig);
  _getLogger_().info('dev_recomputeAndStoreSchemaSignature_', 'Signature recomputed & stored', { signature: sig });
  return sig;
}

// ---------------------------- Migration Planning ----------------------------

/**
 * Build a migration plan from current sheet header → desired OWNERSHIP_CFG.outputColumns.
 * Adds version/signature info {currentVersion, targetVersion, versionCmp, currentSignature, desiredSignature, signatureMatch}.
 *
 * @param {Object} [opts]
 * @param {string} [opts.targetVersion] desired target schema version (semver)
 * @returns {Object} plan
 */
function ownershipMigrationPlan_(opts) {
  const log = _getLogger_();
  const fn  = 'ownershipMigrationPlan_';
  const options = Object(opts || {});
  const currentVersion = getOwnershipSchemaVersion_();
  const targetVersion = String(options.targetVersion || OWNERSHIP_SCHEMA_META.DEFAULT_TARGET);

  const ss = _getSs_();
  const sheetName = OWNERSHIP_CFG.sheetName;
  const sh = ss.getSheetByName(sheetName);

  let current = [];
  let lastRow = 0;
  let lastCol = 0;

  if (!sh) {
    const desired = OWNERSHIP_CFG.outputColumns.slice();
    const planMissing = {
      sheetName,
      current,
      desired,
      missing: desired.slice(),
      extra: [],
      orderMismatch: true,
      rowCount: 0,
      colCount: desired.length,
      currentVersion,
      targetVersion,
      versionCmp: _semverCmp_(currentVersion, targetVersion),
      currentSignature: '',
      desiredSignature: computeHeaderSignature_(desired),
      signatureMatch: false,
      note: 'Sheet does not exist. Will be created if migration is applied.'
    };
    log.info(fn, 'Migration plan (sheet missing)', { ...planMissing, current: undefined, desired: undefined });
    return planMissing;
  }

  lastRow = sh.getLastRow();
  lastCol = sh.getLastColumn();
  current = (lastCol > 0)
    ? (sh.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v || '').trim()))
    : [];

  const desired = OWNERSHIP_CFG.outputColumns.slice();
  const currentSet = new Set(current.map(s => s.toLowerCase()));
  const desiredSet = new Set(desired.map(s => s.toLowerCase()));

  const missing = desired.filter(d => !currentSet.has(d.toLowerCase()));
  const extra   = current.filter(c => !desiredSet.has(c.toLowerCase()));
  const orderMismatch = (current.length !== desired.length) ||
    current.some((c, i) => (String(c || '').trim() !== String(desired[i] || '').trim()));

  const currentSignature = computeHeaderSignature_(current);
  const desiredSignature = computeHeaderSignature_(desired);
  const signatureMatch   = currentSignature && desiredSignature && (currentSignature === desiredSignature);

  const plan = {
    sheetName,
    current,
    desired,
    missing,
    extra,
    orderMismatch,
    rowCount: lastRow,
    colCount: lastCol,
    currentVersion,
    targetVersion,
    versionCmp: _semverCmp_(currentVersion, targetVersion),
    currentSignature,
    desiredSignature,
    signatureMatch
  };

  log.info(fn, 'Migration plan built', { ...plan, current: undefined, desired: undefined });
  return plan;
}

// ----------------------------- Migration Apply ------------------------------

/**
 * Apply the migration safely.
 * Options:
 *  - dryRun (bool): only returns plan, no changes (default false)
 *  - createBackupCopy (bool): duplicate original sheet before modifying (default true)
 *  - backupSuffix (string): suffix for backup sheet name (default '_backup')
 *  - ensureColumns (bool): add any missing columns to existing sheet header (default true)
 *  - reorderToDefault (bool): create new sheet with desired header and copy rows (default true)
 *  - newSheetSuffix (string): suffix for new sheet if reordering (default '_v2')
 *  - targetVersion (string): desired target schema version (default OWNERSHIP_SCHEMA_META.DEFAULT_TARGET)
 *  - setVersionOnApply (bool): write targetVersion to Props if applied (default true)
 *  - setSignatureOnApply (bool): write desired header signature to Props (default true)
 *  - forceApply (bool): ignore version/signature checks and apply anyway (default false)
 *
 * @returns {Object} result { applied, actions, plan, newSheetName?, backupSheetName?, versionUpdated?, signatureUpdated? }
 */
function applyOwnershipMigration_(opts) {
  const log = _getLogger_();
  const fn  = 'applyOwnershipMigration_';

  const options = Object.assign({
    dryRun: false,
    createBackupCopy: true,
    backupSuffix: '_backup',
    ensureColumns: true,
    reorderToDefault: true,
    newSheetSuffix: '_v2',
    targetVersion: OWNERSHIP_SCHEMA_META.DEFAULT_TARGET,
    setVersionOnApply: true,
    setSignatureOnApply: true,
    forceApply: false
  }, opts || {});

  const plan = ownershipMigrationPlan_({ targetVersion: options.targetVersion });
  const result = { applied: false, actions: [], plan, versionUpdated: false, signatureUpdated: false };

  // Guard (unless forced): if version OK AND signatures match → no-op
  if (!options.forceApply) {
    const versionOk  = (plan.versionCmp >= 0);
    const sigMatches = !!plan.signatureMatch;
    if (versionOk && sigMatches) {
      result.actions.push('no_change_needed_version_and_signature_ok');
      if (typeof _metric_ === 'function') _metric_('migration_noop', fn, { reason: 'version_signature_ok' });
      return result;
    }
  }

  if (options.dryRun) {
    result.actions.push('dryRun_only');
    if (typeof _metric_ === 'function') _metric_('migration_dry_run', fn, { currentVersion: plan.currentVersion, targetVersion: plan.targetVersion });
    return result;
  }

  const ss = _getSs_();
  let sh = ss.getSheetByName(plan.sheetName);

  // Create fresh sheet if missing
  if (!sh) {
    const created = ss.insertSheet(plan.sheetName);
    _setupSheet_(created, OWNERSHIP_CFG.outputColumns);
    sh = created;
    result.applied = true;
    result.actions.push('created_sheet_with_desired_header');
    log.info(fn, 'Sheet created with desired header', { sheet: plan.sheetName });
  } else {
    // Optional: backup
    if (options.createBackupCopy) {
      const backupName = uniqueSheetName_(ss, plan.sheetName + options.backupSuffix);
      sh.copyTo(ss).setName(backupName);
      result.backupSheetName = backupName;
      result.actions.push('backup_created');
      log.info(fn, 'Backup created', { backupSheetName: backupName });
    }

    // Only missing columns → append (no reorder)
    if (options.ensureColumns && plan.missing.length > 0 && !options.reorderToDefault) {
      appendMissingColumns_(sh, plan);
      result.actions.push('missing_columns_appended');
      result.applied = true;
    }

    // Rebuild to desired order (and copy rows mapped)
    if (options.reorderToDefault) {
      const newName = uniqueSheetName_(ss, plan.sheetName + options.newSheetSuffix);
      const newSh = ss.insertSheet(newName);
      _setupSheet_(newSh, OWNERSHIP_CFG.outputColumns);

      // Map current header → index
      const currentIndex = {};
      plan.current.forEach((h, i) => { currentIndex[(h || '').toLowerCase()] = i; });

      const rows = sh.getLastRow();
      if (rows > 1) {
        const cols = sh.getLastColumn();
        const data = (cols > 0) ? sh.getRange(2, 1, rows - 1, cols).getValues() : [];
        const desired = OWNERSHIP_CFG.outputColumns;
        const desiredLower = desired.map(h => String(h || '').toLowerCase());

        const out = data.map(r => {
          const line = new Array(desired.length).fill('');
          desiredLower.forEach((hLower, di) => {
            const srcIdx = currentIndex[hLower];
            if (typeof srcIdx === 'number' && srcIdx >= 0 && srcIdx < r.length) {
              line[di] = r[srcIdx];
            } else if (desired[di] === 'Tidspunkt') {
              line[di] = ''; // don’t synthesize timestamps
            }
          });
          return line;
        });

        if (out.length > 0) newSh.getRange(2, 1, out.length, desired.length).setValues(out);
      }

      // Swap: delete old, rename new to original
      const oldName = plan.sheetName;
      ss.deleteSheet(sh);
      newSh.setName(oldName);
      result.newSheetName = oldName;
      result.actions.push('reordered_by_rebuild');
      result.applied = true;

      log.info(fn, 'Reordered by rebuild', { sheet: oldName });
      sh = newSh; // not strictly necessary after rename, but keeps reference clear
    }
  }

  // If we applied changes → optionally persist version & signature
  if (result.applied) {
    if (options.setVersionOnApply && setOwnershipSchemaVersion_(plan.targetVersion)) {
      result.versionUpdated = true;
      result.actions.push('version_updated');
      if (typeof _metric_ === 'function') _metric_('migration_version_updated', fn, { from: plan.currentVersion, to: plan.targetVersion });
    }
    if (options.setSignatureOnApply) {
      const desiredSig = computeHeaderSignature_(OWNERSHIP_CFG.outputColumns);
      if (setOwnershipSchemaSignature_(desiredSig)) {
        result.signatureUpdated = true;
        result.actions.push('signature_updated');
      } else {
        result.actions.push('signature_update_failed');
      }
    }
  }

  if (!result.applied) {
    result.actions.push('no_change_detected');
    if (typeof _metric_ === 'function') _metric_('migration_noop', fn, { reason: 'no_changes_detected' });
  } else {
    if (typeof _metric_ === 'function') _metric_('migration_applied', fn, { actions: result.actions, from: plan.currentVersion, to: plan.targetVersion });
  }

  return result;
}

// -------------------------------- Utilities ---------------------------------

/** Append any missing columns to the end of the existing header (no reordering). */
function appendMissingColumns_(sh, plan) {
  const toAppend = plan.missing.slice();
  if (toAppend.length === 0) return;

  const startCol = (sh.getLastColumn() || plan.current.length) + 1;
  sh.getRange(1, startCol, 1, toAppend.length).setValues([toAppend]);
  sh.getRange(1, startCol, 1, toAppend.length).setFontWeight('bold').setBackground('#E8F0FE');
}

/** Produce a unique sheet name within the spreadsheet. */
function uniqueSheetName_(ss, baseName) {
  let name = baseName;
  let i = 1;
  while (ss.getSheetByName(name)) name = `${baseName}_${i++}`;
  return name;
}
