// =============================================================================
// HMS – Vedlikeholdsplan (plan → TASKS)  [MODERNISERT]
// FILE: 60_HMS_Vedlikeholdsplan.gs
// VERSION: 2.0.0
// UPDATED: 2025-09-26
// CHANGES:
//  - Modernisert til let/const og arrow functions.
//  - Forbedret lesbarhet og kodestruktur.
// =============================================================================

const HMS_PLAN_SHEET = 'HMS_PLAN';
const TASKS_SHEET = 'TASKS';
const SUPPLIERS_SHEET = 'LEVERANDØRER';

const PLAN_HEADER = [
  'PlanID', 'System', 'Komponent', 'Oppgave', 'Beskrivelse', 'Frekvens',
  'PreferertMåned', 'NesteStart', 'AnsvarligRolle', 'Leverandør', 'LeverandørKontakt',
  'Myndighetskrav', 'Standard/Referanse', 'Kritikalitet(1-5)',
  'EstTidTimer', 'EstKost', 'HistoriskKost', 'BudsjettKonto',
  'DokumentasjonURL', 'SjekklisteURL', 'Lokasjon', 'Byggnummer', 'Garantistatus',
  'SesongAvhengig', 'SistUtført', 'Kommentar', 'Aktiv'
];

const TASKS_HEADER = [
  'Tittel', 'Kategori', 'Status', 'Frist', 'Opprettet', 'Ansvarlig',
  'Seksjonsnr', 'PlanID', 'AutoKey', 'System', 'Komponent', 'Lokasjon', 'Byggnummer',
  'Myndighetskrav', 'Kritikalitet', 'Hasteprioritering',
  'EstKost', 'BudsjettKonto', 'FaktiskKost',
  'DokumentasjonURL', 'SjekklisteURL', 'Garantistatus', 'BeboerVarsling', 'Værforhold', 'Leverandør', 'LeverandørKontakt',
  'Kommentar', 'OppdatertAv', 'Oppdatert'
];

function hmsMigrateSchema_v1_1() {
  const ss = SpreadsheetApp.getActive();

  const plan = ss.getSheetByName(HMS_PLAN_SHEET) || ss.insertSheet(HMS_PLAN_SHEET);
  if (plan.getLastRow() === 0) {
    plan.getRange(1, 1, 1, PLAN_HEADER.length).setValues([PLAN_HEADER]).setFontWeight('bold');
    plan.setFrozenRows(1);
  } else {
    _ensureColumns_(plan, PLAN_HEADER);
  }

  const tasks = ss.getSheetByName(TASKS_SHEET) || ss.insertSheet(TASKS_SHEET);
  if (tasks.getLastRow() === 0) {
    tasks.getRange(1, 1, 1, TASKS_HEADER.length).setValues([TASKS_HEADER]).setFontWeight('bold');
    tasks.setFrozenRows(1);
  } else {
    _ensureColumns_(tasks, TASKS_HEADER);
  }

  let sup = ss.getSheetByName(SUPPLIERS_SHEET);
  if (!sup) {
    sup = ss.insertSheet(SUPPLIERS_SHEET);
    sup.getRange(1, 1, 1, 6).setValues([['Kategori/System', 'Komponent', 'Navn', 'Telefon', 'Epost', 'Notat']]).setFontWeight('bold');
    sup.setFrozenRows(1);
  }

  return 'HMS v1.1 migrering ok';
}

function _ensureColumns_(sheet, desiredHeader) {
  const rng = sheet.getRange(1, 1, 1, sheet.getLastColumn() || desiredHeader.length);
  const cur = rng.getValues()[0];
  const map = cur.reduce((acc, h, i) => {
    acc[h] = i + 1;
    return acc;
  }, {});
  const missing = desiredHeader.filter(h => !map[h]);
  if (missing.length) {
    const start = (cur.filter(String).length || 0) + 1;
    sheet.getRange(1, start, 1, missing.length).setValues([missing]).setFontWeight('bold');
  }
}

function hmsEnsurePlanSheet() {
  return hmsMigrateSchema_v1_1();
}

function hmsGenerateTasks(options = {}) {
  const {
    monthsAhead = 12,
    startDate = new Date(),
    kategori = 'HMS',
    statusDefault = 'Åpen',
    buildingFilter = null,
    replaceExisting = true
  } = options;

  const ss = SpreadsheetApp.getActive();
  const plan = ss.getSheetByName(HMS_PLAN_SHEET);
  if (!plan) return { ok: false, error: `Mangler ark: ${HMS_PLAN_SHEET}` };

  let tasks = ss.getSheetByName(TASKS_SHEET) || ss.insertSheet(TASKS_SHEET);
  if (tasks.getLastRow() === 0) {
    tasks.getRange(1, 1, 1, TASKS_HEADER.length).setValues([TASKS_HEADER]).setFontWeight('bold');
  }

  const planValues = plan.getDataRange().getValues();
  if (planValues.length < 2) return { ok: false, error: 'Planen er tom.' };

  const planHeader = planValues.shift();
  const pidx = _byName_(planHeader);

  const tHeader = tasks.getRange(1, 1, 1, tasks.getLastColumn()).getValues()[0];
  const tidx = _byName_(tHeader);
  let existing = tasks.getLastRow() > 1 ? tasks.getRange(2, 1, tasks.getLastRow() - 1, tasks.getLastColumn()).getValues() : [];

  const autoKeySet = new Set();
  if (existing.length && tidx.AutoKey) {
    const akCol = tidx.AutoKey - 1;
    existing.forEach(row => {
      const ak = String(row[akCol] || '').trim();
      if (ak) autoKeySet.add(ak);
    });
  }

  if (replaceExisting && existing.length && tidx.PlanID && tidx.Frist) {
    const endDate = new Date(startDate);
    endDate.setMonth(endDate.getMonth() + monthsAhead);
    const delRows = [];
    const frCol = tidx.Frist - 1;
    const pidCol = tidx.PlanID - 1;

    existing.forEach((row, i) => {
      const d = row[frCol];
      const pid = String(row[pidCol] || '').trim();
      if (d instanceof Date && d >= startDate && d <= endDate && pid) {
        delRows.push(i + 2);
      }
    });

    delRows.sort((a, b) => b - a).forEach(r => tasks.deleteRow(r));

    existing = tasks.getLastRow() > 1 ? tasks.getRange(2, 1, tasks.getLastRow() - 1, tasks.getLastColumn()).getValues() : [];
    autoKeySet.clear();
    if (existing.length && tidx.AutoKey) {
      const akCol2 = tidx.AutoKey - 1;
      existing.forEach(row => {
        const ak2 = String(row[akCol2] || '').trim();
        if (ak2) autoKeySet.add(ak2);
      });
    }
  }

  const out = [];
  let created = 0;
  const end = new Date(startDate);
  end.setMonth(end.getMonth() + monthsAhead);

  for (const row of planValues) {
    const aktiv = _str(row[pidx.Aktiv - 1] || 'Ja').toLowerCase();
    if (['nei', '0', 'false'].includes(aktiv)) continue;

    const planId = _str(row[pidx.PlanID - 1]);
    if (!planId) continue;

    const byggnr = pidx.Byggnummer ? row[pidx.Byggnummer - 1] : '';
    if (buildingFilter) {
      if (Array.isArray(buildingFilter) && !buildingFilter.includes(Number(byggnr))) continue;
      if (!Array.isArray(buildingFilter) && Number(byggnr) !== Number(buildingFilter)) continue;
    }

    const [system, komponent, oppgave, beskrivelse, frek, pref, ansvarlig, lokasjon, mynd, krit, estKost, konto, dok, sjekkl, garanti, lever, leverKontakt] = [
      pidx.System ? _str(row[pidx.System-1]) : '',
      pidx.Komponent ? _str(row[pidx.Komponent-1]) : '',
      pidx.Oppgave ? _str(row[pidx.Oppgave-1]) : '',
      pidx.Beskrivelse ? _str(row[pidx.Beskrivelse-1]) : '',
      _str(row[pidx.Frekvens-1]),
      pidx.PreferertMåned ? _str(row[pidx.PreferertMåned-1]) : (pidx['PreferertMåned'] ? _str(row[pidx['PreferertMåned']-1]) : ''),
      pidx.AnsvarligRolle ? _str(row[pidx.AnsvarligRolle-1]) : '',
      pidx.Lokasjon ? _str(row[pidx.Lokasjon-1]) : '',
      pidx.Myndighetskrav ? _str(row[pidx.Myndighetskrav-1]) : '',
      pidx['Kritikalitet(1-5)'] ? row[pidx['Kritikalitet(1-5)']-1] : '',
      pidx.EstKost ? _num(row[pidx.EstKost-1]) : '',
      pidx.BudsjettKonto ? _str(row[pidx.BudsjettKonto-1]) : '',
      pidx.DokumentasjonURL ? _str(row[pidx.DokumentasjonURL-1]) : '',
      pidx.SjekklisteURL ? _str(row[pidx.SjekklisteURL-1]) : '',
      pidx.Garantistatus ? _str(row[pidx.Garantistatus-1]) : '',
      pidx.Leverandør ? _str(row[pidx.Leverandør-1]) : '',
      pidx.LeverandørKontakt ? _str(row[pidx.LeverandørKontakt-1]) : ''
    ];
    const nextStart = pidx.NesteStart ? _parseDateSafe_(row[pidx.NesteStart - 1]) : null;
    const sesongAvh = pidx.SesongAvhengig ? _str(row[pidx.SesongAvhengig - 1]).toLowerCase() === 'ja' : false;
    const haste = _derivePriorityFromCriticality_(krit);
    const anchor = nextStart || startDate;
    const prefMonths = _parsePreferredMonths_(pref);
    const occurrences = _expandOccurrences_(anchor, end, frek, prefMonths);

    for (const occurrence of occurrences) {
      let due = new Date(occurrence);
      if (_containsIgnoreCase_(oppgave, 'takrenne') && _isWinterMonth_(due)) {
        due = _moveToMonth_(due, 4);
      }
      if (sesongAvh && _isWinterMonth_(due)) {
        due = _moveToMonth_(due, 4);
      }

      const dueStr = Utilities.formatDate(due, Session.getScriptTimeZone() || 'Europe/Oslo', 'yyyy-MM-dd');
      const autoKey = `${planId}::${dueStr}${byggnr ? `::${byggnr}` : ''}`;
      if (autoKeySet.has(autoKey)) continue;
      autoKeySet.add(autoKey);

      const title = `${system ? `${system}: ` : ''}${komponent ? `${komponent} – ` : ''}${oppgave || 'Oppgave'}`;
      const rowOut = Array(tHeader.length).fill('');

      _set(rowOut, tidx, 'Tittel', title);
      _set(rowOut, tidx, 'Kategori', kategori);
      _set(rowOut, tidx, 'Status', statusDefault);
      _set(rowOut, tidx, 'Frist', due);
      _set(rowOut, tidx, 'Opprettet', new Date());
      _set(rowOut, tidx, 'Ansvarlig', ansvarlig);
      _set(rowOut, tidx, 'Seksjonsnr', '');
      _set(rowOut, tidx, 'PlanID', planId);
      _set(rowOut, tidx, 'AutoKey', autoKey);
      _set(rowOut, tidx, 'System', system);
      _set(rowOut, tidx, 'Komponent', komponent);
      _set(rowOut, tidx, 'Lokasjon', lokasjon);
      _set(rowOut, tidx, 'Byggnummer', byggnr);
      _set(rowOut, tidx, 'Myndighetskrav', mynd);
      _set(rowOut, tidx, 'Kritikalitet', krit);
      _set(rowOut, tidx, 'Hasteprioritering', haste);
      _set(rowOut, tidx, 'EstKost', estKost);
      _set(rowOut, tidx, 'BudsjettKonto', konto);
      _set(rowOut, tidx, 'DokumentasjonURL', dok);
      _set(rowOut, tidx, 'SjekklisteURL', sjekkl);
      _set(rowOut, tidx, 'Garantistatus', garanti);
      _set(rowOut, tidx, 'BeboerVarsling', _suggestResidentNotice_(lokasjon, oppgave, mynd, byggnr));
      _set(rowOut, tidx, 'Værforhold', lokasjon && /ute|utendørs/i.test(lokasjon) ? '' : 'N/A');
      _set(rowOut, tidx, 'Leverandør', lever);
      _set(rowOut, tidx, 'LeverandørKontakt', leverKontakt || _lookupSupplierContact_(system, komponent));
      _set(rowOut, tidx, 'Kommentar', beskrivelse);

      out.push(rowOut);
      created++;
    }
  }

  if (out.length) tasks.getRange(tasks.getLastRow() + 1, 1, out.length, tHeader.length).setValues(out);
  return { ok: true, created };
}

function markTaskCompleted(options = {}) {
  const ss = SpreadsheetApp.getActive();
  const tasks = ss.getSheetByName(TASKS_SHEET);
  if (!tasks) return { ok: false, error: `Mangler ark: ${TASKS_SHEET}` };

  const tHeader = tasks.getRange(1, 1, 1, tasks.getLastColumn()).getValues()[0];
  const tidx = _byName_(tHeader);

  let rowIndex = options.row || null;
  const autoKey = options.autoKey || null;
  if (!rowIndex && !autoKey) return { ok: false, error: 'Oppgi row eller autoKey' };

  if (!rowIndex) {
    const values = tasks.getDataRange().getValues();
    values.shift();
    const akCol = tidx.AutoKey - 1;
    rowIndex = values.findIndex(row => String(row[akCol] || '').trim() === autoKey) + 2;
    if (rowIndex < 2) return { ok: false, error: 'Fant ikke oppgave med AutoKey' };
  }

  const now = new Date();
  if (tidx.Status) tasks.getRange(rowIndex, tidx.Status).setValue('Utført');
  if (tidx.OppdatertAv) tasks.getRange(rowIndex, tidx.OppdatertAv).setValue(Session.getEffectiveUser().getEmail());
  if (tidx.Oppdatert) tasks.getRange(rowIndex, tidx.Oppdatert).setValue(now);
  if (tidx.FaktiskKost && options.faktiskKost != null) tasks.getRange(rowIndex, tidx.FaktiskKost).setValue(Number(options.faktiskKost) || 0);

  const planId = tidx.PlanID ? tasks.getRange(rowIndex, tidx.PlanID).getValue() : '';
  if (planId) _updatePlanAfterCompletion_(String(planId), now, Number(options.faktiskKost) || null);

  return { ok: true, row: rowIndex };
}

function _updatePlanAfterCompletion_(planId, dateDone, actualCost) {
  const ss = SpreadsheetApp.getActive();
  const plan = ss.getSheetByName(HMS_PLAN_SHEET);
  if (!plan) return;

  const values = plan.getDataRange().getValues();
  const header = values.shift();
  const pidx = _byName_(header);
  const pidCol = pidx.PlanID - 1;
  const sistCol = pidx.SistUtført - 1;
  const histCol = pidx.HistoriskKost ? pidx.HistoriskKost - 1 : null;
  const estCol = pidx.EstKost ? pidx.EstKost - 1 : null;

  const rowIndex = values.findIndex(row => String(row[pidCol] || '').trim() === planId);
  if (rowIndex !== -1) {
    plan.getRange(rowIndex + 2, sistCol + 1).setValue(dateDone);
    if (histCol != null) {
      const prev = Number(values[rowIndex][histCol] || 0);
      const base = (actualCost != null && !isNaN(actualCost)) ? actualCost : (estCol != null ? Number(values[rowIndex][estCol] || 0) : 0);
      const updated = prev ? Math.round(((prev * 2 + base) / 3) * 100) / 100 : base;
      plan.getRange(rowIndex + 2, histCol + 1).setValue(updated);
    }
  }
}

function _lookupSupplierContact_(system, komponent) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SUPPLIERS_SHEET);
  if (!sh) return '';
  const values = sh.getDataRange().getValues();
  values.shift();
  for (const row of values) {
    const cat = _str(row[0]);
    const comp = _str(row[1]);
    if (cat && _equalsIgnoreCase_(cat, system) && (!comp || _equalsIgnoreCase_(comp, komponent))) {
      const navn = _str(row[2]), tlf = _str(row[3]), ep = _str(row[4]);
      return [navn, tlf, ep].filter(Boolean).join(' ');
    }
  }
  return '';
}

function _suggestResidentNotice_(lokasjon, oppgave, myndighetskrav, byggnr) {
  let out = 'Nei';
  if (/heis|vann|sprinkler|brannalarm|garasjeport/i.test(oppgave)) out = 'Ja';
  if (lokasjon && /ute|utendørs/i.test(lokasjon)) out = 'Vurder';
  if (_str(myndighetskrav).toLowerCase() === 'ja') out = 'Vurder';
  if (byggnr) out += ` (Bygg ${byggnr})`;
  return out;
}

function hmsNotifyUpcomingTasks(daysAhead) {
  const ss = SpreadsheetApp.getActive();
  const tasks = ss.getSheetByName(TASKS_SHEET);
  if (!tasks) return { ok: true, notified: 0 };

  const values = tasks.getDataRange().getValues();
  const header = values.shift();
  const idx = _byName_(header);

  const today = new Date();
  const ahead = Number(daysAhead || 14);
  const limit = new Date(today.getFullYear(), today.getMonth(), today.getDate() + ahead);

  const res = values.filter(r => {
    const status = _str(r[idx.Status - 1] || '');
    const due = r[idx.Frist - 1];
    return status === 'Åpen' && due instanceof Date && due >= today && due <= limit;
  });

  if (res.length === 0) return { ok: true, notified: 0 };

  const to = PropertiesService.getScriptProperties().getProperty('HMS_NOTIFY_EMAIL') || Session.getEffectiveUser().getEmail();
  const lines = res.slice(0, 50).map(r => {
    const t = r[idx.Tittel - 1], d = Utilities.formatDate(r[idx.Frist - 1], Session.getScriptTimeZone() || 'Europe/Oslo', 'yyyy-MM-dd');
    const b = r[idx.Byggnummer - 1] || '', h = r[idx.Hasteprioritering - 1] || '';
    return `• [${h || 'Normal'}] ${t}${b ? ` (Bygg ${b})` : ''} – frist ${d}`;
  });
  MailApp.sendEmail({ to, subject: `HMS: kommende oppgaver (${res.length})`, body: lines.join('\n') });
  return { ok: true, notified: res.length, to };
}

const _byName_ = (header) => header.reduce((map, h, i) => {
  const key = String(h || '').trim();
  if (key) map[key] = i + 1;
  return map;
}, {});

const _emptyRow_ = n => Array(n).fill('');
const _set = (arr, idx, name, val) => { if (idx[name]) arr[idx[name] - 1] = val; };
const _str = v => String(v == null ? '' : v).trim();
const _equalsIgnoreCase_ = (a, b) => String(a || '').trim().toLowerCase() === String(b || '').trim().toLowerCase();
const _containsIgnoreCase_ = (s, needle) => String(s || '').toLowerCase().includes(String(needle || '').toLowerCase());

function _parseDateSafe_(v) {
  if (v instanceof Date) return v;
  const s = _str(v);
  if (!s) return null;
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (!isNaN(d.getTime())) return d;
  }
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

function _parsePreferredMonths_(s) {
  s = _str(s);
  if (!s) return [];
  const tokens = s.split(/[;,\/\s]+/).filter(Boolean);
  const map = { jan: 1, januar: 1, feb: 2, februar: 2, mar: 3, mars: 3, apr: 4, april: 4, mai: 5, jun: 6, juni: 6, jul: 7, juli: 7, aug: 8, august: 8, sep: 9, september: 9, okt: 10, oktober: 10, nov: 11, november: 11, des: 12, desember: 12, 'vår': -1, 'var': -1, 'sommer': -2, 'høst': -3, 'host': -3, 'vinter': -4 };
  let out = [];
  tokens.forEach(t => {
    const token = t.toLowerCase();
    const m = map[token];
    if (m) {
      if (m > 0) out.push(m);
      else if (m === -1) out.push(3, 4, 5);
      else if (m === -2) out.push(6, 7, 8);
      else if (m === -3) out.push(9, 10, 11);
      else if (m === -4) out.push(12, 1, 2);
    } else {
      const n = Number(token);
      if (Number.isInteger(n) && n >= 1 && n <= 12) out.push(n);
    }
  });
  return [...new Set(out)].sort((a, b) => a - b);
}

function _expandOccurrences_(anchorDate, endDate, freq, preferredMonths) {
  const occ = [];
  const a = new Date(anchorDate);
  const e = new Date(endDate);
  const f = _str(freq).toUpperCase().replace(/[ÅØÆ]/g, c => ({ 'Å': 'A', 'Ø': 'O', 'Æ': 'AE' }[c]));

  const makeDate = (y, m, d) => new Date(y, m - 1, d || 15);

  let interval = 12;
  if (f.includes('MND') || f.includes('MANED') || f === 'MÅNEDLIG' || f === 'MANEDLIG') interval = 1;
  else if (f.includes('KVART')) interval = 3;
  else if (f.includes('HALV')) interval = 6;
  else if (f.includes('2AAR') || f === '2ÅR') interval = 24;
  else if (f.includes('3AAR') || f === '3ÅR') interval = 36;
  else if (f.includes('5AAR') || f === '5ÅR') interval = 60;
  else if (f.includes('10AAR') || f === '10ÅR') interval = 120;

  if (preferredMonths && preferredMonths.length) {
    const sy = a.getFullYear(), sm = a.getMonth() + 1;
    const ey = e.getFullYear(), em = e.getMonth() + 1;
    for (let y = sy; y <= ey; y++) {
      for (let m = 1; m <= 12; m++) {
        if ((y === sy && m < sm) || (y === ey && m > em)) continue;
        if (preferredMonths.includes(m)) occ.push(makeDate(y, m, a.getDate() || 15));
      }
    }
  } else {
    let d = new Date(a);
    while (d <= e) {
      occ.push(new Date(d));
      d.setMonth(d.getMonth() + interval);
    }
  }
  return occ.filter(dt => dt >= a && dt <= e);
}

const _isWinterMonth_ = date => {
  const m = date.getMonth() + 1;
  return m === 11 || m === 12 || m <= 3;
};

const _moveToMonth_ = (date, month1to12) => new Date(date.getFullYear(), month1to12 - 1, 15);

function _derivePriorityFromCriticality_(krit) {
  const n = Number(krit || 0);
  if (n >= 5) return 'Kritisk';
  if (n >= 4) return 'Høy';
  if (n >= 3) return 'Normal';
  return 'Lav';
}

function _num(v) {
  if (v === '' || v == null) return '';
  let s = String(v).trim().replace(/\s/g, '');
  const hasC = s.includes(','), hasD = s.includes('.');
  if (hasC && !hasD) s = s.replace(/\./g, '').replace(',', '.');
  else if (hasC && hasD && s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? '' : n;
}