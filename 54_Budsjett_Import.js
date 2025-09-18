/* ====================== Budsjett – Import & Normalisering ======================
 * FILE: 54_Budsjett_Import.gs
 * VERSION: 1.1.0
 * UPDATED: 2025-09-15
 * FORMÅL: Lese budsjett fra mal-ark, validere og normalisere til BUDSJETT-ark
 * ============================================================================== */

const BUDGET_OUTPUT_SHEET = 'BUDSJETT';
const DEFAULT_BUDGET_VERSION = 'main';

const MONTH_ALIASES = {
  '1':'Jan','01':'Jan','jan':'Jan','januar':'Jan',
  '2':'Feb','02':'Feb','feb':'Feb','februar':'Feb',
  '3':'Mar','03':'Mar','mar':'Mar','mars':'Mar',
  '4':'Apr','04':'Apr','apr':'Apr','april':'Apr',
  '5':'Mai','05':'Mai','mai':'Mai',
  '6':'Jun','06':'Jun','jun':'Jun','juni':'Jun',
  '7':'Jul','07':'Jul','jul':'Jul','juli':'Jul',
  '8':'Aug','08':'Aug','aug':'Aug','august':'Aug',
  '9':'Sep','09':'Sep','sep':'Sep','september':'Sep',
  '10':'Okt','okt':'Okt','oktober':'Okt',
  '11':'Nov','nov':'Nov','november':'Nov',
  '12':'Des','des':'Des','desember':'Des'
};
const MONTHS = ['Jan','Feb','Mar','Apr','Mai','Jun','Jul','Aug','Sep','Okt','Nov','Des'];

const BUDGET_COLUMNS = {
  year: ['år','budsjettår','year'],
  version: ['versjon','version'],
  account: ['konto','kontonr','konto nr','account'],
  name: ['navn','tekst','tittel','beskrivelse','name'],
  costCenter: ['kostnadssted','ansvar','enhet','avdeling','costcenter'],
  project: ['prosjekt','anlegg','project'],
  vat: ['mva','mva_kode','mvakode','vat','mva kode'],
  type: ['type','art','kategori','category'],
  annual: ['årsbeløp','år','beløp år','annual'],
  month: ['måned','mnd','month'],
  amount: ['beløp','amount','sum'],
  active: ['aktiv','active','status'],
  comment: ['kommentar','notat','comment']
};

function importBudgetFromSheet(sourceSheetName, options = {}) {
  const res = { ok:false, rows:0, errors:[], warnings:[] };
  const linkAmounts = options.linkAmounts === true;

  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(sourceSheetName);
    if (!sheet) throw new Error(`Fant ikke ark: ${sourceSheetName}`);

    const values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) throw new Error(`Ingen data i ark: ${sourceSheetName}`);

    const header = (values.shift() || []).map(h => String(h || '').trim());
    const col = mapColumns_(header, BUDGET_COLUMNS);
    const isLong = !!(col.month && col.amount);
    const monthIndicesWide = detectMonthColumns_(header);

    if (!col.year || !col.account) throw new Error('Mangler obligatoriske kolonner: År og/eller Konto.');
    if (!isLong && Object.keys(monthIndicesWide).length === 0) throw new Error('Fant ingen månedskolonner (Jan–Des) og heller ikke Måned/Beløp.');

    const output = ensureBudgetOutputSheet_();
    if (options.clearOutput !== false) {
      output.clear();
      const hdr = ['År','Versjon','Konto','Navn','Kostnadssted','Prosjekt','MVA','Type','Måned','Beløp','Kommentar','Kildeark','Rad'];
      output.getRange(1,1,1,hdr.length).setValues([hdr]).setFontWeight('bold'); // 13 kolonner
      output.setFrozenRows(1);
    }

    const rowsOut = [];
    const belopFormulas = []; // for linkAmounts
    const colToA1_ = (n) => { let s=''; while (n>0){const r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26);} return s; };
    const escSheet_ = (name) => "'" + String(name).replace(/'/g, "''") + "'";

    let rIndex = 2; // visningsrad i kildesheetet
    for (const row of values) {
      const year = normalizeYear_(row[col.year-1]);
      if (!year) { res.warnings.push(`Rad ${rIndex}: Ugyldig/blankt År – hopper over.`); rIndex++; continue; }

      const account = String(row[col.account-1] ?? '').trim();
      if (!account) { res.warnings.push(`Rad ${rIndex}: Mangler Konto – hopper over.`); rIndex++; continue; }

      const version = col.version ? String(row[col.version-1] ?? '').trim() || DEFAULT_BUDGET_VERSION : DEFAULT_BUDGET_VERSION;
      const name = col.name ? String(row[col.name-1] ?? '').trim() : '';
      const cc = col.costCenter ? String(row[col.costCenter-1] ?? '').trim() : '';
      const proj = col.project ? String(row[col.project-1] ?? '').trim() : '';
      const vat = col.vat ? String(row[col.vat-1] ?? '').trim() : '';
      const type = col.type ? String(row[col.type-1] ?? '').trim() : '';
      const comment = col.comment ? String(row[col.comment-1] ?? '').trim() : '';
      const active = parseBool_(col.active ? row[col.active-1] : true);
      if (active === false) { rIndex++; continue; }

      if (isLong) {
        const m = normalizeMonth_(row[col.month-1]);
        const amt = toNumber_(row[col.amount-1]);
        if (!m) { res.warnings.push(`Rad ${rIndex}: Ugyldig Måned – hopper over.`); rIndex++; continue; }
        if (!isFinite(amt)) { res.warnings.push(`Rad ${rIndex}: Beløp er ikke tall – hopper over.`); rIndex++; continue; }

        const a1 = `${escSheet_(sourceSheetName)}!${colToA1_(col.amount)}${rIndex}`;
        belopFormulas.push(linkAmounts ? a1 : '');
        rowsOut.push([year, version, account, name, cc, proj, vat, type, m, amt, comment, sourceSheetName, rIndex]);

      } else {
        let foundAny = false;
        for (const m of MONTHS) {
          const idx = monthIndicesWide[m];
          if (!idx) continue;
          const amt = toNumber_(row[idx-1]);
          if ((isFinite(amt) && amt !== 0) || linkAmounts) {
            const a1 = `${escSheet_(sourceSheetName)}!${colToA1_(idx)}${rIndex}`;
            belopFormulas.push(linkAmounts ? a1 : '');
            rowsOut.push([year, version, account, name, cc, proj, vat, type, m, isFinite(amt) ? amt : '', comment, sourceSheetName, rIndex]);
            foundAny = foundAny || (isFinite(amt) && amt !== 0);
          }
        }
        if (!foundAny) {
          const annual = col.annual ? toNumber_(row[col.annual-1]) : NaN;
          if (isFinite(annual) && annual !== 0) {
            const per = Math.round((annual / 12) * 100) / 100;
            for (const m of MONTHS) {
              belopFormulas.push('');
              rowsOut.push([year, version, account, name, cc, proj, vat, type, m, per, comment, sourceSheetName, rIndex]);
            }
            res.warnings.push(`Rad ${rIndex}: Manglet månedstall – fordelte Årsbeløp jevnt (≈ ${per}).`);
          } else {
            res.warnings.push(`Rad ${rIndex}: Ingen månedstall og intet Årsbeløp – hoppet over.`);
          }
        }
      }
      rIndex++;
    }

    if (!rowsOut.length) throw new Error('Ingen gyldige budsjettlinjer å importere.');

    const startRow = output.getLastRow() + 1;
    output.getRange(startRow, 1, rowsOut.length, rowsOut[0].length).setValues(rowsOut);
    if (linkAmounts) {
      const formulas2D = belopFormulas.map(f => [f ? `=${f}` : '']);
      output.getRange(startRow, 10, formulas2D.length, 1).setFormulas(formulas2D);
    }

    res.ok = true; res.rows = rowsOut.length;
    return res;

  } catch (err) {
    res.errors.push(err.message);
    return res;
  }
}

// --- Helpers (import) ---
function ensureBudgetOutputSheet_() {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(BUDGET_OUTPUT_SHEET) || ss.insertSheet(BUDGET_OUTPUT_SHEET);
}
function mapColumns_(header, defs) {
  const lower = header.map(h => String(h || '').trim().toLowerCase());
  const out = {};
  for (const [key, aliases] of Object.entries(defs)) {
    let idx = 0;
    for (const a of aliases) {
      const p = lower.indexOf(String(a).toLowerCase());
      if (p >= 0) { idx = p+1; break; }
    }
    out[key] = idx;
  }
  return out;
}
function detectMonthColumns_(header) {
  const map = {};
  for (let i=0;i<header.length;i++) {
    const raw = String(header[i]||'').trim().toLowerCase();
    const norm = MONTH_ALIASES[raw] || MONTH_ALIASES[raw.replace(/\.$/,'')] || null;
    if (norm && MONTHS.includes(norm)) map[norm] = i+1;
    else {
      const direct = {jan:'Jan',feb:'Feb',mar:'Mar',apr:'Apr',mai:'Mai',jun:'Jun',jul:'Jul',aug:'Aug',sep:'Sep',okt:'Okt',nov:'Nov',des:'Des'}[raw];
      if (direct) map[direct] = i+1;
    }
  }
  return map;
}
function normalizeMonth_(val) {
  const s = String(val ?? '').trim().toLowerCase();
  if (!s) return '';
  if (MONTH_ALIASES[s]) return MONTH_ALIASES[s];
  const n = Number(s);
  if (Number.isInteger(n) && n >= 1 && n <= 12) return MONTHS[n-1];
  return '';
}
function normalizeYear_(val) {
  const n = Number(String(val ?? '').trim());
  return (Number.isInteger(n) && n >= 1900 && n <= 3000) ? n : 0;
}
function toNumber_(v) {
  if (v === '' || v === null || v === undefined) return NaN;
  let s = String(v).trim().replace(/\s/g,'');
  const hasComma = s.includes(','), hasDot = s.includes('.');
  if (hasComma && !hasDot) s = s.replace(/\./g,'').replace(',','.');
  else if (hasComma && hasDot && s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g,'').replace(',','.');
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}
function parseBool_(v) {
  if (typeof v === 'boolean') return v;
  const s = String(v ?? '').trim().toLowerCase();
  if (!s) return true;
  return !(['0','false','nei','no','n','inaktiv'].includes(s));
}
