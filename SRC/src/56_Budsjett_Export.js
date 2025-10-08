/* ====================== Budsjett – Export ======================
 * FILE: 56_Budsjett_Export.gs
 * VERSION: 1.0.0
 * UPDATED: 2025-09-15
 * ================================================================= */

function exportBudgetToCsv(year, version, fileName) {
  const y = Number(year); if (!Number.isInteger(y)) return { ok:false, error:'Ugyldig år' };
  const ver = String(version||'main');
  const res = getBudget(y, ver);
  if (!res.ok) return res;

  const header = ['År','Versjon','Konto','Navn','Kostnadssted','Prosjekt','MVA','Type','Måned','Beløp','Kommentar'];
  const lines = [header.join(';')];
  for (const it of res.items) {
    const row = [
      it.year, it.version, it.account, it.name, it.costCenter, it.project, it.vat, it.type, it.month,
      (Number(it.amount)||0).toString().replace('.',','), // norsk komma
      (it.comment||'').replace(/;/g, ',')
    ].map(v => typeof v === 'string' && v.includes(';') ? `"${v}"` : v);
    lines.push(row.join(';'));
  }

  const blob = Utilities.newBlob(lines.join('\n'), 'text/csv', (fileName || `budsjett_${y}_${ver}.csv`));
  const file = DriveApp.createFile(blob);
  return { ok:true, fileId: file.getId(), name: file.getName(), url: file.getUrl(), rows: res.items.length };
}
