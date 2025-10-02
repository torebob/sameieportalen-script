/* global Sameie */
Sameie.BudgetCore = (function () {
  function parseAmount(x) {
    if (x == null) return 0;
    var s = String(x).replace(/\s/g, '').replace(',', '.');
    var n = Number(s);
    return Number.isFinite(n) ? n : 0;
  }
  function ensureYear(year) {
    var y = Number(year);
    if (!Number.isInteger(y) || y < 2000) throw new Error('Ugyldig år: ' + year);
    return y;
  }
  function mapRow(row, cols) {
    var o = {}, i;
    for (i = 0; i < cols.length; i++) o[cols[i]] = row[i];
    return o;
  }
  function sum(rows, key) {
    var t = 0, i;
    for (i = 0; i < rows.length; i++) t += parseAmount(rows[i][key]);
    return t;
  }
  return { parseAmount: parseAmount, ensureYear: ensureYear, mapRow: mapRow, sum: sum };
})();
