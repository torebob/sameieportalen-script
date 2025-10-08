/* global Sameie, SpreadsheetApp */
function importBudget(sheetName, year) {
  var y = Sameie.BudgetCore.ensureYear(year);
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) throw new Error('Fant ikke ark: ' + sheetName);

  var data = sh.getDataRange().getValues();
  if (!data || data.length < 2) return [];

  var header = data[0];
  var rows = data.slice(1).map(function (row) {
    var r = Sameie.BudgetCore.mapRow(row, header);
    r.amount = Sameie.BudgetCore.parseAmount(r.amount || r.belop || r.sum);
    r.year = y;
    return r;
  });

  var total = Sameie.BudgetCore.sum(rows, 'amount');
  if (!(total >= 0)) throw new Error('Importfeil: total beløp NaN');

  return rows;
}
