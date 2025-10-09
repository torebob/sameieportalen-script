/* global Sameie */
Sameie = typeof Sameie === 'object' ? Sameie : {};
Sameie.Sheets = Sameie.Sheets || {};
Sameie.Sheets.ensureHeader = function (sh, headers) {
  var cur = (sh.getLastRow() > 0)
    ? sh.getRange(1,1,1,Math.max(headers.length, sh.getLastColumn())).getValues()[0]
    : [];
  var mismatch = JSON.stringify(cur) !== JSON.stringify(headers);
  if (sh.getLastRow() === 0 || mismatch) {
    sh.getRange(1,1,1,Math.max(headers.length, sh.getLastColumn())).clearContent();
    sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
    if (sh.getFrozenRows() < 1) sh.setFrozenRows(1);
  }
  return sh;
};
