/* global Sameie */
Sameie = typeof Sameie === 'object' ? Sameie : {};
Sameie.Num = Sameie.Num || {};
Sameie.Num.parseNordicNumber = function (v) {
  var s = String(v == null ? '' : v).trim().replace(/\s/g,'');
  var hasC = s.indexOf(',') >= 0, hasD = s.indexOf('.') >= 0;
  if (hasC && !hasD) s = s.replace(/\./g,'').replace(',', '.');
  else if (hasC && hasD && s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g,'').replace(',', '.');
  var n = Number(s);
  return isNaN(n) ? NaN : n;
};
