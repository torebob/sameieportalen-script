function include(filename, data) {
  var t = HtmlService.createTemplateFromFile(filename);
  if (data && typeof data === 'object') {
    for (var k in data) if (Object.prototype.hasOwnProperty.call(data, k)) t[k] = data[k];
  }
  return t.evaluate().getContent();
}
