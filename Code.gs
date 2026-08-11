function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Facturación')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function respond_(fn) {
  try {
    var result = fn();
    return { ok: true, data: result };
  } catch (err) {
    console.error(err);
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function runTests() {
  var results = [];
  var fns = Object.keys(this).filter(function (k) { return k.indexOf('test_') === 0; });
  fns.forEach(function (k) {
    try {
      this[k]();
      results.push(k + ': PASS');
    } catch (e) {
      results.push(k + ': FAIL - ' + e.message);
    }
  }, this);
  results.forEach(function (r) { Logger.log(r); });
  return results.join('\n');
}

function test_Aux_calcTotales() {
  var r = Aux.calcTotales([{ cantidad: 2, precio_unitario: 100 }, { cantidad: 1, precio_unitario: 50 }], 16);
  if (r.subtotal !== 250) throw new Error('subtotal mal: ' + r.subtotal);
  if (r.iva !== 40) throw new Error('iva mal: ' + r.iva);
  if (r.total !== 290) throw new Error('total mal: ' + r.total);
  if (Aux.requireFields({ a: 'x' }, ['a', 'b']).length !== 1) throw new Error('requireFields falló');
  if (Aux.parseCSV('a, b ,,c').length !== 3) throw new Error('parseCSV falló');
}

function test_ensureSheets() {
  Data.ensureSheets();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Data.SHEET_NAMES.forEach(function (n) {
    if (!ss.getSheetByName(n)) throw new Error('Falta hoja: ' + n);
  });
  if (Data.getSheet_('Config').getFrozenRows() !== 1) throw new Error('Config sin fila congelada');
}
