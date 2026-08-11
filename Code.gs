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

function getConfig() {
  return respond_(function () { return Data.readConfig_(); });
}

function saveConfig(obj) {
  return respond_(function () {
    var requeridos = ['empresa_nombre', 'prefijo_folio'];
    var faltantes = Aux.requireFields(obj, requeridos);
    if (faltantes.length) throw new Error('Faltan campos: ' + faltantes.join(', '));
    if (!/^[0-9]+$/.test(String(obj.contador_folio))) throw new Error('contador_folio debe ser entero');
    if (isNaN(Number(obj.iva_porcentaje))) throw new Error('iva_porcentaje debe ser numérico');
    Data.saveConfig_(obj);
    return true;
  });
}

function test_Config_save() {
  var original = Data.readConfig_();
  var cfg = { empresa_nombre: 'Test SA', prefijo_folio: 'T-', contador_folio: '7', iva_porcentaje: '15' };
  var r = saveConfig(cfg);
  if (!r.ok) throw new Error(r.error);
  var out = Data.readConfig_();
  if (out.prefijo_folio !== 'T-') throw new Error('no guardó prefijo');
  var bad = saveConfig({ prefijo_folio: '' });
  if (bad.ok) throw new Error('debió fallar con prefijo vacío');
  var badNum = saveConfig({ contador_folio: 'abc' });
  if (badNum.ok) throw new Error('debió fallar con contador no numérico');
  saveConfig(original);
}

function listClientes() {
  return respond_(function () { return Data.listClientes_(); });
}

function saveCliente(obj) {
  return respond_(function () {
    var faltantes = Aux.requireFields(obj, ['nombre']);
    if (faltantes.length) throw new Error('El nombre del cliente es obligatorio');
    return Data.saveCliente_(obj);
  });
}

function deleteCliente(id) {
  return respond_(function () { Data.deleteCliente_(id); return true; });
}

function test_Clientes_CRUD() {
  var c = saveCliente({ nombre: 'Juan Test', rfc: 'JUA880101AAA', email: 'j@t.com' });
  if (!c.ok) throw new Error(c.error);
  var id = c.data.id_cliente;
  if (!id) throw new Error('sin id generado');
  var upd = saveCliente({ id_cliente: id, nombre: 'Juan Test 2', rfc: 'JUA880101AAA', email: 'j2@t.com' });
  if (!upd.ok) throw new Error(upd.error);
  var list = listClientes();
  if (!list.ok) throw new Error(list.error);
  var match = list.data.filter(function (x) { return x.id_cliente === id; });
  if (match.length !== 1 || match[0].nombre !== 'Juan Test 2') throw new Error('update falló');
  var del = deleteCliente(id);
  if (!del.ok) throw new Error(del.error);
  list = listClientes();
  if (list.data.some(function (x) { return x.id_cliente === id; })) throw new Error('no eliminó');
}
