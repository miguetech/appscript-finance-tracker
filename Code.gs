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

function createFactura(obj) {
  return respond_(function () { return Data.createFactura_(obj); });
}

function listFacturas(filtro) {
  return respond_(function () { return Data.listFacturas_(filtro); });
}

function getFactura(id) {
  return respond_(function () { return Data.getFactura_(id); });
}

function setEstadoFactura(id, estado) {
  return respond_(function () { Data.setEstadoFactura_(id, estado); return true; });
}

function deleteFactura(id) {
  return respond_(function () { Data.deleteFactura_(id); return true; });
}

function test_Facturas_CRUD() {
  var c = saveCliente({ nombre: 'Fiscal Test' });
  var cid = c.data.id_cliente;
  var cfg = Data.readConfig_();
  var startFolio = Number(cfg.contador_folio);
  var r = createFactura({
    id_cliente: cid,
    items: [{ descripcion: 'Servicio A', cantidad: 2, precio_unitario: 100 }, { descripcion: 'Servicio B', cantidad: 1, precio_unitario: 50 }]
  });
  if (!r.ok) throw new Error(r.error);
  var f = r.data;
  if (!f.folio) throw new Error('sin folio');
  if (f.estado !== 'pendiente') throw new Error('estado inicial mal');
  var cfg2 = Data.readConfig_();
  if (Number(cfg2.contador_folio) !== startFolio + 1) throw new Error('folio no incrementó');
  var g = getFactura(f.id_factura);
  if (!g.ok || g.data.items.length !== 2) throw new Error('items no guardados');
  if (g.data.factura.total !== 290) throw new Error('total mal: ' + g.data.factura.total);
  var mark = setEstadoFactura(f.id_factura, 'pagada');
  if (!mark.ok) throw new Error(mark.error);
  var list = listFacturas({ estado: 'pagada' });
  if (!list.data.some(function (x) { return x.id_factura === f.id_factura; })) throw new Error('no lista pagadas');
  if (listFacturas({ estado: 'pendiente' }).data.some(function (x) { return x.id_factura === f.id_factura; })) throw new Error('sigue pendiente');
  deleteFactura(f.id_factura);
  if (getFactura(f.id_factura).ok) throw new Error('no eliminó factura');
  var items = Data.readRows_(Data.getSheet_('Factura_Items'));
  if (items.some(function (it) { return it.id_factura === f.id_factura; })) throw new Error('no eliminó items');
  deleteCliente(cid);
}

function listGastos(filtro) {
  return respond_(function () { return Data.listGastos_(filtro); });
}

function saveGasto(obj) {
  return respond_(function () {
    var faltantes = Aux.requireFields(obj, ['descripcion', 'monto']);
    if (faltantes.length) throw new Error('Faltan campos: ' + faltantes.join(', '));
    if (Number(obj.monto) < 0) throw new Error('Monto inválido');
    return Data.saveGasto_(obj);
  });
}

function deleteGasto(id) {
  return respond_(function () { Data.deleteGasto_(id); return true; });
}

function test_Gastos_CRUD() {
  var g = saveGasto({ fecha: '2026-08-05', categoria: 'Renta', descripcion: 'Oficina', monto: 1500.5, metodo_pago: 'transferencia', proveedor: 'Inmobiliaria X' });
  if (!g.ok) throw new Error(g.error);
  var id = g.data.id_gasto;
  var upd = saveGasto({ id_gasto: id, fecha: '2026-08-05', categoria: 'Renta', descripcion: 'Oficina mes', monto: 1600, metodo_pago: 'transferencia', proveedor: 'Inmobiliaria X' });
  if (!upd.ok) throw new Error(upd.error);
  var list = listGastos({ mes: '2026-08', categoria: 'Renta' });
  if (list.data.length !== 1 || list.data[0].monto !== 1600) throw new Error('listado/update falló');
  if (listGastos({ mes: '2027-01' }).data.length !== 0) throw new Error('filtro mes falló');
  deleteGasto(id);
  if (listGastos({}).data.some(function (x) { return x.id_gasto === id; })) throw new Error('no eliminó');
}

function getReportes(mes) {
  return respond_(function () { return Data.getReportes_(mes); });
}

function test_Reportes() {
  var c = saveCliente({ nombre: 'Rep Test' });
  var cid = c.data.id_cliente;
  var f = createFactura({ id_cliente: cid, items: [{ descripcion: 'X', cantidad: 1, precio_unitario: 1000 }] }).data;
  setEstadoFactura(f.id_factura, 'pagada');
  saveGasto({ fecha: f.fecha_emision, categoria: 'Renta', descripcion: 'oficina', monto: 200 });
  var mes = String(f.fecha_emision).slice(0, 7);
  var r = getReportes(mes);
  if (!r.ok) throw new Error(r.error);
  var d = r.data;
  if (d.facturado !== 1160) throw new Error('facturado mal: ' + d.facturado);
  if (d.cobrado !== 1160 || d.pendiente !== 0) throw new Error('cobrado/pendiente mal');
  if (d.gastos_total !== 200) throw new Error('gastos mal');
  if (d.utilidad !== 960) throw new Error('utilidad mal: ' + d.utilidad);
  if (d.top_clientes.length !== 1 || d.top_clientes[0].total !== 1160) throw new Error('top clientes mal');
  deleteFactura(f.id_factura);
  deleteCliente(cid);
  var gastos = Data.listGastos_({});
  gastos.forEach(function (g) { Data.deleteGasto_(g.id_gasto); });
}
