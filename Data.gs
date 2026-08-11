var SHEET_NAMES = ['Config', 'Clientes', 'Facturas', 'Factura_Items', 'Gastos'];

var HEADERS = {
  Config: ['clave', 'valor'],
  Clientes: ['id_cliente', 'nombre', 'rfc', 'email', 'telefono', 'direccion', 'fecha_registro'],
  Facturas: ['id_factura', 'folio', 'id_cliente', 'nombre_cliente', 'fecha_emision', 'fecha_vencimiento', 'subtotal', 'iva', 'total', 'estado', 'fecha_pago', 'notas'],
  Factura_Items: ['id_factura', 'descripcion', 'cantidad', 'precio_unitario', 'importe'],
  Gastos: ['id_gasto', 'fecha', 'categoria', 'descripcion', 'monto', 'metodo_pago', 'proveedor']
};

function getSheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADERS[name]);
    sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureSheets() {
  SHEET_NAMES.forEach(getSheet_);
  seedDefaultConfig_();
}

function seedDefaultConfig_() {
  var sh = getSheet_('Config');
  var existing = readRows_(sh).reduce(function (m, r) { m[r.clave] = r.valor; return m; }, {});
  var defaults = {
    empresa_nombre: 'Mi Empresa S.A.',
    empresa_rfc: 'XAXX010101000',
    empresa_direccion: '',
    empresa_telefono: '',
    empresa_email: '',
    empresa_logo: '',
    moneda: 'USD',
    moneda_simbolo: '$',
    prefijo_folio: 'FAC-',
    contador_folio: '1',
    iva_porcentaje: '16',
    categorias_gastos: 'Renta,Internet,Papelería,Servicios'
  };
  Object.keys(defaults).forEach(function (clave) {
    if (existing[clave] === undefined) {
      sh.appendRow([clave, defaults[clave]]);
    }
  });
}

function readRows_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return values.map(function (row) {
    var o = {};
    headers.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
}

function appendRow_(sheet, obj, headers) {
  sheet.appendRow(headers.map(function (h) { return obj[h] === undefined ? '' : obj[h]; }));
}

function findRowBy_(sheet, idField, id, headers) {
  var col = headers.indexOf(idField) + 1;
  if (col === 0) return -1;
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var values = sheet.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function updateRow_(sheet, row, obj, headers) {
  var values = headers.map(function (h) { return obj[h] === undefined ? '' : obj[h]; });
  sheet.getRange(row, 1, 1, headers.length).setValues([values]);
}

function deleteRow_(sheet, row) {
  if (row > 0) sheet.deleteRow(row);
}

function readConfig_() {
  return readRows_(getSheet_('Config')).reduce(function (m, r) { m[r.clave] = r.valor; return m; }, {});
}

function saveConfig_(obj) {
  var sh = getSheet_('Config');
  var headers = HEADERS.Config;
  var allowed = readRows_(sh).map(function (r) { return r.clave; });
  Object.keys(obj).forEach(function (clave) {
    if (allowed.indexOf(clave) === -1) return;
    var row = findRowBy_(sh, 'clave', clave, headers);
    updateRow_(sh, row, { clave: clave, valor: String(obj[clave]) }, headers);
  });
}

var Data = {
  SHEET_NAMES: SHEET_NAMES,
  HEADERS: HEADERS,
  getSheet_: getSheet_,
  ensureSheets: ensureSheets,
  seedDefaultConfig_: seedDefaultConfig_,
  readRows_: readRows_,
  appendRow_: appendRow_,
  findRowBy_: findRowBy_,
  updateRow_: updateRow_,
  deleteRow_: deleteRow_
};

Data.readConfig_ = readConfig_;
Data.saveConfig_ = saveConfig_;

function listClientes_() {
  return readRows_(getSheet_('Clientes'));
}

function saveCliente_(obj) {
  var sh = getSheet_('Clientes');
  var headers = HEADERS.Clientes;
  if (obj.id_cliente) {
    var row = findRowBy_(sh, 'id_cliente', obj.id_cliente, headers);
    if (row < 0) throw new Error('Cliente no encontrado');
    updateRow_(sh, row, obj, headers);
    return obj;
  }
  obj.id_cliente = Aux.uid('cli');
  obj.fecha_registro = obj.fecha_registro || new Date().toISOString().slice(0, 10);
  appendRow_(sh, obj, headers);
  return obj;
}

function hasInvoices_(clienteId) {
  var sh = getSheet_('Facturas');
  return findRowBy_(sh, 'id_cliente', clienteId, HEADERS.Facturas) > 0;
}

function deleteCliente_(id) {
  if (hasInvoices_(id)) throw new Error('No se puede eliminar: el cliente tiene facturas');
  var sh = getSheet_('Clientes');
  deleteRow_(sh, findRowBy_(sh, 'id_cliente', id, HEADERS.Clientes));
}

Data.listClientes_ = listClientes_;
Data.saveCliente_ = saveCliente_;
Data.hasInvoices_ = hasInvoices_;
Data.deleteCliente_ = deleteCliente_;
