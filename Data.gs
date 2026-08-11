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

function mergeFields_(base, obj) {
  var out = {};
  Object.keys(base).forEach(function (k) { out[k] = base[k]; });
  Object.keys(obj).forEach(function (k) { out[k] = obj[k]; });
  return out;
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
  mergeFields_: mergeFields_,
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
    var existing = readRows_(sh).filter(function (c) { return c.id_cliente === obj.id_cliente; })[0] || {};
    updateRow_(sh, row, mergeFields_(existing, obj), headers);
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

function nextFolio_() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var cfg = readConfig_();
    var n = Number(cfg.contador_folio) || 1;
    var folio = String(cfg.prefijo_folio || '') + String(n).padStart(3, '0');
    saveConfig_({ contador_folio: String(n + 1) });
    return folio;
  } finally {
    lock.releaseLock();
  }
}

function createFactura_(obj) {
  var cfg = readConfig_();
  var clienteSh = getSheet_('Clientes');
  var cRow = findRowBy_(clienteSh, 'id_cliente', obj.id_cliente, HEADERS.Clientes);
  if (cRow < 0) throw new Error('Cliente no existe');
  var cliente = readRows_(clienteSh).filter(function (c) { return c.id_cliente === obj.id_cliente; })[0];
  var items = obj.items || [];
  if (!items.length) throw new Error('La factura debe tener al menos un concepto');
  items.forEach(function (it) {
    if (Aux.requireFields(it, ['descripcion', 'cantidad', 'precio_unitario']).length) throw new Error('Concepto incompleto');
    if (isNaN(Number(it.cantidad)) || isNaN(Number(it.precio_unitario))) throw new Error('Cantidad/precio inválido');
    if (Number(it.cantidad) <= 0 || Number(it.precio_unitario) < 0) throw new Error('Cantidad/precio inválido');
  });
  var t = Aux.calcTotales(items, cfg.iva_porcentaje);
  var factura = {
    id_factura: Aux.uid('fac'),
    folio: nextFolio_(),
    id_cliente: obj.id_cliente,
    nombre_cliente: cliente.nombre,
    fecha_emision: obj.fecha_emision || new Date().toISOString().slice(0, 10),
    fecha_vencimiento: obj.fecha_vencimiento || '',
    subtotal: t.subtotal,
    iva: t.iva,
    total: t.total,
    estado: 'pendiente',
    fecha_pago: '',
    notas: obj.notas || ''
  };
  appendRow_(getSheet_('Facturas'), factura, HEADERS.Facturas);
  var itemsSh = getSheet_('Factura_Items');
  items.forEach(function (it) {
    appendRow_(itemsSh, {
      id_factura: factura.id_factura,
      descripcion: it.descripcion,
      cantidad: Number(it.cantidad),
      precio_unitario: Number(it.precio_unitario),
      importe: Aux.round2(Number(it.cantidad) * Number(it.precio_unitario))
    }, HEADERS.Factura_Items);
  });
  return factura;
}

function listFacturas_(filtro) {
  filtro = filtro || {};
  return readRows_(getSheet_('Facturas')).filter(function (f) {
    if (filtro.estado && f.estado !== filtro.estado) return false;
    if (filtro.mes && String(f.fecha_emision).slice(0, 7) !== filtro.mes) return false;
    return true;
  });
}

function getFactura_(id) {
  var sh = getSheet_('Facturas');
  var row = findRowBy_(sh, 'id_factura', id, HEADERS.Facturas);
  if (row < 0) throw new Error('Factura no encontrada');
  var factura = readRows_(sh).filter(function (f) { return f.id_factura === id; })[0];
  var items = readRows_(getSheet_('Factura_Items')).filter(function (it) { return it.id_factura === id; });
  return { factura: factura, items: items };
}

function setEstadoFactura_(id, estado) {
  if (['pendiente', 'pagada'].indexOf(estado) === -1) throw new Error('Estado inválido');
  var sh = getSheet_('Facturas');
  var row = findRowBy_(sh, 'id_factura', id, HEADERS.Facturas);
  if (row < 0) throw new Error('Factura no encontrada');
  var f = readRows_(sh).filter(function (x) { return x.id_factura === id; })[0];
  f.estado = estado;
  f.fecha_pago = estado === 'pagada' ? (f.fecha_pago || new Date().toISOString().slice(0, 10)) : '';
  updateRow_(sh, row, f, HEADERS.Facturas);
}

function deleteFactura_(id) {
  var sh = getSheet_('Facturas');
  deleteRow_(sh, findRowBy_(sh, 'id_factura', id, HEADERS.Facturas));
  var itemsSh = getSheet_('Factura_Items');
  var rows = readRows_(itemsSh);
  rows.forEach(function (r) {
    if (r.id_factura === id) deleteRow_(itemsSh, findRowBy_(itemsSh, 'id_factura', id, HEADERS.Factura_Items));
  });
}

Data.nextFolio_ = nextFolio_;
Data.createFactura_ = createFactura_;
Data.listFacturas_ = listFacturas_;
Data.getFactura_ = getFactura_;
Data.setEstadoFactura_ = setEstadoFactura_;
Data.deleteFactura_ = deleteFactura_;

function listGastos_(filtro) {
  filtro = filtro || {};
  return readRows_(getSheet_('Gastos')).filter(function (g) {
    if (filtro.mes && String(g.fecha).slice(0, 7) !== filtro.mes) return false;
    if (filtro.categoria && g.categoria !== filtro.categoria) return false;
    return true;
  });
}

function saveGasto_(obj) {
  var sh = getSheet_('Gastos');
  var headers = HEADERS.Gastos;
  if (obj.id_gasto) {
    var row = findRowBy_(sh, 'id_gasto', obj.id_gasto, headers);
    if (row < 0) throw new Error('Gasto no encontrado');
    var existing = readRows_(sh).filter(function (g) { return g.id_gasto === obj.id_gasto; })[0] || {};
    updateRow_(sh, row, mergeFields_(existing, obj), headers);
    return obj;
  }
  obj.id_gasto = Aux.uid('gas');
  obj.fecha = obj.fecha || new Date().toISOString().slice(0, 10);
  appendRow_(sh, obj, headers);
  return obj;
}

function deleteGasto_(id) {
  var sh = getSheet_('Gastos');
  deleteRow_(sh, findRowBy_(sh, 'id_gasto', id, HEADERS.Gastos));
}

Data.listGastos_ = listGastos_;
Data.saveGasto_ = saveGasto_;
Data.deleteGasto_ = deleteGasto_;

function getReportes_(mes) {
  var facturas = listFacturas_({ mes: mes });
  var gastos = listGastos_({ mes: mes });
  var facturado = 0, cobrado = 0, pendiente = 0;
  var porCliente = {};
  facturas.forEach(function (f) {
    var total = Number(f.total) || 0;
    facturado += total;
    if (f.estado === 'pagada') cobrado += total; else pendiente += total;
    porCliente[f.nombre_cliente] = (porCliente[f.nombre_cliente] || 0) + total;
  });
  var gastosTotal = 0, gastosPorCat = {};
  gastos.forEach(function (g) {
    var m = Number(g.monto) || 0;
    gastosTotal += m;
    var cat = g.categoria || 'Sin categoría';
    gastosPorCat[cat] = (gastosPorCat[cat] || 0) + m;
  });
  var top = Object.keys(porCliente).map(function (n) { return { nombre: n, total: Aux.round2(porCliente[n]) }; })
    .sort(function (a, b) { return b.total - a.total; }).slice(0, 5);
  return {
    facturado: Aux.round2(facturado),
    cobrado: Aux.round2(cobrado),
    pendiente: Aux.round2(pendiente),
    gastos_total: Aux.round2(gastosTotal),
    gastos_por_categoria: gastosPorCat,
    utilidad: Aux.round2(cobrado - gastosTotal),
    top_clientes: top
  };
}

Data.getReportes_ = getReportes_;
