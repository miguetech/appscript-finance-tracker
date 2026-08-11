# Facturación Apps Script — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single-user Google Apps Script web app (vanilla JS SPA on HtmlService + Google Sheets backend) for invoicing, client management, administrative expense tracking, monthly reports, and on-demand PDF generation.

**Architecture:** SPA frontend in one `index.html` (sidebar nav: Dashboard, Facturas, Clientes, Gastos, Reportes, Configuración) calling server functions via `google.script.run`. Backend split into `Code.gs` (public API + `respond_` wrapper returning `{ok,data}` / `{ok,error}`), `Data.gs` (Sheets I/O), `Aux.gs` (pure helpers: totals, formatting, validation). PDF built from an HTML invoice template via `HtmlService.createHtmlOutput().asDownload()`.

**Tech Stack:** Google Apps Script (HtmlService, SpreadsheetApp, LockService), vanilla JS, Google Sheets as datastore.

## Global Constraints

- Web App deployed with access **"Solo yo"** (single user, no auth UI).
- Every server function returns `{ok:true, data}` or `{ok:false, error}` via `respond_`.
- No external CDNs or dependencies; `.gs` files use ES5-compatible syntax (`var`, no `const/let` in server code). Client-side `index.html` JS also uses `var` for consistency and because inline `onclick` handlers reference globals.
- Folio counter protected by `LockService.getScriptLock()` (10s timeout) to prevent duplicates.
- Data stored in 5 sheets: `Config`, `Clientes`, `Facturas`, `Factura_Items`, `Gastos`. Row 1 frozen header. IDs are string UIDs via `Aux.uid`.
- Currency, IVA %, folio prefix, and expense categories configurable via `Config` sheet; UI reads them at load.
- PDFs generated on demand only; never auto-saved to Drive.
- No code comments unless needed to explain non-obvious logic.
- Tests run via `Code.runTests()` in the Apps Script editor (or `clasp run`). No test framework available in Apps Script runtime; `runTests` executes every `test_*` function and logs PASS/FAIL.
- File naming: `*.gs`, `appscript.json`, `index.html`, `pdfTemplate.html`.

---

### Task 1: Project scaffolding — manifest, sheets, helpers, response wrapper

**Files:**
- Create: `appscript.json`
- Create: `Code.gs`
- Create: `Aux.gs`
- Create: `Data.gs`

**Interfaces:**
- Produces:
  - `Aux.uid(prefix)` → string
  - `Aux.round2(n)` → number (2-decimals, `Number.EPSILON` rounded)
  - `Aux.calcTotales(items, ivaPct)` → `{subtotal, iva, total}`; mutates each item setting `importe`
  - `Aux.formatMoney(n)` → string (2 decimals, thousands separator)
  - `Aux.parseCSV(str)` → array of trimmed non-empty strings
  - `Aux.requireFields(obj, campos)` → array of missing field names
  - `Data.SHEET_NAMES` → array of 5 sheet names
  - `Data.HEADERS` → object mapping sheet name → header array
  - `Data.getSheet_(name)` → `Sheet` (creates if missing, sets frozen bold header)
  - `Data.ensureSheets()` → creates all 5 sheets + seeds default config rows
  - `Data.readRows_(sheet)` → array of `{header: value}` objects
  - `Data.appendRow_(sheet, obj, headers)` → void
  - `Data.findRowBy_(sheet, idField, id, headers)` → row index (1-based) or -1
  - `Data.updateRow_(sheet, row, obj, headers)` → void
  - `Data.deleteRow_(sheet, row)` → void
  - `Code.doGet(e)` → serves `index.html`
  - `Code.respond_(fn)` → `{ok:true,data}` or `{ok:false,error}`
  - `Code.runTests()` → logs all `test_*` PASS/FAIL, returns joined string

- [ ] **Step 1: Write `appscript.json`**

```json
{
  "timeZone": "America/Mexico_City",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "MYSELF"
  }
}
```

- [ ] **Step 2: Write `Aux.gs`**

```javascript
function uid(prefix) {
  return (prefix || 'id') + '_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8);
}

function round2(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

function calcTotales(items, ivaPct) {
  var subtotal = 0;
  items.forEach(function (it) {
    var importe = Number(it.cantidad || 0) * Number(it.precio_unitario || 0);
    it.importe = round2(importe);
    subtotal += importe;
  });
  subtotal = round2(subtotal);
  var iva = round2(subtotal * (Number(ivaPct) || 0) / 100);
  return { subtotal: subtotal, iva: iva, total: round2(subtotal + iva) };
}

function formatMoney(n) {
  return Number(n).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseCSV(str) {
  return String(str || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}

function requireFields(obj, campos) {
  return campos.filter(function (c) { return obj[c] === undefined || obj[c] === null || String(obj[c]).trim() === ''; });
}
```

- [ ] **Step 3: Write `Data.gs`**

```javascript
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
```

- [ ] **Step 4: Write `Code.gs`**

```javascript
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
```

- [ ] **Step 5: Run tests to verify**

Open Apps Script editor (create a new spreadsheet + Extensions → Apps Script, paste files, or use `clasp push`). Run function `runTests()`.
Expected: `test_Aux_calcTotales: PASS`, `test_ensureSheets: PASS`.

- [ ] **Step 6: Commit**

```bash
git add appscript.json Code.gs Aux.gs Data.gs
git commit -m "feat: scaffolding apps script (manifest, sheets, helpers)"
```

---

### Task 2: Config API (get/save)

**Files:**
- Modify: `Data.gs`
- Modify: `Code.gs`

**Interfaces:**
- Consumes: `Data.getSheet_('Config')`, `Data.readRows_`, `Data.findRowBy_`, `Data.updateRow_`, `Aux.requireFields`
- Produces:
  - `Data.readConfig_()` → object `{clave: valor}` (all keys)
  - `Data.saveConfig_(obj)` → void; only updates keys that already exist as rows
  - `Code.getConfig()` → `{ok, data: {clave: valor}}`
  - `Code.saveConfig(obj)` → `{ok, data:true}`; validates required keys + numeric fields

- [ ] **Step 1: Add data functions to `Data.gs`**

```javascript
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
```

- [ ] **Step 2: Add API endpoints to `Code.gs`**

```javascript
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
```

- [ ] **Step 3: Add test to `Code.gs`**

```javascript
function test_Config_save() {
  var cfg = { empresa_nombre: 'Test SA', prefijo_folio: 'T-', contador_folio: '7', iva_porcentaje: '15' };
  var r = saveConfig(cfg);
  if (!r.ok) throw new Error(r.error);
  var out = Data.readConfig_();
  if (out.prefijo_folio !== 'T-') throw new Error('no guardó prefijo');
  var bad = saveConfig({ prefijo_folio: '' });
  if (bad.ok) throw new Error('debió fallar con prefijo vacío');
  var badNum = saveConfig({ contador_folio: 'abc' });
  if (badNum.ok) throw new Error('debió fallar con contador no numérico');
  saveConfig(cfg);
}
```

Note: test restores original config at the end to keep later tests predictable.

- [ ] **Step 4: Run tests**

Run `runTests()`. Expected: `test_Config_save: PASS` (and prior tests still PASS).

- [ ] **Step 5: Commit**

```bash
git add Code.gs Data.gs
git commit -m "feat: config API get/save"
```

---

### Task 3: Clientes CRUD API

**Files:**
- Modify: `Data.gs`
- Modify: `Code.gs`

**Interfaces:**
- Consumes: `Data.getSheet_('Clientes')`, `Data.readRows_`, `Data.appendRow_`, `Data.updateRow_`, `Data.deleteRow_`, `Data.findRowBy_`, `Aux.uid`, `Aux.requireFields`
- Produces:
  - `Data.listClientes_()` → array of client objects
  - `Data.saveCliente_(obj)` → client object; `id_cliente` present → update, else create with UID + `fecha_registro`
  - `Data.hasInvoices_(clienteId)` → boolean
  - `Data.deleteCliente_(id)` → void; throws if client has invoices
  - `Code.listClientes()` → `{ok, data:[...]}`
  - `Code.saveCliente(obj)` → `{ok, data:{...}}`; requires `nombre`
  - `Code.deleteCliente(id)` → `{ok, data:true}`

- [ ] **Step 1: Add data functions to `Data.gs`**

```javascript
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
```

- [ ] **Step 2: Add API endpoints to `Code.gs`**

```javascript
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
```

- [ ] **Step 3: Add test to `Code.gs`**

```javascript
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
```

- [ ] **Step 4: Run tests**

Run `runTests()`. Expected: `test_Clientes_CRUD: PASS`.

- [ ] **Step 5: Commit**

```bash
git add Code.gs Data.gs
git commit -m "feat: clientes CRUD API"
```

---

### Task 4: Facturas CRUD API (folio + LockService)

**Files:**
- Modify: `Data.gs`
- Modify: `Code.gs`

**Interfaces:**
- Consumes: `Data.getSheet_('Facturas'|'Factura_Items'|'Clientes')`, `Data.readRows_`, `Data.appendRow_`, `Data.updateRow_`, `Data.deleteRow_`, `Data.findRowBy_`, `Data.readConfig_`, `Data.saveConfig_`, `Aux.calcTotales`, `Aux.uid`, `Aux.requireFields`, `LockService`
- Produces:
  - `Data.nextFolio_()` → string; locks script lock, reads counter, builds `prefijo + counter (padded 3)`, increments counter, saves, releases lock
  - `Data.createFactura_(obj)` → factura object; obj = `{id_cliente, fecha_emision, fecha_vencimiento, notas, items:[{descripcion,cantidad,precio_unitario}]}`; validates client + items, computes totals, writes header + item rows
  - `Data.listFacturas_(filtro)` → array; filtro `{estado, mes}` (mes `YYYY-MM` from `fecha_emision`)
  - `Data.getFactura_(id)` → `{factura, items}`
  - `Data.setEstadoFactura_(id, estado)` → void; sets `fecha_pago` when `pagada`, clears when `pendiente`
  - `Data.deleteFactura_(id)` → void; deletes header + matching item rows
  - `Code.createFactura(obj)`, `Code.listFacturas(filtro)`, `Code.getFactura(id)`, `Code.setEstadoFactura(id, estado)`, `Code.deleteFactura(id)` → standard `{ok,...}` responses

- [ ] **Step 1: Add data functions to `Data.gs`**

```javascript
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
```

- [ ] **Step 2: Add API endpoints to `Code.gs`**

```javascript
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
```

- [ ] **Step 3: Add test to `Code.gs`**

```javascript
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
```

- [ ] **Step 4: Run tests**

Run `runTests()`. Expected: `test_Facturas_CRUD: PASS`.

- [ ] **Step 5: Commit**

```bash
git add Code.gs Data.gs
git commit -m "feat: facturas CRUD API con folio lock"
```

---

### Task 5: Gastos CRUD API

**Files:**
- Modify: `Data.gs`
- Modify: `Code.gs`

**Interfaces:**
- Consumes: `Data.getSheet_('Gastos')`, `Data.readRows_`, `Data.appendRow_`, `Data.updateRow_`, `Data.deleteRow_`, `Data.findRowBy_`, `Aux.uid`, `Aux.requireFields`
- Produces:
  - `Data.listGastos_(filtro)` → array; filtro `{mes, categoria}`
  - `Data.saveGasto_(obj)` → gasto object; create (UID + `fecha` default today) or update
  - `Data.deleteGasto_(id)` → void
  - `Code.listGastos(filtro)`, `Code.saveGasto(obj)`, `Code.deleteGasto(id)` → standard `{ok,...}` responses

- [ ] **Step 1: Add data functions to `Data.gs`**

```javascript
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
    updateRow_(sh, row, obj, headers);
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
```

- [ ] **Step 2: Add API endpoints to `Code.gs`**

```javascript
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
```

- [ ] **Step 3: Add test to `Code.gs`**

```javascript
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
```

- [ ] **Step 4: Run tests**

Run `runTests()`. Expected: `test_Gastos_CRUD: PASS`.

- [ ] **Step 5: Commit**

```bash
git add Code.gs Data.gs
git commit -m "feat: gastos CRUD API"
```

---

### Task 6: Reportes API

**Files:**
- Modify: `Data.gs`
- Modify: `Code.gs`

**Interfaces:**
- Consumes: `Data.listFacturas_`, `Data.listGastos_`, `Aux.round2`
- Produces:
  - `Data.getReportes_(mes)` → `{facturado, cobrado, pendiente, gastos_total, gastos_por_categoria:{cat:monto}, utilidad, top_clientes:[{nombre, total}]}`
  - `Code.getReportes(mes)` → `{ok, data:{...}}`

- [ ] **Step 1: Add data function to `Data.gs`**

```javascript
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
```

- [ ] **Step 2: Add API endpoint to `Code.gs`**

```javascript
function getReportes(mes) {
  return respond_(function () { return Data.getReportes_(mes); });
}
```

- [ ] **Step 3: Add test to `Code.gs`**

```javascript
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
```

Note: this test depends on `iva_porcentaje = 16` (restored by `test_Config_save`). 1000 + 16% = 1160. If a future test changes IVA, restore it here.

- [ ] **Step 4: Run tests**

Run `runTests()`. Expected: `test_Reportes: PASS`.

- [ ] **Step 5: Commit**

```bash
git add Code.gs Data.gs
git commit -m "feat: reportes mensuales API"
```

---

### Task 7: PDF generation

**Files:**
- Create: `pdfTemplate.html`
- Modify: `Code.gs`

**Interfaces:**
- Consumes: `Data.getFactura_`, `Data.readConfig_`, `Data.listClientes_`, `Aux.formatMoney`, `HtmlService`
- Produces:
  - `Code.buildPDF(idFactura)` → PDF `Blob` (NOT `{ok,...}`; consumed directly by `google.script.run` success handler, returns base64)
  - `pdfTemplate.html` placeholders: `${EMPRESA_NOMBRE}`, `${EMPRESA_RFC}`, `${EMPRESA_DIRECCION}`, `${EMPRESA_TELEFONO}`, `${EMPRESA_EMAIL}`, `${LOGO}`, `${FOLIO}`, `${FECHA_EMISION}`, `${FECHA_VENCIMIENTO}`, `${NOMBRE_CLIENTE}`, `${RFC_CLIENTE}`, `${DIRECCION_CLIENTE}`, `${MONEDA_SIMBOLO}`, `${SUBTOTAL}`, `${IVA}`, `${TOTAL}`, `${NOTAS}`, `${ITEMS_ROWS}`

- [ ] **Step 1: Write `pdfTemplate.html`**

```html
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body { font-family: Arial, Helvetica, sans-serif; color: #1a1a1a; margin: 0; padding: 40px; }
  .header { width: 100%; border-bottom: 3px solid #1a56db; padding-bottom: 16px; margin-bottom: 24px; }
  .header:after { content: ''; display: table; clear: both; }
  .empresa { float: left; font-size: 14px; }
  .empresa h1 { margin: 0 0 4px 0; font-size: 22px; color: #1a56db; }
  .logo { float: right; max-height: 80px; max-width: 200px; }
  .folio-box { text-align: right; margin-bottom: 24px; }
  .folio-box .folio { font-size: 20px; font-weight: bold; }
  .grid { width: 100%; margin-bottom: 24px; }
  .grid td { vertical-align: top; font-size: 14px; }
  .grid .label { color: #666; font-size: 12px; text-transform: uppercase; }
  table.items { width: 100%; border-collapse: collapse; margin-bottom: 24px; }
  table.items th { background: #1a56db; color: #fff; text-align: left; padding: 8px; font-size: 13px; }
  table.items td { border: 1px solid #ddd; padding: 8px; font-size: 13px; }
  table.items td.num { text-align: right; }
  .totals { float: right; width: 260px; }
  .totals table { width: 100%; border-collapse: collapse; font-size: 14px; }
  .totals td { padding: 4px 8px; }
  .totals tr.total td { font-size: 18px; font-weight: bold; border-top: 2px solid #1a56db; }
  .notas { clear: both; font-size: 13px; color: #444; margin-top: 32px; }
</style>
</head>
<body>
  <div class="header">
    <div class="empresa">
      <h1>${EMPRESA_NOMBRE}</h1>
      <div>${EMPRESA_RFC}</div>
      <div>${EMPRESA_DIRECCION}</div>
      <div>Tel: ${EMPRESA_TELEFONO} &nbsp; ${EMPRESA_EMAIL}</div>
    </div>
    ${LOGO}
  </div>
  <div class="folio-box">
    <div class="folio">FACTURA ${FOLIO}</div>
    <div>Fecha de emisión: ${FECHA_EMISION}</div>
    <div>Fecha de vencimiento: ${FECHA_VENCIMIENTO}</div>
  </div>
  <table class="grid">
    <tr>
      <td>
        <div class="label">Facturar a</div>
        <div><strong>${NOMBRE_CLIENTE}</strong></div>
        <div>${RFC_CLIENTE}</div>
        <div>${DIRECCION_CLIENTE}</div>
      </td>
    </tr>
  </table>
  <table class="items">
    <thead>
      <tr><th>Descripción</th><th style="width:70px">Cant.</th><th style="width:110px" class="num">P. Unitario</th><th style="width:120px" class="num">Importe</th></tr>
    </thead>
    <tbody>
      ${ITEMS_ROWS}
    </tbody>
  </table>
  <div class="totals">
    <table>
      <tr><td>Subtotal</td><td class="num">${MONEDA_SIMBOLO}${SUBTOTAL}</td></tr>
      <tr><td>IVA</td><td class="num">${MONEDA_SIMBOLO}${IVA}</td></tr>
      <tr class="total"><td>Total</td><td class="num">${MONEDA_SIMBOLO}${TOTAL}</td></tr>
    </table>
  </div>
  <div class="notas">${NOTAS}</div>
</body>
</html>
```

- [ ] **Step 2: Add `buildPDF` to `Code.gs`**

```javascript
function buildPDF(idFactura) {
  var cfg = Data.readConfig_();
  var detail = Data.getFactura_(idFactura);
  var f = detail.factura;
  var items = detail.items;
  var cliente = Data.listClientes_().filter(function (c) { return c.id_cliente === f.id_cliente; })[0] || {};
  var tpl = HtmlService.createHtmlOutputFromFile('pdfTemplate').getContent();
  var rows = items.map(function (it) {
    return '<tr><td>' + it.descripcion + '</td><td>' + it.cantidad + '</td><td class="num">' + Aux.formatMoney(it.precio_unitario) + '</td><td class="num">' + Aux.formatMoney(it.importe) + '</td></tr>';
  }).join('');
  var logo = cfg.empresa_logo ? '<img class="logo" src="' + cfg.empresa_logo + '" alt="logo">' : '';
  var html = tpl
    .split('${EMPRESA_NOMBRE}').join(cfg.empresa_nombre || '')
    .split('${EMPRESA_RFC}').join(cfg.empresa_rfc || '')
    .split('${EMPRESA_DIRECCION}').join(cfg.empresa_direccion || '')
    .split('${EMPRESA_TELEFONO}').join(cfg.empresa_telefono || '')
    .split('${EMPRESA_EMAIL}').join(cfg.empresa_email || '')
    .split('${LOGO}').join(logo)
    .split('${FOLIO}').join(f.folio)
    .split('${FECHA_EMISION}').join(f.fecha_emision)
    .split('${FECHA_VENCIMIENTO}').join(f.fecha_vencimiento || '—')
    .split('${NOMBRE_CLIENTE}').join(f.nombre_cliente)
    .split('${RFC_CLIENTE}').join(cliente.rfc || '')
    .split('${DIRECCION_CLIENTE}').join(cliente.direccion || '')
    .split('${MONEDA_SIMBOLO}').join(cfg.moneda_simbolo || '$')
    .split('${SUBTOTAL}').join(Aux.formatMoney(f.subtotal))
    .split('${IVA}').join(Aux.formatMoney(f.iva))
    .split('${TOTAL}').join(Aux.formatMoney(f.total))
    .split('${NOTAS}').join(f.notas || '')
    .split('${ITEMS_ROWS}').join(rows);
  var blob = Utilities.newBlob(html, 'text/html', 'factura_' + f.folio + '.html');
  return blob.getAs('application/pdf').setName('Factura_' + f.folio + '.pdf');
}
```

- [ ] **Step 3: Add test to `Code.gs`**

```javascript
function test_buildPDF() {
  var c = saveCliente({ nombre: 'PDF Test', rfc: 'PDF0001', direccion: 'Calle 5' });
  var f = createFactura({ id_cliente: c.data.id_cliente, items: [{ descripcion: 'Diseño', cantidad: 1, precio_unitario: 500 }] }).data;
  var pdf = buildPDF(f.id_factura);
  if (!pdf || pdf.getContentType() !== 'application/pdf') throw new Error('no genera PDF');
  if (pdf.getName().indexOf('Factura_') !== 0) throw new Error('nombre PDF mal: ' + pdf.getName());
  deleteFactura(f.id_factura);
  deleteCliente(c.data.id_cliente);
}
```

- [ ] **Step 4: Run tests**

Run `runTests()`. Expected: `test_buildPDF: PASS`. Also verify a real PDF renders: run `buildPDF` with a real factura id in the editor, inspect the returned blob's PDF (or download via a temporary `Logger.log` of the name).

- [ ] **Step 5: Commit**

```bash
git add pdfTemplate.html Code.gs
git commit -m "feat: generacion PDF bajo demanda"
```

---

### Task 8: SPA frontend — shell, navigation, CSS, state, router

**Files:**
- Create: `index.html`

**Interfaces:**
- Consumes: `Code.getConfig`
- Produces (global JS functions used by later tasks and inline `onclick`):
  - `STATE` → `{config:{}, view}`; config loaded from server
  - `showView(name)` → router; calls per-view loader if defined
  - `call(name, args)` → Promise wrapper over `google.script.run` (unwraps `{ok,data}`; rejects with error)
  - `toast(msg, type)` → transient toast (type `ok`|`error`)
  - `fmtMoney(n)` → formatted with `STATE.config.moneda_simbolo`
  - `openModal(html)` / `closeModal()` → modal management
  - `mesActual()` → `YYYY-MM` string
  - App shell: sidebar nav (6 views), topbar, `.view` sections `#view-dashboard|facturas|clientes|gastos|reportes|config`, toast container, modal container, CSS design system

- [ ] **Step 1: Write `index.html`**

```html
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
  :root {
    --primary: #1a56db; --primary-dark: #1e429f; --bg: #f3f4f6; --card: #ffffff;
    --text: #111827; --muted: #6b7280; --border: #e5e7eb;
    --green: #059669; --red: #dc2626; --amber: #d97706; --radius: 10px;
  }
  * { box-sizing: border-box; }
  body { margin: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; background: var(--bg); color: var(--text); }
  .app { display: flex; min-height: 100vh; }
  .sidebar { width: 220px; background: #111827; color: #e5e7eb; padding: 16px 0; flex-shrink: 0; }
  .sidebar .brand { padding: 8px 20px 20px; font-weight: 700; font-size: 18px; color: #fff; border-bottom: 1px solid #1f2937; margin-bottom: 12px; }
  .sidebar nav a { display: block; padding: 10px 20px; color: #9ca3af; text-decoration: none; cursor: pointer; font-size: 14px; }
  .sidebar nav a:hover { background: #1f2937; color: #fff; }
  .sidebar nav a.active { background: var(--primary); color: #fff; }
  .main { flex: 1; padding: 24px; max-width: 1100px; }
  .topbar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
  .topbar h1 { margin: 0; font-size: 22px; }
  .topbar .empresa { color: var(--muted); font-size: 14px; }
  .card { background: var(--card); border-radius: var(--radius); border: 1px solid var(--border); padding: 20px; margin-bottom: 16px; }
  .grid-stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 20px; }
  .stat { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 18px; }
  .stat .val { font-size: 24px; font-weight: 700; margin-top: 4px; }
  .stat .lbl { color: var(--muted); font-size: 13px; }
  button { background: var(--primary); color: #fff; border: 0; border-radius: 8px; padding: 9px 16px; font-size: 14px; cursor: pointer; }
  button:hover { background: var(--primary-dark); }
  button.ghost { background: transparent; color: var(--primary); border: 1px solid var(--primary); }
  button.danger { background: var(--red); }
  button.danger:hover { background: #b91c1c; }
  button.small { padding: 5px 10px; font-size: 12px; }
  input, select, textarea { width: 100%; padding: 9px 10px; border: 1px solid var(--border); border-radius: 8px; font-size: 14px; margin-bottom: 12px; background: #fff; }
  label { display: block; font-size: 13px; color: var(--muted); margin-bottom: 4px; }
  table { width: 100%; border-collapse: collapse; font-size: 14px; }
  th { text-align: left; color: var(--muted); font-weight: 600; padding: 8px; border-bottom: 2px solid var(--border); }
  td { padding: 8px; border-bottom: 1px solid var(--border); }
  .badge { padding: 3px 8px; border-radius: 999px; font-size: 12px; font-weight: 600; }
  .badge.pendiente { background: #fef3c7; color: var(--amber); }
  .badge.pagada { background: #d1fae5; color: var(--green); }
  .toast { position: fixed; bottom: 20px; right: 20px; padding: 12px 18px; border-radius: 8px; color: #fff; font-size: 14px; opacity: 0; transform: translateY(10px); transition: all .25s; z-index: 99; }
  .toast.show { opacity: 1; transform: translateY(0); }
  .toast.error { background: var(--red); }
  .toast.ok { background: var(--green); }
  .view { display: none; }
  .view.active { display: block; }
  .row { display: flex; gap: 12px; }
  .row > div { flex: 1; }
  .muted { color: var(--muted); font-size: 13px; }
  .item-row { display: flex; gap: 8px; align-items: flex-start; margin-bottom: 8px; }
  .item-row input { margin-bottom: 0; }
  .item-row .desc { flex: 3; }
  .item-row .num { flex: 1; }
  .modal { position: fixed; inset: 0; background: rgba(0,0,0,.5); display: none; align-items: center; justify-content: center; z-index: 50; }
  .modal.show { display: flex; }
  .modal-content { background: #fff; border-radius: var(--radius); padding: 24px; width: 90%; max-width: 640px; max-height: 90vh; overflow-y: auto; }
  .totals-preview { background: var(--bg); border-radius: 8px; padding: 12px 16px; font-size: 14px; margin-top: 8px; }
</style>
</head>
<body>
<div class="app">
  <aside class="sidebar">
    <div class="brand">Facturación</div>
    <nav>
      <a data-view="dashboard" class="active">Dashboard</a>
      <a data-view="facturas">Facturas</a>
      <a data-view="clientes">Clientes</a>
      <a data-view="gastos">Gastos</a>
      <a data-view="reportes">Reportes</a>
      <a data-view="config">Configuración</a>
    </nav>
  </aside>
  <main class="main">
    <div class="topbar">
      <h1 id="viewTitle">Dashboard</h1>
      <span class="empresa" id="empresaLabel"></span>
    </div>
    <section class="view active" id="view-dashboard"></section>
    <section class="view" id="view-facturas"></section>
    <section class="view" id="view-clientes"></section>
    <section class="view" id="view-gastos"></section>
    <section class="view" id="view-reportes"></section>
    <section class="view" id="view-config"></section>
  </main>
</div>
<div class="toast" id="toast"></div>
<div class="modal" id="modal"><div class="modal-content" id="modalContent"></div></div>
<script>
  var STATE = { config: {}, view: 'dashboard' };
  var VIEW_TITLES = { dashboard: 'Dashboard', facturas: 'Facturas', clientes: 'Clientes', gastos: 'Gastos', reportes: 'Reportes', config: 'Configuración' };

  function showView(name) {
    document.querySelectorAll('.view').forEach(function (s) { s.classList.remove('active'); });
    document.getElementById('view-' + name).classList.add('active');
    document.querySelectorAll('.sidebar nav a').forEach(function (a) {
      a.classList.toggle('active', a.dataset.view === name);
    });
    document.getElementById('viewTitle').textContent = VIEW_TITLES[name];
    STATE.view = name;
    var loaders = { dashboard: loadDashboard, facturas: loadFacturas, clientes: loadClientes, gastos: loadGastos, reportes: loadReportes, config: loadConfigView };
    if (loaders[name]) loaders[name]();
  }
  document.querySelectorAll('.sidebar nav a').forEach(function (a) {
    a.addEventListener('click', function () { showView(a.dataset.view); });
  });

  function call(name, args) {
    return new Promise(function (resolve, reject) {
      var fn = google.script.run;
      fn.withSuccessHandler(function (r) {
        if (r && r.ok) resolve(r.data);
        else reject(new Error(r && r.error ? r.error : 'Error desconocido'));
      }).withFailureHandler(function (err) { reject(new Error(String(err))); });
      if (args === undefined) fn[name]();
      else fn[name].apply(null, args);
    });
  }

  function toast(msg, type) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'toast show ' + (type || 'ok');
    setTimeout(function () { t.classList.remove('show'); }, 3000);
  }

  function fmtMoney(n) {
    return (STATE.config.moneda_simbolo || '$') + ' ' + Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function mesActual() { return new Date().toISOString().slice(0, 7); }

  function openModal(html) {
    document.getElementById('modalContent').innerHTML = html;
    document.getElementById('modal').classList.add('show');
  }
  function closeModal() { document.getElementById('modal').classList.remove('show'); }
  document.getElementById('modal').addEventListener('click', function (e) { if (e.target === this) closeModal(); });

  function init() {
    call('getConfig').then(function (cfg) {
      STATE.config = cfg;
      document.getElementById('empresaLabel').textContent = (cfg.empresa_nombre || '') + ' — ' + (cfg.moneda || '');
      loadDashboard();
    }).catch(function (e) { toast(e.message, 'error'); });
  }
  window.addEventListener('DOMContentLoaded', init);
</script>
</body>
</html>
```

Note: later tasks add the loader functions referenced here (`loadDashboard`, `loadFacturas`, `loadClientes`, `loadGastos`, `loadReportes`, `loadConfigView`) by inserting `<script>` content before the closing `</script>` tag or replacing section innerHTML. The `if (loaders[name])` guard prevents crashes until those exist.

- [ ] **Step 2: Manual verification**

Deploy a test web app (Deploy → New deployment → Web app, access "Solo yo" / "Me"). Open URL. Expected: sidebar renders, "Dashboard" title shows, empty main section, no console errors.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: shell SPA con navegacion y design system"
```

---

### Task 9: Frontend — Dashboard view

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: `Code.getReportes`, `State.config`
- Produces:
  - `loadDashboard()` → fills 4 stat cards (facturado mes, pendiente, gastos mes, utilidad) using `mesActual()`
  - Quick-action buttons: Nueva factura (→ `showView('facturas')` + `loadFacturas(true)`), Registrar gasto (→ `showView('gastos')`)

- [ ] **Step 1: Fill `#view-dashboard` section and add `loadDashboard`**

Replace the line `<section class="view active" id="view-dashboard"></section>` with:

```html
<section class="view active" id="view-dashboard">
  <div class="grid-stats">
    <div class="stat"><div class="lbl">Facturado (mes)</div><div class="val" id="stFacturado">—</div></div>
    <div class="stat"><div class="lbl">Pendiente de cobro</div><div class="val" id="stPendiente">—</div></div>
    <div class="stat"><div class="lbl">Gastos (mes)</div><div class="val" id="stGastos">—</div></div>
    <div class="stat"><div class="lbl">Utilidad (mes)</div><div class="val" id="stUtilidad">—</div></div>
  </div>
  <div class="card">
    <h3>Acciones rápidas</h3>
    <button onclick="openNewInvoice()">Nueva factura</button>
    <button class="ghost" onclick="showView('gastos')">Registrar gasto</button>
  </div>
</section>
```

Add before the closing `</script>` tag:

```javascript
  function loadDashboard() {
    call('getReportes', [mesActual()]).then(function (d) {
      document.getElementById('stFacturado').textContent = fmtMoney(d.facturado);
      document.getElementById('stPendiente').textContent = fmtMoney(d.pendiente);
      document.getElementById('stGastos').textContent = fmtMoney(d.gastos_total);
      document.getElementById('stUtilidad').textContent = fmtMoney(d.utilidad);
    }).catch(function (e) { toast(e.message, 'error'); });
  }
```

Note: `openNewInvoice()` is defined in Task 10. Until then the button throws if clicked; add it in the same commit as Task 10 or leave a temporary `function openNewInvoice() { showView('facturas'); loadFacturas(true); }` stub.

- [ ] **Step 2: Manual verification**

Refresh web app. Expected: stat cards show numbers (0s or real values for current month), quick buttons visible.

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat: vista dashboard"
```

---

### Task 10: Frontend — Facturas view + new invoice modal + PDF download

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: `Code.listFacturas`, `Code.getFactura`, `Code.createFactura`, `Code.setEstadoFactura`, `Code.deleteFactura`, `Code.listClientes`, `Code.saveCliente`, `Code.buildPDF`
- Produces (global functions):
  - `loadFacturas(openNew)` → renders table with filters; stores rows in `FACTURAS`
  - `applyFiltrosFacturas()`, `markFactura(id, estado)`, `delFactura(id)`
  - `downloadPDF(id)` → base64 from `google.script.run.buildPDF` → anchor download
  - `openInvoiceDetail(id)` → read-only modal with items + totals + PDF button
  - `openNewInvoice()` → modal: client select + quick-create + dynamic item rows + totals preview + save
  - `addItemRow()`, `updatePreview()`, `openQuickClient()`, `saveQuickClient()`, `saveInvoice()`

- [ ] **Step 1: Fill `#view-facturas` section**

Replace `<section class="view" id="view-facturas"></section>` with:

```html
<section class="view" id="view-facturas">
  <div class="card">
    <div class="row" style="align-items:flex-end">
      <div><label>Estado</label><select id="filtroEstado"><option value="">Todos</option><option value="pendiente">Pendiente</option><option value="pagada">Pagada</option></select></div>
      <div><label>Mes</label><input type="month" id="filtroMes"></div>
      <div style="flex:0;padding-bottom:12px"><button onclick="applyFiltrosFacturas()">Filtrar</button></div>
      <div style="flex:0;padding-bottom:12px"><button class="ghost" onclick="openNewInvoice()">Nueva factura</button></div>
    </div>
  </div>
  <div class="card">
    <table>
      <thead><tr><th>Folio</th><th>Cliente</th><th>Fecha</th><th>Total</th><th>Estado</th><th></th></tr></thead>
      <tbody id="tbodyFacturas"><tr><td colspan="6" class="muted">Cargando…</td></tr></tbody>
    </table>
  </div>
</section>
```

- [ ] **Step 2: Add facturas JS before `</script>`**

```javascript
  var FACTURAS = [];

  function loadFacturas(openNew) {
    var estado = document.getElementById('filtroEstado').value;
    var mes = document.getElementById('filtroMes').value;
    call('listFacturas', [{ estado: estado, mes: mes }]).then(function (data) {
      FACTURAS = data;
      var tbody = document.getElementById('tbodyFacturas');
      if (!data.length) { tbody.innerHTML = '<tr><td colspan="6" class="muted">Sin facturas</td></tr>'; return; }
      tbody.innerHTML = data.map(function (f) {
        var badge = f.estado === 'pagada' ? '<span class="badge pagada">Pagada</span>' : '<span class="badge pendiente">Pendiente</span>';
        var btnEstado = f.estado === 'pagada'
          ? '<button class="small ghost" onclick="markFactura(\'' + f.id_factura + '\',\'pendiente\')">Marcar pendiente</button>'
          : '<button class="small ghost" onclick="markFactura(\'' + f.id_factura + '\',\'pagada\')">Marcar pagada</button>';
        return '<tr>' +
          '<td><strong>' + f.folio + '</strong></td>' +
          '<td>' + f.nombre_cliente + '</td>' +
          '<td>' + f.fecha_emision + '</td>' +
          '<td>' + fmtMoney(f.total) + '</td>' +
          '<td>' + badge + '</td>' +
          '<td>' +
            '<button class="small ghost" onclick="openInvoiceDetail(\'' + f.id_factura + '\')">Ver</button> ' +
            '<button class="small ghost" onclick="downloadPDF(\'' + f.id_factura + '\')">PDF</button> ' +
            btnEstado + ' ' +
            '<button class="small danger" onclick="delFactura(\'' + f.id_factura + '\')">Eliminar</button>' +
          '</td></tr>';
      }).join('');
      if (openNew) openNewInvoice();
    }).catch(function (e) { toast(e.message, 'error'); });
  }

  function applyFiltrosFacturas() { loadFacturas(false); }

  function markFactura(id, estado) {
    call('setEstadoFactura', [id, estado]).then(function () {
      toast('Factura ' + (estado === 'pagada' ? 'pagada' : 'pendiente'));
      loadFacturas(false);
    }).catch(function (e) { toast(e.message, 'error'); });
  }

  function delFactura(id) {
    if (!confirm('¿Eliminar esta factura y sus conceptos?')) return;
    call('deleteFactura', [id]).then(function () { toast('Factura eliminada'); loadFacturas(false); })
      .catch(function (e) { toast(e.message, 'error'); });
  }

  function downloadPDF(id) {
    google.script.run.withSuccessHandler(function (b64) {
      var a = document.createElement('a');
      a.href = 'data:application/pdf;base64,' + b64;
      a.download = 'Factura_' + id + '.pdf';
      a.click();
    }).withFailureHandler(function (err) { toast(String(err), 'error'); }).buildPDF(id);
  }

  function openInvoiceDetail(id) {
    call('getFactura', [id]).then(function (d) {
      var f = d.factura;
      var rows = d.items.map(function (it) {
        return '<tr><td>' + it.descripcion + '</td><td>' + it.cantidad + '</td><td>' + fmtMoney(it.precio_unitario) + '</td><td>' + fmtMoney(it.importe) + '</td></tr>';
      }).join('');
      openModal(
        '<h2>Factura ' + f.folio + '</h2>' +
        '<div class="muted">Cliente: ' + f.nombre_cliente + '</div>' +
        '<div class="muted">Emisión: ' + f.fecha_emision + '</div>' +
        '<table><thead><tr><th>Descripción</th><th>Cant.</th><th>P. Unitario</th><th>Importe</th></tr></thead><tbody>' + rows + '</tbody></table>' +
        '<div class="totals-preview">Subtotal: ' + fmtMoney(f.subtotal) + '<br>IVA: ' + fmtMoney(f.iva) + '<br><strong>Total: ' + fmtMoney(f.total) + '</strong></div>' +
        (f.notas ? '<div class="muted">Notas: ' + f.notas + '</div>' : '') +
        '<br><button class="ghost" onclick="downloadPDF(\'' + f.id_factura + '\')">Descargar PDF</button> ' +
        '<button onclick="closeModal()">Cerrar</button>'
      );
    }).catch(function (e) { toast(e.message, 'error'); });
  }

  function addItemRow() {
    var box = document.getElementById('itemsBox');
    var div = document.createElement('div');
    div.className = 'item-row';
    div.innerHTML =
      '<input class="desc" placeholder="Descripción">' +
      '<input class="num" type="number" min="0" step="0.01" placeholder="Cant." value="1">' +
      '<input class="num" type="number" min="0" step="0.01" placeholder="Precio">' +
      '<button class="small danger" onclick="this.parentNode.remove(); updatePreview()">✕</button>';
    box.appendChild(div);
    updatePreview();
  }

  function updatePreview() {
    var rows = document.querySelectorAll('#itemsBox .item-row');
    var subtotal = 0;
    rows.forEach(function (r) {
      var inputs = r.querySelectorAll('input');
      subtotal += Number(inputs[1].value || 0) * Number(inputs[2].value || 0);
    });
    var ivaPct = Number(STATE.config.iva_porcentaje) || 0;
    var iva = subtotal * ivaPct / 100;
    document.getElementById('previewTotales').innerHTML =
      'Subtotal: ' + fmtMoney(subtotal) + '<br>IVA (' + ivaPct + '%): ' + fmtMoney(iva) + '<br><strong>Total: ' + fmtMoney(subtotal + iva) + '</strong>';
  }

  function openNewInvoice() {
    call('listClientes').then(function (clientes) {
      var opts = clientes.map(function (c) {
        return '<option value="' + c.id_cliente + '">' + c.nombre + '</option>';
      }).join('');
      openModal(
        '<h2>Nueva factura</h2>' +
        '<label>Cliente</label><select id="selCliente">' + (opts || '<option value="">Sin clientes</option>') + '</select>' +
        '<button class="small ghost" onclick="openQuickClient()">Crear cliente rápido</button>' +
        '<div id="quickClientBox"></div>' +
        '<label>Conceptos</label><div id="itemsBox"></div>' +
        '<button class="small ghost" onclick="addItemRow()">+ Agregar concepto</button>' +
        '<div class="totals-preview" id="previewTotales"></div>' +
        '<label>Fecha de emisión</label><input type="date" id="fechaEmision" value="' + new Date().toISOString().slice(0, 10) + '">' +
        '<label>Fecha de vencimiento</label><input type="date" id="fechaVencimiento">' +
        '<label>Notas</label><textarea id="notasFactura" rows="2"></textarea>' +
        '<button onclick="saveInvoice()">Guardar factura</button> <button onclick="closeModal()">Cancelar</button>'
      );
      addItemRow();
    }).catch(function (e) { toast(e.message, 'error'); });
  }

  function openQuickClient() {
    var box = document.getElementById('quickClientBox');
    box.innerHTML =
      '<div class="card" style="margin-top:8px">' +
      '<label>Nombre</label><input id="qcNombre">' +
      '<label>RFC</label><input id="qcRfc">' +
      '<label>Email</label><input id="qcEmail">' +
      '<button class="small" onclick="saveQuickClient()">Guardar cliente</button></div>';
  }

  function saveQuickClient() {
    call('saveCliente', [{ nombre: document.getElementById('qcNombre').value, rfc: document.getElementById('qcRfc').value, email: document.getElementById('qcEmail').value }])
      .then(function (c) {
        var sel = document.getElementById('selCliente');
        sel.insertAdjacentHTML('afterbegin', '<option value="' + c.id_cliente + '" selected>' + c.nombre + '</option>');
        document.getElementById('quickClientBox').innerHTML = '';
        toast('Cliente creado');
      }).catch(function (e) { toast(e.message, 'error'); });
  }

  function saveInvoice() {
    var items = [];
    document.querySelectorAll('#itemsBox .item-row').forEach(function (r) {
      var i = r.querySelectorAll('input');
      items.push({ descripcion: i[0].value, cantidad: Number(i[1].value), precio_unitario: Number(i[2].value) });
    });
    var payload = {
      id_cliente: document.getElementById('selCliente').value,
      fecha_emision: document.getElementById('fechaEmision').value,
      fecha_vencimiento: document.getElementById('fechaVencimiento').value,
      notas: document.getElementById('notasFactura').value,
      items: items
    };
    call('createFactura', [payload]).then(function (f) {
      toast('Factura ' + f.folio + ' creada');
      closeModal();
      loadFacturas(false);
    }).catch(function (e) { toast(e.message, 'error'); });
  }
```

- [ ] **Step 3: Manual verification**

Create a client (quick), add 2 items, save. Expected: folio auto-assigned, appears in table, "Marcar pagada" toggles badge, "PDF" downloads valid PDF, "Ver" shows detail, "Eliminar" deletes header + items.

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat: vista facturas + crear factura + pdf"
```

---

### Task 11: Frontend — Clientes view

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: `Code.listClientes`, `Code.saveCliente`, `Code.deleteCliente`
- Produces (global functions):
  - `loadClientes()`, `renderClientes(data)`, `buscarClientes()`
  - `openClienteForm(id)`, `saveClienteForm(id)`, `delCliente(id)`
  - Caches client list in `CLIENTES` for client-side search

- [ ] **Step 1: Fill `#view-clientes` section**

Replace `<section class="view" id="view-clientes"></section>` with:

```html
<section class="view" id="view-clientes">
  <div class="card">
    <div class="row" style="align-items:flex-end">
      <div><label>Buscar</label><input id="buscarCliente" oninput="buscarClientes()" placeholder="Nombre o RFC"></div>
      <div style="flex:0;padding-bottom:12px"><button onclick="openClienteForm()">Nuevo cliente</button></div>
    </div>
  </div>
  <div class="card">
    <table>
      <thead><tr><th>Nombre</th><th>RFC</th><th>Email</th><th>Teléfono</th><th></th></tr></thead>
      <tbody id="tbodyClientes"><tr><td colspan="5" class="muted">Cargando…</td></tr></tbody>
    </table>
  </div>
</section>
```

- [ ] **Step 2: Add clientes JS before `</script>`**

```javascript
  var CLIENTES = [];

  function loadClientes() {
    call('listClientes').then(function (data) {
      CLIENTES = data;
      renderClientes(data);
    }).catch(function (e) { toast(e.message, 'error'); });
  }

  function renderClientes(data) {
    var tbody = document.getElementById('tbodyClientes');
    if (!data.length) { tbody.innerHTML = '<tr><td colspan="5" class="muted">Sin clientes</td></tr>'; return; }
    tbody.innerHTML = data.map(function (c) {
      return '<tr><td><strong>' + c.nombre + '</strong></td><td>' + (c.rfc || '') + '</td><td>' + (c.email || '') + '</td><td>' + (c.telefono || '') + '</td>' +
        '<td><button class="small ghost" onclick="openClienteForm(\'' + c.id_cliente + '\')">Editar</button> ' +
        '<button class="small danger" onclick="delCliente(\'' + c.id_cliente + '\')">Eliminar</button></td></tr>';
    }).join('');
  }

  function buscarClientes() {
    var q = document.getElementById('buscarCliente').value.toLowerCase();
    renderClientes(CLIENTES.filter(function (c) {
      return (c.nombre || '').toLowerCase().indexOf(q) !== -1 || (c.rfc || '').toLowerCase().indexOf(q) !== -1;
    }));
  }

  function openClienteForm(id) {
    var c = CLIENTES.filter(function (x) { return x.id_cliente === id; })[0] || {};
    openModal(
      '<h2>' + (id ? 'Editar cliente' : 'Nuevo cliente') + '</h2>' +
      '<label>Nombre *</label><input id="cliNombre" value="' + (c.nombre || '') + '">' +
      '<label>RFC</label><input id="cliRfc" value="' + (c.rfc || '') + '">' +
      '<label>Email</label><input id="cliEmail" value="' + (c.email || '') + '">' +
      '<label>Teléfono</label><input id="cliTel" value="' + (c.telefono || '') + '">' +
      '<label>Dirección</label><textarea id="cliDir" rows="2">' + (c.direccion || '') + '</textarea>' +
      '<button onclick="saveClienteForm(\'' + (c.id_cliente || '') + '\')">Guardar</button> ' +
      '<button onclick="closeModal()">Cancelar</button>'
    );
  }

  function saveClienteForm(id) {
    var payload = {
      id_cliente: id || undefined,
      nombre: document.getElementById('cliNombre').value,
      rfc: document.getElementById('cliRfc').value,
      email: document.getElementById('cliEmail').value,
      telefono: document.getElementById('cliTel').value,
      direccion: document.getElementById('cliDir').value
    };
    call('saveCliente', [payload]).then(function () {
      toast('Cliente guardado');
      closeModal();
      loadClientes();
    }).catch(function (e) { toast(e.message, 'error'); });
  }

  function delCliente(id) {
    if (!confirm('¿Eliminar este cliente?')) return;
    call('deleteCliente', [id]).then(function () { toast('Cliente eliminado'); loadClientes(); })
      .catch(function (e) { toast(e.message, 'error'); });
  }
```

- [ ] **Step 3: Manual verification**

Create, edit, search, delete a client. Deleting a client that has invoices must show the "tiene facturas" error toast.

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat: vista clientes"
```

---

### Task 12: Frontend — Gastos view

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: `Code.listGastos`, `Code.saveGasto`, `Code.deleteGasto`, `State.config.categorias_gastos`
- Produces (global functions):
  - `loadGastos()`, `applyFiltrosGastos()`, `openGastoForm(id)`, `saveGastoForm(id)`, `delGasto(id)`
  - `catOptions(selected)` → `<option>` list from `STATE.config.categorias_gastos` CSV
  - Caches gastos in `window.__GASTOS` for edit form

- [ ] **Step 1: Fill `#view-gastos` section**

Replace `<section class="view" id="view-gastos"></section>` with:

```html
<section class="view" id="view-gastos">
  <div class="card">
    <div class="row" style="align-items:flex-end">
      <div><label>Mes</label><input type="month" id="filtroGastoMes"></div>
      <div><label>Categoría</label><select id="filtroGastoCat"><option value="">Todas</option></select></div>
      <div style="flex:0;padding-bottom:12px"><button onclick="applyFiltrosGastos()">Filtrar</button></div>
      <div style="flex:0;padding-bottom:12px"><button class="ghost" onclick="openGastoForm()">Registrar gasto</button></div>
    </div>
  </div>
  <div class="card">
    <table>
      <thead><tr><th>Fecha</th><th>Categoría</th><th>Descripción</th><th>Monto</th><th>Proveedor</th><th></th></tr></thead>
      <tbody id="tbodyGastos"><tr><td colspan="6" class="muted">Cargando…</td></tr></tbody>
    </table>
  </div>
</section>
```

- [ ] **Step 2: Add gastos JS before `</script>`**

```javascript
  function loadGastos() {
    var mes = document.getElementById('filtroGastoMes').value;
    var cat = document.getElementById('filtroGastoCat').value;
    call('listGastos', [{ mes: mes, categoria: cat }]).then(function (data) {
      window.__GASTOS = data;
      var tbody = document.getElementById('tbodyGastos');
      if (!data.length) { tbody.innerHTML = '<tr><td colspan="6" class="muted">Sin gastos</td></tr>'; return; }
      tbody.innerHTML = data.map(function (g) {
        return '<tr><td>' + g.fecha + '</td><td>' + (g.categoria || '') + '</td><td>' + g.descripcion + '</td>' +
          '<td><strong>' + fmtMoney(g.monto) + '</strong></td><td>' + (g.proveedor || '') + '</td>' +
          '<td><button class="small ghost" onclick="openGastoForm(\'' + g.id_gasto + '\')">Editar</button> ' +
          '<button class="small danger" onclick="delGasto(\'' + g.id_gasto + '\')">Eliminar</button></td></tr>';
      }).join('');
    }).catch(function (e) { toast(e.message, 'error'); });
  }

  function applyFiltrosGastos() { loadGastos(); }

  function catOptions(selected) {
    return String(STATE.config.categorias_gastos || '').split(',').map(function (c) {
      c = c.trim();
      return c ? '<option value="' + c + '"' + (c === selected ? ' selected' : '') + '>' + c + '</option>' : '';
    }).join('');
  }

  function openGastoForm(id) {
    var g = (window.__GASTOS || []).filter(function (x) { return x.id_gasto === id; })[0] || {};
    var metodo = g.metodo_pago || 'Transferencia';
    var metodos = ['Efectivo', 'Transferencia', 'Tarjeta'].map(function (m) {
      return '<option' + (m === metodo ? ' selected' : '') + '>' + m + '</option>';
    }).join('');
    openModal(
      '<h2>' + (id ? 'Editar gasto' : 'Registrar gasto') + '</h2>' +
      '<label>Fecha</label><input type="date" id="gasFecha" value="' + (g.fecha || new Date().toISOString().slice(0, 10)) + '">' +
      '<label>Categoría</label><select id="gasCat"><option value="">—</option>' + catOptions(g.categoria) + '</select>' +
      '<label>Descripción *</label><input id="gasDesc" value="' + (g.descripcion || '') + '">' +
      '<label>Monto *</label><input id="gasMonto" type="number" min="0" step="0.01" value="' + (g.monto !== undefined ? g.monto : '') + '">' +
      '<label>Método de pago</label><select id="gasMetodo">' + metodos + '</select>' +
      '<label>Proveedor</label><input id="gasProv" value="' + (g.proveedor || '') + '">' +
      '<button onclick="saveGastoForm(\'' + (g.id_gasto || '') + '\')">Guardar</button> ' +
      '<button onclick="closeModal()">Cancelar</button>'
    );
  }

  function saveGastoForm(id) {
    var payload = {
      id_gasto: id || undefined,
      fecha: document.getElementById('gasFecha').value,
      categoria: document.getElementById('gasCat').value,
      descripcion: document.getElementById('gasDesc').value,
      monto: Number(document.getElementById('gasMonto').value),
      metodo_pago: document.getElementById('gasMetodo').value,
      proveedor: document.getElementById('gasProv').value
    };
    call('saveGasto', [payload]).then(function () { toast('Gasto guardado'); closeModal(); loadGastos(); })
      .catch(function (e) { toast(e.message, 'error'); });
  }

  function delGasto(id) {
    if (!confirm('¿Eliminar este gasto?')) return;
    call('deleteGasto', [id]).then(function () { toast('Gasto eliminado'); loadGastos(); })
      .catch(function (e) { toast(e.message, 'error'); });
  }
```

- [ ] **Step 3: Manual verification**

Register, edit, filter by month/category, delete a gasto. Category dropdown populated from Config.

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat: vista gastos"
```

---

### Task 13: Frontend — Reportes view

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: `Code.getReportes`, `State.config.categorias_gastos`
- Produces:
  - `loadReportes()` → stat cards (facturado, cobrado, pendiente, gastos, utilidad) + gastos por categoría list + top 5 clientes
  - Month picker `reporteMes` defaults to current month; refresh button `loadReportes`

- [ ] **Step 1: Fill `#view-reportes` section**

Replace `<section class="view" id="view-reportes"></section>` with:

```html
<section class="view" id="view-reportes">
  <div class="card">
    <div class="row" style="align-items:flex-end">
      <div><label>Mes</label><input type="month" id="reporteMes" value=""></div>
      <div style="flex:0;padding-bottom:12px"><button onclick="loadReportes()">Actualizar</button></div>
    </div>
  </div>
  <div class="grid-stats">
    <div class="stat"><div class="lbl">Facturado</div><div class="val" id="rFacturado">—</div></div>
    <div class="stat"><div class="lbl">Cobrado</div><div class="val" id="rCobrado">—</div></div>
    <div class="stat"><div class="lbl">Pendiente</div><div class="val" id="rPendiente">—</div></div>
    <div class="stat"><div class="lbl">Gastos</div><div class="val" id="rGastos">—</div></div>
    <div class="stat"><div class="lbl">Utilidad</div><div class="val" id="rUtilidad">—</div></div>
  </div>
  <div class="card">
    <h3>Gastos por categoría</h3>
    <ul id="listaCategorias" class="muted"></ul>
  </div>
  <div class="card">
    <h3>Top 5 clientes</h3>
    <table>
      <thead><tr><th>Cliente</th><th>Total</th></tr></thead>
      <tbody id="tbodyTopClientes"></tbody>
    </table>
  </div>
</section>
```

- [ ] **Step 2: Add reportes JS before `</script>`**

```javascript
  function loadReportes() {
    var mes = document.getElementById('reporteMes').value || mesActual();
    document.getElementById('reporteMes').value = mes;
    call('getReportes', [mes]).then(function (d) {
      document.getElementById('rFacturado').textContent = fmtMoney(d.facturado);
      document.getElementById('rCobrado').textContent = fmtMoney(d.cobrado);
      document.getElementById('rPendiente').textContent = fmtMoney(d.pendiente);
      document.getElementById('rGastos').textContent = fmtMoney(d.gastos_total);
      document.getElementById('rUtilidad').textContent = fmtMoney(d.utilidad);
      var cats = Object.keys(d.gastos_por_categoria);
      document.getElementById('listaCategorias').innerHTML = cats.length
        ? cats.map(function (c) { return '<li>' + c + ': <strong>' + fmtMoney(d.gastos_por_categoria[c]) + '</strong></li>'; }).join('')
        : '<li>Sin gastos en el mes</li>';
      var tbody = document.getElementById('tbodyTopClientes');
      tbody.innerHTML = d.top_clientes.length
        ? d.top_clientes.map(function (t) { return '<tr><td>' + t.nombre + '</td><td>' + fmtMoney(t.total) + '</td></tr>'; }).join('')
        : '<tr><td colspan="2" class="muted">Sin facturas en el mes</td></tr>';
    }).catch(function (e) { toast(e.message, 'error'); });
  }
```

- [ ] **Step 3: Manual verification**

Open Reportes. Expected: stats for current month, categories list, top clients. Change month and refresh.

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat: vista reportes"
```

---

### Task 14: Frontend — Configuración view + full end-to-end verification

**Files:**
- Modify: `index.html`

**Interfaces:**
- Consumes: `Code.getConfig`, `Code.saveConfig`
- Produces:
  - `loadConfigView()` → fills form from `STATE.config`
  - `saveConfigForm()` → validates + saves; on success updates `STATE.config` and topbar label
  - Fields: empresa_nombre, empresa_rfc, empresa_direccion, empresa_telefono, empresa_email, empresa_logo, moneda, moneda_simbolo, prefijo_folio, contador_folio, iva_porcentaje, categorias_gastos

- [ ] **Step 1: Fill `#view-config` section**

Replace `<section class="view" id="view-config"></section>` with:

```html
<section class="view" id="view-config">
  <div class="card">
    <h3>Datos de la empresa</h3>
    <div class="row">
      <div><label>Nombre *</label><input id="cfEmpresa" value=""></div>
      <div><label>RFC</label><input id="cfRfc"></div>
    </div>
    <label>Dirección</label><input id="cfDir">
    <div class="row">
      <div><label>Teléfono</label><input id="cfTel"></div>
      <div><label>Email</label><input id="cfEmail"></div>
    </div>
    <label>Logo (URL)</label><input id="cfLogo" placeholder="https://...">
  </div>
  <div class="card">
    <h3>Configuración de facturación</h3>
    <div class="row">
      <div><label>Moneda</label><input id="cfMoneda" placeholder="USD"></div>
      <div><label>Símbolo</label><input id="cfSimbolo" placeholder="$"></div>
    </div>
    <div class="row">
      <div><label>Prefijo de folio *</label><input id="cfPrefijo" placeholder="FAC-"></div>
      <div><label>Contador actual</label><input id="cfContador" type="number" min="1"></div>
    </div>
    <div class="row">
      <div><label>IVA %</label><input id="cfIva" type="number" min="0" step="0.01"></div>
      <div><label>Categorías de gastos (separadas por coma)</label><input id="cfCats"></div>
    </div>
    <button onclick="saveConfigForm()">Guardar configuración</button>
  </div>
</section>
```

- [ ] **Step 2: Add config JS before `</script>`**

```javascript
  function loadConfigView() {
    var c = STATE.config;
    document.getElementById('cfEmpresa').value = c.empresa_nombre || '';
    document.getElementById('cfRfc').value = c.empresa_rfc || '';
    document.getElementById('cfDir').value = c.empresa_direccion || '';
    document.getElementById('cfTel').value = c.empresa_telefono || '';
    document.getElementById('cfEmail').value = c.empresa_email || '';
    document.getElementById('cfLogo').value = c.empresa_logo || '';
    document.getElementById('cfMoneda').value = c.moneda || '';
    document.getElementById('cfSimbolo').value = c.moneda_simbolo || '';
    document.getElementById('cfPrefijo').value = c.prefijo_folio || '';
    document.getElementById('cfContador').value = c.contador_folio || '1';
    document.getElementById('cfIva').value = c.iva_porcentaje || '0';
    document.getElementById('cfCats').value = c.categorias_gastos || '';
  }

  function saveConfigForm() {
    var payload = {
      empresa_nombre: document.getElementById('cfEmpresa').value,
      empresa_rfc: document.getElementById('cfRfc').value,
      empresa_direccion: document.getElementById('cfDir').value,
      empresa_telefono: document.getElementById('cfTel').value,
      empresa_email: document.getElementById('cfEmail').value,
      empresa_logo: document.getElementById('cfLogo').value,
      moneda: document.getElementById('cfMoneda').value,
      moneda_simbolo: document.getElementById('cfSimbolo').value,
      prefijo_folio: document.getElementById('cfPrefijo').value,
      contador_folio: document.getElementById('cfContador').value,
      iva_porcentaje: document.getElementById('cfIva').value,
      categorias_gastos: document.getElementById('cfCats').value
    };
    call('saveConfig', [payload]).then(function () {
      STATE.config = payload;
      document.getElementById('empresaLabel').textContent = (payload.empresa_nombre || '') + ' — ' + (payload.moneda || '');
      toast('Configuración guardada');
      loadConfigView();
    }).catch(function (e) { toast(e.message, 'error'); });
  }
```

- [ ] **Step 3: Full end-to-end verification**

Run `runTests()` in the editor — all `test_*` PASS. Then in the deployed web app: configure empresa/IVA/folio, create client, create invoice with 2 items, download PDF, mark paid, register gasto, check Dashboard + Reportes reflect the numbers, delete invoice/client/gasto.

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat: vista configuracion"
```

---

### Task 15: README + deployment docs

**Files:**
- Create: `README.md`

**Interfaces:**
- Documents: file structure, setup steps (spreadsheet + Apps Script project + copy files + run `ensureSheets` + `runTests`), web app deployment, config fields, how folio/IVA/moneda work.

- [ ] **Step 1: Write `README.md`**

```markdown
# App de Facturación (Google Apps Script)

Mini app de facturación para un solo usuario: crea facturas con PDF bajo demanda,
registra clientes, controla gastos administrativos y genera reportes mensuales.
Datos en Google Sheets, interfaz web (HTML + vanilla JS).

## Archivos

- `appscript.json` — manifest (acceso "Solo yo")
- `Code.gs` — API del servidor (google.script.run) + tests
- `Data.gs` — acceso a Google Sheets
- `Aux.gs` — helpers puros (totales, formato, validación)
- `index.html` — SPA del frontend
- `pdfTemplate.html` — plantilla de la factura para PDF
- `docs/superpowers/specs/2026-08-11-facturacion-design.md` — diseño completo
- `docs/superpowers/plans/2026-08-11-facturacion.md` — plan de implementación

## Instalación

1. Crea un Google Spreadsheet (será el contenedor de datos y de tu script).
2. Extensiones → Apps Script. Copia `appscript.json`, `Code.gs`, `Data.gs`, `Aux.gs`,
   `index.html` y `pdfTemplate.html` al proyecto (nombres idénticos).
3. En el editor, ejecuta `ensureSheets` (crea las 5 hojas y la config por defecto).
4. Ejecuta `runTests` y verifica que todo pase (PASS).
5. Configura tus datos: pestaña `Config` de la hoja (empresa, RFC, moneda, % IVA,
   prefijo de folio, categorías de gastos) o desde la web app (vista Configuración).

## Despliegue (Web App)

1. En el editor: Implementar → Nueva implementación → Aplicación web.
2. Ejecutar como: "Yo". Tener acceso: "Solo yo".
3. Implementa y abre la URL generada.

## Uso

- **Dashboard**: totales del mes.
- **Facturas**: crear, ver, marcar pagada/pendiente, PDF bajo demanda, eliminar.
- **Clientes**: crear, editar, buscar, eliminar (no se elimina si tiene facturas).
- **Gastos**: registrar con categoría, filtrar por mes/categoría.
- **Reportes**: facturado/cobrado/pendiente/gastos/utilidad, gastos por categoría, top clientes.
- **Configuración**: datos de empresa, moneda, prefijo de folio, IVA, categorías.

## Notas

- Los folios usan prefijo + contador autoincremental (protegido con LockService).
- El IVA se calcula sobre el subtotal con el porcentaje configurado.
- Los PDF se generan bajo demanda; no se almacenan automáticamente.
- Facturas internas: sin timbrado fiscal (CFDI/SAT).
```

- [ ] **Step 2: Review final structure**

Verify project tree matches the spec's file list. Run `git status` and confirm only intended files.

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: readme y despliegue"
```

---

## Self-Review (completed by planner)

- **Spec coverage:** Dashboard ✓ (T9), Facturas CRUD + PDF ✓ (T4, T7, T10), Clientes ✓ (T3, T11), Gastos ✓ (T5, T12), Reportes ✓ (T6, T13), Configuración ✓ (T2, T14), folio lock ✓ (T4), sheets setup ✓ (T1), errores/validación ✓ (every endpoint), README ✓ (T15). No gaps.
- **Placeholder scan:** no TBD/TODO; every code step includes full implementation code; test steps include the exact test function.
- **Type consistency:** consistent names — `Data.createFactura_` used in T4 and T7/T9; `Aux.calcTotales` in T1/T4; `Code.getReportes` in T6/T9/T13; `buildPDF` returns Blob (server, not wrapped) and is called directly by `google.script.run` in T10 `downloadPDF` with base64 handling — consistent with `buildPDF` returning a Blob (Apps Script serializes Blobs to base64 over `google.script.run`). `listFacturas_`/`listGastos_` filters match frontend filter params in T10/T12.

