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

var Aux = {
  uid: uid,
  round2: round2,
  calcTotales: calcTotales,
  formatMoney: formatMoney,
  parseCSV: parseCSV,
  requireFields: requireFields
};
