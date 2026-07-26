/************************************************************************
 *  TOLIMAX · Conector Drive  (base de datos + pedidos + facturas)
 *  Abre tus hojas por ID (funciona también en Unidades compartidas).
 *  Impuestos: ICUI 20% + IVA 5%.
 *
 *  INSTALACIÓN (una vez):
 *   1. https://script.google.com > "Nuevo proyecto", borra todo y pega esto.
 *      IMPORTANTE: entra con la MISMA cuenta de Google dueña de las hojas.
 *   2. "Implementar" > "Nueva implementación" > "Aplicación web":
 *        Ejecutar como: Yo   ·   Con acceso: Cualquier persona
 *   3. Copia la URL /exec y pégala en la app (⚙ Configuración / Drive).
 *
 *  CARGAR LOS 838 CLIENTES: elige la función  importarClientes  y Ejecutar.
 ************************************************************************/

/* IDs de tus hojas de Google (ya creadas en tu Drive) */
var IDS = {
  clientes:      '1AJCUyOZWjujhXdnZKGxQp38De7jDK31ggpZ4-b8fWOQ',
  precios:       '1qvBB-XEDvhoTkayBMVvnadISpy4F24DbNbTNOotPTc0',
  solicitantes:  '11ap5Kn_w-bA5r6IqiGkkYDqnL2XAIUepLitSKW4bcIY',
  despachadores: '1moPr-ABs-6dgLOs-SXGUoA0bq7Cz1oAxpuc0Vhv1Dps'
};
var PEDIDOS_ID = '';   // se llena solo la primera vez (no tocar)
var FOLDER_FACTURAS = 'TOLIMAX - Facturas';
var DATA_URL = 'https://tolimax-neiva.github.io/tolimax/data.json';

function doGet(e) {
  var res = (e && e.parameter && e.parameter.resource) || 'pedidos';
  if (res === 'catalogo') return json(getCatalogo());
  return json(getPedidos());
}
function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action || 'pedido';
    if (action === 'addCliente') return json(addCliente(body.cliente));
    if (action === 'addSolicitante') return json(addPersona('solicitantes', body.nombre));
    if (action === 'addDespachador') return json(addPersona('despachadores', body.nombre));
    return json(guardarPedido(body));
  } catch (err) { return json({ ok: false, error: String(err) }); } finally { lock.releaseLock(); }
}

/* ---- Escrituras ---- */
function guardarPedido(order) {
  var sh = pedidosSheet();
  order.items.forEach(function (it) {
    sh.appendRow([order.folio, order.fecha, order.solicitante, order.despachador,
      order.cliente.cedula, order.cliente.nombre, order.cliente.tier, order.cliente.pago || '',
      it.desc, it.qty, it.unit, it.total, order.subtotal, order.icui, order.iva, order.total, new Date()]);
  });
  guardarFactura(order);
  return { ok: true, folio: order.folio };
}
function addCliente(c) {
  hoja('clientes').appendRow([c.cedula, c.nombre, c.tipo || 'COMERCIAL', c.dir || '', c.pago || '',
    c.tel || '', c.mail || '', c.city || '', c.zip || '', c.country || 'Colombia']);
  return { ok: true, cedula: c.cedula };
}
function addPersona(hojaNom, nombre) {
  hoja(hojaNom).appendRow([nombre, '', 'SI']);
  return { ok: true, nombre: nombre };
}

/* ---- Catálogo ---- */
function getCatalogo() {
  return { clientes: readClientes(), productos: readPrecios(),
    solicitantes: readLista('solicitantes'), despachadores: readLista('despachadores') };
}
function readClientes() {
  var v = valores('clientes'); if (!v.length) return []; v.shift();
  return v.filter(function (r) { return r[0] !== '' && r[0] != null; }).map(function (r) {
    return { cedula: String(r[0]).trim(), nombre: String(r[1] || '').trim(),
      tipo: String(r[2] || 'COMERCIAL').trim(), dir: String(r[3] || '').trim(),
      pago: String(r[4] || '').trim(), tel: r[5] ? String(r[5]).trim() : '',
      mail: String(r[6] || '').trim(), city: String(r[7] || '').trim() }; });
}
function readPrecios() {
  var v = valores('precios'); if (!v.length) return []; v.shift();
  return v.filter(function (r) { return r[0]; }).map(function (r) {
    return { desc: String(r[0]).trim(), gramos: r[1],
      base: { COMERCIAL: Math.round(r[2]), DISTRIBUIDOR: Math.round(r[6]), MAYORISTA: Math.round(r[10]) },
      pvta: { COMERCIAL: Math.round(r[5]), DISTRIBUIDOR: Math.round(r[9]), MAYORISTA: Math.round(r[13]) } }; });
}
function readLista(name) {
  var v = valores(name); if (!v.length) return []; v.shift();
  return v.filter(function (r) { return r[0] && String(r[2] || 'SI').toUpperCase() !== 'NO'; })
          .map(function (r) { return String(r[0]).trim(); });
}
function getPedidos() {
  var sh = pedidosSheet(); var rows = sh.getDataRange().getValues(); rows.shift();
  var by = {};
  rows.forEach(function (r) {
    var folio = r[0]; if (!folio) return;
    if (!by[folio]) by[folio] = { folio: folio, fecha: r[1], solicitante: r[2], despachador: r[3],
      cliente: { cedula: String(r[4]), nombre: r[5], tier: r[6], pago: r[7] }, items: [],
      subtotal: r[12], icui: r[13], iva: r[14], total: r[15] };
    by[folio].items.push({ desc: r[8], qty: r[9], unit: r[10], total: r[11] });
  });
  return Object.keys(by).map(function (k) { return by[k]; });
}

/* ---- Carga masiva de clientes desde la web ---- */
function importarClientes() {
  if (!DATA_URL) throw new Error('Falta DATA_URL.');
  var data = JSON.parse(UrlFetchApp.fetch(DATA_URL).getContentText());
  var sh = hoja('clientes'); sh.clear();
  sh.appendRow(['Cedula','Nombre y apellido','tipo','Address','Tipo de pago','Phone1','E_Mail','City','ZipCode','Country']);
  var rows = data.clientes.map(function (c) { return [c.cedula, c.nombre, c.tipo, c.dir, c.pago, c.tel, c.mail, c.city, '', 'Colombia']; });
  if (rows.length) sh.getRange(2, 1, rows.length, 10).setValues(rows);
  return rows.length;
}

/* ---- Helpers (abren por ID) ---- */
function hoja(name) { return SpreadsheetApp.openById(IDS[name]).getSheets()[0]; }
function valores(name) { try { return hoja(name).getDataRange().getValues(); } catch (e) { return []; } }
function pedidosSheet() {
  var props = PropertiesService.getScriptProperties();
  var id = PEDIDOS_ID || props.getProperty('PEDIDOS_ID');
  if (id) { try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) {} }
  var ss = SpreadsheetApp.create('TOLIMAX_pedidos');
  ss.getSheets()[0].appendRow(['Folio','Fecha','Solicitante','Despachador','Cedula','Cliente','Lista',
    'Forma de pago','Producto','Cantidad','Subtotal.Unit','Subtotal.Linea','Subtotal','ICUI','IVA','Total','Registrado']);
  props.setProperty('PEDIDOS_ID', ss.getId());
  return ss.getSheets()[0];
}
function guardarFactura(order) {
  var it = DriveApp.getFoldersByName(FOLDER_FACTURAS);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_FACTURAS);
  var rows = order.items.map(function (i) {
    return '<tr><td>' + i.desc + '</td><td align="right">' + i.qty + '</td><td align="right">' +
      money(i.unit) + '</td><td align="right">' + money(i.total) + '</td></tr>'; }).join('');
  var html = '<html><head><meta charset="utf-8"><style>body{font-family:Arial;color:#222;padding:24px}' +
    'h1{color:#4A2A18}table{width:100%;border-collapse:collapse;margin-top:10px}th{background:#4A2A18;color:#fff;text-align:left;padding:8px}' +
    'td{padding:8px;border-bottom:1px solid #eee}.r{text-align:right}</style></head><body>' +
    '<h1>TOLIMAX S.A. &mdash; ' + order.folio + '</h1><p>' + order.cliente.nombre + ' &middot; CC ' +
    order.cliente.cedula + ' &middot; Lista: ' + order.cliente.tier + '<br>Fecha: ' + order.fecha + '</p>' +
    '<table><tr><th>Producto</th><th class="r">Cant</th><th class="r">Subtotal U.</th><th class="r">Subtotal</th></tr>' + rows +
    '<tr><td colspan="3" class="r">Subtotal</td><td class="r">' + money(order.subtotal) + '</td></tr>' +
    '<tr><td colspan="3" class="r">ICUI 20%</td><td class="r">' + money(order.icui) + '</td></tr>' +
    '<tr><td colspan="3" class="r">IVA 5%</td><td class="r">' + money(order.iva) + '</td></tr>' +
    '<tr><td colspan="3" class="r"><b>TOTAL</b></td><td class="r"><b>' + money(order.total) + '</b></td></tr></table></body></html>';
  folder.createFile(order.folio + '.html', html, MimeType.HTML);
}
function money(n) { return '$' + Math.round(n).toLocaleString('es-CO'); }
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
