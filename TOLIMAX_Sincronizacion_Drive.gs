/************************************************************************
 *  TOLIMAX · Conector Drive  (base de datos + pedidos + facturas)
 *  Abre tus hojas por ID (funciona también en Unidades compartidas).
 *  Impuestos: ICUI 20% + IVA 5%.
 *
 *  Al actualizar el código: Implementar > Administrar implementaciones >
 *  editar (lápiz) > Versión: Nueva versión > Implementar. (La URL /exec no cambia.)
 ************************************************************************/

var IDS = {
  clientes:      '1AJCUyOZWjujhXdnZKGxQp38De7jDK31ggpZ4-b8fWOQ',
  precios:       '1qvBB-XEDvhoTkayBMVvnadISpy4F24DbNbTNOotPTc0',
  solicitantes:  '11ap5Kn_w-bA5r6IqiGkkYDqnL2XAIUepLitSKW4bcIY',
  despachadores: '1moPr-ABs-6dgLOs-SXGUoA0bq7Cz1oAxpuc0Vhv1Dps',
  facturadores:  '1YiK3UTb-bJ7_Uu2wo8xu3WaD_i_Qgvcc8jUIWU9bgkg'
};
var PEDIDOS_ID = '1xHs5wK4XQp_2Cwly3D4HWvz53xDj_DpRsOBv_6LRQE0';
var FOLDER_ID = '12RiBuZIlOOCyufUnqu9cuF_krFgRqpXU';
var FOLDER_FACTURAS = 'TOLIMAX - Facturas';
var DATA_URL = 'https://tolimax-neiva.github.io/tolimax/data.json';

var HEADER = ['Folio','Fecha','Solicitante','Despachador','Cedula','Cliente','Lista','Forma de pago',
  'Producto','Cantidad','Subtotal.Unit','Subtotal.Linea','Subtotal','ICUI','IVA','Total','Registrado',
  'ObsItem','Estado','Facturador','FechaFactura','ObsFactura'];

function ahora() { return Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss'); }

function doGet(e) {
  var res = (e && e.parameter && e.parameter.resource) || 'pedidos';
  if (res === 'catalogo') return json(getCatalogo());
  return json(getPedidos());
}
function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents);
    var a = body.action || 'pedido';
    if (a === 'addCliente') return json(saveCliente(body.cliente, false));
    if (a === 'updateCliente') return json(saveCliente(body.cliente, true));
    if (a === 'saveProducto') return json(saveProducto(body.producto));
    if (a === 'addSolicitante') return json(addPersona('solicitantes', body.nombre));
    if (a === 'addDespachador') return json(addPersona('despachadores', body.nombre));
    if (a === 'addFacturador') return json(addPersona('facturadores', body.nombre));
    if (a === 'facturar') return json(facturar(body));
    if (a === 'updatePedido') return json(updatePedido(body));
    return json(guardarPedido(body));
  } catch (err) { return json({ ok: false, error: String(err) }); } finally { lock.releaseLock(); }
}

/* ---- Pedidos ---- */
function pedSheet() { var sh = SpreadsheetApp.openById(PEDIDOS_ID).getSheets()[0];
  if (sh.getLastColumn() < HEADER.length) sh.getRange(1, 1, 1, HEADER.length).setValues([HEADER]); return sh; }
function filaPedido(order, it) {
  return [order.folio, order.fecha, order.solicitante, order.despachador, order.cliente.cedula, order.cliente.nombre,
    order.cliente.tier, order.cliente.pago || '', it.desc, it.qty, it.unit, it.total, order.subtotal, order.icui,
    order.iva, order.total, ahora(), it.obs || '', 'PENDIENTE', '', '', ''];
}
function guardarPedido(order) {
  var sh = pedSheet();
  order.items.forEach(function (it) { sh.appendRow(filaPedido(order, it)); });
  guardarFactura(order);
  return { ok: true, folio: order.folio };
}
function updatePedido(order) {
  var sh = pedSheet();
  var v = sh.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(order.folio)) {
    if (String(v[i][18]).toUpperCase() === 'FACTURADO') return { ok: false, error: 'FACTURADO' };
    rows.push(i + 1);
  }
  rows.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  order.items.forEach(function (it) { sh.appendRow(filaPedido(order, it)); });
  return { ok: true, updated: order.folio };
}
function facturar(body) {
  var sh = pedSheet();
  var v = sh.getDataRange().getValues(); var n = 0;
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === String(body.folio)) {
    sh.getRange(i + 1, 19, 1, 4).setValues([['FACTURADO', body.facturador || '', ahora(), body.obs || '']]); n++;
  }
  return { ok: n > 0, folio: body.folio, filas: n };
}

/* ---- Escrituras base ---- */
function saveCliente(c, update) {
  var sh = hoja('clientes');
  var row = [c.cedula, c.nombre, c.tipo || 'COMERCIAL', c.dir || '', c.pago || '', c.tel || '', c.mail || '', c.city || '', c.zip || '', c.country || 'Colombia'];
  if (update) { var v = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim() === String(c.cedula).trim() && (!c._match || String(v[i][1]).trim() === String(c._match).trim())) {
      sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return { ok: true, updated: c.cedula }; } }
  sh.appendRow(row); return { ok: true, added: c.cedula };
}
function saveProducto(p) {
  var sh = hoja('precios');
  function b(base) { base = Math.round(base || 0); return [base, Math.round(base * 0.2), Math.round(base * 0.05), Math.round(base * 1.25)]; }
  var row = [p.desc, p.gramos || ''].concat(b(p.comercial)).concat(b(p.distribuidor)).concat(b(p.mayorista));
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]).trim().toUpperCase() === String(p.descOriginal || p.desc).trim().toUpperCase()) {
    sh.getRange(i + 1, 1, 1, row.length).setValues([row]); return { ok: true, updated: p.desc }; }
  sh.appendRow(row); return { ok: true, added: p.desc };
}
function addPersona(hojaNom, nombre) { hoja(hojaNom).appendRow([nombre, '', 'SI']); return { ok: true, nombre: nombre }; }

/* ---- Catálogo ---- */
function getCatalogo() {
  return { clientes: readClientes(), productos: readPrecios(), solicitantes: readLista('solicitantes'),
    despachadores: readLista('despachadores'), facturadores: readLista('facturadores') };
}
function readClientes() { var v = valores('clientes'); if (!v.length) return []; v.shift();
  return v.filter(function (r) { return r[0] !== '' && r[0] != null; }).map(function (r) {
    return { cedula: String(r[0]).trim(), nombre: String(r[1] || '').trim(), tipo: String(r[2] || 'COMERCIAL').trim(),
      dir: String(r[3] || '').trim(), pago: String(r[4] || '').trim(), tel: r[5] ? String(r[5]).trim() : '',
      mail: String(r[6] || '').trim(), city: String(r[7] || '').trim() }; }); }
function readPrecios() { var v = valores('precios'); if (!v.length) return []; v.shift();
  return v.filter(function (r) { return r[0]; }).map(function (r) {
    return { desc: String(r[0]).trim(), gramos: r[1],
      base: { COMERCIAL: Math.round(r[2]), DISTRIBUIDOR: Math.round(r[6]), MAYORISTA: Math.round(r[10]) },
      pvta: { COMERCIAL: Math.round(r[5]), DISTRIBUIDOR: Math.round(r[9]), MAYORISTA: Math.round(r[13]) } }; }); }
function readLista(name) { var v = valores(name); if (!v.length) return []; v.shift();
  return v.filter(function (r) { return r[0] && String(r[2] || 'SI').toUpperCase() !== 'NO'; }).map(function (r) { return String(r[0]).trim(); }); }

/* ---- Leer pedidos (Panel + columna del despacho) ---- */
function getPedidos() {
  var sh = pedSheet(); var rows = sh.getDataRange().getValues(); rows.shift();
  var by = {};
  rows.forEach(function (r) {
    var folio = r[0]; if (!folio) return;
    if (!by[folio]) by[folio] = { folio: folio, fecha: r[1], solicitante: r[2], despachador: r[3],
      cliente: { cedula: String(r[4]), nombre: r[5], tier: r[6], pago: r[7] }, items: [],
      subtotal: r[12], icui: r[13], iva: r[14], total: r[15],
      estado: r[18] || 'PENDIENTE', facturador: r[19] || '', fechaFactura: r[20] || '', obsFactura: r[21] || '' };
    by[folio].items.push({ desc: r[8], qty: r[9], unit: r[10], total: r[11], obs: r[17] || '' });
  });
  return Object.keys(by).map(function (k) { return by[k]; });
}

/* ---- Carga masiva de clientes ---- */
function importarClientes() {
  var data = JSON.parse(UrlFetchApp.fetch(DATA_URL).getContentText());
  var sh = hoja('clientes'); sh.clear();
  sh.appendRow(['Cedula','Nombre y apellido','tipo','Address','Tipo de pago','Phone1','E_Mail','City','ZipCode','Country']);
  var rows = data.clientes.map(function (c) { return [c.cedula, c.nombre, c.tipo, c.dir, c.pago, c.tel, c.mail, c.city, '', 'Colombia']; });
  if (rows.length) sh.getRange(2, 1, rows.length, 10).setValues(rows); return rows.length;
}
function organizar() {
  var folder = DriveApp.getFolderById(FOLDER_ID);
  DriveApp.getFileById(PEDIDOS_ID).moveTo(folder);
  try { DriveApp.getFileById('1LQYKHTF2AsnBZy0Ih_eso2n1O8DJOyVJCvfYY-J1nvg').setTrashed(true); } catch (e) {}
  return 'listo';
}

/* ---- Helpers ---- */
function hoja(name) { return SpreadsheetApp.openById(IDS[name]).getSheets()[0]; }
function valores(name) { try { return hoja(name).getDataRange().getValues(); } catch (e) { return []; } }
function guardarFactura(order) {
  var it = DriveApp.getFoldersByName(FOLDER_FACTURAS);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_FACTURAS);
  var rows = order.items.map(function (i) {
    return '<tr><td>' + i.desc + (i.obs ? '<br><small style="color:#888">' + i.obs + '</small>' : '') + '</td><td align="right">' + i.qty +
      '</td><td align="right">' + money(i.unit) + '</td><td align="right">' + money(i.total) + '</td></tr>'; }).join('');
  var html = '<html><head><meta charset="utf-8"><style>body{font-family:Arial;color:#222;padding:24px}h1{color:#4A2A18}' +
    'table{width:100%;border-collapse:collapse;margin-top:10px}th{background:#4A2A18;color:#fff;text-align:left;padding:8px}' +
    'td{padding:8px;border-bottom:1px solid #eee}.r{text-align:right}</style></head><body>' +
    '<h1>TOLIMAX S.A. &mdash; ' + order.folio + '</h1><p>' + order.cliente.nombre + ' &middot; CC ' + order.cliente.cedula +
    ' &middot; Lista: ' + order.cliente.tier + '<br>Fecha: ' + order.fecha + '</p>' +
    '<table><tr><th>Producto</th><th class="r">Cant</th><th class="r">Subtotal U.</th><th class="r">Subtotal</th></tr>' + rows +
    '<tr><td colspan="3" class="r">Subtotal</td><td class="r">' + money(order.subtotal) + '</td></tr>' +
    '<tr><td colspan="3" class="r">ICUI 20%</td><td class="r">' + money(order.icui) + '</td></tr>' +
    '<tr><td colspan="3" class="r">IVA 5%</td><td class="r">' + money(order.iva) + '</td></tr>' +
    '<tr><td colspan="3" class="r"><b>TOTAL</b></td><td class="r"><b>' + money(order.total) + '</b></td></tr></table></body></html>';
  folder.createFile(order.folio + '.html', html, MimeType.HTML);
}
function money(n) { return '$' + Math.round(n).toLocaleString('es-CO'); }
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }
