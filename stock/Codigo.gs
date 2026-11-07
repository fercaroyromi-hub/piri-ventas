/**
 * Pirí Stock · conexión entre la app y la planilla "Pirí Stock | App".
 * Pegar este código completo en Extensiones → Apps Script de esa planilla
 * e implementarlo como "Aplicación web" (Ejecutar como: yo · Acceso: cualquier persona).
 */
const ID_PLANILLA = '1Rdg3WjKs_1UPHRjkRioRhZMmSkNReuxPn36m25c3ETY';
const SS = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(ID_PLANILLA);
const TZ = 'America/Montevideo';
const CARPETA_FACTURAS = 'Pirí Stock · Facturas';

function doGet(e) {
  if (e && e.parameter && e.parameter.c) {
    let out;
    try { out = Object.assign({ ok: true }, catalogo()); } catch (err) { out = { ok: false, error: err.message }; }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput('Pirí Stock funcionando ✓');
}

// Correr una vez desde el editor para dar los permisos (Drive para guardar fotos de facturas)
function autorizar() {
  DriveApp.getRootFolder();
  leerConfig();
  Logger.log('Listo ✓');
}

function doPost(e) {
  let out;
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const cfg = leerConfig();
    if (String(d.pin || '') !== cfg.pin) throw new Error('PIN incorrecto');
    const u = String(d.usuario || '').trim() || '—';
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      switch (d.action) {
        case 'init': sincronizarDirectas(cfg); break;
        case 'mov': movimiento(d, u); break;
        case 'conteo': conteo(d, u); break;
        case 'venta': venta(d, u); break;
        case 'anular': anular(d, u); break;
        case 'producto': guardarProducto(d); break;
        case 'colores': separarColores(d); break;
        case 'kit': guardarKit(d); break;
        case 'kitBorrar': borrarKit(d.nombre); break;
        case 'asignar': asignar(d); sincronizarDirectas(cfg); break;
        case 'asignarColor': asignarColor(d, u, cfg); break;
        case 'factura': factura(d, u); break;
        case 'gasto': gasto(d, u); break;
        case 'gastoBorrar': borrarGasto(d.id); break;
        case 'costo': ponerCostoMov(d); break;
        case 'foto': guardarFoto(d); break;
        case 'config': guardarConfig(d); break;
        default: throw new Error('Acción desconocida');
      }
      SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
    out = Object.assign({ ok: true }, estado(cfg));
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- lectura ---------- */
function hoja(n) {
  let sh = SS.getSheetByName(n);
  if (!sh) {
    const H = { 'Facturas': ['ID', 'Fecha de carga', 'Fecha factura', 'Proveedor', 'Ítems', 'Total', 'Usuario', 'Foto'],
      'Alias factura': ['Texto en la factura', 'ID producto'],
      'Gastos': ['ID', 'Fecha', 'Usuario', 'Concepto', 'Monto', 'Nota', 'Origen'] }[n];
    if (!H) throw new Error('Falta la pestaña ' + n);
    sh = SS.insertSheet(n); sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold'); sh.setFrozenRows(1);
  }
  return sh;
}
function filas(n) {
  const sh = hoja(n); const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, Math.max(1, sh.getLastColumn())).getValues();
}
// "Fer: 094914223, Caro: 095110574" → [{nombre:'Fer', numero:'59894914223'}, …]
function telefonos(txt) {
  return String(txt || '').split(/[,;\n]/).map(x => x.trim()).filter(String).map(x => {
    const m = x.match(/^(.*?)[:\s]*([+\d][\d\s-]{6,})$/);
    let n = (m ? m[2] : x).replace(/\D/g, ''), nom = m ? m[1].trim() : '';
    if (/^0\d{8}$/.test(n)) n = '598' + n.slice(1); else if (/^9\d{7}$/.test(n)) n = '598' + n;
    return { nombre: nom, numero: n };
  }).filter(t => t.numero.length >= 8);
}
function leerConfig() {
  const c = {};
  filas('Config').forEach(r => c[String(r[0]).trim()] = r[1]);
  let desde = c['Descontar ventas directas desde'];
  desde = desde instanceof Date ? desde : (desde ? new Date(String(desde) + 'T00:00:00') : new Date());
  return {
    whatsapp: telefonos(c['WhatsApp pedidos']),
    nombreCat: String(c['Nombre del catálogo'] || 'Pirí · Diseño Láser').trim(),
    mostrarCant: /^s/i.test(String(c['Mostrar cantidades en catálogo'] || 'No')),
    pin: String(c['PIN'] || '').trim(),
    usuarios: String(c['Usuarios'] || 'Fer').split(',').map(s => s.trim()).filter(String),
    desde: desde,
    directas: String(c['ID planilla ventas directas'] || '').trim()
  };
}
function productos() {
  return filas('Productos').filter(r => r[0]).map(r => ({
    id: String(r[0]), nombre: String(r[1]), cat: String(r[2] || 'Otros'), unidad: String(r[3] || 'unidades'),
    min: Number(r[4]) || 0, costo: r[5] === '' ? '' : Number(r[5]) || 0, color: String(r[8] || '').trim(),
    precio: r[9] === '' || r[9] == null ? '' : Number(r[9]) || 0, foto: String(r[10] || '').trim(),
    catalogo: r[11] === '' || r[11] == null ? !!String(r[10] || '').trim() : /^s/i.test(String(r[11])),
    catRaw: String(r[11] == null ? '' : r[11]).trim()
  }));
}
const nombreCompleto = p => p.nombre + (p.color ? ' · ' + p.color : '');
function movimientos() {
  return filas('Movimientos').filter(r => r[0]).map(r => ({
    id: String(r[0]), fecha: r[1] instanceof Date ? r[1].getTime() : new Date(r[1]).getTime(), usuario: String(r[2]),
    tipo: String(r[3]), pid: String(r[4]), producto: String(r[5]), cant: Number(r[6]) || 0,
    canal: String(r[7] || ''), nota: String(r[8] || ''), origen: String(r[9] || ''),
    costo: r[10] === '' || r[10] == null ? '' : Number(r[10])
  }));
}
function kits() {
  const k = {};
  filas('Kits').forEach(r => {
    const n = String(r[0]).trim(); if (!n || !r[1]) return;
    (k[n] = k[n] || []).push({ id: String(r[1]).trim(), cant: Number(r[2]) || 1 });
  });
  return k;
}
function equivalencias() {
  const m = {};
  filas('Equivalencias').forEach(r => { if (r[0]) m[norm(r[0])] = String(r[1] || '').trim(); });
  return m;
}
function aliasFactura() {
  const m = {};
  filas('Alias factura').forEach(r => { if (r[0] && r[1]) m[norm(r[0])] = String(r[1]).trim(); });
  return m;
}
function norm(s) {
  return String(s).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/(?<=[A-Z])0\b/g, 'O').replace(/[^A-Z0-9]+/g, ' ').trim();
}

function estado(cfg) {
  const P = productos(), M = movimientos(), stock = {};
  M.forEach(m => stock[m.pid] = (stock[m.pid] || 0) + m.cant);
  P.forEach(p => p.stock = Math.round((stock[p.id] || 0) * 1000) / 1000);
  const F = filas('Facturas').filter(r => r[0]).slice(-60).reverse().map(r => ({
    id: String(r[0]), fecha: r[1] instanceof Date ? r[1].getTime() : 0, fechaFac: String(r[2] || ''),
    proveedor: String(r[3] || ''), items: String(r[4] || ''), total: Number(r[5]) || 0, usuario: String(r[6] || ''), foto: String(r[7] || '')
  }));
  const anul = {}; M.forEach(m => { if (m.origen.indexOf('ANULA:') === 0) anul[m.origen.slice(6)] = 1; });
  const compras = M.filter(m => m.tipo === 'Entrada' && !anul[m.id]).map(m => ({ id: m.id, fecha: m.fecha, pid: m.pid, producto: m.producto,
    cant: m.cant, costo: m.costo, nota: m.nota, origen: m.origen, usuario: m.usuario })).reverse();
  const gastos = filas('Gastos').filter(r => r[0]).map(r => ({ id: String(r[0]), fecha: r[1] instanceof Date ? r[1].getTime() : new Date(r[1]).getTime(),
    usuario: String(r[2]), concepto: String(r[3]), monto: Number(r[4]) || 0, nota: String(r[5] || ''), origen: String(r[6] || '') })).reverse();
  return {
    config: { whatsapp: cfg.whatsapp, nombreCat: cfg.nombreCat, mostrarCant: cfg.mostrarCant },
    productos: P, kits: kits(), usuarios: cfg.usuarios, alias: aliasFactura(), facturas: F, compras: compras, gastos: gastos,
    movimientos: M.slice(-500).reverse(),
    pendientes: pendientesDirectas(cfg, P)
  };
}

/* ---------- escritura ---------- */
let _n = 0;
function nuevoId(pre) { return (pre || 'M') + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + (++_n) + Math.floor(Math.random() * 90 + 10); }
function prod(pid) {
  const p = productos().find(x => x.id === pid);
  if (!p) throw new Error('Producto no encontrado');
  return p;
}
function agregar(rows) {
  if (!rows.length) return;
  rows = rows.map(r => { while (r.length < 12) r.push(''); return r; });
  const sh = hoja('Movimientos');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
}
function fila(u, tipo, p, cant, canal, nota, origen, costo) {
  const c = costo === '' || costo == null || isNaN(Number(costo)) ? '' : Number(costo);
  return [nuevoId(), new Date(), u, tipo, p.id, nombreCompleto(p), cant, canal || '', nota || '', origen || '', c, c === '' ? '' : Math.round(c * Math.abs(cant) * 100) / 100];
}
function positivo(n) {
  n = Number(n);
  if (!(n > 0)) throw new Error('La cantidad tiene que ser mayor a 0');
  return n;
}
function ponerCosto(pid, costo) {
  if (costo === '' || costo == null || isNaN(Number(costo))) return;
  const sh = hoja('Productos'), P = filas('Productos');
  const i = P.findIndex(r => String(r[0]) === pid);
  if (i >= 0) sh.getRange(i + 2, 6).setValue(Number(costo));
}
// tipo: Entrada (+) · Salida (−)
function movimiento(d, u) {
  const p = prod(d.pid), n = positivo(d.cant);
  if (d.tipo === 'Entrada') {
    agregar([fila(u, 'Entrada', p, n, '', d.nota, '', d.costo)]);
    ponerCosto(p.id, d.costo);
  } else {
    agregar([fila(u, d.motivo || 'Salida', p, -n, d.canal, d.nota)]);
  }
}
function conteo(d, u) {
  const p = prod(d.pid), real = Number(d.cant);
  if (!(real >= 0)) throw new Error('Poné cuántas hay (0 o más)');
  const actual = movimientos().filter(m => m.pid === p.id).reduce((a, m) => a + m.cant, 0);
  const dif = real - actual;
  if (dif === 0) return;
  agregar([fila(u, 'Conteo', p, dif, '', 'Conté ' + real + ' (antes figuraban ' + actual + ')' + (d.nota ? ' · ' + d.nota : ''))]);
}
function componentes(destino, cant) {
  if (destino.indexOf('K:') === 0) {
    const k = kits()[destino.slice(2)];
    if (!k) throw new Error('Kit no encontrado');
    return k.map(c => ({ p: prod(c.id), cant: c.cant * cant }));
  }
  return [{ p: prod(destino.replace(/^P:/, '')), cant: cant }];
}
function venta(d, u) {
  const n = positivo(d.cant);
  const comp = componentes(d.destino, n);
  const etiqueta = d.destino.indexOf('K:') === 0 ? d.destino.slice(2) + ' × ' + n : '';
  agregar(comp.map(c => fila(u, 'Venta', c.p, -c.cant, d.canal || 'Otro', [etiqueta, d.nota].filter(String).join(' · '))));
}
function anular(d, u) {
  const M = movimientos(), m = M.find(x => x.id === d.id);
  if (!m) throw new Error('Movimiento no encontrado');
  if (m.tipo === 'Anulación' || M.some(x => x.origen === 'ANULA:' + m.id)) throw new Error('Ya estaba anulado');
  agregar([[nuevoId(), new Date(), u, 'Anulación', m.pid, m.producto, -m.cant, m.canal, 'Anula: ' + m.tipo + ' del ' +
    Utilities.formatDate(new Date(m.fecha), TZ, 'dd/MM HH:mm'), 'ANULA:' + m.id]]);
}
function siguienteId(P) {
  const num = P.reduce((a, r) => Math.max(a, parseInt(String(r[0]).replace(/\D/g, '')) || 0), 0) + 1;
  return 'P' + (num < 10 ? '0' : '') + num;
}
function crearProducto(x) {
  const nombre = String(x.nombre || '').trim();
  if (!nombre) throw new Error('Falta el nombre');
  const color = String(x.color || '').trim();
  const sh = hoja('Productos'), P = filas('Productos');
  if (P.some(r => norm(r[1]) === norm(nombre) && norm(r[8] || '') === norm(color))) throw new Error('Ya existe ' + nombre + (color ? ' · ' + color : ''));
  const id = siguienteId(P), r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, 12).setValues([[id, nombre, String(x.cat || 'Otros').trim(), String(x.unidad || 'unidades').trim(), Number(x.min) || 0,
    x.costo === '' || x.costo == null ? '' : Number(x.costo), '', '', color,
    x.precio === '' || x.precio == null ? '' : Number(x.precio), '', x.catalogo === false ? 'No' : '']]);
  sh.getRange(r, 7, 1, 2).setFormulas([['=SUMIFS(Movimientos!G:G,Movimientos!E:E,A' + r + ')',
    '=IF(G' + r + '<=0,"Sin stock",IF(G' + r + '<=E' + r + ',"Reponer","OK"))']]);
  return id;
}
function guardarProducto(d) {
  const nombre = String(d.nombre || '').trim();
  if (!nombre) throw new Error('Falta el nombre');
  if (d.id) {
    const sh = hoja('Productos'), P = filas('Productos');
    const i = P.findIndex(r => String(r[0]) === d.id);
    if (i < 0) throw new Error('Producto no encontrado');
    const viejo = String(P[i][1]);
    const vals = [nombre, String(d.cat || 'Otros').trim(), String(d.unidad || 'unidades').trim(), Number(d.min) || 0, d.costo === '' || d.costo == null ? '' : Number(d.costo)];
    sh.getRange(i + 2, 2, 1, 5).setValues([vals]);
    sh.getRange(i + 2, 9).setValue(String(d.color || '').trim());
    if ('precio' in d) sh.getRange(i + 2, 10).setValue(d.precio === '' || d.precio == null ? '' : Number(d.precio));
    if ('catalogo' in d) sh.getRange(i + 2, 12).setValue(d.catalogo ? 'Sí' : 'No');
    // si cambió el nombre, los otros colores del mismo producto también
    if (d.todosLosColores && norm(viejo) !== norm(nombre)) P.forEach((r, j) => { if (j !== i && norm(r[1]) === norm(viejo)) sh.getRange(j + 2, 2, 1, 4).setValues([vals.slice(0, 4)]); });
    return;
  }
  const colores = (d.colores || []).map(c => String(c).trim()).filter(String);
  if (!colores.length) return crearProducto(d);
  colores.forEach(c => crearProducto(Object.assign({}, d, { color: c })));
}
// Pasa un producto sin color a tener colores: el producto actual toma el primer color (conserva su stock)
function separarColores(d) {
  const p = prod(d.id);
  const colores = (d.colores || []).map(c => String(c).trim()).filter(String);
  if (!colores.length) throw new Error('Escribí al menos un color');
  let lista = colores;
  if (!p.color) {
    const sh = hoja('Productos'), P = filas('Productos');
    const i = P.findIndex(r => String(r[0]) === p.id);
    sh.getRange(i + 2, 9).setValue(colores[0]);
    lista = colores.slice(1);
  }
  lista.forEach(c => crearProducto({ nombre: p.nombre, cat: p.cat, unidad: p.unidad, min: p.min, costo: p.costo, precio: p.precio, color: c }));
}
function guardarKit(d) {
  const nombre = String(d.nombre || '').trim();
  if (!nombre) throw new Error('Falta el nombre del kit');
  const items = (d.items || []).filter(x => x.id && Number(x.cant) > 0);
  if (!items.length) throw new Error('El kit necesita al menos un producto');
  if (d.anterior && d.anterior !== nombre) borrarKit(d.anterior, nombre);
  borrarKit(nombre);
  const sh = hoja('Kits');
  sh.getRange(sh.getLastRow() + 1, 1, items.length, 3).setValues(items.map(x => [nombre, x.id, Number(x.cant)]));
}
function borrarKit(nombre, renombrarA) {
  const sh = hoja('Kits'), F = filas('Kits');
  for (let i = F.length - 1; i >= 0; i--) if (String(F[i][0]).trim() === nombre) sh.deleteRow(i + 2);
  if (renombrarA) {
    const e = hoja('Equivalencias'), E = filas('Equivalencias');
    E.forEach((r, i) => { if (String(r[1]).trim() === 'K:' + nombre) e.getRange(i + 2, 2).setValue('K:' + renombrarA); });
  }
}
function asignar(d) {
  const texto = String(d.texto || '').trim();
  if (!texto) throw new Error('Falta el producto');
  const sh = hoja('Equivalencias'), E = filas('Equivalencias');
  const i = E.findIndex(r => norm(r[0]) === norm(texto));
  if (i >= 0) sh.getRange(i + 2, 2).setValue(d.destino);
  else sh.appendRow([texto, d.destino]);
}

/* ---------- facturas de compra ---------- */
function factura(d, u) {
  const items = (d.items || []).filter(x => Number(x.cant) > 0 && (x.pid || x.nuevo));
  if (!items.length && !(d.gastos || []).length) throw new Error('No hay nada para cargar');
  const fid = nuevoId('F');
  const proveedor = String(d.proveedor || '').trim(), fechaFac = String(d.fecha || '').trim();
  let foto = '';
  if (d.foto) {
    try {
      const it = DriveApp.getFoldersByName(CARPETA_FACTURAS);
      const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_FACTURAS);
      const blob = Utilities.newBlob(Utilities.base64Decode(String(d.foto).replace(/^data:image\/\w+;base64,/, '')), 'image/jpeg',
        'Factura ' + (proveedor || 'sin proveedor') + ' ' + (fechaFac || Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy')) + '.jpg');
      foto = carpeta.createFile(blob).getUrl();
    } catch (e) { foto = ''; }
  }
  const aliasSh = hoja('Alias factura'), A = aliasFactura();
  const filasMov = [], resumen = [];
  let total = 0;
  items.forEach(x => {
    let pid = x.pid;
    if (!pid && x.nuevo) pid = crearProducto(x.nuevo);
    const p = prod(pid), n = Number(x.cant), c = x.costo === '' || x.costo == null ? '' : Number(x.costo);
    filasMov.push(fila(u, 'Entrada', p, n, '', 'Factura' + (proveedor ? ' ' + proveedor : '') + (fechaFac ? ' · ' + fechaFac : ''), 'FAC:' + fid, c));
    if (c !== '') { ponerCosto(pid, c); total += c * n; }
    resumen.push(n + ' × ' + nombreCompleto(p) + (c !== '' ? ' ($' + c + ')' : ''));
    const t = String(x.texto || '').trim();
    if (t && norm(t).length > 2 && A[norm(t)] !== pid) { aliasSh.appendRow([t, pid]); A[norm(t)] = pid; }
  });
  agregar(filasMov);
  const G = (d.gastos || []).filter(g => Number(g.monto) > 0);
  if (G.length) {
    const sh = hoja('Gastos');
    sh.getRange(sh.getLastRow() + 1, 1, G.length, 7).setValues(G.map(g => [nuevoId('G'), new Date(), u, String(g.concepto || 'Otro gasto').trim(),
      Number(g.monto), 'Factura' + (proveedor ? ' ' + proveedor : ''), 'FAC:' + fid]));
    G.forEach(g => { total += Number(g.monto); resumen.push(String(g.concepto || 'Otro gasto') + ' ($' + g.monto + ')'); });
  }
  hoja('Facturas').appendRow([fid, new Date(), fechaFac, proveedor, resumen.join(' · '), Math.round(total * 100) / 100, u, foto]);
}

/* ---------- ventas de la app Pirí Ventas ---------- */
function ventasDirectas(cfg) {
  if (!cfg.directas) return [];
  let v;
  try { v = SpreadsheetApp.openById(cfg.directas).getSheetByName('Ventas').getDataRange().getValues(); }
  catch (e) { return []; }
  const h = v.shift().map(x => String(x).trim());
  const c = n => h.indexOf(n);
  const desde = new Date(cfg.desde); desde.setHours(0, 0, 0, 0);
  const out = [], ocur = {};
  v.forEach(r => {
    const id = String(r[c('ID')] || '').trim(); if (!id) return;
    let f = r[c('Fecha')];
    if (!(f instanceof Date)) { const p = String(f).split('/'); f = new Date(+p[2], +p[1] - 1, +p[0]); }
    if (isNaN(f) || f < desde) return;
    const prodTxt = String(r[c('Producto')] || '').trim(), cant = Number(String(r[c('Cantidad')]).replace(/[^\d.,]/g, '').replace(',', '.')) || 1;
    const base = id + '|' + norm(prodTxt); ocur[base] = (ocur[base] || 0) + 1;
    out.push({ clave: 'VD:' + base + '|' + ocur[base], id: id, fecha: f.getTime(), producto: prodTxt, cant: cant,
      cliente: String(r[c('Cliente')] || ''), vendedor: String(r[c('Vendedor')] || ''),
      anulado: /anulad/i.test(String(r[c('Estado')] || '')) });
  });
  return out;
}
// Si el producto tiene varios colores, no se puede saber cuál se vendió: queda para elegir en la app
function colores(P, pid) {
  const p = P.find(x => x.id === pid); if (!p) return [];
  const g = P.filter(x => norm(x.nombre) === norm(p.nombre));
  return g.length > 1 ? g : [];
}
function sincronizarDirectas(cfg) {
  const V = ventasDirectas(cfg); if (!V.length) return;
  const eq = equivalencias(), M = movimientos(), P = productos(), hechos = {};
  M.forEach(m => { if (m.origen.indexOf('VD:') === 0) hechos[m.origen] = (hechos[m.origen] || 0) + m.cant; });
  const nuevas = [];
  V.forEach(v => {
    const dest = eq[norm(v.producto)];
    if (!dest || dest === 'IGNORAR') return;
    const ya = Object.prototype.hasOwnProperty.call(hechos, v.clave);
    const neto = hechos[v.clave] || 0;
    if (!v.anulado && !ya) {
      if (dest.indexOf('P:') === 0 && colores(P, dest.slice(2)).length) return; // espera que elijan el color
      try {
        componentes(dest, v.cant).forEach(c => nuevas.push(fila(v.vendedor || 'App', 'Venta', c.p, -c.cant, 'Venta directa',
          v.id + ' · ' + v.producto + ' × ' + v.cant + (v.cliente ? ' · ' + v.cliente : ''), v.clave)));
      } catch (e) { /* kit o producto borrado: queda pendiente */ }
    } else if (v.anulado && ya && neto !== 0) {
      M.filter(m => m.origen === v.clave).forEach(m => nuevas.push([nuevoId(), new Date(), 'App', 'Anulación', m.pid, m.producto,
        -m.cant, 'Venta directa', 'Venta ' + v.id + ' anulada en Pirí Ventas', v.clave]));
    }
  });
  agregar(nuevas);
  const sh = hoja('Config'), F = filas('Config');
  const i = F.findIndex(r => String(r[0]).trim() === 'Última sincronización');
  if (i >= 0) sh.getRange(i + 2, 2).setValue(Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'));
}
function pendientesDirectas(cfg, P) {
  const eq = equivalencias(), vistos = {}, hechos = {};
  movimientos().forEach(m => { if (m.origen.indexOf('VD:') === 0) hechos[m.origen] = 1; });
  const out = [];
  ventasDirectas(cfg).forEach(v => {
    if (v.anulado) return;
    const dest = eq[norm(v.producto)];
    if (!dest) {
      const k = norm(v.producto); if (vistos[k]) { vistos[k].veces++; return; }
      vistos[k] = v; v.veces = 1; v.tipo = 'asignar'; out.push(v); return;
    }
    if (dest.indexOf('P:') === 0 && !hechos[v.clave]) {
      const g = colores(P, dest.slice(2));
      if (g.length) { v.tipo = 'color'; v.opciones = g.map(x => ({ id: x.id, color: x.color || 'Sin color', stock: x.stock })); v.nombre = g[0].nombre; out.push(v); }
    }
  });
  return out;
}
function asignarColor(d, u, cfg) {
  const v = ventasDirectas(cfg).find(x => x.clave === d.clave);
  if (!v) throw new Error('No encontré esa venta');
  if (movimientos().some(m => m.origen === v.clave)) return;
  const p = prod(d.pid);
  agregar([fila(v.vendedor || u, 'Venta', p, -v.cant, 'Venta directa', v.id + ' · ' + v.producto + ' × ' + v.cant + (v.cliente ? ' · ' + v.cliente : ''), v.clave)]);
}

/* ---------- gastos que no son mercadería ---------- */
function gasto(d, u) {
  const monto = Number(d.monto);
  if (!(monto > 0)) throw new Error('Poné cuánto gastaste');
  const concepto = String(d.concepto || '').trim();
  if (!concepto) throw new Error('Poné en qué gastaste');
  hoja('Gastos').appendRow([nuevoId('G'), new Date(), u, concepto, monto, String(d.nota || '').trim(), '']);
}
function borrarGasto(id) {
  const sh = hoja('Gastos'), F = filas('Gastos');
  const i = F.findIndex(r => String(r[0]) === id);
  if (i < 0) throw new Error('Gasto no encontrado');
  sh.deleteRow(i + 2);
}
// Poner o corregir lo que costó una compra ya cargada
function ponerCostoMov(d) {
  const c = Number(d.costo);
  if (!(c >= 0) || d.costo === '' || d.costo == null) throw new Error('Poné el costo por unidad');
  const sh = hoja('Movimientos'), F = filas('Movimientos');
  const i = F.findIndex(r => String(r[0]) === d.id);
  if (i < 0 || String(F[i][3]) !== 'Entrada') throw new Error('Compra no encontrada');
  sh.getRange(i + 2, 11, 1, 2).setValues([[c, Math.round(c * Math.abs(Number(F[i][6])) * 100) / 100]]);
  // si es la compra más nueva de ese producto, también actualiza el costo del producto
  const pid = String(F[i][4]);
  const ultima = F.map((r, j) => ({ r, j })).filter(x => String(x.r[3]) === 'Entrada' && String(x.r[4]) === pid).pop();
  if (ultima && ultima.j === i) ponerCosto(pid, c);
}

/* ---------- fotos y catálogo ---------- */
const CARPETA_FOTOS = 'Pirí Stock · Fotos de productos';
function carpeta(n) { const it = DriveApp.getFoldersByName(n); return it.hasNext() ? it.next() : DriveApp.createFolder(n); }
function guardarFoto(d) {
  const p = prod(d.id);
  if (!d.imagen) throw new Error('Falta la foto');
  const blob = Utilities.newBlob(Utilities.base64Decode(String(d.imagen).replace(/^data:image\/\w+;base64,/, '')), 'image/jpeg', nombreCompleto(p) + '.jpg');
  const f = carpeta(CARPETA_FOTOS).createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const sh = hoja('Productos'), P = filas('Productos');
  const i = P.findIndex(r => String(r[0]) === p.id);
  const vieja = String(P[i][10] || '').trim();
  sh.getRange(i + 2, 11).setValue(f.getId());
  if (String(P[i][11] || '') === '') sh.getRange(i + 2, 12).setValue('Sí');
  if (vieja) { try { DriveApp.getFileById(vieja).setTrashed(true); } catch (e) {} }
}
function guardarConfig(d) {
  const sh = hoja('Config'), F = filas('Config');
  const poner = (clave, valor, ayuda) => {
    const i = F.findIndex(r => String(r[0]).trim() === clave);
    if (i >= 0) sh.getRange(i + 2, 2).setValue(valor);
    else { sh.appendRow([clave, valor, ayuda]); F.push([clave, valor]); }
  };
  if ('whatsapp' in d) poner('WhatsApp pedidos', String(d.whatsapp || '').trim(), 'Quién recibe los pedidos del catálogo. Ej.: Fer: 094914223, Caro: 095110574');
  if ('nombreCat' in d) poner('Nombre del catálogo', String(d.nombreCat || '').trim(), 'Título que ven tus clientes en el catálogo.');
  if ('mostrarCant' in d) poner('Mostrar cantidades en catálogo', d.mostrarCant ? 'Sí' : 'No', 'Sí = el cliente ve cuántas quedan. No = solo Disponible / Agotado.');
}
function catalogo() {
  const cfg = leerConfig(), P = productos(), stock = {};
  movimientos().forEach(m => stock[m.pid] = (stock[m.pid] || 0) + m.cant);
  return {
    nombre: cfg.nombreCat, whatsapp: cfg.whatsapp,
    productos: P.filter(p => p.catalogo).map(p => {
      const st = Math.max(0, stock[p.id] || 0);
      const o = { id: p.id, nombre: p.nombre, color: p.color, cat: p.cat, precio: p.precio, foto: p.foto, disponible: st > 0 };
      if (cfg.mostrarCant) o.stock = st;
      return o;
    })
  };
}
