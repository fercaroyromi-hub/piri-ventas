/**
 * Pirí Stock · conexión entre la app y la planilla "Pirí Stock | App".
 * Pegar este código completo en Extensiones → Apps Script de esa planilla
 * e implementarlo como "Aplicación web" (Ejecutar como: yo · Acceso: cualquier persona).
 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
const TZ = 'America/Montevideo';

function doGet() {
  return ContentService.createTextOutput('Pirí Stock funcionando ✓');
}

function doPost(e) {
  let out;
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const cfg = leerConfig();
    if (String(d.pin || '') !== cfg.pin) throw new Error('PIN incorrecto');
    const u = String(d.usuario || '').trim() || '—';
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      switch (d.action) {
        case 'init': sincronizarDirectas(cfg); break;
        case 'mov': movimiento(d, u); break;
        case 'conteo': conteo(d, u); break;
        case 'venta': venta(d, u); break;
        case 'anular': anular(d, u); break;
        case 'producto': guardarProducto(d); break;
        case 'kit': guardarKit(d); break;
        case 'kitBorrar': borrarKit(d.nombre); break;
        case 'asignar': asignar(d); sincronizarDirectas(cfg); break;
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
function hoja(n) { return SS.getSheetByName(n); }
function filas(n) {
  const sh = hoja(n); const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}
function leerConfig() {
  const c = {};
  filas('Config').forEach(r => c[String(r[0]).trim()] = r[1]);
  let desde = c['Descontar ventas directas desde'];
  desde = desde instanceof Date ? desde : (desde ? new Date(String(desde) + 'T00:00:00') : new Date());
  return {
    pin: String(c['PIN'] || '').trim(),
    usuarios: String(c['Usuarios'] || 'Fer').split(',').map(s => s.trim()).filter(String),
    desde: desde,
    directas: String(c['ID planilla ventas directas'] || '').trim()
  };
}
function productos() {
  return filas('Productos').filter(r => r[0]).map(r => ({
    id: String(r[0]), nombre: String(r[1]), cat: String(r[2] || 'Otros'), unidad: String(r[3] || 'unidades'),
    min: Number(r[4]) || 0, costo: r[5] === '' ? '' : Number(r[5]) || 0
  }));
}
function movimientos() {
  return filas('Movimientos').filter(r => r[0]).map(r => ({
    id: String(r[0]), fecha: r[1] instanceof Date ? r[1].getTime() : new Date(r[1]).getTime(), usuario: String(r[2]),
    tipo: String(r[3]), pid: String(r[4]), producto: String(r[5]), cant: Number(r[6]) || 0,
    canal: String(r[7] || ''), nota: String(r[8] || ''), origen: String(r[9] || '')
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
function norm(s) {
  return String(s).toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/(?<=[A-Z])0\b/g, 'O').replace(/[^A-Z0-9]+/g, ' ').trim();
}

function estado(cfg) {
  const P = productos(), M = movimientos(), stock = {};
  M.forEach(m => stock[m.pid] = (stock[m.pid] || 0) + m.cant);
  P.forEach(p => p.stock = Math.round((stock[p.id] || 0) * 1000) / 1000);
  return {
    productos: P, kits: kits(), usuarios: cfg.usuarios,
    movimientos: M.slice(-400).reverse(),
    pendientes: pendientesDirectas(cfg)
  };
}

/* ---------- escritura ---------- */
function nuevoId() { return 'M' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10); }
function prod(pid) {
  const p = productos().find(x => x.id === pid);
  if (!p) throw new Error('Producto no encontrado');
  return p;
}
function agregar(rows) {
  if (!rows.length) return;
  const sh = hoja('Movimientos');
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
}
function fila(u, tipo, p, cant, canal, nota, origen) {
  return [nuevoId(), new Date(), u, tipo, p.id, p.nombre, cant, canal || '', nota || '', origen || ''];
}
function positivo(n) {
  n = Number(n);
  if (!(n > 0)) throw new Error('La cantidad tiene que ser mayor a 0');
  return n;
}
// tipo: Entrada (+) · Salida (−)
function movimiento(d, u) {
  const p = prod(d.pid), n = positivo(d.cant);
  const signo = d.tipo === 'Entrada' ? 1 : -1;
  agregar([fila(u, d.tipo === 'Entrada' ? 'Entrada' : (d.motivo || 'Salida'), p, signo * n, d.canal, d.nota)]);
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
  const m = movimientos().find(x => x.id === d.id);
  if (!m) throw new Error('Movimiento no encontrado');
  if (m.tipo === 'Anulación' || movimientos().some(x => x.origen === 'ANULA:' + m.id)) throw new Error('Ya estaba anulado');
  agregar([[nuevoId(), new Date(), u, 'Anulación', m.pid, m.producto, -m.cant, m.canal, 'Anula: ' + m.tipo + ' del ' +
    Utilities.formatDate(new Date(m.fecha), TZ, 'dd/MM HH:mm'), 'ANULA:' + m.id]]);
}
function guardarProducto(d) {
  const nombre = String(d.nombre || '').trim();
  if (!nombre) throw new Error('Falta el nombre');
  const sh = hoja('Productos'), P = filas('Productos');
  const vals = [nombre, String(d.cat || 'Otros').trim(), String(d.unidad || 'unidades').trim(), Number(d.min) || 0, d.costo === '' || d.costo == null ? '' : Number(d.costo)];
  let i = d.id ? P.findIndex(r => String(r[0]) === d.id) : -1;
  if (i >= 0) { sh.getRange(i + 2, 2, 1, 5).setValues([vals]); return; }
  if (P.some(r => norm(r[1]) === norm(nombre))) throw new Error('Ya existe un producto con ese nombre');
  const num = P.reduce((a, r) => Math.max(a, parseInt(String(r[0]).replace(/\D/g, '')) || 0), 0) + 1;
  const id = 'P' + ('0' + num).slice(-2), r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, 8).setValues([[id].concat(vals, [
    '=SUMIFS(Movimientos!G:G,Movimientos!E:E,A' + r + ')',
    '=IF(G' + r + '<=0,"Sin stock",IF(G' + r + '<=E' + r + ',"Reponer","OK"))'])]);
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
function sincronizarDirectas(cfg) {
  const V = ventasDirectas(cfg); if (!V.length) return;
  const eq = equivalencias(), M = movimientos(), hechos = {};
  M.forEach(m => { if (m.origen.indexOf('VD:') === 0) hechos[m.origen] = (hechos[m.origen] || 0) + m.cant; });
  const nuevas = [];
  V.forEach(v => {
    const dest = eq[norm(v.producto)];
    if (!dest || dest === 'IGNORAR') return;
    const ya = Object.prototype.hasOwnProperty.call(hechos, v.clave);
    const neto = hechos[v.clave] || 0;
    if (!v.anulado && !ya) {
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
function pendientesDirectas(cfg) {
  const eq = equivalencias(), vistos = {};
  return ventasDirectas(cfg).filter(v => !v.anulado && !eq[norm(v.producto)]).filter(v => {
    const k = norm(v.producto); if (vistos[k]) { vistos[k].veces++; return false; }
    vistos[k] = v; v.veces = 1; return true;
  });
}
