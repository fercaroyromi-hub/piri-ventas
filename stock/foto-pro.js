// Pirí Stock · foto profesional automática
// Recorta el producto, corrige luz y color, y lo pone centrado sobre fondo blanco con una sombra suave.
// El recorte usa un modelo de IA que corre en el mismo celular (se descarga la primera vez).

const LADO = 1200;           // foto final cuadrada 1200 × 1200
const MARGEN = 0.12;         // aire alrededor del producto

let _quitar = null;
async function quitarFondo(blob, avance) {
  if (!_quitar) _quitar = (await import('./vendor/bgremoval.js')).removeBackground;
  return _quitar(blob, {
    model: 'medium',
    output: { format: 'image/png', quality: 1 },
    progress: (k, hecho, total) => {
      if (!avance) return;
      if (k.startsWith('fetch')) avance('Descargando el recortador (solo la primera vez)…', total ? hecho / total : 0);
      else avance('Recortando el producto…', null);
    }
  });
}

function cargar(src) {
  return new Promise((ok, mal) => {
    const im = new Image();
    im.onload = () => ok(im);
    im.onerror = () => mal(new Error('No pude abrir la foto'));
    im.src = typeof src === 'string' ? src : URL.createObjectURL(src);
  });
}
function lienzo(w, h) { const c = document.createElement('canvas'); c.width = w; c.height = h; return c; }
function achicar(im, max) {
  const k = Math.min(1, max / Math.max(im.naturalWidth || im.width, im.naturalHeight || im.height));
  const c = lienzo(Math.round((im.naturalWidth || im.width) * k), Math.round((im.naturalHeight || im.height) * k));
  c.getContext('2d').drawImage(im, 0, 0, c.width, c.height);
  return c;
}
const aBlob = (c, tipo, q) => new Promise(ok => c.toBlob(ok, tipo, q));

// ---------- corrección de luz y color (solo sobre el producto) ----------
function percentil(hist, total, p) { let a = 0; for (let i = 0; i < 256; i++) { a += hist[i]; if (a >= total * p) return i; } return 255; }
function mejorar(c) {
  const x = c.getContext('2d', { willReadFrequently: true }), d = x.getImageData(0, 0, c.width, c.height), a = d.data;
  const H = [new Uint32Array(256), new Uint32Array(256), new Uint32Array(256)], L = new Uint32Array(256);
  let n = 0;
  for (let i = 0; i < a.length; i += 4) {
    if (a[i + 3] < 200) continue;
    const r = a[i], g = a[i + 1], b = a[i + 2];
    H[0][r]++; H[1][g]++; H[2][b]++; L[(r * 77 + g * 150 + b * 29) >> 8]++;
    n++;
  }
  if (n < 50) return c;
  // 1) balance de blancos: las zonas más claras del producto (reflejos, partes blancas) deberían ser neutras
  const lim = percentil(L, n, 0.9);
  let br = 0, bg = 0, bb = 0, bn = 0;
  for (let i = 0; i < a.length; i += 4) {
    if (a[i + 3] < 200) continue;
    const r = a[i], g = a[i + 1], b = a[i + 2];
    if (((r * 77 + g * 150 + b * 29) >> 8) < lim) continue;
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b);
    if (mx - mn > 70) continue; // ignorar colores fuertes (un escudo rojo no es "luz")
    br += r; bg += g; bb += b; bn++;
  }
  let wb = [1, 1, 1];
  if (bn > 30) {
    const mm = (br + bg + bb) / 3;
    wb = [br, bg, bb].map(s => Math.max(0.9, Math.min(1.12, 1 + (mm / s - 1) * 0.7)));
  }
  // 2) niveles: estirar negros y blancos (sin quemar)
  const lo0 = Math.min(...H.map(h => percentil(h, n, 0.004))), hi0 = Math.max(...H.map(h => percentil(h, n, 0.996)));
  const lo = [lo0, lo0, lo0], hi = [hi0, hi0, hi0]; // mismos niveles en los 3 canales: no cambia el color, solo la luz
  // 3) exposición: llevar la luz media a un valor agradable
  const med = percentil(L, n, 0.5);
  const gamma = Math.max(0.7, Math.min(1.35, Math.log(0.56) / Math.log(Math.max(0.05, Math.min(0.95, med / 255)))));
  const T = [0, 1, 2].map(ch => {
    const t = new Uint8ClampedArray(256), l = lo[ch], h = Math.max(l + 30, hi[ch]);
    for (let v = 0; v < 256; v++) {
      let y = Math.max(0, Math.min(1, (v * wb[ch] - l * 0.85) / (h - l * 0.85)));
      y = Math.pow(y, 1 / gamma);
      y = y + 0.12 * Math.sin(Math.PI * (y - 0.5)) * (1 - Math.abs(2 * y - 1)); // contraste suave (curva S)
      t[v] = Math.round(Math.max(0, Math.min(1, y)) * 255);
    }
    return t;
  });
  for (let i = 0; i < a.length; i += 4) {
    if (a[i + 3] === 0) continue;
    let r = T[0][a[i]], g = T[1][a[i + 1]], b = T[2][a[i + 2]];
    const gr = (r + g + b) / 3, s = 1.1; // un poquito más de color
    a[i] = gr + (r - gr) * s; a[i + 1] = gr + (g - gr) * s; a[i + 2] = gr + (b - gr) * s;
  }
  x.putImageData(d, 0, 0);
  return enfocar(c, 0.45);
}
// nitidez (máscara de enfoque)
function enfocar(c, k) {
  const w = c.width, h = c.height, x = c.getContext('2d');
  const d = x.getImageData(0, 0, w, h), a = d.data, o = new Uint8ClampedArray(a);
  for (let y = 1; y < h - 1; y++) for (let X = 1; X < w - 1; X++) {
    const i = (y * w + X) * 4; if (a[i + 3] < 128) continue;
    for (let ch = 0; ch < 3; ch++) {
      const j = i + ch;
      const blur = (o[j - 4] + o[j + 4] + o[j - w * 4] + o[j + w * 4] + o[j] * 4) / 8;
      a[j] = o[j] + (o[j] - blur) * k * 2;
    }
  }
  x.putImageData(d, 0, 0);
  return c;
}
// caja del producto: limpia restos sueltos del fondo y busca dónde está el producto
function caja(c) {
  const w = c.width, h = c.height, x = c.getContext('2d', { willReadFrequently: true });
  const d = x.getImageData(0, 0, w, h), a = d.data;
  for (let i = 3; i < a.length; i += 4) if (a[i] < 60) a[i] = 0; // bruma casi transparente → fuera
  x.putImageData(d, 0, 0);
  const fil = new Uint32Array(h), col = new Uint32Array(w);
  for (let y = 0; y < h; y++) for (let X = 0; X < w; X++) if (a[(y * w + X) * 4 + 3] > 128) { fil[y]++; col[X]++; }
  const mf = Math.max(3, w * 0.006), mc = Math.max(3, h * 0.006);
  let y0 = 0, y1 = h - 1, x0 = 0, x1 = w - 1;
  while (y0 < h && fil[y0] < mf) y0++; while (y1 > y0 && fil[y1] < mf) y1--;
  while (x0 < w && col[x0] < mc) x0++; while (x1 > x0 && col[x1] < mc) x1--;
  if (y0 >= y1 || x0 >= x1) return null;
  // lo que quedó fuera de la caja se borra (manchitas sueltas)
  for (let y = 0; y < h; y++) for (let X = 0; X < w; X++) if (y < y0 || y > y1 || X < x0 || X > x1) a[(y * w + X) * 4 + 3] = 0;
  x.putImageData(d, 0, 0);
  return { x: x0, y: y0, w: x1 - x0 + 1, h: y1 - y0 + 1 };
}
// armar la foto final: fondo blanco, producto centrado, sombra suave abajo
function componer(prod, b) {
  const f = lienzo(LADO, LADO), x = f.getContext('2d');
  x.fillStyle = '#fff'; x.fillRect(0, 0, LADO, LADO);
  const util = LADO * (1 - 2 * MARGEN), k = Math.min(util / b.w, util / b.h);
  const w = b.w * k, h = b.h * k, px = (LADO - w) / 2, py = (LADO - h) / 2 + LADO * 0.02;
  // sombra de apoyo
  const sx = LADO / 2, sy = py + h - h * 0.012, rx = w * 0.46, ry = Math.max(8, w * 0.06);
  const g = x.createRadialGradient(sx, sy, 0, sx, sy, rx);
  g.addColorStop(0, 'rgba(0,0,0,0.20)'); g.addColorStop(0.55, 'rgba(0,0,0,0.08)'); g.addColorStop(1, 'rgba(0,0,0,0)');
  x.save(); x.translate(sx, sy); x.scale(1, ry / rx); x.translate(-sx, -sy);
  x.fillStyle = g; x.beginPath(); x.arc(sx, sy, rx, 0, Math.PI * 2); x.fill(); x.restore();
  // sombra de contacto suave
  x.save(); x.shadowColor = 'rgba(0,0,0,0.16)'; x.shadowBlur = LADO * 0.03; x.shadowOffsetY = LADO * 0.012;
  x.drawImage(prod, b.x, b.y, b.w, b.h, px, py, w, h); x.restore();
  return f;
}
// versión sin recorte: solo mejora luz y color, cuadrada
function cuadrada(c) {
  const f = lienzo(LADO, LADO), x = f.getContext('2d'), s = Math.min(c.width, c.height);
  x.fillStyle = '#fff'; x.fillRect(0, 0, LADO, LADO);
  x.drawImage(c, (c.width - s) / 2, (c.height - s) / 2, s, s, 0, 0, LADO, LADO);
  return f;
}

/**
 * Convierte una foto cualquiera en foto de catálogo.
 * Devuelve { pro, simple, recortada } — blobs JPEG (pro = fondo blanco; simple = original mejorada).
 */
export async function fotoProfesional(file, avance) {
  const im = await cargar(file);
  const base = achicar(im, 1600);
  avance && avance('Preparando la foto…', null);
  const simple = await aBlob(cuadrada(mejorar(achicar(im, 1600))), 'image/jpeg', 0.9);
  let pro = null, recortada = false;
  try {
    const png = await quitarFondo(await aBlob(base, 'image/png'), avance);
    avance && avance('Corrigiendo luz y color…', null);
    const cut = achicar(await cargar(png), 1600);
    const b = caja(cut);
    if (b && b.w * b.h > cut.width * cut.height * 0.02) {
      pro = await aBlob(componer(mejorar(cut), b), 'image/jpeg', 0.9);
      recortada = true;
    }
  } catch (e) { console.warn('[fotoPro] sin recorte:', e); }
  return { pro: pro || simple, simple, recortada };
}
