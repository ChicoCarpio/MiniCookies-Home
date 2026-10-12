// Compara la fórmula de la norma (HLS con HLSMAX=255) con lo que da Excel (medidas/base.json).
const j = JSON.parse(require('fs').readFileSync(__dirname + '/base.json', 'utf8').replace(/^\uFEFF/, ''));
const base = Object.fromEntries(j.tema.map(t => [t.indice, t.rgb]));
function aHls(r, g, b) { r /= 255; g /= 255; b /= 255; const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0, l = (mx + mn) / 2;
  if (mx !== mn) { const d = mx - mn; s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; } return [h, s, l]; }
function aRgb(h, s, l) { const f = (p, q, t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < .5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  if (!s) return [l, l, l].map(v => Math.round(v * 255)); const q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; return [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)].map(v => Math.round(v * 255)); }
const hex = a => a.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
function tinte(rgb, t, cuantizar) {
  let [h, s, l] = aHls(parseInt(rgb.slice(0, 2), 16), parseInt(rgb.slice(2, 4), 16), parseInt(rgb.slice(4), 16));
  if (cuantizar) { h = Math.round(h * 255) / 255; s = Math.round(s * 255) / 255; l = Math.round(l * 255) / 255; }
  l = t < 0 ? l * (1 + t) : l * (1 - t) + (1 - (1 - t));
  if (cuantizar) l = Math.round(l * 255) / 255;
  return hex(aRgb(h, s, Math.min(1, Math.max(0, l))));
}
for (const cuant of [false, true]) {
  let mal = 0, maxd = 0, peor = null;
  for (const m of j.tint) { const e = tinte(base[m.tema], m.tint, cuant); const d = Math.max(...[0, 2, 4].map(i => Math.abs(parseInt(e.substr(i, 2), 16) - parseInt(m.rgb.substr(i, 2), 16)))); if (d) mal++; if (d > maxd) { maxd = d; peor = [m.tema, m.tint, e, m.rgb]; } }
  console.log('cuantizado a 0..255:', cuant, '→ distintos', mal, 'de', j.tint.length, 'máx. diferencia por canal', maxd, peor);
}
