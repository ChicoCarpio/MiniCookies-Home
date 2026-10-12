const j = JSON.parse(require('fs').readFileSync(__dirname + '/base.json', 'utf8').replace(/^\uFEFF/, ''));
const base = Object.fromEntries(j.tema.map(t => [t.indice, t.rgb]));
const hex = a => a.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
function aHls(r, g, b, M) { r /= 255; g /= 255; b /= 255; const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0, l = (mx + mn) / 2;
  if (mx !== mn) { const d = mx - mn; s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; } return [h, s, l]; }
function aRgb(h, s, l) { const f = (p, q, t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < .5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  if (!s) return [l, l, l]; const q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; return [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)]; }
function variante(rgb, t, M, qHS, qL, redondeoSalida) {
  let [h, s, l] = aHls(parseInt(rgb.slice(0, 2), 16), parseInt(rgb.slice(2, 4), 16), parseInt(rgb.slice(4), 16));
  const q = (v, modo) => modo === 'r' ? Math.round(v * M) / M : modo === 't' ? Math.floor(v * M) / M : v;
  h = q(h, qHS); s = q(s, qHS); l = q(l, qHS);
  l = t < 0 ? l * (1 + t) : l * (1 - t) + (1 - (1 - t));
  l = q(Math.min(1, Math.max(0, l)), qL);
  return hex(aRgb(h, s, l).map(v => redondeoSalida === 'r' ? Math.round(v * 255) : Math.floor(v * 255 + 1e-9)));
}
const res = [];
for (const M of [240, 255, 256, 100000])
  for (const qHS of ['n', 'r', 't']) for (const qL of ['n', 'r', 't']) for (const so of ['r', 't']) {
    let mal = 0, maxd = 0;
    for (const m of j.tint) { const e = variante(base[m.tema], m.tint, M, qHS, qL, so); const d = Math.max(...[0, 2, 4].map(i => Math.abs(parseInt(e.substr(i, 2), 16) - parseInt(m.rgb.substr(i, 2), 16)))); if (d) mal++; maxd = Math.max(maxd, d); }
    res.push([mal, maxd, M, qHS, qL, so]);
  }
res.sort((a, b) => a[0] - b[0]); console.log(res.slice(0, 8).map(r => 'distintos ' + r[0] + ' máx ' + r[1] + ' HLSMAX ' + r[2] + ' cuantiza H/S ' + r[3] + ' L ' + r[4] + ' salida ' + r[5]).join('\n'));
