/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Geometría de las formas de PowerPoint: evalúa las fórmulas de DrawingML
// (avLst, gdLst, pathLst) y devuelve los trazados ya medidos, con los arcos
// pasados a curvas Bézier (así el mismo trazado sirve para Canvas y para PDF).
//
// Las ~186 formas predefinidas (rect, roundRect, flechas, llamadas…) vienen de
// la norma, convertidas sin cambios: vendor/ecma-376-geometrias/5ed-2016/
// geometrias.js (window.GeometriasEcma). Las formas libres (custGeom) traen
// sus propias fórmulas en el XML y se evalúan igual.
//
// LICENCIA: esas formas están bajo el aviso de copyright de Ecma, aprobado
// por el usuario con la condición de cumplirlo al pie de la letra (CLAUDE.md
// de la raíz, "Formas de PowerPoint de ECMA-376"). No se edita geometrias.js:
// cualquier corrección de una forma va acá, en código propio.
//
// Uso:
//   const def = PptxGeometria.predefinida('roundRect') | PptxGeometria.desdeXml(custGeom)
//   const { rutas, texto } = PptxGeometria.trazar(def, ancho, alto, ajustes)
//   rutas: [{ cmds: [['M',x,y], ['L',x,y], ['C',x1,y1,x2,y2,x,y], ['Z']], relleno, trazo }]
//   texto: { l, t, r, b } — el rectángulo del texto, en las mismas unidades.
// Unidades: las de ancho/alto (las fórmulas son proporciones, salvo las
// rutas de formas libres, que se escalan desde su propio w/h).

const PptxGeometria = (() => {
  const GRADO = 60000; // los ángulos de DrawingML van en 1/60000 de grado
  const aRad = (a) => (a / GRADO) * Math.PI / 180;
  const aAng = (r) => (r * 180 / Math.PI) * GRADO;

  function variablesBase(w, h) {
    const ss = Math.min(w, h), ls = Math.max(w, h);
    const v = {
      w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, ss, ls,
      cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000,
    };
    for (const d of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) { v['wd' + d] = w / d; v['hd' + d] = h / d; v['ssd' + d] = ss / d; }
    return v;
  }

  function evaluarFormula(fmla, v) {
    const p = fmla.trim().split(/\s+/);
    const x = (i) => { const s = p[i]; if (s == null) return 0; if (s in v) return v[s]; const n = Number(s); return Number.isFinite(n) ? n : 0; };
    switch (p[0]) {
      case '*/': return x(3) === 0 ? 0 : (x(1) * x(2)) / x(3);
      case '+-': return x(1) + x(2) - x(3);
      case '+/': return x(3) === 0 ? 0 : (x(1) + x(2)) / x(3);
      case '?:': return x(1) > 0 ? x(2) : x(3);
      case 'abs': return Math.abs(x(1));
      case 'at2': return aAng(Math.atan2(x(2), x(1)));
      case 'cat2': return x(1) * Math.cos(Math.atan2(x(3), x(2)));
      case 'sat2': return x(1) * Math.sin(Math.atan2(x(3), x(2)));
      case 'cos': return x(1) * Math.cos(aRad(x(2)));
      case 'sin': return x(1) * Math.sin(aRad(x(2)));
      case 'tan': return x(1) * Math.tan(aRad(x(2)));
      case 'max': return Math.max(x(1), x(2));
      case 'min': return Math.min(x(1), x(2));
      case 'mod': return Math.sqrt(x(1) * x(1) + x(2) * x(2) + x(3) * x(3));
      case 'pin': return x(2) < x(1) ? x(1) : x(2) > x(3) ? x(3) : x(2);
      case 'sqrt': return Math.sqrt(Math.max(0, x(1)));
      case 'val': return x(1);
      default: return 0;
    }
  }

  // Un arco de elipse como curvas Bézier (de a 90° como máximo).
  function arcoABezier(cx, cy, rx, ry, t0, dt, cmds) {
    const partes = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9));
    const paso = dt / partes;
    const k = (4 / 3) * Math.tan(paso / 4);
    let t = t0;
    for (let i = 0; i < partes; i++) {
      const c0 = Math.cos(t), s0 = Math.sin(t), c1 = Math.cos(t + paso), s1 = Math.sin(t + paso);
      cmds.push(['C',
        cx + rx * (c0 - k * s0), cy + ry * (s0 + k * c0),
        cx + rx * (c1 + k * s1), cy + ry * (s1 - k * c1),
        cx + rx * c1, cy + ry * s1]);
      t += paso;
    }
  }
  // El ángulo de DrawingML es el visual (el que se ve sobre la elipse); para
  // dibujarla hace falta el paramétrico.
  const parametrico = (ang, rx, ry) => Math.atan2(rx * Math.sin(ang), ry * Math.cos(ang));

  function trazar(def, w, h, ajustes = {}) {
    const v = variablesBase(w, h);
    for (const [nombre, fmla] of def.av || []) v[nombre] = nombre in ajustes ? ajustes[nombre] : evaluarFormula(fmla, v);
    for (const [nombre, fmla] of def.gd || []) v[nombre] = evaluarFormula(fmla, v);
    const val = (s) => { if (s in v) return v[s]; const n = Number(s); return Number.isFinite(n) ? n : 0; };

    const rutas = [];
    for (const ruta of def.p || []) {
      const sx = ruta.w ? w / ruta.w : 1, sy = ruta.h ? h / ruta.h : 1;
      const cmds = [];
      let x = 0, y = 0, x0 = 0, y0 = 0;
      for (const c of ruta.c) {
        switch (c[0]) {
          case 'M': x = x0 = val(c[1]) * sx; y = y0 = val(c[2]) * sy; cmds.push(['M', x, y]); break;
          case 'L': x = val(c[1]) * sx; y = val(c[2]) * sy; cmds.push(['L', x, y]); break;
          case 'Q': {
            const qx = val(c[1]) * sx, qy = val(c[2]) * sy, ex = val(c[3]) * sx, ey = val(c[4]) * sy;
            cmds.push(['C', x + (2 / 3) * (qx - x), y + (2 / 3) * (qy - y), ex + (2 / 3) * (qx - ex), ey + (2 / 3) * (qy - ey), ex, ey]);
            x = ex; y = ey; break;
          }
          case 'C': cmds.push(['C', ...c.slice(1).map((s, i) => val(s) * (i % 2 ? sy : sx))]); x = cmds[cmds.length - 1][5]; y = cmds[cmds.length - 1][6]; break;
          case 'A': {
            const rx = val(c[1]) * sx, ry = val(c[2]) * sy, st = aRad(val(c[3])), sw = aRad(val(c[4]));
            if (!cmds.length) cmds.push(['M', x, y]);
            if (rx === 0 || ry === 0 || sw === 0) break;
            const t0 = parametrico(st, rx, ry);
            // Vueltas completas aparte; el resto, llevado al paramétrico con el
            // mismo sentido de giro.
            const vueltas = Math.trunc(sw / (2 * Math.PI));
            const resto = sw - vueltas * 2 * Math.PI;
            let dt = 0;
            if (Math.abs(resto) > 1e-9) {
              dt = parametrico(st + resto, rx, ry) - t0;
              if (resto > 0) { while (dt <= 1e-9) dt += 2 * Math.PI; } else { while (dt >= -1e-9) dt -= 2 * Math.PI; }
            }
            dt += vueltas * 2 * Math.PI;
            const cx = x - rx * Math.cos(t0), cy = y - ry * Math.sin(t0);
            arcoABezier(cx, cy, rx, ry, t0, dt, cmds);
            x = cx + rx * Math.cos(t0 + dt); y = cy + ry * Math.sin(t0 + dt);
            break;
          }
          case 'Z': cmds.push(['Z']); x = x0; y = y0; break;
        }
      }
      rutas.push({ cmds, relleno: ruta.fill || 'norm', trazo: !(ruta.stroke === 'false' || ruta.stroke === '0' || ruta.stroke === 0) });
    }
    const r = def.rect;
    const texto = r ? { l: val(r[0]), t: val(r[1]), r: val(r[2]), b: val(r[3]) } : { l: 0, t: 0, r: w, b: h };
    return { rutas, texto };
  }

  function predefinida(nombre) {
    const g = typeof GeometriasEcma !== 'undefined' ? GeometriasEcma : null;
    return (g && g[nombre]) || null;
  }

  // custGeom (o prstGeom con avLst) del XML → la misma forma que las predefinidas.
  const LETRAS = { moveTo: 'M', lnTo: 'L', arcTo: 'A', quadBezTo: 'Q', cubicBezTo: 'C', close: 'Z' };
  function desdeXml(custGeom) {
    const hijo = (el, n) => el && [...el.children].find((h) => h.localName === n);
    const guias = (n) => { const l = hijo(custGeom, n); return l ? [...l.children].filter((g) => g.localName === 'gd').map((g) => [g.getAttribute('name'), g.getAttribute('fmla')]) : []; };
    const rect = hijo(custGeom, 'rect');
    const pathLst = hijo(custGeom, 'pathLst');
    return {
      av: guias('avLst'), gd: guias('gdLst'),
      rect: rect ? ['l', 't', 'r', 'b'].map((a) => rect.getAttribute(a)) : null,
      p: (pathLst ? [...pathLst.children].filter((p) => p.localName === 'path') : []).map((p) => ({
        w: +p.getAttribute('w') || 0, h: +p.getAttribute('h') || 0,
        fill: p.getAttribute('fill') || 'norm', stroke: p.getAttribute('stroke') || 'true',
        c: [...p.children].filter((c) => LETRAS[c.localName]).map((c) => {
          const l = LETRAS[c.localName];
          if (l === 'A') return [l, c.getAttribute('wR'), c.getAttribute('hR'), c.getAttribute('stAng'), c.getAttribute('swAng')];
          return [l, ...[...c.children].flatMap((pt) => [pt.getAttribute('x'), pt.getAttribute('y')])];
        }),
      })),
    };
  }
  // Los ajustes de una forma (avLst de su prstGeom): { adj: 25000, … }.
  function ajustesDeXml(geom) {
    const av = geom && [...geom.children].find((h) => h.localName === 'avLst');
    const salida = {};
    for (const gd of av ? [...av.children] : []) {
      const m = /^\s*val\s+(-?[\d.]+)/.exec(gd.getAttribute('fmla') || '');
      if (m) salida[gd.getAttribute('name')] = Number(m[1]);
    }
    return salida;
  }

  return { trazar, predefinida, desdeXml, ajustesDeXml, evaluarFormula };
})();
