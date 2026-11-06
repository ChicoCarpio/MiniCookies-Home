/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Edición de texto sobre el modelo de texto de pptx-modelo.js (Fase 4, 4.3): dado el
// texto nuevo de una forma, devuelve los párrafos nuevos con el formato que le dejaría
// PowerPoint. Es la misma regla que presentacion-editable.js (opTexto) pero sobre el
// modelo ya leído, sin XML: sirve para mostrar el texto con su formato mientras se
// escribe, sin tocar el archivo. Sin DOM.
//
// Texto plano de la forma: los párrafos unidos con \n; un salto de línea suave
// (trozo con salto) cuenta como U+2028.
//
// Reglas (iguales a opTexto):
//   - Los párrafos iguales del principio y del final se devuelven tal cual (el mismo
//     objeto: el motor de texto no los vuelve a medir).
//   - En un párrafo que cambió, lo igual del principio y del final del texto queda
//     como estaba; lo escrito toma el formato del primer carácter reemplazado o, si
//     solo se inserta, el del carácter anterior (al principio del párrafo, el de la
//     primera corrida; sin corridas, el del final de párrafo).
//   - Un párrafo nuevo parte del párrafo que ya estaba en ese lugar (o del anterior).
//
// Uso: const nuevos = EdicionTexto.aplicar(parrafos, nuevoPlano)
//      EdicionTexto.formatearRango(parrafos, ini, fin, cambios)   el mismo formato en una parte del texto: las corridas que el rango corta se parten (cambios: objeto, o función del trozo)
//      EdicionTexto.rangosMarcados(parrafos, 'marca')             [{ ini, fin, valor }] de los trozos con esa propiedad (lo escrito hereda las propiedades del trozo)

const EdicionTexto = (() => {
  const SALTO = ' ';
  const textoTrozo = (t) => t.salto ? SALTO : (t.texto || '');
  const textoParrafo = (p) => p.trozos.map(textoTrozo).join('');
  const plano = (parrafos) => parrafos.map(textoParrafo).join('\n');

  // El formato de un trozo, sin su texto, para fabricar trozos nuevos con él.
  const formatoDe = (t) => { const { texto, salto, campo, ...f } = t; return f; };
  // Como en opTexto, el formato de lo escrito sale solo de corridas: no de un salto ni de un campo (<a:fld>, trozo con campo)
  const esCorrida = (t) => !t.salto && !t.campo;
  const trozoNuevo = (formato, texto) => ({ ...formato, texto, salto: false });
  const trozoSalto = (formato) => ({ ...formato, texto: '', salto: true });
  // Texto con saltos suaves → trozos con ese formato.
  function trozosDe(formato, s) {
    const salida = [];
    s.split(SALTO).forEach((parte, i) => {
      if (i) salida.push(trozoSalto(formato));
      if (parte) salida.push(trozoNuevo(formato, parte));
    });
    return salida;
  }

  function editarParrafo(p, nuevo) {
    const items = p.trozos.map((t) => ({ t, texto: textoTrozo(t) }));
    const viejo = items.map((x) => x.texto).join('');
    let cp = 0; while (cp < viejo.length && cp < nuevo.length && viejo[cp] === nuevo[cp]) cp++;
    let cs = 0; while (cs < viejo.length - cp && cs < nuevo.length - cp && viejo[viejo.length - 1 - cs] === nuevo[nuevo.length - 1 - cs]) cs++;
    const finBorrado = viejo.length - cs, insertado = nuevo.slice(cp, nuevo.length - cs);

    // Formato de lo escrito.
    let base = null, off = 0;
    for (const x of items) {
      const s = off, e = off + x.texto.length; off = e;
      if (esCorrida(x.t) && (finBorrado > cp ? s <= cp && cp < e : s < cp && cp <= e)) base = x;
    }
    if (!base) base = items.find((x) => esCorrida(x.t)) || null;
    const formatoNuevo = base ? formatoDe(base.t) : formatoDe(p.fin);

    const salida = [];
    let colocado = false, pos = 0;
    const ponerInsertado = () => { if (!colocado) { colocado = true; if (insertado) salida.push(...trozosDe(formatoNuevo, insertado)); } };
    for (const x of items) {
      const s = pos, e = pos + x.texto.length; pos = e;
      // se escribe al final de la corrida: sigue siendo una sola
      if (e === cp && x === base && finBorrado === cp && !insertado.includes(SALTO)) { salida.push({ ...x.t, texto: x.texto + insertado }); colocado = true; continue; }
      if (e <= cp) { salida.push(x.t); continue; }
      if (s >= finBorrado) { ponerInsertado(); salida.push(x.t); continue; }
      // el elemento toca la parte cambiada
      if (!x.t.salto) {
        const antes = s < cp ? x.texto.slice(0, cp - s) : '', despues = e > finBorrado ? x.texto.slice(finBorrado - s) : '';
        if (x === base && !insertado.includes(SALTO)) { salida.push({ ...x.t, texto: antes + insertado + despues }); colocado = true; continue; }
        if (antes) salida.push({ ...x.t, texto: antes });
        ponerInsertado();
        if (despues) salida.push({ ...x.t, texto: despues });
      } else ponerInsertado();   // un salto borrado o reemplazado
    }
    ponerInsertado();
    return { ...p, trozos: salida.filter((t) => t.salto || t.texto) };
  }

  function aplicar(parrafos, nuevoPlano) {
    const lineas = String(nuevoPlano).split('\n');
    const viejos = parrafos, textosViejos = viejos.map(textoParrafo);
    const n = lineas.length, m = viejos.length;
    let pre = 0; while (pre < n && pre < m && textosViejos[pre] === lineas[pre]) pre++;
    let suf = 0; while (suf < n - pre && suf < m - pre && textosViejos[m - 1 - suf] === lineas[n - 1 - suf]) suf++;
    return lineas.map((l, k) => {
      if (k < pre) return viejos[k];
      if (k >= n - suf) return viejos[m - n + k];
      if (k - pre < m - suf - pre) return editarParrafo(viejos[k], l);
      // párrafo nuevo: parte del que estaba en ese lugar; si solo se agregaron, del de antes (o del de después si es al principio)
      const modelo = m - suf > pre ? viejos[Math.min(k, m - suf - 1)] : viejos[pre > 0 ? pre - 1 : Math.min(pre, m - 1)];
      const primera = modelo.trozos.find(esCorrida);
      // El archivo se escribe de una vez con el texto final: el párrafo nuevo lleva el formato de la primera corrida del modelo aunque se
      // escriba después de partirlo con Enter; por eso, vacío, su formato de fin de párrafo (el que usará lo que se escriba) es el de esa corrida.
      const formato = formatoDe(primera || modelo.fin);
      return { ...modelo, trozos: l ? trozosDe(formato, l) : [], fin: l ? modelo.fin : { ...modelo.fin, ...formato } };
    });
  }

  function formatearRango(parrafos, ini, fin, cambios) {
    const de = (t) => (typeof cambios === 'function' ? cambios(t) : cambios);
    let base = 0;
    return parrafos.map((p) => {
      const largo = textoParrafo(p).length, lo = Math.max(ini, base) - base, hi = Math.min(fin, base + largo) - base;
      base += largo + 1;
      if (hi <= lo) return p;
      const trozos = [];
      let pos = 0;
      for (const t of p.trozos) {
        const texto = textoTrozo(t), s = pos, e = pos + texto.length; pos = e;
        if (e <= lo || s >= hi || t.salto) { trozos.push(t); continue; }
        if (t.campo) { trozos.push({ ...t, ...de(t) }); continue; }   // un campo es una sola pieza: si el rango lo toca, se formatea entero
        const a = texto.slice(0, Math.max(0, lo - s)), m = texto.slice(Math.max(0, lo - s), hi - s), d = texto.slice(Math.min(texto.length, hi - s));
        if (a) trozos.push({ ...t, texto: a });
        trozos.push({ ...t, ...de(t), texto: m });
        if (d) trozos.push({ ...t, texto: d });
      }
      return { ...p, trozos };
    });
  }
  function rangosMarcados(parrafos, marca) {
    const salida = [];
    let base = 0;
    for (const p of parrafos) {
      let pos = base;
      for (const t of p.trozos) {
        const n = textoTrozo(t).length, v = t[marca];
        if (v && !t.salto) {
          const u = salida[salida.length - 1];
          if (u && u.fin === pos && u.valor === v) u.fin += n; else salida.push({ ini: pos, fin: pos + n, valor: v });
        }
        pos += n;
      }
      base += textoParrafo(p).length + 1;
    }
    return salida;
  }

  return { aplicar, plano, textoParrafo, formatearRango, rangosMarcados };
})();
