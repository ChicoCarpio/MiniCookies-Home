/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Diálogos de las aplicaciones de oficina (arquetipo G, núcleo de la suite): un formulario armado desde una descripción de campos y una lista
// con acciones por fila. Un solo <dialog class="of-dialogo"> por llamada; se cierra con Escape o con los botones y devuelve una promesa.
// Los textos llegan ya traducidos (quien llama usa su función de i18n).
//
//   await Dialogo.formulario({ titulo, campos, aceptar: 'Aceptar', cancelar: 'Cancelar' }) → { nombre: valor, … } | null
//     campo = { nombre, etiqueta, tipo: 'texto'|'numero'|'select'|'color'|'check'|'area', valor, opciones: [[valor, texto]],
//               visible: (valores) => bool, ayuda, requerido }
//   await Dialogo.lista({ titulo, vacio, filas: [{ texto, detalle, acciones: [{ valor, texto }] }], botones: [{ valor, texto }], cerrar })
//     → { accion, indice } | null        (indice = -1 para los botones de abajo)
(function () {
  const crear = (tag, props, hijos) => {
    const el = document.createElement(tag);
    for (const k in props || {}) { if (k === 'texto') el.textContent = props[k]; else if (k in el && typeof props[k] !== 'string') el[k] = props[k]; else el.setAttribute(k, props[k]); }
    for (const h of hijos || []) if (h) el.append(h);
    return el;
  };
  let contador = 0;
  function abrir(dlg, alCerrar) {
    document.body.append(dlg);
    return new Promise((resolver) => {
      let resultado = null;
      dlg.addEventListener('close', () => { dlg.remove(); resolver(resultado); if (alCerrar) alCerrar(); });
      dlg._fin = (r) => { resultado = r; dlg.close(); };
      dlg.showModal();
      const primero = dlg.querySelector('input:not([type=hidden]), select, textarea, button');
      if (primero) primero.focus();
    });
  }

  function formulario(o) {
    const dlg = crear('dialog', { class: 'of-dialogo', 'aria-label': o.titulo });
    const form = crear('form', { class: 'of-formulario', method: 'dialog', novalidate: true });
    form.append(crear('h2', { texto: o.titulo }));
    const controles = new Map(), filas = new Map();
    for (const c of o.campos) {
      const id = 'of-campo-' + (++contador);
      let ctl;
      if (c.tipo === 'select') { ctl = crear('select', { id, name: c.nombre }); for (const [v, t] of c.opciones || []) ctl.append(crear('option', { value: String(v), texto: t })); ctl.value = c.valor === undefined || c.valor === null ? String((c.opciones[0] || [''])[0]) : String(c.valor); }
      else if (c.tipo === 'check') { ctl = crear('input', { id, name: c.nombre, type: 'checkbox' }); ctl.checked = !!c.valor; }
      else if (c.tipo === 'area') { ctl = crear('textarea', { id, name: c.nombre, rows: 3 }); ctl.value = c.valor || ''; }
      else { ctl = crear('input', { id, name: c.nombre, type: c.tipo === 'numero' ? 'number' : c.tipo === 'color' ? 'color' : 'text', autocomplete: 'off', spellcheck: false }); if (c.tipo === 'numero') ctl.step = 'any'; ctl.value = c.valor === undefined || c.valor === null ? (c.tipo === 'color' ? '#000000' : '') : String(c.valor); }
      if (c.requerido) ctl.required = true;
      controles.set(c.nombre, ctl);
      const etiqueta = crear('label', { for: id, class: c.tipo === 'check' ? 'of-campo of-campo-check' : 'of-campo' });
      if (c.tipo === 'check') etiqueta.append(ctl, crear('span', { texto: c.etiqueta }));
      else etiqueta.append(crear('span', { texto: c.etiqueta }), ctl);
      if (c.ayuda) etiqueta.append(crear('small', { texto: c.ayuda }));
      filas.set(c.nombre, etiqueta);
      form.append(etiqueta);
    }
    const valores = () => { const v = {}; for (const c of o.campos) { const ctl = controles.get(c.nombre); v[c.nombre] = c.tipo === 'check' ? ctl.checked : c.tipo === 'numero' ? (ctl.value === '' ? null : +ctl.value) : c.tipo === 'color' ? ctl.value.replace(/^#/, '').toUpperCase() : ctl.value; } return v; };
    const refrescar = () => { const v = valores(); for (const c of o.campos) if (c.visible) filas.get(c.nombre).hidden = !c.visible(v); };
    form.addEventListener('input', refrescar); form.addEventListener('change', refrescar);
    refrescar();
    const error = crear('p', { class: 'of-formulario-error', role: 'alert' });
    form.append(error);
    const aceptar = crear('button', { type: 'submit', class: 'boton', value: 'aceptar', texto: o.aceptar || 'OK' });
    const cancelar = crear('button', { type: 'button', class: 'boton boton-secundario', value: 'cancelar', texto: o.cancelar || 'Cancel' });
    form.append(crear('div', { class: 'of-formulario-botones' }, [aceptar, cancelar]));
    cancelar.addEventListener('click', () => dlg._fin(null));
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const v = valores();
      const falta = o.campos.find((c) => c.requerido && (!c.visible || c.visible(v)) && (v[c.nombre] === '' || v[c.nombre] === null));
      if (falta) { error.textContent = falta.etiqueta; controles.get(falta.nombre).focus(); return; }
      const malo = o.validar ? o.validar(v) : null;
      if (malo) { error.textContent = malo; return; }
      dlg._fin(v);
    });
    dlg.append(form);
    return abrir(dlg, o.alCerrar);
  }

  function lista(o) {
    const dlg = crear('dialog', { class: 'of-dialogo', 'aria-label': o.titulo });
    const caja = crear('div', { class: 'of-formulario' });
    caja.append(crear('h2', { texto: o.titulo }));
    const ul = crear('ul', { class: 'of-lista' });
    if (!o.filas.length && o.vacio) ul.append(crear('li', { class: 'of-lista-vacia', texto: o.vacio }));
    o.filas.forEach((f, i) => {
      const li = crear('li', {}, [crear('div', { class: 'of-lista-texto' }, [crear('strong', { texto: f.texto }), f.detalle ? crear('span', { texto: f.detalle }) : null])]);
      const acciones = crear('div', { class: 'of-lista-acciones' });
      for (const a of f.acciones || []) { const b = crear('button', { type: 'button', class: 'boton boton-secundario', texto: a.texto, disabled: !!a.deshabilitado }); b.addEventListener('click', () => dlg._fin({ accion: a.valor, indice: i })); acciones.append(b); }
      li.append(acciones);
      ul.append(li);
    });
    caja.append(ul);
    const botones = crear('div', { class: 'of-formulario-botones' });
    for (const b of o.botones || []) { const el = crear('button', { type: 'button', class: 'boton', texto: b.texto }); el.addEventListener('click', () => dlg._fin({ accion: b.valor, indice: -1 })); botones.append(el); }
    const cerrar = crear('button', { type: 'button', class: 'boton boton-secundario', texto: o.cerrar || 'Close' });
    cerrar.addEventListener('click', () => dlg._fin(null));
    botones.append(cerrar);
    caja.append(botones);
    dlg.append(caja);
    return abrir(dlg, o.alCerrar);
  }

  globalThis.Dialogo = { formulario, lista };
})();
