/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Piezas de interfaz de las aplicaciones de oficina (arquetipo G, núcleo de la suite): la cinta con pestañas armada desde una descripción,
// el menú "⋯", la paleta de comandos (Ctrl+K) y la línea de estado honesta. Los botones solo llevan data-comando: lo que hacen, si están
// habilitados y si están apretados lo dice el registro de comandos (comandos.js).
//
//   Oficina.cinta(pestanasEl, cintaEl, descripcion, comandos)
//     descripcion = [{ id, texto, grupos: [{ texto, items: [{ comando, icono, texto?, tipo: 'boton'|'color'|'select', opciones?: [[valor, texto]], solo?: 'pc' }] }] }]
//   Oficina.menu(boton, items, comandos)            items = [{ comando } | '-' ]
//   Oficina.paleta(comandos)                        abre la paleta (también con Ctrl+K si se registra el comando)
//   Oficina.aviso(el, texto, ms)                    un aviso breve en la barra de estado
//   Oficina.montar(contenedor, { nombre, inicio, riel, inspector, extra })   arma el esqueleto de la página del editor (devuelve las referencias)
//   Oficina.riel(el, [{ comando, icono, texto? }], comandos)                 modos de herramienta: una activa a la vez, cada una es un comando
//   Oficina.inspector(el) → { mostrar(secciones), vaciar(), abrir(), cerrar(), alternar(), visible() }   propiedades de lo seleccionado
// Los textos son claves de i18n (window.t).
(function () {
  const tx = (k) => (typeof globalThis.t === 'function' ? globalThis.t(k) : k);
  const crearEl = (tag, props, hijos) => {
    const el = document.createElement(tag);
    for (const k in props || {}) {
      if (k === 'dataset') Object.assign(el.dataset, props[k]);
      else if (k in el && typeof props[k] !== 'string') el[k] = props[k];
      else el.setAttribute(k, props[k]);
    }
    for (const h of hijos || []) el.append(h);
    return el;
  };

  function cinta(pestanasEl, cintaEl, descripcion, comandos) {
    pestanasEl.setAttribute('role', 'tablist');
    let actual = null;
    const paneles = new Map();
    for (const p of descripcion) {
      const boton = crearEl('button', { type: 'button', role: 'tab', id: 'of-tab-' + p.id, 'aria-controls': 'of-panel-' + p.id, 'data-i18n': p.texto }, [tx(p.texto)]);
      boton.addEventListener('click', () => mostrar(p.id));
      pestanasEl.append(boton);
      const panel = crearEl('div', { role: 'tabpanel', id: 'of-panel-' + p.id, 'aria-labelledby': 'of-tab-' + p.id, class: 'of-panel', hidden: true });
      panel.style.display = 'contents';
      for (const g of p.grupos) {
        const items = crearEl('div', { class: 'of-grupo-items' });
        for (const it of g.items) items.append(item(it, comandos));
        const grupo = crearEl('div', { class: 'of-grupo' + (g.solo === 'pc' ? ' of-solo-pc' : ''), role: 'group', 'aria-label': tx(g.texto) }, [items, crearEl('div', { class: 'of-grupo-titulo', 'data-i18n': g.texto }, [tx(g.texto)])]);
        panel.append(grupo);
      }
      cintaEl.append(panel);
      paneles.set(p.id, { boton, panel });
    }
    function mostrar(id) {
      actual = id;
      for (const [k, x] of paneles) { const si = k === id; x.boton.setAttribute('aria-selected', String(si)); x.boton.tabIndex = si ? 0 : -1; x.panel.hidden = !si; x.panel.style.display = si ? 'contents' : 'none'; }
    }
    pestanasEl.addEventListener('keydown', (e) => {
      if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
      const ids = descripcion.map((p) => p.id), k = ids.indexOf(actual);
      const sig = ids[(k + (e.key === 'ArrowRight' ? 1 : ids.length - 1)) % ids.length];
      mostrar(sig); paneles.get(sig).boton.focus(); e.preventDefault();
    });
    mostrar(descripcion[0].id);
    comandos.conectarBotones(cintaEl);
    return { mostrar };
  }

  function item(it, comandos) {
    const c = comandos.obtener(it.comando);
    const etiqueta = tx(it.texto || (c && c.texto) || it.comando);
    if (it.tipo === 'select') {
      const s = crearEl('select', { 'data-comando': it.comando, 'aria-label': etiqueta, title: etiqueta, class: it.solo === 'pc' ? 'of-solo-pc' : '' });
      for (const [v, t] of it.opciones) s.append(crearEl('option', { value: v, 'data-i18n': t }, [tx(t)]));
      return s;
    }
    if (it.tipo === 'color') {
      const muestra = crearEl('span', { class: 'of-muestra' });
      muestra.style.background = it.porDefecto || '#000000';
      const entrada = crearEl('input', { type: 'color', value: it.porDefecto || '#000000', 'aria-label': etiqueta, 'data-comando': it.comando });
      entrada.addEventListener('input', () => { muestra.style.background = entrada.value; });
      const caja = crearEl('label', { class: 'of-color', title: etiqueta }, [crearEl('span', { class: 'of-color-boton' }, [crearEl('span', { class: 'of-icono', 'aria-hidden': 'true' }, [it.icono || '■']), muestra]), entrada]);
      return caja;
    }
    const b = crearEl('button', { type: 'button', 'data-comando': it.comando, 'aria-label': etiqueta, class: it.solo === 'pc' ? 'of-solo-pc' : '' });
    if (it.valor !== undefined) b.dataset.valor = it.valor;
    b.append(crearEl('span', { class: 'of-icono', 'aria-hidden': 'true' }, [it.icono || '•']));
    if (it.conTexto) b.append(crearEl('span', { 'data-i18n': it.texto || (c && c.texto) }, [etiqueta]));
    return b;
  }

  // ---- menú desplegable -------------------------------------------------------------------------------------------
  function menu(boton, items, comandos) {
    const caja = crearEl('div', { class: 'of-menu', role: 'menu', hidden: true });
    document.body.append(caja);
    function armar() {
      caja.textContent = '';
      for (const it of items) {
        if (it === '-') { caja.append(crearEl('hr')); continue; }
        const c = comandos.obtener(it.comando);
        if (!c) continue;
        const b = crearEl('button', { type: 'button', role: 'menuitem', 'data-comando': it.comando, disabled: !comandos.habilitado(it.comando) }, [crearEl('span', {}, [tx(c.texto)]), crearEl('kbd', {}, [c.atajo ? comandos.textoAtajo([].concat(c.atajo)[0]) : ''])]);
        caja.append(b);
      }
    }
    function abrir() {
      armar();
      caja.hidden = false;
      const r = boton.getBoundingClientRect();
      caja.style.top = Math.min(r.bottom + 4, innerHeight - caja.offsetHeight - 8) + 'px';
      caja.style.left = Math.max(8, Math.min(r.right - caja.offsetWidth, innerWidth - caja.offsetWidth - 8)) + 'px';
      boton.setAttribute('aria-expanded', 'true');
      const primero = caja.querySelector('button:not(:disabled)'); if (primero) primero.focus();
    }
    function cerrar(devolver) { if (caja.hidden) return; caja.hidden = true; boton.setAttribute('aria-expanded', 'false'); if (devolver) boton.focus(); }
    boton.setAttribute('aria-haspopup', 'menu'); boton.setAttribute('aria-expanded', 'false');
    boton.addEventListener('click', () => (caja.hidden ? abrir() : cerrar()));
    caja.addEventListener('click', (e) => { const b = e.target.closest('button[data-comando]'); if (!b) return; cerrar(); comandos.ejecutar(b.dataset.comando); });
    caja.addEventListener('keydown', (e) => {
      const lista = [...caja.querySelectorAll('button:not(:disabled)')], k = lista.indexOf(document.activeElement);
      if (e.key === 'Escape') { e.preventDefault(); cerrar(true); }
      else if (e.key === 'ArrowDown') { e.preventDefault(); (lista[k + 1] || lista[0]).focus(); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); (lista[k - 1] || lista[lista.length - 1]).focus(); }
    });
    document.addEventListener('pointerdown', (e) => { if (!caja.hidden && !caja.contains(e.target) && !boton.contains(e.target)) cerrar(); });
    return { abrir, cerrar };
  }

  // ---- paleta de comandos ------------------------------------------------------------------------------------------
  let dialogoPaleta = null;
  function paleta(comandos, alCerrar) {
    if (!dialogoPaleta) {
      dialogoPaleta = crearEl('dialog', { class: 'of-dialogo of-paleta', 'aria-label': tx('of-paleta') });
      dialogoPaleta.append(crearEl('input', { type: 'search', placeholder: tx('of-paleta-buscar'), 'aria-label': tx('of-paleta-buscar'), autocomplete: 'off', spellcheck: 'false' }), crearEl('ul', { role: 'listbox' }));
      document.body.append(dialogoPaleta);
    }
    const entrada = dialogoPaleta.querySelector('input'), lista = dialogoPaleta.querySelector('ul');
    const normal = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
    let visibles = [], marcada = 0;
    function pintar() {
      const q = normal(entrada.value.trim());
      visibles = comandos.lista().filter((c) => comandos.habilitado(c.id) && (!q || normal(tx(c.texto)).includes(q) || c.id.includes(q))).slice(0, 60);
      marcada = Math.min(marcada, Math.max(0, visibles.length - 1));
      lista.textContent = '';
      visibles.forEach((c, i) => {
        const li = crearEl('li', { role: 'option', 'aria-selected': String(i === marcada) }, [crearEl('span', {}, [tx(c.texto)]), crearEl('kbd', {}, [c.atajo ? comandos.textoAtajo([].concat(c.atajo)[0]) : ''])]);
        li.addEventListener('click', () => elegir(i));
        lista.append(li);
      });
    }
    function elegir(i) { const c = visibles[i]; dialogoPaleta.close(); if (c) comandos.ejecutar(c.id); }
    entrada.oninput = () => { marcada = 0; pintar(); };
    entrada.onkeydown = (e) => {
      if (e.key === 'ArrowDown') { e.preventDefault(); marcada = Math.min(visibles.length - 1, marcada + 1); pintar(); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); marcada = Math.max(0, marcada - 1); pintar(); }
      else if (e.key === 'Enter') { e.preventDefault(); elegir(marcada); }
    };
    dialogoPaleta.onclose = () => { if (alCerrar) alCerrar(); };
    entrada.value = ''; marcada = 0; pintar();
    dialogoPaleta.showModal();
    entrada.focus();
  }

  function aviso(el, texto, ms) {
    el.textContent = texto; el.classList.add('of-aviso');
    clearTimeout(el._tiempo);
    el._tiempo = setTimeout(() => { el.textContent = ''; el.classList.remove('of-aviso'); }, ms || 4000);
  }

  // ---- esqueleto de la aplicación (arquetipo G) -------------------------------------------------------------------------
  // Arma la página del editor abierto: barra superior, pestañas y cinta, [extra], centro (riel de herramientas, lienzo, inspector) y barra
  // de estado. Los ids son los de siempre (app, nombreDoc, estadoGuardado, botonModo, botonMas, pestanas, cinta, inspector, lineaHonesta,
  // avisoEstado) para que el código de cada aplicación los busque igual. Devuelve las referencias.
  //   Oficina.montar(contenedor, { nombre: 'nombre-app', inicio: './', riel: false, inspector: true, extra: Node, ayudaInicio: 'volver-inicio' })
  function montar(contenedor, op) {
    op = Object.assign({ nombre: 'nombre-app', inicio: './', riel: false, inspector: true, extra: null }, op);
    contenedor.classList.add('of-app');
    contenedor.textContent = '';
    const icono = (t) => crearEl('span', { class: 'of-icono', 'aria-hidden': 'true' }, [t]);
    const boton = (id, comando, hijos, clase) => crearEl('button', Object.assign({ type: 'button', 'data-comando': comando, class: 'of-boton' + (clase ? ' ' + clase : '') }, id ? { id } : {}), hijos);
    const marca = crearEl('a', { class: 'of-marca', href: op.inicio, id: 'enlaceInicio', 'data-i18n-attr': 'title:volver-inicio', title: tx('volver-inicio') }, [crearEl('span', {}, ['Mini Cookies']), ' ', crearEl('b', { 'data-i18n': op.nombre }, [tx(op.nombre)])]);
    const nombre = crearEl('input', { class: 'of-nombre', id: 'nombreDoc', 'aria-label': tx('nombre-archivo'), 'data-i18n-attr': 'aria-label:nombre-archivo', spellcheck: 'false' });
    const guardado = crearEl('span', { class: 'of-guardado', id: 'estadoGuardado', 'aria-live': 'polite' });
    const botonMas = crearEl('button', { type: 'button', id: 'botonMas', class: 'of-boton', 'aria-label': tx('mas-opciones'), 'data-i18n-attr': 'aria-label:mas-opciones' }, [icono('⋯')]);
    const botonModo = boton('botonModo', 'modoEdicion', [], 'of-solo-celular');
    const acciones = crearEl('div', { class: 'of-acciones' }, [
      botonModo, boton(null, 'deshacer', [icono('↶')]), boton(null, 'rehacer', [icono('↷')]),
      boton(null, 'descargar', [icono('⤓'), crearEl('span', { class: 'of-solo-pc', 'data-i18n': 'cmd-descargar-corto' }, [tx('cmd-descargar-corto')])], 'of-principal'), botonMas,
    ]);
    const barra = crearEl('header', { class: 'of-barra' }, [marca, crearEl('div', { class: 'of-documento' }, [nombre, guardado]), acciones]);
    const pestanas = crearEl('nav', { class: 'of-pestanas', id: 'pestanas', 'aria-label': tx('cinta'), 'data-i18n-attr': 'aria-label:cinta' });
    const cintaEl = crearEl('div', { class: 'of-cinta', id: 'cinta' });
    const lienzo = crearEl('div', { class: 'of-lienzo' });
    const centro = crearEl('div', { class: 'of-centro' }, [lienzo]);
    let rielEl = null, inspectorEl = null;
    if (op.riel) { rielEl = crearEl('nav', { class: 'of-riel', id: 'riel', role: 'toolbar', 'aria-orientation': 'vertical', 'aria-label': tx('of-herramientas'), 'data-i18n-attr': 'aria-label:of-herramientas' }); centro.prepend(rielEl); }
    if (op.inspector) { inspectorEl = crearEl('aside', { class: 'of-inspector', id: 'inspector', 'aria-label': tx('inspector'), 'data-i18n-attr': 'aria-label:inspector' }); centro.append(inspectorEl); }
    const honesta = crearEl('span', { class: 'of-honesta', id: 'lineaHonesta' });
    const avisoEl = crearEl('span', { id: 'avisoEstado', role: 'status', 'aria-live': 'polite' });
    contenedor.append(barra, pestanas, cintaEl);
    if (op.extra) contenedor.append(op.extra);
    contenedor.append(centro, crearEl('footer', { class: 'of-estado' }, [honesta, avisoEl]));
    return { raiz: contenedor, barra, nombre, guardado, botonMas, botonModo, pestanas, cinta: cintaEl, lienzo, centro, riel: rielEl, inspector: inspectorEl, honesta, aviso: avisoEl };
  }

  // ---- riel de modos de herramienta -----------------------------------------------------------------------------------------
  // Una columna de iconos (en el celular, una tira debajo del lienzo) donde solo una herramienta está activa a la vez. Cada una es un comando:
  // su `activo()` dice cuál está apretada, así que cambiar de modo por teclado, por la paleta o con el riel da el mismo resultado.
  //   Oficina.riel(el, [{ comando, icono, texto? }], comandos)
  function riel(el, herramientas, comandos) {
    el.textContent = '';
    for (const h of herramientas) {
      const c = comandos.obtener(h.comando);
      const etiqueta = tx(h.texto || (c && c.texto) || h.comando);
      el.append(crearEl('button', { type: 'button', class: 'of-herramienta', 'data-comando': h.comando, 'aria-label': etiqueta, title: etiqueta }, [crearEl('span', { class: 'of-icono', 'aria-hidden': 'true' }, [h.icono || '•']), crearEl('span', { class: 'of-herramienta-texto', 'data-i18n': h.texto || (c && c.texto) || h.comando }, [etiqueta])]));
    }
    el.addEventListener('keydown', (e) => {
      const lista = [...el.querySelectorAll('button:not(:disabled)')], k = lista.indexOf(document.activeElement);
      const vertical = el.getAttribute('aria-orientation') === 'vertical' && matchMedia('(min-width: 901px)').matches;
      const sig = vertical ? 'ArrowDown' : 'ArrowRight', ant = vertical ? 'ArrowUp' : 'ArrowLeft';
      if (k < 0 || (e.key !== sig && e.key !== ant)) return;
      e.preventDefault();
      lista[(k + (e.key === sig ? 1 : lista.length - 1)) % lista.length].focus();
    });
    comandos.conectarBotones(el);
    return { el };
  }

  // ---- inspector contextual ---------------------------------------------------------------------------------------------------
  // Muestra solo las propiedades de lo seleccionado. En escritorio es la columna de la derecha; en el celular es una hoja que sube desde abajo
  // (se abre con el comando, se cierra con su botón o con Escape): ningún gesto es la única forma de hacerlo.
  //   const insp = Oficina.inspector(el); insp.mostrar([{ titulo, filas: [[clave_i18n, valor], …] }]); insp.vaciar(); insp.alternar(); insp.visible()
  function inspector(el) {
    const celular = () => matchMedia('(max-width: 900px)').matches;
    let abiertoCelular = false, cerradoPc = false, secciones = [];
    const cerrar = crearEl('button', { type: 'button', class: 'of-inspector-cerrar', 'aria-label': tx('of-cerrar'), 'data-i18n-attr': 'aria-label:of-cerrar' }, ['✕']);
    cerrar.addEventListener('click', () => poner(false));
    el.addEventListener('keydown', (e) => { if (e.key === 'Escape' && celular()) { poner(false); e.stopPropagation(); } });
    function visible() { return celular() ? abiertoCelular : !cerradoPc; }
    function poner(si) {
      if (celular()) { abiertoCelular = si; el.classList.toggle('of-hoja-abierta', si); if (si) cerrar.focus(); } else { cerradoPc = !si; el.hidden = !si; }
    }
    function pintar() {
      el.textContent = '';
      el.append(cerrar);
      secciones.forEach((s, i) => {
        const titulo = crearEl('h2', {}, [s.titulo]);
        if (i) titulo.className = 'of-inspector-sep';
        el.append(titulo);
        const dl = crearEl('dl');
        for (const [k, v] of s.filas) dl.append(crearEl('dt', {}, [tx(k)]), crearEl('dd', {}, [String(v)]));
        el.append(dl);
      });
    }
    addEventListener('resize', () => { el.classList.toggle('of-hoja-abierta', celular() && abiertoCelular); el.hidden = !celular() && cerradoPc; });
    return {
      mostrar(s) { secciones = s; pintar(); },
      vaciar() { secciones = []; pintar(); },
      abrir: () => poner(true), cerrar: () => poner(false), alternar: () => poner(!visible()), visible,
    };
  }

  // textos comunes de la interfaz G
  const pagina = globalThis.i18nPagina || (globalThis.i18nPagina = {});
  pagina.es = Object.assign({ 'of-paleta': 'Paleta de comandos', 'of-paleta-buscar': 'Buscar un comando…', 'of-herramientas': 'Herramientas', 'of-cerrar': 'Cerrar' }, pagina.es);
  pagina.en = Object.assign({ 'of-paleta': 'Command palette', 'of-paleta-buscar': 'Search for a command…', 'of-herramientas': 'Tools', 'of-cerrar': 'Close' }, pagina.en);

  globalThis.Oficina = { cinta, menu, paleta, aviso, crearEl, montar, riel, inspector };
})();
