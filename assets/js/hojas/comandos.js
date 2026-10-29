/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Registro único de comandos de los editores de oficina (núcleo de la suite, fase 1 del PLAN). La cinta, los menús, la paleta de comandos,
// los atajos de teclado y la barra del celular salen todos de aquí: un comando se define una vez y cada botón solo dice su id.
//
// Uso:
//   const cmd = Comandos.crear();
//   cmd.registrar({ id: 'negrita', texto: 'cmd-negrita', icono: 'B', atajo: 'mod+b', ejecutar: () => …, habilitado: () => true, activo: () => false,
//                   enTexto: false })   // enTexto: el atajo vale también mientras se escribe en un campo
//   cmd.ejecutar('negrita')
//   cmd.conectarTeclado(document)         // los atajos ("mod" = Ctrl en Windows/Linux, ⌘ en Mac)
//   cmd.conectarBotones(raiz)             // cada [data-comando="id"] ejecuta el comando; refrescar() pone disabled y aria-pressed
//   cmd.lista() → los comandos (para la paleta)
//   cmd.textoAtajo('mod+b') → 'Ctrl+B' | '⌘B'
// `texto` es una clave de i18n; se traduce con window.t si existe.
(function () {
  const esMac = () => typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent || '');
  const SIN_TEXTO = ['range', 'checkbox', 'radio', 'button', 'color', 'file', 'submit', 'reset'];
  function enCampoDeTexto(e) {
    const t = e.target;
    if (!t || !t.tagName) return false;
    if (t.isContentEditable || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT') return true;
    return t.tagName === 'INPUT' && !SIN_TEXTO.includes(t.type);
  }
  // 'mod+shift+z' → clave normalizada; también la que produce un evento de teclado
  function normalizar(atajo) {
    const partes = atajo.toLowerCase().split('+');
    const tecla = partes.pop();
    const mods = new Set(partes.map((p) => (p === 'mod' ? (esMac() ? 'meta' : 'ctrl') : p)));
    return ['ctrl', 'meta', 'alt', 'shift'].filter((m) => mods.has(m)).join('+') + (mods.size ? '+' : '') + tecla;
  }
  function claveDeEvento(e) {
    let tecla = e.key.length === 1 ? e.key.toLowerCase() : e.key.toLowerCase();
    if (e.code && /^Key[A-Z]$/.test(e.code) && (e.ctrlKey || e.metaKey || e.altKey)) tecla = e.code.slice(3).toLowerCase();   // Ctrl+Z con teclados no latinos
    if (e.code && /^Digit\d$/.test(e.code) && (e.ctrlKey || e.metaKey)) tecla = e.code.slice(5);
    const mods = [];
    if (e.ctrlKey) mods.push('ctrl'); if (e.metaKey) mods.push('meta'); if (e.altKey) mods.push('alt'); if (e.shiftKey && (e.key.length > 1 || e.ctrlKey || e.metaKey || e.altKey)) mods.push('shift');
    return mods.join('+') + (mods.length ? '+' : '') + tecla;
  }
  function textoAtajo(atajo) {
    if (!atajo) return '';
    const mac = esMac();
    return atajo.split('+').map((p) => {
      const l = p.toLowerCase();
      if (l === 'mod') return mac ? '⌘' : 'Ctrl';
      if (l === 'shift') return mac ? '⇧' : 'Shift';
      if (l === 'alt') return mac ? '⌥' : 'Alt';
      return p.length === 1 ? p.toUpperCase() : p.charAt(0).toUpperCase() + p.slice(1);
    }).join(mac ? '' : '+');
  }
  const traducir = (clave) => (typeof globalThis.t === 'function' ? globalThis.t(clave) : clave);

  function crear() {
    const porId = new Map(), porAtajo = new Map(), alRefrescar = [];
    let raices = [];

    function registrar(c) {
      if (!c.id || typeof c.ejecutar !== 'function') throw new Error('comando sin id o sin ejecutar');
      porId.set(c.id, c);
      for (const a of [].concat(c.atajo || [])) porAtajo.set(normalizar(a), c);
      return c;
    }
    const habilitado = (c) => !c.habilitado || !!c.habilitado();
    function ejecutar(id, ...args) {
      const c = porId.get(id);
      if (!c || !habilitado(c)) return false;
      const r = c.ejecutar(...args);
      refrescar();
      return r === undefined ? true : r;
    }
    function conectarTeclado(destino) {
      const fn = (e) => {
        if (e.defaultPrevented || e.isComposing) return;
        const c = porAtajo.get(claveDeEvento(e));
        if (!c) return;
        if (enCampoDeTexto(e) && !c.enTexto) return;
        if (!habilitado(c)) return;
        e.preventDefault();
        ejecutar(c.id);
      };
      destino.addEventListener('keydown', fn);
      return () => destino.removeEventListener('keydown', fn);
    }
    function conectarBotones(raiz) {
      raices.push(raiz);
      raiz.addEventListener('click', (e) => {
        const b = e.target.closest('[data-comando]');
        if (!b || !raiz.contains(b) || b.disabled) return;
        if (b.tagName === 'SELECT' || b.tagName === 'INPUT') return;   // estos avisan con change/input
        ejecutar(b.dataset.comando, b.dataset.valor);
      });
      raiz.addEventListener('change', (e) => {
        const b = e.target.closest('[data-comando]');
        if (!b || (b.tagName !== 'SELECT' && b.tagName !== 'INPUT')) return;
        ejecutar(b.dataset.comando, b.value);
      });
      for (const b of raiz.querySelectorAll('[data-comando]')) {
        const c = porId.get(b.dataset.comando);
        if (c && c.atajo && !b.title) b.title = traducir(c.texto) + ' (' + textoAtajo([].concat(c.atajo)[0]) + ')';
      }
      refrescar();
    }
    function refrescar() {
      raices = raices.filter((r) => r.isConnected !== false);
      for (const raiz of raices) {
        for (const b of raiz.querySelectorAll('[data-comando]')) {
          const c = porId.get(b.dataset.comando);
          if (!c) continue;
          b.disabled = !habilitado(c);
          if (c.activo) b.setAttribute('aria-pressed', String(!!c.activo()));
          if (c.valor && (b.tagName === 'SELECT' || b.tagName === 'INPUT') && document.activeElement !== b) { const v = c.valor(); if (v !== undefined && v !== null) b.value = v; }
        }
      }
      for (const fn of alRefrescar) fn();
    }
    return {
      registrar, ejecutar, conectarTeclado, conectarBotones, refrescar,
      alRefrescar: (fn) => alRefrescar.push(fn),
      obtener: (id) => porId.get(id) || null,
      lista: () => [...porId.values()].filter((c) => !c.oculto),
      habilitado: (id) => { const c = porId.get(id); return !!c && habilitado(c); },
      textoAtajo,
    };
  }

  globalThis.Comandos = { crear, textoAtajo, normalizar, claveDeEvento, enCampoDeTexto };
})();
