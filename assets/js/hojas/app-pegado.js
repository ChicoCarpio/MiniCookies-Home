/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Mini Cookies Hojas: la interfaz del pegado especial. El menú «Opciones de pegado» de la cinta, la ventana «Pegado especial», el botón que aparece junto a lo pegado y
// la brocha de «Copiar formato». Se engancha a app.js como extensión; lo que hace cada tipo está en pegado-especial.js.
(function () {
  'use strict';
  const ES = {
    'cmd-pegar-menu': 'Opciones de pegado', 'cmd-pegar-especial': 'Pegado especial…', 'cmd-pegar-formulas': 'Fórmulas', 'cmd-pegar-formulas-num': 'Fórmulas y formato de números', 'cmd-pegar-formatos': 'Formatos',
    'cmd-pegar-sin-bordes': 'Todo excepto bordes', 'cmd-pegar-anchos': 'Anchos de columna', 'cmd-pegar-transponer': 'Transponer', 'cmd-pegar-valores-num': 'Valores y formato de números', 'cmd-pegar-vinculo': 'Pegar vínculo',
    'cmd-pegar-notas': 'Notas y comentarios', 'cmd-pegar-validacion': 'Validación', 'cmd-pegar-fusion': 'Combinar formato condicional', 'cmd-pegar-todo': 'Pegar', 'cmd-pegar-texto': 'Solo texto', 'cmd-copiar-formato': 'Copiar formato',
    'pe-titulo': 'Pegado especial', 'pe-pegar': 'Pegar', 'pe-operacion': 'Operación', 'pe-ninguna': 'Ninguna', 'pe-sumar': 'Sumar', 'pe-restar': 'Restar', 'pe-multiplicar': 'Multiplicar', 'pe-dividir': 'Dividir', 'pe-blancos': 'Omitir celdas en blanco', 'pe-transponer': 'Transponer',
    'pe-todo': 'Todo', 'pe-valores': 'Valores', 'pe-aceptar': 'Aceptar', 'pe-cancelar': 'Cancelar', 'pe-como': 'Pegar como', 'pe-con-formato': 'Con su formato', 'pe-solo-texto': 'Solo texto',
    'aviso-pegar-externo': 'Ese tipo de pegado es para lo copiado aquí en Hojas; desde otro programa se pega con su formato o solo el texto',
    'aviso-pegar-limite': 'No se puede pegar aquí: el área copiada y el área de pegado no tienen el mismo tamaño y forma',
    'aviso-pegar-solape': 'Al transponer, el área copiada y el área de pegado no pueden superponerse',
    'aviso-pegar-combinada': 'No se puede pegar aquí: se partiría una celda combinada',
    'aviso-pegar-matriz': 'No se puede cambiar una parte de una matriz',
    'aviso-pegar-sin-copia': 'Primero copia algo',
    'aviso-brocha': 'Elige las celdas a las que quieres dar ese formato. Esc cancela',
    'aviso-brocha-fija': 'Copiar formato queda activo hasta que pulses Esc o el botón otra vez',
  };
  const EN = {
    'cmd-pegar-menu': 'Paste options', 'cmd-pegar-especial': 'Paste special…', 'cmd-pegar-formulas': 'Formulas', 'cmd-pegar-formulas-num': 'Formulas and number formatting', 'cmd-pegar-formatos': 'Formatting',
    'cmd-pegar-sin-bordes': 'All except borders', 'cmd-pegar-anchos': 'Column widths', 'cmd-pegar-transponer': 'Transpose', 'cmd-pegar-valores-num': 'Values and number formatting', 'cmd-pegar-vinculo': 'Paste link',
    'cmd-pegar-notas': 'Notes and comments', 'cmd-pegar-validacion': 'Validation', 'cmd-pegar-fusion': 'Merge conditional formatting', 'cmd-pegar-todo': 'Paste', 'cmd-pegar-texto': 'Text only', 'cmd-copiar-formato': 'Format painter',
    'pe-titulo': 'Paste special', 'pe-pegar': 'Paste', 'pe-operacion': 'Operation', 'pe-ninguna': 'None', 'pe-sumar': 'Add', 'pe-restar': 'Subtract', 'pe-multiplicar': 'Multiply', 'pe-dividir': 'Divide', 'pe-blancos': 'Skip blanks', 'pe-transponer': 'Transpose',
    'pe-todo': 'All', 'pe-valores': 'Values', 'pe-aceptar': 'OK', 'pe-cancelar': 'Cancel', 'pe-como': 'Paste as', 'pe-con-formato': 'With its formatting', 'pe-solo-texto': 'Text only',
    'aviso-pegar-externo': 'That kind of paste is for what you copied here in Sheets; from another program you can paste with its formatting or as text only',
    'aviso-pegar-limite': 'You can’t paste here: the copy area and the paste area aren’t the same size and shape',
    'aviso-pegar-solape': 'When transposing, the copy area and the paste area can’t overlap',
    'aviso-pegar-combinada': 'You can’t paste here: a merged cell would be split',
    'aviso-pegar-matriz': 'You can’t change part of an array',
    'aviso-pegar-sin-copia': 'Copy something first',
    'aviso-brocha': 'Pick the cells that should get this formatting. Esc cancels',
    'aviso-brocha-fija': 'Format painter stays on until you press Esc or the button again',
  };
  const pagina = window.i18nPagina || (window.i18nPagina = {});
  pagina.es = Object.assign(pagina.es || {}, ES);
  pagina.en = Object.assign(pagina.en || {}, EN);

  (window.HojasExtensiones = window.HojasExtensiones || []).push((A) => {
    const { tx, reg, comandos, avisar } = A;
    const internos = [
      ['pegarFormulas', 'cmd-pegar-formulas', 'formulas'], ['pegarFormulasNum', 'cmd-pegar-formulas-num', 'formulasNum'], ['pegarFormatos', 'cmd-pegar-formatos', 'formatos'], ['pegarSinBordes', 'cmd-pegar-sin-bordes', 'sinBordes'],
      ['pegarAnchos', 'cmd-pegar-anchos', 'anchos'], ['pegarValoresNum', 'cmd-pegar-valores-num', 'valoresNum'], ['pegarNotas', 'cmd-pegar-notas', 'notas'], ['pegarValidacion', 'cmd-pegar-validacion', 'validacion'], ['pegarFusion', 'cmd-pegar-fusion', 'fusionCF'],
    ];
    // desde el botón que aparece junto a lo pegado se cambia lo que ya se pegó; desde la cinta se pega de nuevo
    let cambiar = false;
    const pegarCon = (o) => {
      const repetir = cambiar;
      cambiar = false;
      return repetir ? A.repegar(o) : A.pegarDesdeBoton(o);
    };
    for (const [id, texto, tipo] of internos) reg({ id, texto, ejecutar: () => pegarCon({ tipo }) });
    reg({ id: 'pegarTodo', texto: 'cmd-pegar-todo', ejecutar: () => pegarCon({ tipo: 'todo' }) });
    reg({ id: 'pegarSoloValores', texto: 'cmd-pegar-valores', ejecutar: () => pegarCon({ tipo: 'valores' }) });
    reg({ id: 'pegarTranspuesto', texto: 'cmd-pegar-transponer', ejecutar: () => pegarCon({ tipo: 'todo', trans: true }) });
    reg({ id: 'pegarVinculo', texto: 'cmd-pegar-vinculo', ejecutar: () => pegarCon({ tipo: 'todo', vinculo: true }) });
    reg({ id: 'pegarSoloTexto', texto: 'cmd-pegar-texto', ejecutar: () => pegarCon({ tipo: 'valores', soloTexto: true }) });
    reg({ id: 'pegarMenu', texto: 'cmd-pegar-menu', ejecutar: () => {} });
    const OPCIONES = [{ comando: 'pegarTodo' }, { comando: 'pegarFormulas' }, { comando: 'pegarFormulasNum' }, { comando: 'pegarSinBordes' }, { comando: 'pegarTranspuesto' }, { comando: 'pegarAnchos' }, '-', { comando: 'pegarSoloValores' }, { comando: 'pegarValoresNum' }, { comando: 'pegarSoloTexto' }, '-', { comando: 'pegarFormatos' }, { comando: 'pegarVinculo' }, { comando: 'pegarNotas' }, { comando: 'pegarValidacion' }, { comando: 'pegarFusion' }, '-', { comando: 'pegarEspecial' }];

    // ---- la ventana ------------------------------------------------------------------------------------------------------
    const CON_OPERACION = new Set(['todo', 'formulas', 'valores', 'formulasNum', 'valoresNum']);
    reg({
      id: 'pegarEspecial', texto: 'cmd-pegar-especial', atajo: 'mod+alt+v',
      ejecutar: async () => {
        const repetir = cambiar;
        cambiar = false;
        const interno = A.hayCopiaInterna();
        const comun = { aceptar: tx('pe-aceptar'), cancelar: tx('pe-cancelar') };
        const v = await Dialogo.formulario(Object.assign({}, comun, interno ? {
          titulo: tx('pe-titulo'),
          campos: [
            { nombre: 'tipo', etiqueta: tx('pe-pegar'), tipo: 'select', opciones: [['todo', tx('pe-todo')], ['formulas', tx('cmd-pegar-formulas')], ['valores', tx('pe-valores')], ['formatos', tx('cmd-pegar-formatos')], ['notas', tx('cmd-pegar-notas')], ['validacion', tx('cmd-pegar-validacion')], ['sinBordes', tx('cmd-pegar-sin-bordes')], ['anchos', tx('cmd-pegar-anchos')], ['formulasNum', tx('cmd-pegar-formulas-num')], ['valoresNum', tx('cmd-pegar-valores-num')], ['fusionCF', tx('cmd-pegar-fusion')]] },
            { nombre: 'op', etiqueta: tx('pe-operacion'), tipo: 'select', opciones: [['', tx('pe-ninguna')], ['sumar', tx('pe-sumar')], ['restar', tx('pe-restar')], ['multiplicar', tx('pe-multiplicar')], ['dividir', tx('pe-dividir')]], visible: (x) => CON_OPERACION.has(x.tipo) },
            { nombre: 'blancos', etiqueta: tx('pe-blancos'), tipo: 'check', visible: (x) => x.tipo !== 'anchos' },
            { nombre: 'trans', etiqueta: tx('pe-transponer'), tipo: 'check' },
          ],
          extras: [{ valor: 'vinculo', texto: tx('cmd-pegar-vinculo') }],
        } : {
          titulo: tx('pe-titulo'),
          campos: [{ nombre: 'como', etiqueta: tx('pe-como'), tipo: 'select', opciones: [['formato', tx('pe-con-formato')], ['texto', tx('pe-solo-texto')]] }],
        }));
        A.rejilla && A.rejilla.enfocar();
        if (!v) return;
        const o = !interno ? (v.como === 'texto' ? { tipo: 'valores', soloTexto: true } : { tipo: 'todo' })
          : v._boton === 'vinculo' ? { tipo: 'todo', vinculo: true }
            : { tipo: v.tipo, op: CON_OPERACION.has(v.tipo) && v.op ? v.op : null, blancos: v.tipo !== 'anchos' && v.blancos, trans: v.trans };
        return repetir ? A.repegar(o) : A.pegarDesdeBoton(o);
      },
    });

    // ---- el botón junto a lo pegado ---------------------------------------------------------------------------------------
    const flotante = Oficina.crearEl('button', { type: 'button', class: 'of-boton hc-opciones-pegado', hidden: true, 'aria-label': tx('cmd-pegar-menu'), title: tx('cmd-pegar-menu'), 'data-i18n-attr': 'aria-label:cmd-pegar-menu;title:cmd-pegar-menu' }, ['📋▾']);
    document.body.append(flotante);
    let lugar = null;
    const menuFlotante = Oficina.menu(flotante, OPCIONES, comandos);
    flotante.addEventListener('pointerdown', () => { cambiar = true; });
    flotante.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') cambiar = true; });
    function colocar() {
      if (!lugar) return;
      const r = A.rectEnPantalla(lugar.f1, lugar.c1);
      const dentro = r && r.x + r.w > r.vista.left && r.y + r.h > r.vista.top && r.x + r.w < r.vista.right && r.y + r.h < r.vista.bottom;
      flotante.hidden = !dentro;
      if (dentro) { flotante.style.left = (r.x + r.w + 2) + 'px'; flotante.style.top = (r.y + r.h + 2) + 'px'; }
    }
    function ocultar() { lugar = null; flotante.hidden = true; cambiar = false; menuFlotante.cerrar(); }
    A.ganchos.alPegar.push((destino) => { lugar = destino; requestAnimationFrame(colocar); });
    A.ganchos.alSeleccionar.push(() => { if (!lugar) return; const s = A.seleccion(); if (s.f0 !== lugar.f0 || s.c0 !== lugar.c0 || s.f1 !== lugar.f1 || s.c1 !== lugar.c1) ocultar(); });
    A.ganchos.alCambiar.push(() => { if (lugar && !A.pegando()) ocultar(); });
    A.ganchos.alEscape.push(() => { ocultar(); soltarBrocha(); });
    document.addEventListener('scroll', () => colocar(), true);
    window.addEventListener('resize', () => colocar());

    // ---- copiar formato ----------------------------------------------------------------------------------------------------
    let brocha = null, ultimoClic = 0;
    function soltarBrocha() {
      if (!brocha) return;
      brocha = null;
      A.rejilla && A.rejilla.marcarCorte(null);
      document.documentElement.classList.remove('hc-brocha');
      comandos.refrescar();
    }
    function pintar() {
      if (!brocha) return;
      const P = brocha.P;
      const seguir = brocha.fijo;
      if (!seguir) soltarBrocha();
      A.pegarInterno(P, { tipo: 'formatos', brocha: true });
    }
    reg({
      id: 'copiarFormato', texto: 'cmd-copiar-formato',
      activo: () => !!brocha,
      ejecutar: () => {
        const ahora = Date.now();
        // un clic: una vez; dos seguidos: hasta que se cancele (también con el dedo, donde no hay doble clic fiable)
        if (brocha) {
          if (!brocha.fijo && ahora - ultimoClic < 450) { brocha.fijo = true; avisar(tx('aviso-brocha-fija')); ultimoClic = 0; comandos.refrescar(); return; }
          soltarBrocha();
          return;
        }
        ultimoClic = ahora;
        const P = A.capturarSeleccion();
        if (!P) return;
        brocha = { P, fijo: false };
        A.rejilla.marcarCorte(Object.assign({ h: P.h }, P.r));
        document.documentElement.classList.add('hc-brocha');
        avisar(tx('aviso-brocha'));
        comandos.refrescar();
      },
    });
    A.ganchos.alSoltar.push(pintar);
    document.addEventListener('keydown', (e) => {
      if (!brocha || e.key !== 'Enter' || !e.target.closest || !e.target.closest('.hc-rejilla')) return;
      e.preventDefault(); e.stopPropagation();
      pintar();
    }, true);

    // Oficina.menu escucha el clic en el botón: el de la cinta abre el mismo menú que el que aparece junto a lo pegado
    A.despuesDeCinta.push(() => { const b = document.querySelector('[data-comando="pegarMenu"]'); if (b) Oficina.menu(b, OPCIONES, comandos); });
  });
})();
