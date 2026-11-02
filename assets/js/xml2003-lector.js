/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Lector de "Hoja de cálculo XML 2003" (SpreadsheetML, el .xml o .xls en XML que guardan Excel 2002/2003 y muchos sistemas) del motor propio de Excel (etapa 9).
// Requiere xml-ligero.js, xlsx-modelo.js (para los errores) y xlsx-formato.js (fechas). No depende del navegador.
//
//   Xml2003Lector.parece(bytes)   → true si el archivo es SpreadsheetML (declara urn:schemas-microsoft-com:office:spreadsheet al principio)
//   Xml2003Lector.leer(bytes)     → { hojas: [{ nombre, filas, anchos, combinadas, estado, columnasOcultas, filasOcultas }] }   (mismo formato que OdsLector; filas puede tener huecos)
//
// Qué lee: valores (Number, String, DateTime, Boolean, Error), el formato de número de cada celda (los nombres de Excel —"Fixed", "Short Date"…— y los
// códigos propios), combinadas (MergeAcross/MergeDown), anchos, columnas y filas ocultas, hojas ocultas, y los saltos de Index/Span. Las fórmulas no se
// leen (ss:Formula está en R1C1): queda el último valor. Estilos de celda (color, borde) no se llevan.
(function () {
  const MAX_FILAS = 1048576, MAX_COLS = 16384, MAX_TEXTO = 32767, MAX_HOJAS = 1000, MAX_CELDAS = 10000000, MAX_COMBINADAS = 500000;
  const ERRORES = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A']);

  // Los nombres de formato que escribe Excel 2003 (son los de VBA) y el código de Excel que les corresponde.
  const FORMATOS = {
    'general': undefined, 'general number': undefined, 'general date': 'm/d/yyyy h:mm', 'long date': 'dddd, mmmm dd, yyyy', 'medium date': 'dd-mmm-yy', 'short date': 'm/d/yyyy',
    'long time': 'h:mm:ss AM/PM', 'medium time': 'h:mm AM/PM', 'short time': 'hh:mm', 'currency': '"$"#,##0.00', 'euro currency': '[$€-2] #,##0.00',
    'fixed': '0.00', 'standard': '#,##0.00', 'percent': '0.00%', 'scientific': '0.00E+00', 'yes/no': '"Yes";"Yes";"No"', 'true/false': '"True";"True";"False"', 'on/off': '"On";"On";"Off"',
  };

  const decodificar = (bytes) => {
    let inicio = 0, etiqueta = 'utf-8';
    if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) inicio = 3;
    else if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) etiqueta = 'utf-16le';
    else if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) etiqueta = 'utf-16be';
    else {
      let cabecera = '';
      for (let k = 0, n = Math.min(bytes.length, 200); k < n; k++) cabecera += String.fromCharCode(bytes[k]);
      const m = /<\?xml[^>]*encoding\s*=\s*["']([\w.-]+)["']/i.exec(cabecera);
      if (m) etiqueta = m[1];
    }
    try { return new TextDecoder(etiqueta).decode(bytes.subarray(inicio)); } catch (e) { return new TextDecoder('utf-8').decode(bytes.subarray(inicio)); }
  };

  function parece(bytes) {
    let s = '';
    const n = Math.min(bytes.length, 4096);
    if (n >= 2 && ((bytes[0] === 0xFF && bytes[1] === 0xFE) || (bytes[0] === 0xFE && bytes[1] === 0xFF))) s = decodificar(bytes.subarray(0, n)).slice(0, 4096);
    else for (let k = 0; k < n; k++) { if (bytes[k] === 0) return false; s += String.fromCharCode(bytes[k]); }
    return /<\?xml|<Workbook\b/i.test(s) && /urn:schemas-microsoft-com:office:spreadsheet/i.test(s) && /<(\w+:)?Workbook\b/i.test(s);
  }

  // "2024-01-15T13:30:00.000" → número de serie de Excel (base 1900). Una hora sola ("1899-12-31T13:30:00.000") queda como fracción de día.
  function serieDeFecha(texto) {
    const m = /^(\d{1,4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/.exec(texto.trim());
    if (!m) return null;
    const a = +m[1], mes = +m[2], d = +m[3];
    const hora = m[4] === undefined ? 0 : (+m[4] * 3600 + +m[5] * 60 + parseFloat(m[6] || '0')) / 86400;
    if (a === 1899 && mes === 12 && d >= 30) return { serie: (d - 31) + hora, soloHora: m[4] !== undefined };
    if (a === 1900 && mes === 2 && d === 29) return { serie: 60 + hora, soloHora: false };
    const t = new Date(0); t.setUTCFullYear(a, mes - 1, d); t.setUTCHours(0, 0, 0, 0);
    let dias = Math.round(t.getTime() / 86400000) + 25569;
    if (dias < 61) dias -= 1;   // antes del 1900-03-01 Excel cuenta un día menos
    return { serie: dias + hora, soloHora: false, conHora: m[4] !== undefined && hora > 0 };
  }

  const entero = (s, defecto) => { const n = parseInt(s, 10); return n > 0 ? n : defecto; };
  const ancho = (puntos) => { const px = puntos / 0.75; return Math.max(0, Math.round(((px - 5) / 7) * 100) / 100); };   // puntos → caracteres de la fuente Normal (Calibri 11)

  function leer(bytes) {
    const texto = decodificar(bytes);
    const estilos = Object.create(null);   // ID → { formato }
    const hojas = [];
    let hoja = null, fila = null, celda = null, dato = null, estiloActual = null, enOpciones = false, enComentario = 0, enNombres = 0;
    let colActual = 0, filaActual = -1, celdas = 0, textoVisible = null, opcionActual = '';
    const pila = [];

    const error = (c, m) => new LibroXlsx.ErrorXlsx(c, m);

    function valorDeDato(tipo, t, formato) {
      if (tipo === 'Number') { const n = Number(t); return Number.isFinite(n) ? n : (t === '' ? undefined : t); }
      if (tipo === 'Boolean') return t === '1' || /^true$/i.test(t);
      if (tipo === 'Error') return { error: ERRORES.has(t) ? t : '#VALUE!' };
      if (tipo === 'DateTime') {
        const f = serieDeFecha(t);
        if (!f) return t === '' ? undefined : t;
        return { valor: f.serie, formato: formato === undefined ? (f.soloHora ? 'h:mm:ss' : (f.serie % 1 ? 'm/d/yyyy h:mm' : 'm/d/yyyy')) : formato };
      }
      return t === '' ? undefined : (t.length > MAX_TEXTO ? t.slice(0, MAX_TEXTO) : t);
    }

    XmlLigero.recorrer(texto, {
      abrir(nombre, a, vacio) {
        pila.push(nombre);
        if (enComentario) { if (nombre === 'Comment' || nombre === 'ss:Comment') enComentario++; return; }
        switch (nombre) {
          case 'Comment': enComentario = 1; return;
          case 'Names': enNombres++; return;
          case 'Style': if (!hoja) { estiloActual = { id: a.ID, formato: undefined }; if (a.ID !== undefined) estilos[a.ID] = estiloActual; } return;
          case 'NumberFormat':
            if (estiloActual && a.Format !== undefined) { const f = a.Format, k = f.toLowerCase(); estiloActual.formato = k in FORMATOS ? FORMATOS[k] : f; }
            return;
          case 'Worksheet':
            if (hojas.length >= MAX_HOJAS) throw error('demasiadasHojas', 'demasiadas hojas');
            hoja = { nombre: a.Name || ('Hoja' + (hojas.length + 1)), filas: [], anchos: [], combinadas: [], estado: 'visible', columnasOcultas: [], filasOcultas: [], _col: 0, _ancho: 0 };
            hojas.push(hoja); filaActual = -1; return;
          case 'Column':
            if (hoja && !enNombres) {
              const desde = a.Index !== undefined ? entero(a.Index, 1) - 1 : hoja._col;
              const hasta = Math.min(MAX_COLS - 1, desde + (a.Span !== undefined ? Math.max(0, parseInt(a.Span, 10) || 0) : 0));
              const w = a.Width !== undefined ? ancho(parseFloat(a.Width)) : 0;
              const oculta = a.Hidden === '1';
              // _ancho cuenta las columnas recorridas: una hoja bien formada no pasa de 16.384; miles de <Column Span="16383"> repetidas no se recorren
              if (hoja._ancho <= 4 * MAX_COLS) for (let c = desde; c <= hasta && c < MAX_COLS; c++) { hoja._ancho++; if (w > 0) hoja.anchos[c] = w; if (oculta) hoja.columnasOcultas.push(c); }
              hoja._col = hasta + 1;
            }
            return;
          case 'Row':
            if (hoja) {
              filaActual = a.Index !== undefined ? entero(a.Index, 1) - 1 : filaActual + 1;
              if (filaActual >= MAX_FILAS) throw error('limite', 'demasiadas filas');
              fila = hoja.filas[filaActual] || (hoja.filas[filaActual] = []);
              if (a.Hidden === '1') hoja.filasOcultas.push(filaActual);
              colActual = 0;
            }
            return;
          case 'Cell':
            if (hoja && fila) {
              colActual = a.Index !== undefined ? entero(a.Index, 1) - 1 : colActual;
              if (colActual >= MAX_COLS) { celda = null; return; }
              celda = { col: colActual, formato: a.StyleID !== undefined && estilos[a.StyleID] ? estilos[a.StyleID].formato : undefined, tipo: null, texto: '', ma: Math.max(0, parseInt(a.MergeAcross, 10) || 0), md: Math.max(0, parseInt(a.MergeDown, 10) || 0) };
              if (++celdas > MAX_CELDAS) throw error('limite', 'demasiadas celdas');
            }
            return;
          case 'Data':
            if (celda) { dato = true; celda.tipo = a.Type || 'String'; }
            return;
          case 'WorksheetOptions': enOpciones = true; return;
          case 'Visible': opcionActual = 'Visible'; return;
          default:
        }
      },
      cerrar(nombre) {
        pila.pop();
        if (enComentario) { if (nombre === 'Comment' || nombre === 'ss:Comment') enComentario--; return; }
        switch (nombre) {
          case 'Names': if (enNombres) enNombres--; return;
          case 'Style': estiloActual = null; return;
          case 'Data': dato = null; return;
          case 'Cell':
            if (celda) {
              const v = celda.tipo ? valorDeDato(celda.tipo, celda.texto, celda.formato) : undefined;
              let valor = v;
              if (v !== undefined && celda.formato !== undefined && celda.tipo !== 'DateTime' && !(v && typeof v === 'object')) valor = (typeof v === 'number') ? { valor: v, formato: celda.formato } : v;
              if (valor !== undefined) fila[celda.col] = valor;
              if (celda.ma || celda.md) {
                const c2 = Math.min(MAX_COLS - 1, celda.col + celda.ma), f2 = Math.min(MAX_FILAS - 1, filaActual + celda.md);
                if (hoja.combinadas.length < MAX_COMBINADAS) hoja.combinadas.push(LibroXlsx.numeroACol(celda.col) + (filaActual + 1) + ':' + LibroXlsx.numeroACol(c2) + (f2 + 1));
                colActual = c2 + 1;
              } else colActual = celda.col + 1;
              celda = null;
            }
            return;
          case 'Row': fila = null; return;
          case 'Visible': opcionActual = ''; return;
          case 'WorksheetOptions': enOpciones = false; return;
          case 'Worksheet': hoja = null; fila = null; celda = null; return;
          default:
        }
      },
      texto(t) {
        if (enComentario) return;
        if (dato && celda) celda.texto += t;
        else if (enOpciones && opcionActual === 'Visible' && hoja) {
          const s = t.trim();
          if (s === 'SheetHidden') hoja.estado = 'oculta'; else if (s === 'SheetVeryHidden') hoja.estado = 'muyOculta';
        }
      },
    });

    if (!hojas.length) throw error('sinLibro', 'sin hojas');
    for (const h of hojas) {
      delete h._col; delete h._ancho;
      // las filas que no vienen quedan como huecos (XlsxEscritor los salta): llenarlas con [] gastaba ~45 MB por hoja con un Row Index="1048576"
      h.columnasOcultas = [...new Set(h.columnasOcultas)];
      h.filasOcultas = [...new Set(h.filasOcultas)];
    }
    return { hojas };
  }

  globalThis.Xml2003Lector = { parece, leer };
})();
