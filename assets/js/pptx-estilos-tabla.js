/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Los 74 estilos de tabla integrados de PowerPoint, escritos como reglas
// propias: 11 familias, cada una una receta con el color de énfasis como
// parámetro. Hacen falta porque PowerPoint muchas veces guarda en el .pptx
// solo el GUID del estilo, sin su definición (así venía la charla del
// usuario, guardada por PowerPoint 16).
//
// Decisión del usuario (2026-10-10): reglas propias, no copiar el XML de
// Microsoft. Las reglas se comprobaron una por una contra las 74
// definiciones que PowerPoint escribe al aplicar cada estilo
// (comparar-pptx/corpus/08-tablas.pptx, ppt/tableStyles.xml; prueba en
// npm run probar). Los GUID son los identificadores que usa el formato.
//
// Uso: PptxEstilosTabla.definicion(guid) → elemento <a:tblStyle> con el
// mismo vocabulario de la norma (lo lee pptx-modelo.js igual que uno que
// viene en el archivo), o null si el GUID no es de un estilo integrado.
// PptxEstilosTabla.xml(guid) da el texto (para la prueba).

const PptxEstilosTabla = (() => {
  const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  // Colores: 'accent1', 'dk1', 'A' (el énfasis), con un modificador opcional:
  // /t20 = tint 20 %, /s60 = shade 60 %, /a40 = alpha 40 %.
  const MODS = { t: 'tint', s: 'shade', a: 'alpha' };
  function color(spec) {
    const [nombre, mod] = spec.split('/');
    if (!mod) return `<a:schemeClr val="${nombre}"/>`;
    return `<a:schemeClr val="${nombre}"><a:${MODS[mod[0]]} val="${parseInt(mod.slice(1), 10) * 1000}"/></a:schemeClr>`;
  }
  // Líneas: '-' ninguna; 'r1:color' la línea 1 del tema; '12700:color' de ese
  // ancho en EMU; con 'd' al final del ancho, doble.
  function linea(spec) {
    if (spec === '-') return '<a:ln><a:noFill/></a:ln>';
    const [a, c] = spec.split(':');
    if (a[0] === 'r') return `<a:lnRef idx="${a.slice(1)}">${color(c)}</a:lnRef>`;
    const doble = a.endsWith('d');
    return `<a:ln w="${parseInt(a, 10)}" cmpd="${doble ? 'dbl' : 'sng'}"><a:solidFill>${color(c)}</a:solidFill></a:ln>`;
  }
  // Rellenos: 'none', 'r2:color' (relleno 2 del tema) o un color.
  const relleno = (spec) => spec === 'none' ? '<a:fill><a:noFill/></a:fill>'
    : spec[0] === 'r' && spec[2] === ':' ? `<a:fillRef idx="${spec[1]}">${color(spec.slice(3))}</a:fillRef>`
      : `<a:fill><a:solidFill>${color(spec)}</a:solidFill></a:fill>`;
  const LADOS = [['l', 'left'], ['r', 'right'], ['t', 'top'], ['b', 'bottom'], ['h', 'insideH'], ['v', 'insideV']];

  // Una parte: { t: { b, c } texto (negrita, color), bd: { l r t b h v } bordes, f: relleno }.
  function parte(nombre, p) {
    let x = '';
    if (p.t) x += `<a:tcTxStyle${p.t.b ? ' b="on"' : ''}>${p.t.c ? `<a:fontRef idx="minor"/>${color(p.t.c)}` : ''}</a:tcTxStyle>`;
    let estilo = '';
    if (p.bd) estilo += `<a:tcBdr>${LADOS.filter(([k]) => p.bd[k]).map(([k, n]) => `<a:${n}>${linea(p.bd[k])}</a:${n}>`).join('')}</a:tcBdr>`;
    if (p.f) estilo += relleno(p.f);
    if (estilo) x += `<a:tcStyle>${estilo}</a:tcStyle>`;
    return `<a:${nombre}>${x}</a:${nombre}>`;
  }
  const todos = (v) => ({ l: v, r: v, t: v, b: v, h: v, v });
  const N = { b: true };

  // Las recetas. A y B: colores de énfasis (B solo en Oscuro 2).
  const FAMILIAS = {
    claro1: (A) => ({
      wholeTbl: { t: { c: 'tx1' }, bd: { l: '-', r: '-', t: `12700:${A}`, b: `12700:${A}`, h: '-', v: '-' }, f: 'none' },
      band1H: { f: `${A}/a20` }, band1V: { f: `${A}/a20` }, lastCol: { t: N }, firstCol: { t: N },
      lastRow: { t: N, bd: { t: `12700:${A}` }, f: 'none' }, firstRow: { t: N, bd: { b: `12700:${A}` }, f: 'none' },
    }),
    claro2: (A) => ({
      wholeTbl: { t: { c: 'tx1' }, bd: { l: `r1:${A}`, r: `r1:${A}`, t: `r1:${A}`, b: `r1:${A}`, h: '-', v: '-' }, f: 'none' },
      band1H: { bd: { t: `r1:${A}`, b: `r1:${A}` } }, band1V: { bd: { l: `r1:${A}`, r: `r1:${A}` } }, band2V: { bd: { l: `r1:${A}`, r: `r1:${A}` } },
      lastCol: { t: N }, firstCol: { t: N }, lastRow: { t: N, bd: { t: `50800d:${A}` } },
      firstRow: { t: { b: true, c: 'bg1' }, bd: {}, f: `r1:${A}` },
    }),
    claro3: (A) => ({
      wholeTbl: { t: { c: 'tx1' }, bd: todos(`12700:${A}`), f: 'none' },
      band1H: { f: `${A}/a20` }, band1V: { f: `${A}/a20` }, lastCol: { t: N }, firstCol: { t: N },
      lastRow: { t: N, bd: { t: `50800d:${A}` }, f: 'none' }, firstRow: { t: N, bd: { b: `25400:${A}` }, f: 'none' },
    }),
    medio1: (A) => ({
      wholeTbl: { t: { c: 'dk1' }, bd: { ...todos(`12700:${A}`), v: '-' }, f: 'lt1' },
      band1H: { f: `${A}/t20` }, band1V: { f: `${A}/t20` }, lastCol: { t: N }, firstCol: { t: N },
      lastRow: { t: N, bd: { t: `50800d:${A}` }, f: 'lt1' }, firstRow: { t: { b: true, c: 'lt1' }, f: A },
    }),
    medio2: (A) => ({
      wholeTbl: { t: { c: 'dk1' }, bd: todos('12700:lt1'), f: `${A}/t20` },
      band1H: { f: `${A}/t40` }, band1V: { f: `${A}/t40` },
      lastCol: { t: { b: true, c: 'lt1' }, f: A }, firstCol: { t: { b: true, c: 'lt1' }, f: A },
      lastRow: { t: { b: true, c: 'lt1' }, bd: { t: '38100:lt1' }, f: A }, firstRow: { t: { b: true, c: 'lt1' }, bd: { b: '38100:lt1' }, f: A },
    }),
    medio3: (A) => ({
      wholeTbl: { t: { c: 'dk1' }, bd: { l: '-', r: '-', t: '25400:dk1', b: '25400:dk1', h: '-', v: '-' }, f: 'lt1' },
      band1H: { f: 'dk1/t20' }, band1V: { f: 'dk1/t20' },
      lastCol: { t: { b: true, c: 'lt1' }, f: A }, firstCol: { t: { b: true, c: 'lt1' }, f: A },
      lastRow: { t: N, bd: { t: '50800d:dk1' }, f: 'lt1' },
      seCell: { t: { b: true, c: 'dk1' } }, swCell: { t: { b: true, c: 'dk1' } },
      firstRow: { t: { b: true, c: 'lt1' }, bd: { b: '25400:dk1' }, f: A },
    }),
    medio4: (A) => ({
      wholeTbl: { t: { c: 'dk1' }, bd: todos(`12700:${A}`), f: `${A}/t20` },
      band1H: { f: `${A}/t40` }, band1V: { f: `${A}/t40` }, lastCol: { t: N }, firstCol: { t: N },
      lastRow: { t: N, bd: { t: `25400:${A}` }, f: `${A}/t20` }, firstRow: { t: N, f: `${A}/t20` },
    }),
    // Oscuro 1: con énfasis, el color y sus sombras; sin énfasis, dk1 aclarado.
    oscuro1: (A) => {
      const [base, banda, columna, ultima] = A === 'dk1' ? ['dk1/t20', 'dk1/t40', 'dk1/t60', 'dk1/t60'] : [A, `${A}/s60`, `${A}/s60`, `${A}/s40`];
      return {
        wholeTbl: { t: { c: 'lt1' }, bd: todos('-'), f: base },
        band1H: { f: banda }, band1V: { f: banda },
        lastCol: { t: N, bd: { l: '25400:lt1' }, f: columna }, firstCol: { t: N, bd: { r: '25400:lt1' }, f: columna },
        lastRow: { t: N, bd: { t: '25400:lt1' }, f: ultima },
        seCell: { bd: { l: '-' } }, swCell: { bd: { r: '-' } },
        firstRow: { t: N, bd: { b: '25400:lt1' }, f: 'dk1' },
        neCell: { bd: { l: '-' } }, nwCell: { bd: { r: '-' } },
      };
    },
    oscuro2: (A, B) => ({
      wholeTbl: { t: { c: 'dk1' }, bd: todos('-'), f: `${A}/t20` },
      band1H: { f: `${A}/t40` }, band1V: { f: `${A}/t40` }, lastCol: { t: N }, firstCol: { t: N },
      lastRow: { t: N, bd: { t: '50800d:dk1' }, f: `${A}/t20` }, firstRow: { t: { b: true, c: 'lt1' }, f: B },
    }),
    tematico1: (A) => ({
      tblBg: `r2:${A}`,
      wholeTbl: { t: { c: 'dk1' }, bd: todos(`r1:${A}`), f: 'none' },
      band1H: { f: `${A}/a40` }, band1V: { bd: { t: `r1:${A}`, b: `r1:${A}` }, f: `${A}/a40` },
      lastCol: { t: N, bd: { l: `r2:${A}`, r: `r1:${A}`, t: `r1:${A}`, b: `r1:${A}`, h: `r1:${A}`, v: '-' } },
      firstCol: { t: N, bd: { l: `r1:${A}`, r: `r2:${A}`, t: `r1:${A}`, b: `r1:${A}`, h: `r1:${A}`, v: '-' } },
      lastRow: { t: N, bd: { l: `r1:${A}`, r: `r1:${A}`, t: `r2:${A}`, b: `r2:${A}`, h: '-', v: '-' }, f: 'none' },
      firstRow: { t: { b: true, c: 'lt1' }, bd: { l: `r1:${A}`, r: `r1:${A}`, t: `r1:${A}`, b: 'r2:lt1', h: '-', v: '-' }, f: A },
    }),
    tematico2: (A) => ({
      tblBg: `r3:${A}`,
      wholeTbl: { t: { c: 'lt1' }, bd: { l: `r1:${A}/t50`, r: `r1:${A}/t50`, t: `r1:${A}/t50`, b: `r1:${A}/t50`, h: '-', v: '-' }, f: 'none' },
      band1H: { f: 'lt1/a20' }, band1V: { f: 'lt1/a20' },
      lastCol: { t: N, bd: { l: 'r2:lt1' } }, firstCol: { t: N, bd: { r: 'r2:lt1' } },
      lastRow: { t: N, bd: { t: 'r2:lt1' }, f: 'none' },
      seCell: { bd: { l: '-', t: '-' } }, swCell: { bd: { r: '-', t: '-' } },
      firstRow: { t: N, bd: { b: 'r3:lt1' }, f: 'none' }, neCell: { bd: { b: '-' } },
    }),
    // Las variantes sin énfasis de las familias "temáticas" son otra cosa:
    // sin estilo, con o sin cuadrícula.
    sinCuadricula: () => ({ wholeTbl: { t: { c: 'tx1' }, bd: todos('-'), f: 'none' } }),
    cuadricula: () => ({ wholeTbl: { t: { c: 'tx1' }, bd: todos('12700:tx1'), f: 'none' } }),
  };

  // GUID → [familia, A, B]. Por fila: sin énfasis y énfasis 1 a 6.
  const ACENTOS = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
  const GUID = {};
  const fila = (familia, sinEnfasis, ids) => ids.forEach((id, k) => { if (id) GUID[id] = k === 0 ? sinEnfasis : [familia, ACENTOS[k - 1]]; });
  fila('tematico1', ['sinCuadricula'], ['2D5ABB26-0587-4C30-8999-92F81FD0307C', '3C2FFA5D-87B4-456A-9821-1D502468CF0F', '284E427A-3D55-4303-BF80-6455036E1DE7', '69C7853C-536D-4A76-A0AE-DD22124D55A5', '775DCB02-9BB8-47FD-8907-85C794F793BA', '35758FB7-9AC5-4552-8A53-C91805E547FA', '08FB837D-C827-4EFA-A057-4D05807E0F7C']);
  fila('tematico2', ['cuadricula'], ['5940675A-B579-460E-94D1-54222C63F5DA', 'D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7', '18603FDC-E32A-4AB5-989C-0864C3EAD2B8', '306799F8-075E-4A3A-A7F6-7FBC6576F1A4', 'E269D01E-BC32-4049-B463-5C60D7B0CCD2', '327F97BB-C833-4FB7-BDE5-3F7075034690', '638B1855-1B75-4FBE-930C-398BA8C253C6']);
  fila('claro1', ['claro1', 'tx1'], ['9D7B26C5-4107-4FEC-AEDC-1716B250A1EF', '3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5', '0E3FDE45-AF77-4B5C-9715-49D594BDF05E', 'C083E6E3-FA7D-4D7B-A595-EF9225AFEA82', 'D27102A9-8310-4765-A935-A1911B00CA55', '5FD0F851-EC5A-4D38-B0AD-8093EC10F338', '68D230F3-CF80-4859-8CE7-A43EE81993B5']);
  fila('claro2', ['claro2', 'tx1'], ['7E9639D4-E3E2-4D34-9284-5A2195B3D0D7', '69012ECD-51FC-41F1-AA8D-1B2483CD663E', '72833802-FEF1-4C79-8D5D-14CF1EAF98D9', 'F2DE63D5-997A-4646-A377-4702673A728D', '17292A2E-F333-43FB-9621-5CBBE7FDCDCB', '5A111915-BE36-4E01-A7E5-04B1672EAD32', '912C8C85-51F0-491E-9774-3900AFEF0FD7']);
  fila('claro3', ['claro3', 'tx1'], ['616DA210-FB5B-4158-B5E0-FEB733F419BA', 'BC89EF96-8CEA-46FF-86C4-4CE0E7609802', '5DA37D80-6434-44D0-A028-1B22A696006F', '8799B23B-EC83-4686-B30A-512413B5E67A', 'ED083AE6-46FA-4A59-8FB0-9F97EB10719F', 'BDBED569-4797-4DF1-A0F4-6AAB3CD982D8', 'E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3']);
  fila('medio1', ['medio1', 'dk1'], ['793D81CF-94F2-401A-BA57-92F5A7B2D0C5', 'B301B821-A1FF-4177-AEE7-76D212191A09', '9DCAF9ED-07DC-4A11-8D7F-57B35C25682E', '1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90', '1E171933-4619-4E11-9A3F-F7608DF75F80', 'FABFCF23-3B69-468F-B69F-88F6DE6A72F2', '10A1B5D5-9B99-4C35-A422-299274C87663']);
  fila('medio2', ['medio2', 'dk1'], ['073A0DAA-6AF3-43AB-8588-CEC1D06C72B9', '5C22544A-7EE6-4342-B048-85BDC9FD1C3A', '21E4AEA4-8DFA-4A89-87EB-49C32662AFE0', 'F5AB1C69-6EDB-4FF4-983F-18BD219EF322', '00A15C55-8517-42AA-B614-E9B94910E393', '7DF18680-E054-41AD-8BC1-D1AEF772440D', '93296810-A885-4BE3-A3E7-6D5BEEA58F35']);
  fila('medio3', ['medio3', 'dk1'], ['8EC20E35-A176-4012-BC5E-935CFFF8708E', '6E25E649-3F16-4E02-A733-19D2CDBF48F0', '85BE263C-DBD7-4A20-BB59-AAB30ACAA65A', 'EB344D84-9AFB-497E-A393-DC336BA19D2E', 'EB9631B5-78F2-41C9-869B-9F39066F8104', '74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6', '2A488322-F2BA-4B5B-9748-0D474271808F']);
  fila('medio4', ['medio4', 'dk1'], ['D7AC3CCA-C797-4891-BE02-D94E43425B78', '69CF1AB2-1976-4502-BF36-3FF5EA218861', '8A107856-5554-42FB-B03E-39F5DBC370BA', '0505E3EF-67EA-436B-97B2-0124C06EBD24', 'C4B1156A-380E-4F78-BDF5-A606A8083BF9', '22838BEF-8BB2-4498-84A7-C5851F593DF1', '16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4']);
  fila('oscuro1', ['oscuro1', 'dk1'], ['E8034E78-7F5D-4C2E-B375-FC64B27BC917', '125E5076-3810-47DD-B79F-674D7AD40C01', '37CE84F3-28C3-443E-9E96-99CF82512B78', 'D03447BB-5D67-496B-8E87-E561075AD55C', 'E929F9F4-4A8F-4326-A1B4-22849713DDAB', '8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B', 'AF606853-7671-496A-8E4F-DF71F8EC918B']);
  // Oscuro 2 va de a pares de énfasis (el segundo, en la fila de encabezado).
  GUID['5202B0CA-FC54-4496-8BCA-5EF66A818D29'] = ['oscuro2', 'dk1', 'dk1'];
  GUID['0660B408-B3CF-4A94-85FC-2B1E0A45F4A2'] = ['oscuro2', 'accent1', 'accent2'];
  GUID['91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28'] = ['oscuro2', 'accent3', 'accent4'];
  GUID['46F890A9-2807-4EBB-B81D-B2AA78EC7F39'] = ['oscuro2', 'accent5', 'accent6'];

  const ORDEN = ['tblBg', 'wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'lastCol', 'firstCol', 'lastRow', 'seCell', 'swCell', 'firstRow', 'neCell', 'nwCell'];
  const limpiar = (guid) => String(guid || '').replace(/[{}\s]/g, '').toUpperCase();

  function xml(guid) {
    const id = limpiar(guid);
    const g = GUID[id];
    if (!g) return null;
    const partes = FAMILIAS[g[0]](g[1], g[2]);
    let x = '';
    for (const n of ORDEN) {
      if (!partes[n]) continue;
      x += n === 'tblBg' ? `<a:tblBg>${relleno(partes[n])}</a:tblBg>` : parte(n, partes[n]);
    }
    return `<a:tblStyle xmlns:a="${NS_A}" styleId="{${id}}">${x}</a:tblStyle>`;
  }

  const cache = new Map();
  function definicion(guid) {
    const id = limpiar(guid);
    if (!cache.has(id)) {
      const x = xml(id);
      cache.set(id, x ? new DOMParser().parseFromString(x, 'application/xml').documentElement : null);
    }
    return cache.get(id);
  }

  return { definicion, xml, cantidad: () => Object.keys(GUID).length };
})();
