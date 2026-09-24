/*! © 2026 Herramientas Ofimáticas (Mini Cookies). Código propio, todos los derechos reservados. Uso sujeto a los términos: https://www.herramientasofimaticas.minicookies.net/menus/terminos.html */
// Descargar como archivo el texto que genera una herramienta (además de poder
// copiarlo). Compartido por las herramientas que generan texto.
//
// descargarTexto(texto, nombreArchivo, tipo, conBom)
// - tipo: tipo MIME ('text/plain', 'application/json', 'text/csv'...).
// - conBom: true para CSV, así Excel lo abre con los acentos bien.
function descargarTexto(texto, nombreArchivo, tipo = 'text/plain', conBom = false) {
  const blob = new Blob([conBom ? '﻿' : '', texto], { type: `${tipo};charset=utf-8` });
  descargarBlob(blob, nombreArchivo);
}

function descargarBlob(blob, nombreArchivo) {
  const url = URL.createObjectURL(blob);
  const enlace = document.createElement('a');
  enlace.href = url;
  enlace.download = nombreArchivo;
  document.body.appendChild(enlace);
  enlace.click();
  enlace.remove();
  // Se libera después: si se revoca enseguida, algunos navegadores cancelan la descarga.
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

// Nombre de archivo con la fecha y hora, para que dos descargas seguidas no se pisen:
// nombreConFecha('json-formateado', 'json') → "json-formateado-2026-09-24-1530.json"
function nombreConFecha(base, extension) {
  const d = new Date();
  const dos = (n) => String(n).padStart(2, '0');
  return `${base}-${d.getFullYear()}-${dos(d.getMonth() + 1)}-${dos(d.getDate())}-${dos(d.getHours())}${dos(d.getMinutes())}.${extension}`;
}
