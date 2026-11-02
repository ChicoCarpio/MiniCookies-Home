/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// El velo animado "la IA está trabajando" (.trabajando-ia, en
// assets/css/editor-imagen.css) sobre la imagen, para que no parezca que la
// página se trabó mientras un modelo se descarga o analiza. Lo usan el Editor
// avanzado y el Generador de stickers (estaba escrito dentro del editor; los
// stickers lo pidieron el 2026-10-08).
//
// Uso:
//   const velo = TrabajandoIA.crear(document.getElementById('trabajandoIA'));
//   const trabajo = velo.empezar('Analizando la foto…');
//   trabajo.texto('Buscando el objeto…');   // cambia el texto
//   trabajo.terminar();
// Puede haber varios trabajos a la vez: el velo muestra el texto del más
// reciente y se va cuando no queda ninguno. El elemento lleva un
// <span class="trabajando-ia-texto"> adentro (markup en SISTEMA-DE-INTERFAZ.md).

const TrabajandoIA = (() => {
  function crear(velo) {
    const trabajos = [];
    const texto = velo.querySelector('.trabajando-ia-texto');
    function pintar() {
      const ultimo = trabajos[trabajos.length - 1];
      velo.hidden = !ultimo;
      if (ultimo) texto.textContent = ultimo.texto;
    }
    return {
      empezar(textoInicial) {
        const trabajo = { texto: textoInicial };
        trabajos.push(trabajo);
        pintar();
        return {
          texto(nuevo) { trabajo.texto = nuevo; pintar(); },
          terminar() {
            const i = trabajos.indexOf(trabajo);
            if (i >= 0) trabajos.splice(i, 1);
            pintar();
          },
        };
      },
    };
  }

  // El texto del avance de una descarga: al 100 % ya bajó y lo que sigue es
  // analizar la foto. "t" es la función de traducción de la página.
  function textoDeDescarga(t, claveDescarga, claveAnalisis, porcentaje) {
    return porcentaje >= 100 ? t(claveAnalisis) : t(claveDescarga).replace('{p}', porcentaje);
  }

  return { crear, textoDeDescarga };
})();
