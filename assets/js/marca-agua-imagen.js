/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Marca de agua sobre una imagen (texto o logo), sin interfaz: separada de marca-agua-imagen.html para que la usen
// la página (vista previa y descarga) y Flujos con el mismo dibujo.
//
//   MarcaAguaImagen.dibujar(lienzo, imagen, opciones)
//     Deja en `lienzo` la imagen (a su tamaño natural) con la marca encima. opciones:
//       tipo: 'texto' | 'logo'      texto, tamano (20–150, pensado para ~1000 px de lado)
//       logo (imagen ya cargada), tamanoLogo (% del ancho)
//       opacidad (0–1), angulo (grados), posicion ('repartida' | 'arriba-izquierda' … 'abajo-derecha' | 'libre')
//       repeticiones (con 'repartida'), libre: { x, y } (centro en fracción del ancho y del alto, con 'libre')
//     Sin logo cargado con tipo 'logo', deja solo la imagen y devuelve false.
(function () {
  // Reparte "cantidad" copias en una cuadrícula lo más cuadrada posible
  // (ej. 4 -> 2x2, 6 -> 3x2, 1 -> 1x1) para distribuirlas parejo por la imagen.
  function calcularFilasColumnas(cantidad) {
    const columnas = Math.ceil(Math.sqrt(cantidad));
    const filas = Math.ceil(cantidad / columnas);
    return { filas, columnas };
  }

  // Lo que ocupa una marca de ancho × alto girada `radianes`.
  function cajaGirada(ancho, alto, radianes) {
    const seno = Math.abs(Math.sin(radianes)), coseno = Math.abs(Math.cos(radianes));
    return { ancho: ancho * coseno + alto * seno, alto: ancho * seno + alto * coseno };
  }

  // La marca (texto o logo) del tamaño pedido, achicada si girada no entra en
  // el espacio disponible. Devuelve su ancho y alto sin girar y cómo dibujarla.
  function medirMarca(contexto, o, espacioAncho, espacioAlto, radianes) {
    const lienzo = contexto.canvas;
    const margen = 0.9;
    const lado = Math.max(lienzo.width, lienzo.height);
    if (o.tipo === 'logo') {
      const logo = o.logo;
      if (!logo) return null;
      let ancho = lienzo.width * o.tamanoLogo / 100;
      // Un SVG sin ancho ni alto propios da 0: se lo toma cuadrado.
      const proporcion = logo.naturalWidth && logo.naturalHeight ? logo.naturalHeight / logo.naturalWidth : 1;
      let alto = ancho * proporcion;
      const caja = cajaGirada(ancho, alto, radianes);
      const escala = Math.min(1, espacioAncho * margen / caja.ancho, espacioAlto * margen / caja.alto);
      ancho *= escala; alto *= escala;
      return { ancho, alto, dibujar: () => contexto.drawImage(logo, -ancho / 2, -alto / 2, ancho, alto) };
    }
    const texto = o.texto;
    // El tamaño está pensado para una imagen de ~1000px de lado: se escala
    // al tamaño real, si no en una foto de celular (4000px) hasta el máximo
    // quedaba diminuto.
    let tamano = Math.round(o.tamano * lado / 1000);
    for (; tamano > 6; tamano -= 2) {
      contexto.font = `bold ${tamano}px Helvetica, Arial, sans-serif`;
      const caja = cajaGirada(contexto.measureText(texto).width, tamano, radianes);
      if (caja.ancho <= espacioAncho * margen && caja.alto <= espacioAlto * margen) break;
    }
    contexto.font = `bold ${tamano}px Helvetica, Arial, sans-serif`;
    const ancho = contexto.measureText(texto).width;
    return {
      ancho, alto: tamano,
      dibujar: () => {
        contexto.textAlign = 'center';
        contexto.textBaseline = 'middle';
        contexto.fillStyle = '#ffffff';
        contexto.strokeStyle = 'rgba(0,0,0,0.6)';
        contexto.lineWidth = Math.max(1, tamano / 25);
        contexto.strokeText(texto, 0, 0);
        contexto.fillText(texto, 0, 0);
      },
    };
  }

  // Centros donde va la marca, según la posición elegida.
  function centrosDeLaMarca(lienzo, o, marca, radianes) {
    const W = lienzo.width, H = lienzo.height;
    if (o.posicion === 'libre') return [[o.libre.x * W, o.libre.y * H]];
    const caja = cajaGirada(marca.ancho, marca.alto, radianes);
    const margen = 0.03 * Math.min(W, H);
    const [fila, columna] = o.posicion.split('-');
    const x = { izquierda: margen + caja.ancho / 2, centro: W / 2, derecha: W - margen - caja.ancho / 2 }[columna];
    const y = { arriba: margen + caja.alto / 2, centro: H / 2, abajo: H - margen - caja.alto / 2 }[fila];
    return [[x, y]];
  }

  function dibujar(lienzo, imagen, o) {
    lienzo.width = imagen.naturalWidth || imagen.width;
    lienzo.height = imagen.naturalHeight || imagen.height;
    const contexto = lienzo.getContext('2d');
    contexto.drawImage(imagen, 0, 0);

    const radianes = o.angulo * Math.PI / 180;
    let marca, centros;
    if (o.posicion === 'repartida') {
      const cantidad = o.repeticiones;
      const { filas, columnas } = calcularFilasColumnas(cantidad);
      const anchoCelda = lienzo.width / columnas, altoCelda = lienzo.height / filas;
      marca = medirMarca(contexto, o, anchoCelda, altoCelda, radianes);
      centros = [];
      for (let i = 0; i < cantidad; i++) {
        centros.push([anchoCelda * (i % columnas) + anchoCelda / 2, altoCelda * Math.floor(i / columnas) + altoCelda / 2]);
      }
    } else {
      marca = medirMarca(contexto, o, lienzo.width, lienzo.height, radianes);
      if (marca) centros = centrosDeLaMarca(lienzo, o, marca, radianes);
    }
    if (!marca) return false;

    contexto.save();
    contexto.globalAlpha = o.opacidad;
    for (const [x, y] of centros) {
      // Se gira alrededor del centro de la marca.
      contexto.save();
      contexto.translate(x, y);
      contexto.rotate(radianes);
      marca.dibujar();
      contexto.restore();
    }
    contexto.restore();
    return true;
  }

  window.MarcaAguaImagen = { dibujar };
})();
