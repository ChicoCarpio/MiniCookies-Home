/*! © 2026 Mini Cookies. Código propio, todos los derechos reservados. Uso sujeto a los términos: https://minicookies.net/terminos.html */
// Las marcas numeradas para elegir varios objetos con la IA (SlimSAM,
// objeto-por-toque.js): cada toque deja una marca, tocar una marca la quita, y
// el modelo busca el objeto de cada una. Lo usan el Generador de stickers
// ("Los objetos que marco", 2026-10-08) y "Elegir objeto" del Editor avanzado
// (2026-10-09).
//
// Las marcas se guardan en 0..1 de la imagen; acá se reciben ya pasadas a
// píxeles del lienzo donde se dibujan, con el radio con que se ven.
//   MarcasObjetos.dibujar(ctx, [{ x, y }, …], radio);
//   const i = MarcasObjetos.cercana(puntos, x, y, radio); // -1 si ninguna

const MarcasObjetos = (() => {
  // Círculo blanco con borde y número rojos: se ve sobre cualquier foto.
  function dibujar(ctx, puntos, radio) {
    ctx.save();
    ctx.font = `bold ${Math.round(radio * 1.1)}px system-ui, sans-serif`;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.lineWidth = Math.max(2, radio * 0.22);
    puntos.forEach((m, i) => {
      ctx.beginPath();
      ctx.arc(m.x, m.y, radio, 0, Math.PI * 2);
      ctx.fillStyle = '#ffffff';
      ctx.fill();
      ctx.strokeStyle = '#e11d48';
      ctx.stroke();
      ctx.fillStyle = '#e11d48';
      ctx.fillText(String(i + 1), m.x, m.y + 1);
    });
    ctx.restore();
  }

  // La marca que está bajo (x, y), con un poco de margen para el dedo.
  function cercana(puntos, x, y, radio) {
    return puntos.findIndex((m) => Math.hypot(m.x - x, m.y - y) < radio * 1.3);
  }

  return { dibujar, cercana };
})();
