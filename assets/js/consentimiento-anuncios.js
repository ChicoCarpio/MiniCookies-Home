/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// Consentimiento de anuncios fuera de Europa (2026-11-15, revisión legal, Ley 8968 de Costa Rica).
// Google solo muestra su propio aviso en Europa (UE, Reino Unido, Suiza) y en algunos estados de EE. UU.; en el resto del mundo
// los anuncios saldrían personalizados sin preguntar. El build pone antes de adsbygoogle.js una línea que PAUSA los pedidos de
// anuncios (pauseAdRequests = 1) y esta pieza:
//   - si el visitante ya eligió (localStorage 'mc-anuncios' = 'si' | 'no'), aplica la elección y reanuda;
//   - si está donde Google muestra su aviso europeo (TCF: gdprApplies), reanuda y deja que decida el aviso de Google;
//   - si no, muestra un aviso abajo con dos botones: personalizados o no personalizados (requestNonPersonalizedAds).
// Sin elección no se pide ningún anuncio. window.ConsentimientoAnuncios.abrir() vuelve a mostrar el aviso (página de privacidad).
(function () {
  'use strict';
  const CLAVE = 'mc-anuncios';
  const ingles = () => /^en/i.test(document.documentElement.lang || '');
  const TEXTOS = {
    es: {
      titulo: 'Mini Cookies te pide permiso para personalizar los anuncios',
      texto: 'Mini Cookies es gratis gracias a los anuncios de Google. Si aceptás, Google usa cookies para mostrarte anuncios según tus intereses. Si no, igual ves anuncios, pero no personalizados. Podés cambiarlo cuando quieras en la página de privacidad.',
      si: 'Aceptar', no: 'No personalizados', mas: 'Más información', logo: 'Mini Cookies',
    },
    en: {
      titulo: 'Mini Cookies asks for your permission to personalise ads',
      texto: 'Mini Cookies is free thanks to Google ads. If you accept, Google uses cookies to show you ads based on your interests. If not, you still see ads, just not personalised ones. You can change this at any time on the privacy page.',
      si: 'Accept', no: 'Not personalised', mas: 'Learn more', logo: 'Mini Cookies',
    },
  };

  // Raíz del sitio (este archivo está en assets/js/), tomada al cargar: después document.currentScript ya no existe.
  const raiz = new URL('../../', document.currentScript ? document.currentScript.src : location.href);

  function leer() { try { return localStorage.getItem(CLAVE); } catch (e) { return null; } }
  function guardar(v) { try { localStorage.setItem(CLAVE, v); } catch (e) { /* modo privado: se pregunta otra vez la próxima visita */ } }

  function aplicar(v) {
    const a = (window.adsbygoogle = window.adsbygoogle || []);
    a.requestNonPersonalizedAds = v === 'si' ? 0 : 1;
    a.pauseAdRequests = 0;
  }
  function reanudarSinTocar() { (window.adsbygoogle = window.adsbygoogle || []).pauseAdRequests = 0; }

  let caja = null;
  function cerrar() { if (caja) { caja.remove(); caja = null; } }
  function abrir() {
    if (caja) return;
    const t = TEXTOS[ingles() ? 'en' : 'es'];
    // Tarjeta centrada con el mismo aire que el aviso europeo de Google (logo, título, texto y dos botones), más chica.
    caja = document.createElement('div');
    caja.className = 'aviso-anuncios-fondo';
    const tarjeta = document.createElement('div');
    tarjeta.className = 'aviso-anuncios';
    tarjeta.setAttribute('role', 'dialog');
    tarjeta.setAttribute('aria-modal', 'true');
    tarjeta.setAttribute('aria-labelledby', 'avisoAnunciosTitulo');
    const logo = document.createElement('img');
    logo.src = new URL('marca-logo/mini-cookies-logo-horizontal.png', raiz).href;
    logo.alt = t.logo; logo.className = 'aviso-anuncios-logo';
    const h = document.createElement('h2');
    h.id = 'avisoAnunciosTitulo'; h.textContent = t.titulo;
    const p = document.createElement('p');
    p.textContent = t.texto + ' ';
    const a = document.createElement('a');
    a.href = new URL((ingles() ? 'en/' : '') + 'privacidad#h-terceros', raiz).href;
    a.textContent = t.mas;
    p.appendChild(a);
    const botones = document.createElement('div');
    botones.className = 'aviso-anuncios-botones';
    let primero = null;
    for (const [v, etiqueta, clase] of [['no', t.no, 'aviso-anuncios-no'], ['si', t.si, 'aviso-anuncios-si']]) {
      const b = document.createElement('button');
      b.type = 'button'; b.className = clase; b.textContent = etiqueta;
      b.addEventListener('click', () => { guardar(v); aplicar(v); cerrar(); });
      botones.appendChild(b);
      if (!primero) primero = b;
    }
    tarjeta.append(logo, h, p, botones);
    caja.appendChild(tarjeta);
    document.body.appendChild(caja);
    primero.focus({ preventScroll: true });
  }

  // ¿Muestra Google su aviso europeo a este visitante? Lo dice la API de TCF que carga el propio aviso de Google.
  // Si en 3 s no aparece, se decide por la zona horaria (Europa = Google se encarga).
  function enEuropa(listo) {
    let resuelto = false;
    const fin = (v) => { if (!resuelto) { resuelto = true; listo(v); } };
    const preguntar = () => {
      try {
        window.__tcfapi('getTCData', 2, (datos, ok) => { if (ok && datos && typeof datos.gdprApplies === 'boolean') fin(datos.gdprApplies); });
      } catch (e) { /* sigue esperando */ }
    };
    const inicio = Date.now();
    (function esperar() {
      if (resuelto) return;
      if (typeof window.__tcfapi === 'function') preguntar();
      if (Date.now() - inicio > 3000) {
        let zona = '';
        try { zona = Intl.DateTimeFormat().resolvedOptions().timeZone || ''; } catch (e) { /* sin zona */ }
        return fin(/^Europe\//.test(zona) || /^Atlantic\/(Canary|Madeira|Azores|Reykjavik|Faroe)$/.test(zona));
      }
      setTimeout(esperar, 250);
    })();
  }

  function iniciar() {
    const elegido = leer();
    if (elegido === 'si' || elegido === 'no') return aplicar(elegido);
    enEuropa((europa) => { if (europa) reanudarSinTocar(); else abrir(); });
  }

  window.ConsentimientoAnuncios = { abrir, elegido: leer };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar);
  else iniciar();
})();
