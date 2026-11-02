/*! © 2026 Daniels Torres Solano (Mini Cookies). Todos los derechos reservados. Código fuente bajo la Licencia de Código Fuente de Mini Cookies: https://www.minicookies.net/terminos#codigo-fuente */
// La cámara de las herramientas del arquetipo H (SISTEMA-DE-INTERFAZ.md): abrirla, enfocar,
// cambiar de cámara y sacar un cuadro. Va junto con assets/css/camara.css.
//
//   const cam = await Camara.abrir(video);        // trasera si hay; lanza error si no se puede
//   cam.enfocarEn(x, y);                          // x, y en fracción (0..1) del cuadro de video
//   cam.seguir(x, y);                             // enfoque continuo apuntado a un objetivo que se mueve
//   const p = cam.puntoEnVideo(evento);           // dónde se tocó, en fracción, o null (franja negra)
//   const lienzo = cam.capturar(4096);            // cuadro actual en un canvas (lado mayor ≤ 4096)
//   await cam.cambiar();                          // pasa a la siguiente cámara (delantera/trasera…)
//   cam.cuantas                                   // cuántas cámaras hay (para mostrar el botón de cambiar)
//   cam.apagar();
//
// Enfoque: al abrir se pide el enfoque continuo si la cámara lo permite (celulares con Chrome o
// Edge). Tocar enfoca ese punto (puntos de interés) y a los 2,5 s vuelve al continuo. Donde el
// navegador no deja controlar el enfoque, la cámara usa el suyo y no pasa nada más.
// El video tiene que mostrarse con object-fit: contain (lo hace camara.css): puntoEnVideo lo asume.

window.Camara = (function () {
  'use strict';

  function modosDeEnfoque(pista) {
    try { return (pista.getCapabilities && pista.getCapabilities().focusMode) || []; } catch { return []; }
  }

  async function abrir(video, opciones = {}) {
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
      const error = new Error('sin-camara');
      error.name = 'SinCamara';
      throw error;
    }
    let flujo = null, idActual = opciones.idDispositivo || null, temporizador = 0;

    async function encender(id) {
      if (flujo) flujo.getTracks().forEach((p) => p.stop());
      flujo = await navigator.mediaDevices.getUserMedia({
        audio: false,
        video: {
          width: { ideal: 4096 }, height: { ideal: 2160 },
          ...(id ? { deviceId: { exact: id } } : { facingMode: { ideal: 'environment' } }),
        },
      });
      video.srcObject = flujo;
      video.muted = true;
      video.setAttribute('playsinline', '');
      await video.play();
      idActual = flujo.getVideoTracks()[0].getSettings().deviceId || id;
      continuo();
    }

    const pista = () => flujo && flujo.getVideoTracks()[0];

    async function continuo() {
      const p = pista();
      if (!p || !modosDeEnfoque(p).includes('continuous')) return;
      try { await p.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }); } catch { /* sin enfoque controlable */ }
    }

    await encender(idActual);
    // Los nombres y la cantidad de cámaras solo se conocen después de dar el permiso.
    let camaras = [];
    try { camaras = (await navigator.mediaDevices.enumerateDevices()).filter((d) => d.kind === 'videoinput'); } catch { /* sin lista */ }

    return {
      video,
      get cuantas() { return camaras.length; },
      get id() { return idActual; },
      async cambiar() {
        if (camaras.length < 2) return;
        const i = camaras.findIndex((c) => c.deviceId === idActual);
        await encender(camaras[(i + 1) % camaras.length].deviceId);
      },
      async enfocarEn(x, y) {
        const p = pista();
        if (!p) return;
        try {
          await p.applyConstraints({ advanced: [{ pointsOfInterest: [{ x, y }], ...(modosDeEnfoque(p).includes('single-shot') ? { focusMode: 'single-shot' } : {}) }] });
        } catch { /* la cámara no acepta enfocar un punto */ }
        clearTimeout(temporizador);
        temporizador = setTimeout(continuo, 2500);
      },
      // Seguir un objetivo: el enfoque sigue siendo continuo, pero en esa zona del cuadro.
      // Pensado para llamarse cada tanto mientras el objetivo se mueve (no después de enfocarEn,
      // que espera su propio enfoque).
      async seguir(x, y) {
        const p = pista();
        if (!p) return;
        try {
          await p.applyConstraints({ advanced: [{ pointsOfInterest: [{ x, y }], ...(modosDeEnfoque(p).includes('continuous') ? { focusMode: 'continuous' } : {}) }] });
        } catch { /* la cámara no acepta puntos de interés */ }
      },
      puntoEnVideo(evento) {
        if (!video.videoWidth) return null;
        const caja = video.getBoundingClientRect();
        const k = Math.min(caja.width / video.videoWidth, caja.height / video.videoHeight);
        const x = (evento.clientX - caja.left - (caja.width - video.videoWidth * k) / 2) / (video.videoWidth * k);
        const y = (evento.clientY - caja.top - (caja.height - video.videoHeight * k) / 2) / (video.videoHeight * k);
        return x < 0 || x > 1 || y < 0 || y > 1 ? null : { x, y };
      },
      capturar(ladoMax = 4096) {
        const escala = Math.min(1, ladoMax / Math.max(video.videoWidth, video.videoHeight));
        const lienzo = document.createElement('canvas');
        lienzo.width = Math.round(video.videoWidth * escala);
        lienzo.height = Math.round(video.videoHeight * escala);
        lienzo.getContext('2d').drawImage(video, 0, 0, lienzo.width, lienzo.height);
        return lienzo;
      },
      apagar() {
        clearTimeout(temporizador);
        if (flujo) flujo.getTracks().forEach((p) => p.stop());
        flujo = null;
        video.srcObject = null;
      },
    };
  }

  return { abrir };
})();
