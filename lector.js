"use strict";
/* Muestra un correo con su formato dentro de un marco aislado.
   Esta página no tiene sesión ni acceso al buzón: solo recibe el HTML del correo,
   le quita todo lo que no sea contenido y lo pinta. No puede hacer peticiones de red
   salvo cargar imágenes, y no ejecuta nada que venga en el correo. */
(function () {
  const PROHIBIDO = "script,iframe,frame,frameset,object,embed,applet,form,input,button,select,textarea,link,meta,base,audio,video,source,track,svg,math,noscript,template,dialog,portal,title";
  let pintado = false;

  function limpiar(doc) {
    doc.querySelectorAll(PROHIBIDO).forEach((n) => n.remove());
    doc.querySelectorAll("*").forEach((n) => {
      for (const a of Array.from(n.attributes)) {
        const k = a.name.toLowerCase();
        if (k.startsWith("on") || k === "srcdoc" || k === "formaction" || k === "ping" || k === "srcset") n.removeAttribute(a.name);
      }
    });
    doc.querySelectorAll("a, area").forEach((a) => {
      let vale = false;
      try { vale = ["https:", "http:", "mailto:"].includes(new URL(a.getAttribute("href")).protocol); } catch (e) {}
      if (!vale) { a.removeAttribute("href"); return; }
      a.setAttribute("target", "_blank");
      a.setAttribute("rel", "noopener noreferrer");
    });
    doc.querySelectorAll("img").forEach((i) => {
      const s = i.getAttribute("src") || "";
      if (!/^(https:|data:image\/)/i.test(s)) i.remove();
    });
  }

  // Marca dónde empieza cada mensaje anterior del hilo y quita el aviso de «fuente externa».
  function separar() {
    const reCita = /^\s*(De|From|Von|Da)\s?:[\s\S]{3,300}?(Enviat el|Enviado el|Enviado|Enviat|Sent|Envoyé|Date|Fecha|Data|Gesendet)\s?:/i;
    const reAviso = /^\s*Precaución: Este correo[\s\S]{0,220}archivos adjuntos\.\s*$/i;
    const bloques = Array.from(document.body.querySelectorAll("div, p, td, blockquote"));
    for (const n of bloques) {
      const t = n.textContent || "";
      if (t.length > 900) continue;
      const interior = Array.from(n.querySelectorAll("div, p, td")).some((h) => (reCita.test(h.textContent || "") || reAviso.test(h.textContent || "")));
      if (interior) continue;
      if (reAviso.test(t)) {
        // El aviso suele venir en una tabla amarilla: se quita la tabla entera, no solo el texto.
        let quitar = n;
        for (let p = n.parentElement; p && p !== document.body; p = p.parentElement) if (reAviso.test(p.textContent || "")) quitar = p; else break;
        quitar.remove();
      }
      else if (reCita.test(t)) n.setAttribute("style", "border:0;border-top:1px solid #CBD3DC;border-left:5px solid #C2935A;background:#E9EDF1;color:#55606E;padding:9px 12px;margin:22px 0 12px;font-size:14px;line-height:1.4");
    }
    document.body.querySelectorAll("hr").forEach((h) => h.remove());
  }

  function ajustar() {
    const b = document.body;
    b.style.zoom = "";
    const ancho = document.documentElement.scrollWidth, hueco = window.innerWidth;
    if (hueco > 0 && ancho > hueco + 2) b.style.zoom = String(Math.max(0.3, hueco / ancho));
    parent.postMessage({tipo: "alto", alto: Math.ceil(document.documentElement.scrollHeight)}, "*");
  }

  window.addEventListener("message", (e) => {
    if (e.source !== parent || pintado) return;
    const d = e.data;
    if (!d || d.tipo !== "correo" || typeof d.html !== "string") return;
    pintado = true;
    const doc = new DOMParser().parseFromString(d.html, "text/html");
    limpiar(doc);
    doc.querySelectorAll("style").forEach((s) => document.head.append(document.adoptNode(s)));
    document.body.replaceChildren(...Array.from(doc.body.childNodes).map((n) => document.adoptNode(n)));
    separar();
    document.querySelectorAll("img").forEach((i) => { i.addEventListener("load", ajustar); i.addEventListener("error", ajustar); });
    ajustar();
    setTimeout(ajustar, 400);
    setTimeout(ajustar, 2000);
  });
  window.addEventListener("resize", ajustar);
  parent.postMessage({tipo: "listo"}, "*");
})();
