// Copia de reserva de la aplicación en el propio dispositivo.
// Regla: SIEMPRE se intenta primero la red; lo guardado solo se usa cuando el servidor de la app no responde
// (hay redes que lo rechazan). Así nunca se queda una versión vieja por culpa de la copia.
// Solo toca los ficheros de la propia app: Outlook, Microsoft y la IA pasan sin intervenir.
"use strict";
const COPIA = "agente-app-1";
const INICIO = new URL("./", self.location).href;
const ESPERA = 7000;

self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => e.waitUntil((async () => {
  for (const k of await caches.keys()) if (k !== COPIA && k.startsWith("agente-app-")) await caches.delete(k);
  await self.clients.claim();
})()));

const clave = (peticion) => (peticion.mode === "navigate" ? INICIO : peticion.url);

async function avisarCopia() {
  for (const c of await self.clients.matchAll({type: "window"})) c.postMessage({tipo: "copia"});
}

self.addEventListener("fetch", (e) => {
  const p = e.request;
  if (p.method !== "GET") return;
  const u = new URL(p.url);
  if (u.origin !== self.location.origin || !u.href.startsWith(INICIO)) return;
  // El instalador, su aviso de versión y la comprobación de versión nueva van siempre a la red.
  if (u.pathname.includes("/apk/") || u.pathname.endsWith("/apk.json") || (u.pathname.endsWith("/version.js") && u.searchParams.has("t"))) return;
  e.respondWith((async () => {
    const copia = await caches.open(COPIA);
    const guardada = () => copia.match(clave(p));
    const red = fetch(p).then((r) => {
      if (r && r.ok && r.type === "basic") copia.put(clave(p), r.clone()).catch(() => {});
      return r;
    });
    try {
      // Si la red tarda demasiado y hay copia, se sirve la copia; si no la hay, se sigue esperando a la red.
      const ganador = await Promise.race([red, new Promise((ok) => setTimeout(() => ok(null), ESPERA))]);
      if (ganador) {
        if (ganador.ok || ganador.type !== "basic") return ganador;
        const g = await guardada();          // el servidor responde con error: mejor la copia, si existe
        if (g) { avisarCopia(); return g; }
        return ganador;
      }
      const g = await guardada();
      if (g) { red.catch(() => {}); avisarCopia(); return g; }
      return await red;
    } catch (err) {
      const g = await guardada();
      if (g) { avisarCopia(); return g; }
      throw err;
    }
  })());
});

// La página pide guardar de antemano los ficheros que todavía no se han usado (el visor de PDF, por ejemplo).
self.addEventListener("message", (e) => {
  const d = e.data || {};
  if (d.tipo !== "guardar" || !Array.isArray(d.urls)) return;
  e.waitUntil((async () => {
    const copia = await caches.open(COPIA);
    for (const dir of d.urls.slice(0, 60)) {
      try {
        const u = new URL(dir, INICIO);
        if (u.origin !== self.location.origin || !u.href.startsWith(INICIO)) continue;
        const k = u.href === INICIO || u.pathname.endsWith("/index.html") ? INICIO : u.href;
        const r = await fetch(u.href, {cache: "no-cache"});
        if (r.ok) await copia.put(k, r);
      } catch (err) { /* sin red ahora: se guardará en la próxima visita */ }
    }
  })());
});
