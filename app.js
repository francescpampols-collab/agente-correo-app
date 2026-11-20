"use strict";
// La app no se muestra dentro de otra página (evita que la superpongan para robar pulsaciones).
if (window.top !== window.self) {
  document.documentElement.textContent = "";
  throw new Error("No se permite abrir la app dentro de un marco");
}

const GRAPH = "https://graph.microsoft.com/v1.0";
const PROP = "String {00020329-0000-0000-C000-000000000046} Name AgenteCorreo";
const PROP_LISTA = "String {00020329-0000-0000-C000-000000000046} Name AgenteCorreoSiempre";
const CAT = {resp: "Agente: responder", news: "Agente: boletín", arch: "Agente: archivar", pend: "Agente: pendiente", baja: "Agente: baja", manual: "Agente: baja manual"};
const PESTANAS = [["resp", "Responder"], ["news", "Boletines"], ["arch", "Archivar"], ["pend", "Tuyos"]];
const SCOPE = "openid profile offline_access https://graph.microsoft.com/Mail.ReadWrite https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read";
const REDIR = location.origin + location.pathname;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const $ = (id) => document.getElementById(id);
const estado = {tab: "resp", sel: null, datos: {resp: [], news: [], arch: [], pend: []}, lista: [], carpetas: null, cuerpos: {}, confirma: false};

// ---------- utilidades ----------
function guardar(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
function leer(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
function borrar(k) { try { localStorage.removeItem(k); } catch (e) {} }
// Todo el contenido de los correos entra en la página como texto, nunca como HTML.
function el(tag, attrs, ...hijos) {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (k === "onclick") n.addEventListener("click", v);
    else if (k === "class") n.className = v;
    else if (v !== false && v != null) n.setAttribute(k, v);
  }
  for (const h of hijos) if (h != null) n.append(h);
  return n;
}
let temporizador;
function avisar(texto) {
  $("aviso").textContent = texto;
  $("aviso").classList.remove("oculto");
  clearTimeout(temporizador);
  temporizador = setTimeout(() => $("aviso").classList.add("oculto"), 5000);
}
function b64url(bytes) {
  let s = "";
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function escapar(t) { return t.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;"); }
function textoAHtml(t) {
  return t.trim().split(/\n\s*\n/).filter((b) => b.trim()).map((b) => "<p>" + escapar(b.trim()).replace(/\n/g, "<br>") + "</p>").join("");
}
function insertarEnCuerpo(cuerpo, bloque) {
  const r = /(<body[^>]*>)/i;
  return r.test(cuerpo || "") ? cuerpo.replace(r, (m) => m + bloque) : bloque + (cuerpo || "");
}
function prop(m) {
  try { return JSON.parse((m.singleValueExtendedProperties || [])[0].value) || {}; } catch (e) { return {}; }
}
function de(m) {
  const e = (m.from || {}).emailAddress || {};
  return e.name && e.address ? `${e.name} <${e.address}>` : e.address || e.name || "(sin remitente)";
}
// Quita del extracto el aviso que el servidor de correo antepone a los mensajes externos.
function extracto(m) {
  return (m.bodyPreview || "").replace(/^\s*Precaución:[\s\S]*?archivos adjuntos\.\s*/i, "").trim();
}
function remitente(m, etiqueta) {
  const e = (m.from || {}).emailAddress || {};
  return el(etiqueta || "span", {class: "de"}, el("b", {}, e.name || e.address || "(sin remitente)"), e.name && e.address ? " " + e.address : null);
}
function dir(m) { return (((m.from || {}).emailAddress || {}).address || "").toLowerCase(); }
function enlaceSeguro(u) {
  try { const x = new URL(u); return x.protocol === "https:" && !x.username && !x.password ? x : null; } catch (e) { return null; }
}

// ---------- sesión con Microsoft (código de autorización + PKCE) ----------
function config() {
  const c = leer("agente.cfg");
  return c && GUID.test(c.c || "") && GUID.test(c.t || "") ? c : null;
}
async function entrar() {
  const c = config();
  borrar("agente.salio");
  const verificador = b64url(crypto.getRandomValues(new Uint8Array(48)));
  const reto = b64url(new Uint8Array(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verificador))));
  const st = b64url(crypto.getRandomValues(new Uint8Array(16)));
  sessionStorage.setItem("agente.pkce", JSON.stringify({verificador, st}));
  const q = new URLSearchParams({client_id: c.c, response_type: "code", redirect_uri: REDIR, response_mode: "query", scope: SCOPE, code_challenge: reto, code_challenge_method: "S256", state: st});
  location.assign(`https://login.microsoftonline.com/${c.t}/oauth2/v2.0/authorize?${q}`);
}
async function pedirToken(params) {
  const c = config();
  const r = await fetch(`https://login.microsoftonline.com/${c.t}/oauth2/v2.0/token`, {
    method: "POST", headers: {"Content-Type": "application/x-www-form-urlencoded"},
    body: new URLSearchParams({client_id: c.c, scope: SCOPE, redirect_uri: REDIR, ...params}),
  });
  const j = await r.json();
  if (!r.ok) throw new Error(j.error || "Error de sesión");
  const anterior = leer("agente.tok") || {};
  guardar("agente.tok", {a: j.access_token, r: j.refresh_token || anterior.r, exp: Date.now() + (j.expires_in - 60) * 1000});
  return j.access_token;
}
async function token() {
  const t = leer("agente.tok");
  if (t && t.a && t.exp > Date.now()) return t.a;
  if (t && t.r) {
    try { return await pedirToken({grant_type: "refresh_token", refresh_token: t.r}); } catch (e) { borrar("agente.tok"); }
  }
  await entrar();
  return new Promise(() => {});
}
async function api(metodo, ruta, cuerpo, tipoCuerpo) {
  const cab = {Authorization: "Bearer " + await token(), Prefer: `IdType="ImmutableId", outlook.body-content-type="${tipoCuerpo || "text"}"`};
  if (cuerpo !== undefined) cab["Content-Type"] = "application/json";
  let r;
  // Outlook admite pocas peticiones a la vez por buzón: si pide esperar, se espera y se repite.
  for (let intento = 0; ; intento++) {
    r = await fetch(GRAPH + ruta, {method: metodo, headers: cab, body: cuerpo !== undefined ? JSON.stringify(cuerpo) : undefined});
    if ((r.status !== 429 && r.status !== 503) || intento >= 4) break;
    const espera = Math.min(Number(r.headers && r.headers.get && r.headers.get("Retry-After")) || 2 * (intento + 1), 15);
    await new Promise((ok) => setTimeout(ok, espera * 1000));
  }
  if (r.status === 401) { borrar("agente.tok"); await entrar(); return new Promise(() => {}); }
  if (!r.ok) throw new Error(`Outlook ha respondido ${r.status}`);
  return r.status === 204 || r.status === 202 ? null : r.json().catch(() => null);
}
const id = (x) => encodeURIComponent(x);

// ---------- datos ----------
async function listarTodo() {
  const etiquetas = [CAT.resp, CAT.news, CAT.arch, CAT.pend, CAT.manual];
  const q = new URLSearchParams({
    "$filter": etiquetas.map((e) => `categories/any(c:c eq '${e}')`).join(" or "),
    "$select": "id,subject,from,bodyPreview,receivedDateTime,categories,webLink",
    "$expand": `singleValueExtendedProperties($filter=id eq '${PROP}')`,
    "$top": "200",
  });
  const j = await api("GET", "/me/mailFolders/inbox/messages?" + q);
  const lista = (j.value || []).sort((a, b) => (b.receivedDateTime || "").localeCompare(a.receivedDateTime || ""));
  await completar(lista);
  const de = (e) => lista.filter((m) => (m.categories || []).includes(e));
  return {resp: de(CAT.resp), news: de(CAT.news), arch: de(CAT.arch), pend: de(CAT.pend).concat(de(CAT.manual))};
}
// En los listados Outlook corta la propiedad a 255 caracteres. Las que llegan
// cortadas se vuelven a pedir enteras, de cuatro en cuatro (el tope de Outlook).
async function completar(lista) {
  const cortadas = lista.filter((m) => {
    const v = ((m.singleValueExtendedProperties || [])[0] || {}).value || "";
    if (!v) return false;
    try { JSON.parse(v); return false; } catch (e) { return true; }
  });
  const exp = encodeURIComponent(`singleValueExtendedProperties($filter=id eq '${PROP}')`);
  for (let i = 0; i < cortadas.length; i += 4) {
    const grupo = cortadas.slice(i, i + 4);
    const r = await api("POST", "/$batch", {requests: grupo.map((m, k) => ({id: String(k), method: "GET", url: `/me/messages/${id(m.id)}?$select=id&$expand=${exp}`, headers: {Prefer: 'IdType="ImmutableId"'}}))});
    for (const x of (r && r.responses) || []) {
      const m = grupo[Number(x.id)];
      if (m && x.status === 200 && x.body && x.body.singleValueExtendedProperties) m.singleValueExtendedProperties = x.body.singleValueExtendedProperties;
    }
  }
}
// Remitentes de «archivar siempre»: se guardan en la propia Bandeja de entrada de Outlook.
async function leerLista() {
  const q = new URLSearchParams({"$select": "id", "$expand": `singleValueExtendedProperties($filter=id eq '${PROP_LISTA}')`});
  const f = await api("GET", "/me/mailFolders/inbox?" + q);
  try {
    const v = JSON.parse((f.singleValueExtendedProperties || [])[0].value);
    return Array.isArray(v) ? v.filter((x) => typeof x === "string") : [];
  } catch (e) { return []; }
}
async function guardarLista(lista) {
  await api("PATCH", "/me/mailFolders/inbox", {singleValueExtendedProperties: [{id: PROP_LISTA, value: JSON.stringify(lista)}]});
  estado.lista = lista;
}
// Carpetas de primer nivel del buzón, para archivar donde indique el motor (internos o externos).
async function leerCarpetas() {
  const j = await api("GET", "/me/mailFolders?$top=200&$select=id,displayName");
  const mapa = {};
  for (const f of j.value || []) mapa[f.displayName] = f.id;
  return mapa;
}
function carpetaDe(m) {
  const d = prop(m).destino;
  return typeof d === "string" && estado.carpetas && estado.carpetas[d] ? {nombre: d, id: estado.carpetas[d]} : {nombre: "Archivo", id: "archive"};
}
async function cargar() {
  $("recargar").disabled = true;
  try {
    if (!estado.carpetas) estado.carpetas = await leerCarpetas();
    estado.datos = await listarTodo();
    estado.lista = await leerLista();
    if (estado.sel && !todos().some((m) => m.id === estado.sel)) cerrar();
  } catch (e) { avisar("No se ha podido leer el buzón: " + e.message); }
  $("recargar").disabled = false;
  pintar();
}
function todos() { return [].concat(estado.datos.resp, estado.datos.news, estado.datos.arch, estado.datos.pend); }
function quitar(ids) {
  for (const k of Object.keys(estado.datos)) estado.datos[k] = estado.datos[k].filter((m) => !ids.includes(m.id));
  if (ids.includes(estado.sel)) cerrar();
}
function grupos() {
  const g = new Map();
  for (const m of estado.datos.news) {
    const k = dir(m) || de(m);
    if (!g.has(k)) g.set(k, {nombre: ((m.from || {}).emailAddress || {}).name || dir(m) || "(sin remitente)", dir: dir(m), items: []});
    g.get(k).items.push(m);
  }
  return [...g.values()];
}

// ---------- acciones sobre el buzón ----------
async function archivar(m) {
  await api("PATCH", `/me/messages/${id(m.id)}`, {categories: (m.categories || []).filter((c) => !c.startsWith("Agente:"))});
  await api("POST", `/me/messages/${id(m.id)}/move`, {destinationId: carpetaDe(m).id});
}
async function enviar(m, texto) {
  const p = prop(m);
  if (texto != null && texto.trim() !== (p.respuesta || "").trim()) {
    const nuevo = await api("POST", `/me/messages/${id(m.id)}/${p.todos ? "createReplyAll" : "createReply"}`, undefined, "html");
    const cuerpo = insertarEnCuerpo((nuevo.body || {}).content || "", textoAHtml(texto));
    await api("PATCH", `/me/messages/${id(nuevo.id)}`, {body: {contentType: "html", content: cuerpo}});
    await api("POST", `/me/messages/${id(nuevo.id)}/send`);
    if (p.draftId) await api("DELETE", `/me/messages/${id(p.draftId)}`).catch(() => {});
  } else {
    if (!p.draftId) throw new Error("No encuentro el borrador de este correo");
    await api("POST", `/me/messages/${id(p.draftId)}/send`);
  }
  await archivar(m);
}
async function descartar(m) {
  const p = prop(m);
  if (p.draftId) await api("DELETE", `/me/messages/${id(p.draftId)}`).catch(() => {});
  const cats = (m.categories || []).filter((c) => !c.startsWith("Agente:")).concat(CAT.pend);
  await api("PATCH", `/me/messages/${id(m.id)}`, {categories: cats});
}
// Corrección de Francesc: esto no es para archivar, lo tiene que gestionar él.
async function pasarATuyos(m) {
  const cats = (m.categories || []).filter((c) => !c.startsWith("Agente:")).concat(CAT.pend);
  await api("PATCH", `/me/messages/${id(m.id)}`, {categories: cats});
}
async function pedirBaja(grupo) {
  const m = grupo.items[0];
  const cats = (m.categories || []).filter((c) => !c.startsWith("Agente:")).concat(CAT.baja);
  await api("PATCH", `/me/messages/${id(m.id)}`, {categories: cats});
}
async function archivarSiempre(grupo) {
  const lista = await leerLista();
  if (!lista.includes(grupo.dir)) await guardarLista(lista.concat(grupo.dir));
  for (const m of grupo.items) await archivar(m);
}
async function hacer(tarea, ids, hecho, recargar) {
  const botones = document.querySelectorAll("button");
  botones.forEach((b) => (b.disabled = true));
  try {
    await tarea();
    quitar(ids);
    if (hecho) avisar(hecho);
  } catch (e) { avisar("No se ha podido completar: " + e.message); }
  botones.forEach((b) => (b.disabled = false));
  pintar();
  if (recargar) await cargar();
}

// ---------- pantalla ----------
function abrir(mid) {
  estado.sel = mid;
  document.body.classList.add("viendo");
  pintar();
  if (!estado.cuerpos[mid]) cargarCorreo(mid);
}
// Trae el correo entero (texto y lista de adjuntos) para leerlo dentro de la app.
async function cargarCorreo(mid) {
  try {
    const j = await api("GET", `/me/messages/${id(mid)}?$select=subject,body,toRecipients,ccRecipients,hasAttachments,receivedDateTime`);
    const c = {texto: (((j || {}).body || {}).content || "").replace(/^\s*Precaución:[\s\S]*?archivos adjuntos\.\s*/i, "").replace(/\r/g, "").replace(/\[cid:[^\]]*\]\s*/g, ""), para: [], adjuntos: [], fecha: (j || {}).receivedDateTime || ""};
    for (const r of ((j || {}).toRecipients || []).concat((j || {}).ccRecipients || [])) if (r.emailAddress && r.emailAddress.address) c.para.push(r.emailAddress.address);
    if ((j || {}).hasAttachments) {
      const a = await api("GET", `/me/messages/${id(mid)}/attachments?$select=id,name,size,contentType`);
      c.adjuntos = ((a || {}).value || []).filter((x) => x.name);
    }
    c.texto = limpiarTexto(c.texto);
    estado.cuerpos[mid] = c;
  } catch (e) { estado.cuerpos[mid] = {error: e.message, texto: "", para: [], adjuntos: []}; }
  if (estado.sel === mid) pintarDetalle();
}
async function descargar(mid, adj) {
  const r = await fetch(`${GRAPH}/me/messages/${id(mid)}/attachments/${id(adj.id)}/$value`, {headers: {Authorization: "Bearer " + await token(), Prefer: 'IdType="ImmutableId"'}});
  if (!r.ok) throw new Error(`Outlook ha respondido ${r.status}`);
  const url = URL.createObjectURL(await r.blob());
  const a = el("a", {href: url, download: adj.name});
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
function alPulsarCarta(mid) {
  return (e) => { if (!e.target.closest("button, a")) abrir(mid); };
}
function titulo(m) {
  return el("button", {class: "asunto", type: "button", onclick: () => abrir(m.id)}, m.subject || "(sin asunto)");
}
function cerrar() { estado.sel = null; document.body.classList.remove("viendo"); }
function faltaDato(m, texto) { return /\[FALTA/i.test(texto != null ? texto : prop(m).respuesta || ""); }
function destinatarios(p) { return Array.isArray(p.para) && p.para.length ? "Se envía a: " + p.para.slice(0, 6).join(", ") : null; }

function cartaRespuesta(m) {
  const p = prop(m);
  return el("div", {class: "carta" + (m.id === estado.sel ? " sel" : ""), onclick: alPulsarCarta(m.id)},
    remitente(m),
    titulo(m),
    p.pide ? el("p", {class: "pide"}, "Pide: " + p.pide) : null,
    el("p", {class: "resp"}, p.respuesta || ""),
    destinatarios(p) ? el("span", {class: "de"}, destinatarios(p)) : null,
    el("span", {class: "de"}, "Al enviar se archiva en " + carpetaDe(m).nombre),
    p.aviso ? el("span", {class: "aviso"}, "Revisar bien: " + p.aviso) : null,
    faltaDato(m) ? el("span", {class: "aviso"}, "Falta un dato") : null,
    el("div", {class: "fila"},
      el("button", {class: "plano", type: "button", onclick: () => abrir(m.id)}, "Ver y editar"),
      el("button", {class: "lleno", type: "button", disabled: faltaDato(m), onclick: () => hacer(() => enviar(m, null), [m.id], "Enviado y archivado en " + carpetaDe(m).nombre)}, "Enviar")));
}
function cartaGrupo(g) {
  const ids = g.items.map((m) => m.id);
  const posible = g.items.some((m) => prop(m).baja_posible);
  return el("div", {class: "carta" + (g.items.some((m) => m.id === estado.sel) ? " sel" : ""), onclick: alPulsarCarta(g.items[0].id)},
    el("span", {class: "asunto"}, g.nombre),
    g.dir && g.dir !== g.nombre ? el("span", {class: "de"}, g.dir) : null,
    ...g.items.slice(0, 5).map((m) => el("button", {class: "pide enlace", type: "button", onclick: () => abrir(m.id)}, m.subject || "(sin asunto)")),
    g.items.length > 5 ? el("span", {class: "pide"}, "y " + (g.items.length - 5) + " más") : null,
    el("div", {class: "fila"},
      el("button", {class: "plano archivo", type: "button", onclick: () => hacer(async () => { for (const m of g.items) await archivar(m); }, ids, "Archivado en " + carpetaDe(g.items[0]).nombre)}, "Archivar en " + carpetaDe(g.items[0]).nombre),
      g.dir ? el("button", {class: "plano", type: "button", onclick: () => hacer(() => archivarSiempre(g), ids, "Archivado. Lo nuevo de este remitente se archivará solo en cada pasada")}, "Archivar siempre") : null,
      el("button", {class: "plano peligro", type: "button", onclick: () => hacer(() => pedirBaja(g), ids, posible ? "La baja se tramita en la próxima pasada" : "Este boletín no admite baja automática: te aparecerá en Tuyos")}, "Dar de baja")));
}
function cartaSimple(m) {
  const p = prop(m);
  const baja = p.enlace ? enlaceSeguro(p.enlace) : null;
  const outlook = m.webLink ? enlaceSeguro(m.webLink) : null;
  return el("div", {class: "carta" + (m.id === estado.sel ? " sel" : ""), onclick: alPulsarCarta(m.id)},
    remitente(m),
    titulo(m),
    p.pide ? el("p", {class: "pide"}, p.pide) : null,
    p.aviso ? el("span", {class: "aviso"}, p.aviso) : null,
    el("div", {class: "fila"},
      baja ? el("a", {class: "plano", href: baja.href, target: "_blank", rel: "noopener noreferrer"}, "Abrir la baja en " + baja.hostname) : null,
      outlook ? el("a", {class: "plano", href: outlook.href, target: "_blank", rel: "noopener noreferrer"}, "Abrir en Outlook") : null,
      estado.tab === "arch" ? el("button", {class: "plano", type: "button", onclick: () => hacer(() => pasarATuyos(m), [m.id], "Pasado a Tuyos: se queda en la bandeja", true)}, "No archivar, es mío") : null,
      el("button", {class: "plano archivo", type: "button", onclick: () => hacer(() => archivar(m), [m.id], "Archivado en " + carpetaDe(m).nombre)}, "Archivar en " + carpetaDe(m).nombre)));
}
// Limpia el texto de un correo para leerlo: fuera caracteres invisibles, líneas de relleno y huecos.
function limpiarTexto(t) {
  return (t || "")
    .replace(/[\u034f\u00ad\u200b-\u200f\u2060\ufeff]/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/<mailto:[^>\s]*>/gi, "")
    .split("\n").map((l) => l.replace(/[ \t]+/g, " ").trim())
    .filter((l) => !/^[-–—_=|*·.•\s]{1,}$/.test(l))
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}
// Convierte «[Texto]<https://…>» y «<https://…>» en enlaces cortos, en vez de mostrar direcciones kilométricas.
function nodosDeTexto(texto) {
  const nodos = [];
  const re = /(\[([^\]\n]{1,120})\]\s?)?<(https?:\/\/[^>\s]+)>|(https?:\/\/[^\s<>]{50,})/g;
  let ultimo = 0, m;
  while ((m = re.exec(texto))) {
    if (m.index > ultimo) nodos.push(document.createTextNode(texto.slice(ultimo, m.index)));
    const u = enlaceSeguro(m[3] || m[4]);
    const etiqueta = (m[2] || "").trim();
    if (u) nodos.push(el("a", {href: u.href, target: "_blank", rel: "noopener noreferrer"}, etiqueta || (m[3] ? " (enlace: " + u.hostname + ")" : u.hostname + "/…")));
    else if (etiqueta) nodos.push(document.createTextNode(etiqueta));
    ultimo = re.lastIndex;
  }
  if (ultimo < texto.length) nodos.push(document.createTextNode(texto.slice(ultimo)));
  return nodos;
}
// Como append, pero sin los huecos: append(null) escribiría la palabra «null».
function poner(nodo, ...hijos) { for (const h of hijos) if (h != null) nodo.append(h); }
function pintarDetalle() {
  const d = $("detalle");
  d.replaceChildren();
  const m = todos().find((x) => x.id === estado.sel);
  if (!m) {
    poner(d, el("div", {class: "vacio"}, el("b", {}, "Elige un correo"), "Pulsa cualquier correo de la lista para leerlo aquí."));
    return;
  }
  const p = prop(m);
  const c = estado.cuerpos[m.id];
  const outlook = m.webLink ? enlaceSeguro(m.webLink) : null;
  const esRespuesta = estado.datos.resp.includes(m);
  const esArchivar = estado.datos.arch.includes(m) || estado.datos.news.includes(m);
  poner(d,
    el("button", {class: "plano izquierda", id: "volver", type: "button", onclick: () => { cerrar(); pintar(); }}, "‹ Bandeja"),
    el("div", {}, remitente(m, "div"), el("div", {class: "asunto grande"}, m.subject || "(sin asunto)"),
      c && c.para.length ? el("div", {class: "de"}, "Para: " + c.para.slice(0, 8).join(", ")) : null),
    !esRespuesta && p.pide ? el("p", {class: "pide"}, p.pide) : null,
    p.aviso ? el("span", {class: "aviso"}, (esRespuesta ? "Revisar bien: " : "") + p.aviso) : null,
    el("div", {class: "carta"},
      el("span", {class: "eti"}, "Correo"),
      !c ? el("p", {class: "pide"}, "Cargando…")
        : c.error ? el("p", {class: "pide"}, "No se ha podido cargar el correo: " + c.error)
        : el("p", {class: "cuerpo"}, ...nodosDeTexto(c.texto || "(sin texto)")),
      ...((c && c.adjuntos) || []).map((a) => el("button", {class: "plano adjunto", type: "button", onclick: () => hacer(() => descargar(m.id, a), [], "Descargado: " + a.name)}, "Adjunto: " + a.name)),
      outlook ? el("a", {href: outlook.href, target: "_blank", rel: "noopener noreferrer"}, "Abrir en Outlook") : null));
  if (esRespuesta) {
    const area = el("textarea", {id: "borrador", "aria-label": "Tu respuesta"});
    area.value = p.respuesta || "";
    const botonEnviar = el("button", {class: "lleno", type: "button", onclick: () => hacer(() => enviar(m, area.value), [m.id], "Enviado y archivado en " + carpetaDe(m).nombre)}, "Enviar y archivar en " + carpetaDe(m).nombre);
    const comprobar = () => { botonEnviar.disabled = faltaDato(m, area.value) || !area.value.trim(); };
    area.addEventListener("input", comprobar);
    comprobar();
    poner(d,
      el("span", {class: "eti"}, "Tu respuesta" + (p.todos ? " (a todos)" : "")),
      destinatarios(p) ? el("span", {class: "de"}, destinatarios(p)) : null,
      area,
      faltaDato(m) ? el("span", {class: "pide"}, "Completa los huecos [FALTA: …] para poder enviar.") : null,
      el("div", {class: "fila"},
        el("button", {class: "plano", type: "button", onclick: () => hacer(() => descartar(m), [m.id], "Borrador descartado; el correo queda en Tuyos")}, "Descartar"),
        botonEnviar));
  } else {
    poner(d, el("div", {class: "fila"},
      esArchivar ? el("button", {class: "plano", type: "button", onclick: () => hacer(() => pasarATuyos(m), [m.id], "Pasado a Tuyos: se queda en la bandeja", true)}, "No archivar, es mío") : null,
      el("button", {class: "lleno archivo", type: "button", onclick: () => hacer(() => archivar(m), [m.id], "Archivado en " + carpetaDe(m).nombre)}, "Archivar en " + carpetaDe(m).nombre)));
  }
}
function pintar() {
  const d = estado.datos;
  const g = grupos();
  const n = {resp: d.resp.length, news: g.length, arch: d.arch.length, pend: d.pend.length};
  $("resumen").textContent = `${n.resp} ${n.resp === 1 ? "respuesta lista" : "respuestas listas"} y ${d.news.length + n.arch} para archivar`;
  $("pestanas").replaceChildren(...PESTANAS.map(([k, t]) => el("button", {type: "button", "aria-pressed": String(estado.tab === k), onclick: () => { estado.tab = k; estado.confirma = false; pintar(); }}, t, el("span", {class: "cuenta"}, String(n[k])))));
  let cartas;
  if (estado.tab === "resp") cartas = d.resp.map(cartaRespuesta);
  else if (estado.tab === "news") {
    cartas = g.map(cartaGrupo);
    if (estado.lista.length) cartas.push(el("div", {class: "carta"},
      el("span", {class: "eti"}, "Se archivan siempre"),
      ...estado.lista.map((d) => el("div", {class: "fila centro"},
        el("span", {class: "pide normal"}, d),
        el("button", {class: "plano corto", type: "button", onclick: () => hacer(() => guardarLista(estado.lista.filter((x) => x !== d)), [], "Quitado de la lista")}, "Quitar")))));
  }
  else cartas = d[estado.tab].map(cartaSimple);
  $("lista").replaceChildren(...(cartas.length ? cartas : [el("div", {class: "vacio"}, el("b", {}, "Nada pendiente aquí"), "La bandeja de entrada solo guarda lo que queda por resolver.")]));
  pintarDetalle();

  // El envío en bloque deja fuera lo marcado para revisar y lo que tiene huecos: eso se envía de uno en uno.
  const enviables = d.resp.filter((m) => { const p = prop(m); return p.draftId && !p.aviso && !faltaDato(m); });
  const paraArchivar = estado.tab === "news" ? d.news : estado.tab === "arch" ? d.arch : [];
  let texto = "", accion = null, ids = [];
  if (estado.tab === "resp" && enviables.length) {
    texto = estado.confirma ? `Confirmar: enviar ${enviables.length}` : `Enviar ${enviables.length} sin avisos y archivar`;
    accion = async () => { for (const m of enviables) await enviar(m, null); };
    ids = enviables.map((m) => m.id);
  } else if (paraArchivar.length) {
    texto = estado.confirma ? `Confirmar: archivar ${paraArchivar.length}` : `Archivar todos (${paraArchivar.length})`;
    accion = async () => { for (const m of paraArchivar) await archivar(m); };
    ids = paraArchivar.map((m) => m.id);
  }
  $("pie").classList.toggle("oculto", !accion);
  $("todo").classList.toggle("archivo", estado.tab !== "resp");
  $("todo").textContent = texto;
  $("todo").onclick = !accion ? null : () => {
    if (!estado.confirma) { estado.confirma = true; pintar(); return; }
    estado.confirma = false;
    hacer(accion, ids, estado.tab === "resp" ? "Enviado y archivado" : "Archivado, cada correo en su carpeta (Interns o Externs)", true);
  };
}

// ---------- arranque ----------
function pantallaInicio(pedirConfig) {
  const hijos = [el("h1", {}, "Agente de correo")];
  if (pedirConfig) {
    const c = el("input", {type: "text", id: "c", autocomplete: "off", spellcheck: "false"});
    const t = el("input", {type: "text", id: "t", autocomplete: "off", spellcheck: "false"});
    hijos.push(
      el("p", {}, "Primera vez en este dispositivo: pega los dos identificadores de la aplicación registrada en Microsoft 365."),
      el("label", {}, "Application (client) ID", c),
      el("label", {}, "Directory (tenant) ID", t),
      el("button", {class: "lleno", type: "button", onclick: () => {
        const cfg = {c: c.value.trim(), t: t.value.trim()};
        if (!GUID.test(cfg.c) || !GUID.test(cfg.t)) { avisar("Los dos identificadores tienen el formato 00000000-0000-0000-0000-000000000000"); return; }
        guardar("agente.cfg", cfg);
        entrar();
      }}, "Entrar con Microsoft"));
  } else {
    hijos.push(el("p", {}, "Sesión cerrada en este dispositivo."), el("button", {class: "lleno", type: "button", onclick: entrar}, "Entrar con Microsoft"));
  }
  $("app").classList.add("oculto");
  $("inicio").replaceChildren(...hijos);
  $("inicio").classList.remove("oculto");
}
function salir() {
  borrar("agente.tok");
  guardar("agente.salio", true);
  estado.datos = {resp: [], news: [], arch: [], pend: []};
  cerrar();
  pantallaInicio(false);
}
async function arrancar() {
  const q = new URLSearchParams(location.search);
  if (location.search) history.replaceState(null, "", REDIR);
  if (!config()) { borrar("agente.cfg"); borrar("agente.tok"); pantallaInicio(true); return; }
  if (q.get("error")) { pantallaInicio(false); avisar("Microsoft ha rechazado el acceso (" + String(q.get("error")).slice(0, 60) + ")"); return; }
  if (q.get("code")) {
    let pk = null;
    try { pk = JSON.parse(sessionStorage.getItem("agente.pkce")); } catch (e) {}
    sessionStorage.removeItem("agente.pkce");
    if (!pk || !pk.st || pk.st !== q.get("state")) { pantallaInicio(false); avisar("La sesión no coincide. Vuelve a entrar."); return; }
    try { await pedirToken({grant_type: "authorization_code", code: q.get("code"), code_verifier: pk.verificador}); }
    catch (e) { pantallaInicio(false); avisar("No se ha podido iniciar sesión (" + e.message.slice(0, 60) + ")"); return; }
  }
  if (leer("agente.salio")) { pantallaInicio(false); return; }
  $("app").classList.remove("oculto");
  $("recargar").onclick = cargar;
  $("salir").onclick = salir;
  pintar();
  await cargar();
  setInterval(() => { if (!document.hidden && !estado.sel && !leer("agente.salio")) cargar(); }, 120000);
}
arrancar();
