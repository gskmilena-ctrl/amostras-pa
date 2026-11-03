/* Service worker: guarda a "casca" do app para abrir rápido e funcionar como aplicativo instalado.
   Os dados vêm sempre do Firebase (não passam por aqui).
   Sempre busca a versão mais nova na internet; a cópia guardada só é usada sem conexão. */
const CACHE = "amostras-pa-v3";
const SHELL = ["./", "index.html", "config.js", "manifest.webmanifest", "icones/icone-192.png", "icones/icone-512.png", "icones/apple-touch-icon.png", "icones/logo-meripobo.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL.map(u => new Request(u, { cache: "reload" })))).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  // "no-cache": pergunta sempre ao GitHub se há versão nova (ignora a cópia de 10 minutos do navegador).
  e.respondWith(fetch(e.request, { cache: "no-cache" }).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); return r; })
    .catch(() => caches.match(e.request).then(r => r || caches.match("index.html"))));
});
