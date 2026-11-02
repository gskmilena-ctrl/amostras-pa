/* Service worker: guarda a "casca" do app para abrir rápido e funcionar como aplicativo instalado.
   Os dados vêm sempre do Firebase (não passam por aqui). */
const CACHE = "amostras-pa-v1";
const SHELL = ["./", "index.html", "config.js", "manifest.webmanifest", "icones/icone-192.png", "icones/icone-512.png", "icones/apple-touch-icon.png"];
self.addEventListener("install", e => { e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting())); });
self.addEventListener("activate", e => { e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim())); });
self.addEventListener("fetch", e => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  // Rede primeiro (sempre a versão mais nova); sem internet, usa a cópia guardada.
  e.respondWith(fetch(e.request).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put(e.request, cp)); return r; })
    .catch(() => caches.match(e.request).then(r => r || caches.match("index.html"))));
});
