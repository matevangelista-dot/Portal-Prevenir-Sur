/* Service worker: permite abrir la app sin conexión en el lugar de medición.
   Cambiá VERSION en cada despliegue para forzar la actualización. */
const VERSION = "v3";
const CACHE = "iluminacion-" + VERSION;
const SHELL = ["/", "/index.html", "/styles.css", "/app.js", "/estudios.js", "/estudios2.js", "/config.js", "/manifest.webmanifest",
  "/supabase.js", "/591.supabase.js", "/xlsx-js-style.js", "/icon-192.png", "/icon-512.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return; // Supabase y otros: siempre red
  const isLib = /^\/(supabase\.js|591\.supabase\.js|xlsx-js-style\.js|icon-\d+\.png)$/.test(url.pathname);
  if (isLib) {
    e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
    return;
  }
  // App: primero la red (para recibir actualizaciones), si no hay conexión usa la copia guardada
  e.respondWith(
    fetch(e.request).then(res => {
      if (res.ok) { const copy = res.clone(); caches.open(CACHE).then(c => c.put(e.request, copy)); }
      return res;
    }).catch(() => caches.match(e.request).then(r => r || caches.match("/index.html")))
  );
});
