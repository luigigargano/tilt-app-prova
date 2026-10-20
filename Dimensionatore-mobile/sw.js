/* Service worker - Dimensionatore mobile - versione 2026.10.18-REL
   Strategia: navigazione network-first (sempre l'ultima versione quando c'e' rete, copia in cache se offline);
   file statici cache-first con versione nel nome della cache. Nessun dato utente in cache. */
const CACHE = 'stima-fv-dim-2026.10.18-REL';
const STATICI = ["./", "index.html", "dim-core.js", "comuni_dim.json", "manifest.webmanifest", "icons/apple-touch-icon-180.png", "icons/icon-192.png", "icons/icon-512.png"];
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(STATICI)));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k.startsWith('stima-fv-dim-') && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener('message', e => { if (e.data === 'AGGIORNA') self.skipWaiting(); });
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  if (req.mode === 'navigate') {
    e.respondWith(fetch(req).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put('./', cp)); return r; })
      .catch(() => caches.match('./').then(r => r || caches.match('./index.html'))));
    return;
  }
  if (req.url.endsWith('listino_app.json')) {   // listino: sempre dalla rete se possibile
    e.respondWith(fetch(req).then(r => { const cp = r.clone(); caches.open(CACHE).then(c => c.put(req, cp)); return r; })
      .catch(() => caches.match(req)));
    return;
  }
  e.respondWith(caches.match(req).then(r => r || fetch(req)));
});
