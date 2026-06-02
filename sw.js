// ============================================================
// Service Worker — Palma Grande Cosecha
// Permite usar la app sin conexión a internet
// ============================================================

const CACHE_NAME = 'cosecha-v1';
const ASSETS = [
  './',
  './index.html',
  'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css',
];

// Instalar: cachear recursos esenciales
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME).then(function(cache) {
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

// Activar: limpiar caches viejos
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(keys) {
      return Promise.all(
        keys.filter(function(k) { return k !== CACHE_NAME; })
            .map(function(k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

// Fetch: cache-first para assets, network-first para el GAS
self.addEventListener('fetch', function(e) {
  var url = e.request.url;

  // Peticiones al GAS siempre van a la red (no cachear datos)
  if (url.includes('script.google.com')) {
    e.respondWith(
      fetch(e.request).catch(function() {
        return new Response(JSON.stringify({ ok: false, error: 'Sin conexión' }),
          { headers: { 'Content-Type': 'application/json' } });
      })
    );
    return;
  }

  // Para todo lo demás: intentar red primero, caer a caché si falla
  e.respondWith(
    fetch(e.request)
      .then(function(response) {
        // Actualizar caché con respuesta fresca
        var clone = response.clone();
        caches.open(CACHE_NAME).then(function(cache) {
          cache.put(e.request, clone);
        });
        return response;
      })
      .catch(function() {
        return caches.match(e.request);
      })
  );
});
