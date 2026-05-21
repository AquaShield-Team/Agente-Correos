// Service Worker — AquaShield Agente de Correos
const CACHE_NAME = 'aquashield-correos-v1';
const STATIC_ASSETS = [
    '/',
    '/static/styles.css',
    '/static/app.js',
    '/static/logo_aquachile_dark.png',
    '/static/icon-192.png',
    '/static/favicon.ico'
];

// Instalar: cachear recursos estáticos
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS))
    );
    self.skipWaiting();
});

// Activar: limpiar caches antiguos
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((keys) => 
            Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
        )
    );
    self.clients.claim();
});

// Fetch: network-first para API, cache-first para estáticos
self.addEventListener('fetch', (event) => {
    if (event.request.url.includes('/api/')) {
        // API: siempre ir a la red
        event.respondWith(fetch(event.request));
    } else {
        // Estáticos: cache primero, luego red
        event.respondWith(
            caches.match(event.request).then((cached) => cached || fetch(event.request))
        );
    }
});
