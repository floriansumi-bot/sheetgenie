/* SheetGenie service worker
 * - Cache-first for the app shell (works offline / installable).
 * - NETWORK-ONLY for any /api/ request: API responses are never cached.
 * Bump CACHE_VERSION whenever a shell asset changes.
 */
const CACHE_VERSION = 'sheetgenie-v5';

const SHELL = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './manifest.webmanifest',
  './icons/icon.svg',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon-180.png',
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_VERSION).then((cache) => cache.addAll(SHELL)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_VERSION).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;

  // Only handle GET; let everything else (POST to /api/*) hit the network untouched.
  if (req.method !== 'GET') return;

  const url = new URL(req.url);

  // NETWORK-ONLY for the API. Never read from or write to cache.
  if (url.pathname.startsWith('/api/')) {
    event.respondWith(fetch(req));
    return;
  }

  // Cache-first for the shell; fall back to network, then to cached index for navigations.
  event.respondWith(
    caches.match(req).then((cached) => {
      if (cached) return cached;
      return fetch(req).catch(() => {
        if (req.mode === 'navigate') return caches.match('./index.html');
        return Response.error();
      });
    })
  );
});
