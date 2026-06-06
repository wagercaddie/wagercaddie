// WagerCaddie Service Worker
// Caches app shell for fast loads. Data always fetched live from GAS.

const CACHE_NAME = 'wagercaddie-v1';
const APP_SHELL = [
  '/',
  '/index.html',
  'https://fonts.googleapis.com/css2?family=Bebas+Neue&family=Montserrat:wght@400;600;700;800&display=swap'
];

// Install — cache app shell
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(APP_SHELL);
    })
  );
  self.skipWaiting();
});

// Activate — clean up old caches
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch strategy:
// - GAS API calls (script.google.com): always network, never cache
// - Everything else: cache first, then network
self.addEventListener('fetch', event => {
  const url = new URL(event.request.url);

  // Never cache GAS API calls — always live data
  if (url.hostname.includes('script.google.com') ||
      url.hostname.includes('groupme.com') ||
      url.hostname.includes('golfgenius.com')) {
    event.respondWith(fetch(event.request));
    return;
  }

  // App shell: cache first
  event.respondWith(
    caches.match(event.request).then(cached => {
      return cached || fetch(event.request).then(response => {
        // Cache successful GET responses
        if (event.request.method === 'GET' && response.status === 200) {
          const clone = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(event.request, clone));
        }
        return response;
      });
    }).catch(() => {
      // Offline fallback — return cached index.html
      return caches.match('/index.html');
    })
  );
});
