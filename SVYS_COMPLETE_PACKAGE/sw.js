const CACHE_NAME = 'svys-cache-v1';
const STATIC_ASSETS = [
  './',
  './index.html',
  './styles.css',
  './auth.css',
  './dashboard.js',
  './auth.js',
  './config.js',
  './live-data.js',
  './manifest.json',
  './assets/logo.png',
  './assets/upi-qr.png',
  './assets/qrcode.min.js'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      return cache.addAll(STATIC_ASSETS).catch(err => {
        console.warn('SVYS ServiceWorker precache partial note:', err);
      });
    })
  );
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys => {
      return Promise.all(
        keys.map(k => {
          if (k !== CACHE_NAME) return caches.delete(k);
        })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);

  // Network-only for live Apps Script and Firebase APIs
  if (url.origin.includes('google') || url.origin.includes('firebase') || req.method !== 'GET') {
    return;
  }

  // Network-first for local assets with fallback to cache
  event.respondWith(
    fetch(req)
      .then(networkRes => {
        if (networkRes && networkRes.status === 200) {
          const resClone = networkRes.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(req, resClone));
        }
        return networkRes;
      })
      .catch(() => caches.match(req).then(cached => cached || caches.match('./index.html')))
  );
});
