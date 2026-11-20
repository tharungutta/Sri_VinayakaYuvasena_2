const CACHE_NAME = 'svys-cache-v6';
const STATIC_ASSETS = [
  '/',
  '/index.html',
  '/styles.css',
  '/auth.css',
  '/dashboard.js',
  '/auth.js',
  '/config.js',
  '/live-data.js',
  '/manifest.json',
  '/assets/logo.png',
  '/assets/upi-qr.png',
  '/assets/qrcode.min.js'
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

  // Handle navigation requests (when user opens the PWA from home screen or browser navigation)
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then(networkRes => {
          // If server returns a valid page (status 200-299), cache and return it
          if (networkRes && networkRes.ok) {
            const resClone = networkRes.clone();
            caches.open(CACHE_NAME).then(cache => cache.put(req, resClone));
            return networkRes;
          }
          // If server returns 404, 500 or redirect issue, fall back to cached index.html immediately!
          return caches.match('/').then(cached => cached || caches.match('/index.html') || networkRes);
        })
        .catch(() => {
          // Offline fallback
          return caches.match('/').then(cached => cached || caches.match('/index.html'));
        })
    );
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
      .catch(() => caches.match(req))
  );
});

// Push & Notification Click Listener (WhatsApp-style notification handling)
self.addEventListener('notificationclick', event => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clientList => {
      for (const client of clientList) {
        if (client.url && 'focus' in client) {
          client.postMessage({ action: 'openChat' });
          return client.focus();
        }
      }
      if (clients.openWindow) {
        return clients.openWindow('/?action=openChat');
      }
    })
  );
});
