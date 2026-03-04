// RAJ Financial Services — Service Worker
const CACHE = 'raj-finance-v1';
const ASSETS = [
  '/',
  '/static/icon-192.png',
  '/static/icon-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js',
  'https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800;900&family=Barlow:wght@400;500;600;700&family=JetBrains+Mono:wght@300;400;600;700&display=swap'
];

// Install — cache assets
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE).then(cache => {
      return cache.addAll(ASSETS).catch(() => {});
    })
  );
  self.skipWaiting();
});

// Activate — clean old caches
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Fetch — network first, cache fallback
self.addEventListener('fetch', e => {
  // Skip API calls — always network
  if(e.request.url.includes('/analyze') ||
     e.request.url.includes('/compare') ||
     e.request.url.includes('/news') ||
     e.request.url.includes('/export') ||
     e.request.url.includes('/ai')){
    return;
  }
  e.respondWith(
    fetch(e.request)
      .then(res => {
        const clone = res.clone();
        caches.open(CACHE).then(c => c.put(e.request, clone));
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});

// Push notifications (for price alerts)
self.addEventListener('push', e => {
  const data = e.data ? e.data.json() : {};
  e.waitUntil(
    self.registration.showNotification(data.title || '🔔 RAJ Finance Alert', {
      body: data.body || 'Price alert triggered!',
      icon: '/static/icon-192.png',
      badge: '/static/icon-192.png',
      vibrate: [200, 100, 200],
      tag: 'price-alert'
    })
  );
});
