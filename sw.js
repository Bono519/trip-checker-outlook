/* sw.js — Trip Conflict Checker Outlook 版 Service Worker */
const CACHE_NAME = 'trip-checker-outlook-v1';
const ASSETS = [
  '/trip-checker-outlook/',
  '/trip-checker-outlook/index.html',
  '/trip-checker-outlook/style.css',
  '/trip-checker-outlook/app.js',
  '/trip-checker-outlook/manifest.json',
];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  // Microsoft 相關請求不走 cache
  if (e.request.url.includes('microsoft') ||
      e.request.url.includes('microsoftonline') ||
      e.request.url.includes('graph.microsoft') ||
      e.request.url.includes('msauth')) return;
  e.respondWith(
    caches.match(e.request).then(r => r || fetch(e.request))
  );
});
