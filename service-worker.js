/* Stores KPI & Appraisal — Service Worker
   Purpose: installability + app-shell caching ONLY.
   Rule: live business data (Login, Dashboard, Working Register, Target,
   Approval, HR Appraisal, Attendance, Reports, Master Data, Workflow
   Details) must never be served from cache — every such request goes to
   the Apps Script backend (script.google.com / script.googleusercontent.com)
   live, every time. */

const CACHE_NAME = 'stores-kpi-pwa-v1';

const APP_SHELL = [
  './',
  './index.html',
  './manifest.json'
];

// Any request to the Apps Script backend (the /exec API and its redirect
// targets) is treated as live data and is NEVER cached or served from cache.
function isBackendRequest(url) {
  return url.hostname === 'script.google.com' ||
         url.hostname === 'script.googleusercontent.com';
}

self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then(cache => cache.addAll(APP_SHELL))
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(key => key !== CACHE_NAME)
          .map(key => caches.delete(key))
      )
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);

  // Live backend/API calls: always network, never touch the cache.
  if (isBackendRequest(url) || req.method !== 'GET') {
    event.respondWith(fetch(req));
    return;
  }

  // Everything else (app shell, CDN libs, fonts): network-first,
  // falling back to cache only when offline.
  event.respondWith(
    fetch(req)
      .then(response => {
        if (response && response.status === 200) {
          const copy = response.clone();
          caches.open(CACHE_NAME).then(cache => cache.put(req, copy));
        }
        return response;
      })
      .catch(() => caches.match(req))
  );
});
