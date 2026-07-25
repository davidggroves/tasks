const CACHE_NAME = 'tasks-app-v3';
const SHELL_FILES = [
  './',
  './index.html',
  './app.js',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

// App shell: cache-first (so the UI itself always loads offline).
// Anything else (Sheets/Drive API calls, Google auth) goes straight to network —
// app.js handles its own offline data caching via localStorage, not this worker.
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  const isShellRequest = SHELL_FILES.some((f) => url.pathname.endsWith(f.replace('./', '')));

  if (!isShellRequest) return; // let network calls (Google APIs) pass through untouched

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((response) => {
        const clone = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, clone));
        return response;
      });
    })
  );
});
