const CACHE_NAME = 'tasks-app-2026.07.29.2';
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

  // Never touch cross-origin requests (Sheets API, Google auth) — only same-origin
  // app-shell files are candidates for caching here.
  if (url.origin !== self.location.origin) return;

  // Cache API only supports GET — writes (PUT/POST to Sheets) must always pass through untouched.
  if (event.request.method !== 'GET') return;

  const isShellRequest = SHELL_FILES.some((f) => {
    const path = f.replace('./', '');
    return path !== '' && url.pathname.endsWith(path);
  }) || url.pathname.endsWith('/');

  if (!isShellRequest) return;

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
