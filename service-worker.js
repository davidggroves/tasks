const CACHE_NAME = 'tasks-app-2026.08.30.2';

const SHELL_FILES = [
  './',
  './index.html',
  './app.js',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', event => {
  event.waitUntil(
    caches
      .open(CACHE_NAME)
      .then(cache =>
        cache.addAll(SHELL_FILES)
      )
  );

  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches
      .keys()
      .then(names =>
        Promise.all(
          names
            .filter(
              name =>
                name !== CACHE_NAME
            )
            .map(
              name =>
                caches.delete(name)
            )
        )
      )
  );

  self.clients.claim();
});

self.addEventListener('fetch', event => {
  const url =
    new URL(event.request.url);

  if (
    url.origin !==
    self.location.origin
  ) {
    return;
  }

  if (
    event.request.method !== 'GET'
  ) {
    return;
  }

  const isShellRequest =
    SHELL_FILES.some(file => {
      const path =
        file.replace('./', '');

      return (
        path !== '' &&
        url.pathname.endsWith(path)
      );
    }) ||
    url.pathname.endsWith('/');

  if (!isShellRequest) return;

  event.respondWith(
    caches
      .match(event.request)
      .then(cached => {
        if (cached) {
          return cached;
        }

        return fetch(event.request)
          .then(response => {
            const clone =
              response.clone();

            caches
              .open(CACHE_NAME)
              .then(cache =>
                cache.put(
                  event.request,
                  clone
                )
              );

            return response;
          });
      })
  );
});
