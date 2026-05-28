// Service worker: makes the site work offline after the first visit and
// removes the CDN round-trip on repeat visits.
//
// Strategy:
//   - Our own files (html/js/py): stale-while-revalidate — serve the cached
//     copy instantly, then fetch a fresh copy in the background so updates
//     reach the user on their next visit (avoids the "stuck on old version" trap).
//   - Pyodide CDN files: cache-first — they live at versioned, immutable URLs,
//     so once saved they never need re-fetching.
//
// Bump CACHE_VERSION whenever you want to force every saved copy to refresh.

const CACHE_VERSION = "v2";
const APP_CACHE = `expense-app-${CACHE_VERSION}`;
const RUNTIME_CACHE = `expense-runtime-${CACHE_VERSION}`;

const APP_ASSETS = [
  "./",
  "index.html",
  "main.js",
  "clean_data.py",
  "build_excel.py",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(APP_CACHE).then((cache) => cache.addAll(APP_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  const keep = new Set([APP_CACHE, RUNTIME_CACHE]);
  event.waitUntil(
    caches
      .keys()
      .then((keys) =>
        Promise.all(keys.filter((k) => !keep.has(k)).map((k) => caches.delete(k)))
      )
      .then(() => self.clients.claim())
  );
});

async function staleWhileRevalidate(request, cacheName) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(request);
  const network = fetch(request)
    .then((resp) => {
      if (resp && resp.ok) cache.put(request, resp.clone());
      return resp;
    })
    .catch(() => null);
  return cached || (await network) || Response.error();
}

async function cacheFirst(request, cacheName) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(request);
  if (cached) return cached;
  try {
    const resp = await fetch(request);
    if (resp && (resp.ok || resp.type === "opaque")) cache.put(request, resp.clone());
    return resp;
  } catch (e) {
    return Response.error();
  }
}

self.addEventListener("fetch", (event) => {
  const { request } = event;
  if (request.method !== "GET") return;
  const url = new URL(request.url);

  if (url.origin === self.location.origin) {
    event.respondWith(staleWhileRevalidate(request, APP_CACHE));
  } else if (url.hostname === "cdn.jsdelivr.net") {
    event.respondWith(cacheFirst(request, RUNTIME_CACHE));
  }
});
