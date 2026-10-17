// Bump ASSET_VERSION when publishing CSS/JS changes. Bump IMAGE_VERSION only
// when replacing existing images at the same URL, so code releases retain photos.
const ASSET_VERSION = "20261017-1";
const IMAGE_VERSION = "1";
const CACHE_PREFIX = "xlance-public-";
const ASSET_CACHE = `${CACHE_PREFIX}assets-${ASSET_VERSION}`;
const IMAGE_CACHE = `${CACHE_PREFIX}images-${IMAGE_VERSION}`;

function cacheNameFor(request) {
  if (request.method !== "GET") return null;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return null;
  if (/^\/assets\/(?:img|img-webp)\/.+\.(?:avif|gif|ico|jpe?g|png|svg|webp)$/i.test(url.pathname)) {
    return IMAGE_CACHE;
  }
  if (/^\/(?:css\/.+\.css|js\/.+\.js)$/.test(url.pathname)) return ASSET_CACHE;
  return null;
}

self.addEventListener("install", event => {
  event.waitUntil(self.skipWaiting());
});

self.addEventListener("activate", event => {
  event.waitUntil((async () => {
    const names = await caches.keys();
    await Promise.all(names
      .filter(name => name.startsWith(CACHE_PREFIX) && name !== ASSET_CACHE && name !== IMAGE_CACHE)
      .map(name => caches.delete(name)));
    await self.clients.claim();
  })());
});

async function cachedResponse(request, cacheName, event) {
  let cache;
  try {
    cache = await caches.open(cacheName);
    const cached = await cache.match(request);
    if (cached) return cached;
  } catch (_error) {
    // Storage restrictions must not prevent a normal network load.
  }

  // Revalidate a miss so a new image version cannot permanently store an old
  // response from the browser's separate HTTP cache. Hits never use the network.
  const response = await fetch(request.cache === "force-cache"
    ? request : new Request(request, { cache: "no-cache" }));
  const contentType = response.headers.get("content-type") || "";
  const expectedType = cacheName === IMAGE_CACHE
    ? /^image\//i.test(contentType)
    : /(?:text\/css|(?:text|application)\/(?:javascript|x-javascript|ecmascript))/i.test(contentType);
  // Redirects, errors and HTML fallbacks must never become cached static files.
  if (cache && response.status === 200 && !response.redirected && expectedType &&
      !/\b(?:no-store|private)\b/i.test(response.headers.get("cache-control") || "")) {
    const stored = cache.put(request, response.clone()).catch(() => {});
    event.waitUntil(stored);
  }
  return response;
}

self.addEventListener("fetch", event => {
  const cacheName = cacheNameFor(event.request);
  if (cacheName) event.respondWith(cachedResponse(event.request, cacheName, event));
});

// CSS and scripts on the first visit loaded before the worker took control.
// Seed just those files from the HTTP cache; photos load through the worker.
self.addEventListener("message", event => {
  if (event.data?.type !== "CACHE_PAGE_ASSETS" || !Array.isArray(event.data.urls)) return;
  event.waitUntil(Promise.all(event.data.urls.map(async url => {
    try {
      const request = new Request(url, { cache: "force-cache", credentials: "same-origin" });
      if (cacheNameFor(request) === ASSET_CACHE) {
        await cachedResponse(request, ASSET_CACHE, event);
      }
    } catch (_error) {
      // Warming the cache is optional and must not affect the page.
    }
  })));
});
