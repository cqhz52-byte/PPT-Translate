const CACHE_NAME = "ppt-translator-pwa-v112";
const CORE_ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./styles.css?v=112",
  "./app.js",
  "./app.js?v=112",
  "./version.json",
  "./assets/curaway-logo.png",
  "./manifest.webmanifest",
  "./icons/icon.svg",
  "./icons/icon-192.png",
  "./icons/icon-512.png",
];

const OPTIONAL_ASSETS = [
  "./vendor/jszip.min.js",
  "./vendor/pdf-lib.esm.min.js",
  "./vendor/fontkit.bundle.js",
  "./vendor/pako.esm.js",
  "./vendor/fonts/NotoSansCJKsc-Regular.otf",
];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(CORE_ASSETS))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))
    )
  );
  self.clients.claim();
});

self.addEventListener("message", (event) => {
  if (event.data?.type === "SKIP_WAITING") {
    self.skipWaiting();
    return;
  }

  if (event.data?.type === "WARM_OPTIONAL_CACHE") {
    event.waitUntil(warmOptionalCache());
  }
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;

  if (shouldUseNetworkFirst(event.request)) {
    event.respondWith(
      fetch(event.request)
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          return response;
        })
        .catch(() => caches.match(event.request))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => {
      if (cached) return cached;
      return fetch(event.request).then((response) => {
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
        return response;
      });
    })
  );
});

function shouldUseNetworkFirst(request) {
  const url = new URL(request.url);
  return url.pathname.endsWith("/version.json") || request.mode === "navigate" || ["document", "script", "style"].includes(request.destination);
}

async function warmOptionalCache() {
  const cache = await caches.open(CACHE_NAME);
  await Promise.allSettled(
    OPTIONAL_ASSETS.map(async (asset) => {
      const cached = await cache.match(asset);
      if (cached) return;
      const response = await fetch(asset, { cache: "no-cache" });
      if (response.ok) await cache.put(asset, response);
    })
  );
}
