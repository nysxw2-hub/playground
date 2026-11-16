// 金月軒 PWA service worker
// 頁面：先連網、失敗時用快取（離線也能開）；圖片與 CDN 程式庫：先用快取、背景更新。
const VERSION = 'v1';
const SHELL = `gm-shell-${VERSION}`;
const RUNTIME = `gm-runtime-${VERSION}`;
const SHELL_FILES = [
  './',
  './index.html',
  './manifest.webmanifest',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(SHELL).then((c) => c.addAll(SHELL_FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== SHELL && k !== RUNTIME).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

async function networkFirst(request) {
  try {
    const fresh = await fetch(request);
    const cache = await caches.open(SHELL);
    cache.put('./index.html', fresh.clone());
    return fresh;
  } catch (e) {
    return (await caches.match(request)) || (await caches.match('./index.html'));
  }
}

async function staleWhileRevalidate(event) {
  const cache = await caches.open(RUNTIME);
  const cached = await cache.match(event.request);
  const update = fetch(event.request)
    .then((res) => {
      if (res && (res.ok || res.type === 'opaque')) cache.put(event.request, res.clone());
      return res;
    })
    .catch(() => cached);
  if (cached) {
    event.waitUntil(update);
    return cached;
  }
  return update;
}

self.addEventListener('fetch', (event) => {
  const { request } = event;
  if (request.method !== 'GET') return;
  const url = new URL(request.url);
  if (request.mode === 'navigate') {
    event.respondWith(networkFirst(request));
    return;
  }
  const sameOrigin = url.origin === self.location.origin;
  const cacheable =
    sameOrigin ||
    request.destination === 'image' ||
    request.destination === 'script' ||
    request.destination === 'style' ||
    request.destination === 'font';
  if (!cacheable) return;
  event.respondWith(staleWhileRevalidate(event));
});
