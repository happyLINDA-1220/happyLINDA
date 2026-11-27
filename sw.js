// 金灶鍋貼 PWA Service Worker
// 更新網站內容後，把版本號加 1，舊快取就會被清掉。
const VERSION = 'jinzao-v1';
const APP_SHELL = [
  './',
  './index.html',
  './manifest.webmanifest',
  './images/hero-poster.jpg',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/maskable-512.png',
  './icons/apple-touch-icon.png',
  './icons/favicon-32.png'
];

self.addEventListener('install', (event) => {
  event.waitUntil(caches.open(VERSION).then((cache) => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (event) => {
  const req = event.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);

  // 頁面本身：先抓最新版，離線時用快取
  if (req.mode === 'navigate') {
    event.respondWith(
      fetch(req)
        .then((res) => {
          const copy = res.clone();
          caches.open(VERSION).then((c) => c.put('./index.html', copy));
          return res;
        })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  // 其他資源（圖片、圖示、CDN 上的 Tailwind／Alpine／three.js／字型）：
  // 先回快取讓畫面秒開，同時在背景更新
  const cacheable = url.origin === self.location.origin ||
    /(cdn\.tailwindcss\.com|cdn\.jsdelivr\.net|cdnjs\.cloudflare\.com|fonts\.googleapis\.com|fonts\.gstatic\.com|i\.postimg\.cc)$/.test(url.hostname);
  if (!cacheable) return;

  event.respondWith(
    caches.open(VERSION).then((cache) =>
      cache.match(req).then((cached) => {
        const network = fetch(req)
          .then((res) => {
            if (res && (res.ok || res.type === 'opaque')) cache.put(req, res.clone());
            return res;
          })
          .catch(() => cached);
        return cached || network;
      })
    )
  );
});
