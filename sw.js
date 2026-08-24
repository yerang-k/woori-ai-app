// 우리 아이 앱 서비스워커 — 네트워크 우선(항상 최신), 오프라인 시 캐시 폴백
const CACHE = 'woori-cache-v1';

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (e) => e.waitUntil(self.clients.claim()));

self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET') return;                     // POST(AI·동기화)는 통과
  const url = new URL(req.url);
  if (url.origin !== self.location.origin) return;      // 외부(폰트 등)는 그대로
  e.respondWith(
    fetch(req)
      .then((res) => { const c = res.clone(); caches.open(CACHE).then((ca) => ca.put(req, c)); return res; })
      .catch(() => caches.match(req))                   // 오프라인이면 캐시
  );
});
