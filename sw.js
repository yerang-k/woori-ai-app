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

// 가족 푸시 알림 수신(앱이 꺼져 있어도 동작)
self.addEventListener('push', (e) => {
  let data = {};
  try { data = e.data ? e.data.json() : {}; } catch (err) {}
  const title = data.title || '우리 아이';
  const body = data.body || '새 소식이 있어요.';
  e.waitUntil(self.registration.showNotification(title, {
    body,
    icon: '/icons/icon-192.png',
    badge: '/icons/icon-192.png',
    tag: 'wooriai-push'
  }));
});
self.addEventListener('notificationclick', (e) => {
  e.notification.close();
  e.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then((list) => {
      for (const c of list) { if ('focus' in c) return c.focus(); }
      if (self.clients.openWindow) return self.clients.openWindow('/');
    })
  );
});
