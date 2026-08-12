// 캐시 버전 — 변경 시 구 캐시 자동 삭제
const CACHE = 'stocktracker-v2';

self.addEventListener('install', e => {
  // 설치 즉시 활성화 (대기 없음)
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  // 이전 버전 캐시 모두 삭제
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  const url = new URL(e.request.url);

  // API 요청 → 캐시 안 함 (브라우저 기본 fetch)
  if (url.pathname.startsWith('/api/')) return;

  // 메인 페이지(/, /index.html) → 캐시 안 함, 항상 서버에서 가져옴
  if (url.pathname === '/' || url.pathname === '/index.html') return;

  // 나머지 정적 자원(sw.js, 아이콘 등) → 네트워크 우선, 실패 시 캐시
  e.respondWith(
    fetch(e.request).catch(() => caches.match(e.request))
  );
});
