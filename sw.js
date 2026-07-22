// sw.js - Versi Ultra-Hybrid v55 (Professional UI Refresh)
// Dibuat untuk: E-HADIR PWA
// Strategi: Network First (Data Terkini) -> Fallback Cache (Offline)

const CACHE_NAME = 'ehadir-hybrid-v55-ui-refresh';
const URLS_TO_CACHE = [
  './',
  './index.html',
  './manifest.json',
  './logo.png'
];

// 1. INSTALL: Simpan 'kulit' aplikasi (fail asas) ke dalam telefon
self.addEventListener('install', (event) => {
  self.skipWaiting(); // Paksa SW baru ambil alih segera tanpa lengah
  event.waitUntil(
    caches.open(CACHE_NAME)
      .then((cache) => {
        console.log('Opened cache v54-integrity');
        return cache.addAll(URLS_TO_CACHE);
      })
  );
});

// 2. ACTIVATE: Buang cache versi lama sepenuhnya (bersihkan memori telefon)
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((cacheNames) => {
      return Promise.all(
        cacheNames.map((cacheName) => {
          if (cacheName !== CACHE_NAME) {
            console.log('Memadam cache memori lama secara mandatori:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    })
  );
  return self.clients.claim();
});

// 3. FETCH: Strategi menapis akses ke internet atau cache
self.addEventListener('fetch', (event) => {
  const request = event.request;
  const url = new URL(request.url);

  // A. JIKA HANTAR DATA (POST) - Cth: Tekan Butang Masuk/Keluar
  if (request.method === 'POST') {
    event.respondWith(
      fetch(request.clone())
        .catch(() => {
          // Jika internet tiada, pulangkan isyarat 'Offline'
          return new Response(JSON.stringify({ offline: true }), {
            headers: { 'Content-Type': 'application/json' }
          });
        })
    );
    return;
  }

  // B. JIKA MINTA DATA LIST DARI GOOGLE APPS SCRIPT (GET List)
  if (url.hostname.includes('script.google.com') || url.hostname.includes('googleusercontent.com')) {
    event.respondWith(
      fetch(request)
        .then((response) => {
          const resClone = response.clone();
          caches.open(CACHE_NAME).then((cache) => {
            cache.put(request, resClone);
          });
          return response;
        })
        .catch(() => {
          // Fallback offline cache jika tiada internet
          return caches.match(request).then((cachedRes) => {
            if (cachedRes) return cachedRes;
            return new Response(JSON.stringify([]), {
              headers: { 'Content-Type': 'application/json' }
            });
          });
        })
    );
    return;
  }

  // C. JIKA MINTA FAIL HTML UTAMA (Network First + Paksa bypass Cache HTTP Browser)
  if (request.mode === 'navigate' || request.url.includes('.html')) {
    event.respondWith(
      // fetch dengan { cache: 'no-store' } memaksa enjin mendapatkan fail terus dari pelayan hosting, mengabaikan sekatan cache lama
      fetch(request, { cache: 'no-store' }).then((networkResponse) => {
        return caches.open(CACHE_NAME).then((cache) => {
          cache.put(request, networkResponse.clone());
          return networkResponse;
        });
      }).catch(() => {
        return caches.match('./index.html'); // Mod offline fallback
      })
    );
    return;
  }

  // D. JIKA MINTA FAIL BIASA LAIN (Gambar, Logo dll)
  event.respondWith(
    caches.match(request).then((cachedResponse) => {
      return cachedResponse || fetch(request, { cache: 'no-store' }).then((networkResponse) => {
         return caches.open(CACHE_NAME).then((cache) => {
            cache.put(request, networkResponse.clone());
            return networkResponse;
         });
      });
    }).catch(() => {})
  );
});
