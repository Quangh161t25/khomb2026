// ─────────────────────────────────────────────────────────────
// UPMISA Service Worker — auto-update + offline cache
// Tăng APP_VERSION mỗi khi deploy để trigger auto-reload
// ─────────────────────────────────────────────────────────────
const APP_VERSION = 'upmisa-v4';
const CACHE_NAME = APP_VERSION;

const STATIC_ASSETS = [
    'index.html',
    'login.html',
    'styles.css',
    'icon.png',
    'manifest.json',
    'js/config.js',
    'js/state.js',
    'js/utils.js',
    'js/api.js',
    'js/auth.js',
    'js/main.js',
    'js/shared/sheet-ranges.js',
    'js/modules/donhang.js',
    'js/modules/bandon.js',
    'js/modules/sanpham.js',
    'js/modules/upmisa.js',
    'js/modules/baocao.js',
    'js/modules/bchh.js',
    'js/modules/dhct.js',
    'js/modules/dhct_form.js',
    'js/modules/unique_dh_ct.js',
    'js/modules/inventory.js',
    'js/modules/hanghoan.js',
    'js/modules/hh_shop_dien.js',
];

// ── INSTALL: cache toàn bộ app shell ──────────────────────────
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => cache.addAll(STATIC_ASSETS))
    );
    // Kích hoạt SW mới ngay, không chờ tab cũ đóng
    self.skipWaiting();
});

// ── ACTIVATE: xoá cache cũ + thông báo client reload ─────────
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((names) =>
            Promise.all(
                names
                    .filter((name) => name !== CACHE_NAME)
                    .map((name) => caches.delete(name))
            )
        ).then(() => {
            // Chiếm quyền kiểm soát tất cả tab ngay
            return self.clients.claim();
        }).then(() => {
            // Báo cho tất cả tab biết có version mới → tự reload
            return self.clients.matchAll({ type: 'window' }).then((clients) => {
                clients.forEach((client) => {
                    client.postMessage({ type: 'SW_UPDATED', version: APP_VERSION });
                });
            });
        })
    );
});

// ── FETCH: network-first với fallback cache ───────────────────
self.addEventListener('fetch', (event) => {
    if (event.request.method !== 'GET') return;

    const url = new URL(event.request.url);

    // Bỏ qua Google APIs — không cache
    if (url.hostname.includes('googleapis.com') ||
        url.hostname.includes('accounts.google.com') ||
        url.hostname.includes('fonts.googleapis.com') ||
        url.hostname.includes('fonts.gstatic.com') ||
        url.hostname.includes('cdnjs.cloudflare.com') ||
        url.hostname.includes('cdn.jsdelivr.net') ||
        url.hostname.includes('cdn.tailwindcss.com') ||
        url.hostname.includes('cdn.sheetjs.com')) {
        return; // Để browser tự xử lý
    }

    // HTML, JS, CSS → network-first (luôn lấy mới nhất, fallback cache)
    const isAppFile = ['.html', '.js', '.css'].some(ext => url.pathname.endsWith(ext));
    if (isAppFile) {
        event.respondWith(
            fetch(event.request)
                .then((response) => {
                    if (response && response.status === 200) {
                        const copy = response.clone();
                        caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
                    }
                    return response;
                })
                .catch(() => caches.match(event.request))
        );
        return;
    }

    // Các asset khác (icon, manifest) → cache-first
    event.respondWith(
        caches.match(event.request).then((cached) => cached || fetch(event.request))
    );
});
