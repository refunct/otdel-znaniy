// sw.js
const CACHE_NAME = 'knowledge-base-v1';
const STATIC_ASSETS = [
    '/',
    '/index.html',
    '/style.css',
    '/app.js',
    '/xlsx.full.min.js',
    '/guide.xlsx',
    '/tests.xlsx'
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(STATIC_ASSETS))
            .then(() => self.skipWaiting())
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then(keys => 
            Promise.all(
                keys.filter(key => key !== CACHE_NAME)
                    .map(key => caches.delete(key))
            )
        ).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);
    
    // Кешируем только статические файлы проекта
    if (STATIC_ASSETS.some(asset => url.pathname.endsWith(asset.replace(/^\//, '')))) {
        event.respondWith(
            caches.match(event.request)
                .then(response => response || fetch(event.request))
        );
        return;
    }
    
    // Для Excel-файлов тоже кешируем
    if (url.pathname.endsWith('.xlsx')) {
        event.respondWith(
            caches.match(event.request)
                .then(response => response || fetch(event.request).then(fetchResponse => {
                    return caches.open(CACHE_NAME).then(cache => {
                        cache.put(event.request, fetchResponse.clone());
                        return fetchResponse;
                    });
                }))
        );
        return;
    }
    
    // Остальное - network first с fallback на офлайн
    event.respondWith(
        fetch(event.request)
            .catch(() => caches.match(event.request))
    );
});