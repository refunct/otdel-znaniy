const CACHE_NAME = 'knowledge-base-v6';
const STATIC_ASSETS = [
    './',
    './index.html',
    './style.css',
    './app.js',
    './xlsx.full.min.js',
    './manifest.json'
];

self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => {
                return Promise.allSettled(
                    STATIC_ASSETS.map(url => 
                        cache.add(url).catch(err => console.warn('Failed to cache:', url))
                    )
                );
            })
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
    
    event.respondWith(
        fetch(event.request)
            .catch(() => caches.match(event.request))
    );
});