// sw.js
let CACHE_NAME = 'knowledge-base-v1';

const STATIC_ASSETS = [
    './',
    './index.html',
    './style.css',
    './app.js',
    './xlsx.full.min.js',
    './manifest.json',
    './guide.xlsx',
    './tests.xlsx'
];

async function hashBuffer(buffer) {
    const hashBuffer = await crypto.subtle.digest('SHA-256', buffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

async function getCombinedHash(hashes) {
    const combined = hashes.join('');
    const encoder = new TextEncoder();
    const data = encoder.encode(combined);
    const hashBuffer = await crypto.subtle.digest('SHA-256', data);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('').substring(0, 16);
}

self.addEventListener('install', (event) => {
    event.waitUntil(
        (async () => {
            try {
                const hashes = [];
                const filesToCache = new Map();
                
                for (const asset of STATIC_ASSETS) {
                    try {
                        const response = await fetch(asset);
                        if (response.ok) {
                            const buffer = await response.arrayBuffer();
                            const hash = await hashBuffer(buffer);
                            hashes.push(hash);
                            filesToCache.set(asset, buffer);
                        } else {
                            hashes.push(asset + Date.now());
                        }
                    } catch (err) {
                        console.warn(`Failed to load ${asset}:`, err);
                        hashes.push(asset + Date.now());
                    }
                }
                
                const combinedHash = await getCombinedHash(hashes);
                CACHE_NAME = `knowledge-base-${combinedHash}`;
                
                const cache = await caches.open(CACHE_NAME);
                
                for (const [asset, buffer] of filesToCache) {
                    await cache.put(asset, new Response(buffer));
                }
                
                console.log('SW: Cached', filesToCache.size, 'files in', CACHE_NAME);
                await self.skipWaiting();
            } catch (err) {
                console.error('SW Install error:', err);
            }
        })()
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then(keys => 
            Promise.all(
                keys.filter(key => key.startsWith('knowledge-base-'))
                    .map(key => {
                        console.log('SW: Deleting old cache', key);
                        return caches.delete(key);
                    })
            )
        ).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);
    
    const isStaticAsset = STATIC_ASSETS.some(asset => 
        url.pathname.endsWith(asset.replace('./', '')) || 
        url.pathname.endsWith(asset)
    );
    
    if (isStaticAsset) {
        event.respondWith(
            caches.match(event.request)
                .then(response => {
                    if (response) {
                        return response;
                    }
                    return fetch(event.request).then(fetchResponse => {
                        if (!fetchResponse || fetchResponse.status !== 200) {
                            return fetchResponse;
                        }
                        return caches.open(CACHE_NAME).then(cache => {
                            cache.put(event.request, fetchResponse.clone());
                            return fetchResponse;
                        });
                    });
                })
                .catch(() => {
                    return caches.match(event.request);
                })
        );
        return;
    }
    
    event.respondWith(
        fetch(event.request)
            .catch(() => caches.match(event.request))
    );
});