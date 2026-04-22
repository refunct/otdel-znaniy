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

// Функция для вычисления хеша из ArrayBuffer
async function hashBuffer(buffer) {
    const hashBuffer = await crypto.subtle.digest('SHA-256', buffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

// Функция для вычисления общего хеша из нескольких хешей
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
            const hashes = [];
            const filesToCache = new Map();
            
            // Загружаем все файлы и вычисляем их хеши
            for (const asset of STATIC_ASSETS) {
                try {
                    const response = await fetch(asset);
                    const buffer = await response.arrayBuffer();
                    const hash = await hashBuffer(buffer);
                    hashes.push(hash);
                    filesToCache.set(asset, buffer);
                } catch (err) {
                    console.warn(`Failed to load ${asset}:`, err);
                    // Если файл не загрузился, используем путь как часть хеша
                    hashes.push(asset);
                }
            }
            
            // Создаём общий хеш из всех хешей
            const combinedHash = await getCombinedHash(hashes);
            CACHE_NAME = `knowledge-base-${combinedHash}`;
            
            const cache = await caches.open(CACHE_NAME);
            
            // Кешируем все файлы
            for (const [asset, buffer] of filesToCache) {
                await cache.put(asset, new Response(buffer));
            }
            
            await self.skipWaiting();
        })()
    );
});

self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then(keys => 
            Promise.all(
                keys.filter(key => key.startsWith('knowledge-base-'))
                    .map(key => caches.delete(key))
            )
        ).then(() => self.clients.claim())
    );
});

self.addEventListener('fetch', (event) => {
    const url = new URL(event.request.url);
    
    // Проверяем, есть ли файл в списке статики
    const isStaticAsset = STATIC_ASSETS.some(asset => 
        url.pathname.endsWith(asset.replace('./', '')) || 
        url.pathname.endsWith(asset)
    );
    
    if (isStaticAsset) {
        event.respondWith(
            caches.match(event.request)
                .then(response => response || fetch(event.request))
        );
        return;
    }
    
    // Остальное - network first
    event.respondWith(
        fetch(event.request)
            .catch(() => caches.match(event.request))
    );
});