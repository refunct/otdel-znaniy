const STATIC_ASSETS = [
    './',
    './index.html',
    './styles/style.css',
    './js/app.js',
    './js/guides.js',
    './js/tests.js',
    './js/xlsx.full.min.js',
    './docs/guide.xlsx',
    './docs/tests.xlsx'
];
let CACHE_NAME = 'knowledge-base-v1';

async function hashBuffer(buffer) {
    const hash = await crypto.subtle.digest('SHA-256', buffer);
    return Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2,'0')).join('');
}

self.addEventListener('install', e => {
    e.waitUntil((async () => {
        const hashes = [];
        const files = new Map();
        for (const url of STATIC_ASSETS) {
            try {
                const res = await fetch(url);
                if (res.ok) {
                    const buf = await res.arrayBuffer();
                    hashes.push(await hashBuffer(buf));
                    files.set(url, buf);
                } else hashes.push(url + Date.now());
            } catch { hashes.push(url + Date.now()); }
        }
        const combined = hashes.join('');
        const encoder = new TextEncoder();
        const hash = await crypto.subtle.digest('SHA-256', encoder.encode(combined));
        const hashStr = Array.from(new Uint8Array(hash)).map(b => b.toString(16).padStart(2,'0')).join('').slice(0,16);
        CACHE_NAME = `knowledge-${hashStr}`;
        const cache = await caches.open(CACHE_NAME);
        for (const [url, buf] of files) await cache.put(url, new Response(buf));
        await self.skipWaiting();
    })());
});

self.addEventListener('activate', e => {
    e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => k.startsWith('knowledge-') && k!==CACHE_NAME).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});

self.addEventListener('fetch', e => {
    const url = new URL(e.request.url);
    if (STATIC_ASSETS.some(a => url.pathname.endsWith(a.replace('./','')))) {
        e.respondWith(caches.match(e.request).then(r => r || fetch(e.request)));
    } else if (e.request.mode === 'navigate') {
        e.respondWith(fetch(e.request).catch(() => caches.match('./index.html')));
    } else {
        e.respondWith(fetch(e.request).catch(() => caches.match(e.request)));
    }
});