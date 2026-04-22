self.addEventListener('install',e=>{
  e.waitUntil(
    caches.open('v6').then(c=>c.addAll([
      './',
      './index.html',
      './style.css',
      './app.js',
      './guide.xlsx',
      './tests.xlsx',
      './xlsx.full.min.js'
    ]))
  )
})
self.addEventListener('fetch',e=>{
  const url=e.request.url

  if(url.includes('/img/')||url.includes('/video/')||url.includes('/download/')){
    e.respondWith(new Response('Недоступно в офлайн режиме'))
    return
  }

  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request)))
})