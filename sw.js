self.addEventListener("install", e=>{
  e.waitUntil(
    caches.open("v14").then(c=>{
      return c.addAll([
        "./",
        "./index.html",
        "./style.css",
        "./app.js",
        "./xlsx.full.min.js",
        "./guide.xlsx",
        "./tests.xlsx"
      ])
    })
  )
})

self.addEventListener("fetch", e=>{
  e.respondWith(
    caches.match(e.request).then(r=> r || fetch(e.request))
  )
})