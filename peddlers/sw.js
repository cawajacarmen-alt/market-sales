// Network first so updates show up; falls back to the saved copy when offline.
// Each branch page has its own cache; only this page's old caches are removed.
const P="app-pdl-v", C=P+"3", FILES=["./", "index.html", "manifest.webmanifest", "../icon.svg", "../icon-192.png", "../icon-512.png", "../lib/xlsx.full.min.js"];
self.addEventListener("install",e=>{ e.waitUntil(caches.open(C).then(c=>c.addAll(FILES)).then(()=>self.skipWaiting())); });
self.addEventListener("activate",e=>{ e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x.startsWith(P)&&x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim())); });
self.addEventListener("fetch",e=>{ if(e.request.method!=="GET"||new URL(e.request.url).origin!==location.origin) return;
  e.respondWith(fetch(e.request).then(r=>{ const cp=r.clone(); caches.open(C).then(c=>c.put(e.request,cp)); return r; }).catch(()=>caches.match(e.request).then(r=>r||caches.match("index.html")))); });
