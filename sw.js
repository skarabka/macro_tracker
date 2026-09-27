const CACHE='macro-tracker-v3';
const ASSETS=['./','index.html','styles.css','app.js','config.js','manifest.webmanifest','icon-192.png','icon-512.png'];
self.addEventListener('install',e=>{self.skipWaiting();e.waitUntil(caches.open(CACHE).then(c=>c.addAll(ASSETS)))});
self.addEventListener('activate',e=>e.waitUntil(Promise.all([caches.keys().then(keys=>Promise.all(keys.filter(k=>k!==CACHE).map(k=>caches.delete(k)))),self.clients.claim()])));
self.addEventListener('fetch',e=>{
  const u=new URL(e.request.url);
  if(u.origin===location.origin && u.pathname.endsWith('/config.js')){
    e.respondWith(fetch(e.request).catch(()=>caches.match(e.request))); return;
  }
  e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request)));
});
