const C='form-v9';
const LIBS=['https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js','https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js'];
self.addEventListener('install',e=>{self.skipWaiting();e.waitUntil(caches.open(C).then(c=>Promise.all(['./','./index.html',...LIBS].map(u=>c.add(u).catch(()=>{})))))});
self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==C).map(k=>caches.delete(k)))).then(()=>self.clients.claim())));
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!='GET')return;const u=new URL(r.url);
 const sameOrigin=u.origin===location.origin;
 if(sameOrigin){ // network-first for our own files
  e.respondWith(fetch(r,{cache:'no-store'}).then(res=>{const cl=res.clone();caches.open(C).then(c=>c.put(r,cl));return res}).catch(()=>caches.match(r).then(x=>x||caches.match('./index.html'))));return}
 e.respondWith(caches.match(r).then(x=>x||fetch(r).then(res=>{if(res.ok||res.type=='opaque'){const cl=res.clone();caches.open(C).then(c=>c.put(r,cl))}return res})))});
