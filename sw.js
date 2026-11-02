const C='form-v4';const PRE=['./','./index.html','./manifest.json','./icon-180.png','https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js','https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/tesseract.min.js'];
self.addEventListener('install',e=>{self.skipWaiting();e.waitUntil(caches.open(C).then(c=>Promise.all(PRE.map(u=>c.add(u).catch(()=>{})))))});
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('fetch',e=>{if(e.request.method!='GET')return;e.respondWith(caches.match(e.request).then(r=>r||fetch(e.request).then(res=>{if(res.ok||res.type=='opaque'){const cl=res.clone();caches.open(C).then(c=>c.put(e.request,cl))}return res})))});
