// Offline support for the price guide. Version changes on every new upload.
const CACHE='pricelist-56b71bcbba';
const CORE=['./','index.html','manifest.webmanifest','icon-192.png','icon-512.png'];
const PDFS=["BRC010-Solar-Pathway.pdf","BVC010-Solar-Flood.pdf","BVC080-G2-Solar-Flood.pdf","BVP150-100W-CW-G4.pdf","BVP150-10W-CW-G4.pdf","BVP150-150W-CW-G4.pdf","BVP150-200W-CW-G4.pdf","BVP150-20W-CW-G4.pdf","BVP150-30W-CW-G4.pdf","BVP150-50W-CW-G4.pdf","BVP150-70W-CW-G4.pdf","GU10-CorePro-4.6W-830.pdf","GU10-CorePro-4W-830-DIM.pdf","GU10-Master-4.9W-930-DIM.pdf","RC048B-840-G3.pdf","RC048B-865-G3.pdf"];
self.addEventListener('install',e=>{self.skipWaiting();e.waitUntil(caches.open(CACHE).then(async c=>{await c.addAll(CORE);for(const f of PDFS){try{await c.add(f)}catch(_){}}}))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k!==CACHE).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('fetch',e=>{
  const r=e.request; if(r.method!=='GET'||new URL(r.url).origin!==location.origin) return;
  if(r.mode==='navigate'||r.url.endsWith('.html')){ // network first so updates show up; cached copy when offline
    e.respondWith(fetch(r).then(res=>{const cp=res.clone();caches.open(CACHE).then(c=>c.put(r,cp));return res})
      .catch(()=>caches.match(r,{ignoreSearch:true}).then(h=>h||caches.match('index.html'))));
    return;
  }
  e.respondWith(caches.match(r).then(hit=>hit||fetch(r).then(res=>{if(res.ok){const cp=res.clone();caches.open(CACHE).then(c=>c.put(r,cp))}return res})));
});
