/* Service worker: makes the system installable and keeps the app shell available offline.
   Data is not cached here; Firestore keeps its own offline cache and syncs writes when the connection returns (offline.js).
   Strategy: app files network-first (so updates arrive) with a short timeout and cache fallback, so a weak connection
   does not leave the page hanging; vendored libraries, icons and Google Fonts cache-first. */
importScripts('version.js');
const CACHE='brega-'+self.APP_VERSION;
const FONTS='brega-fonts';
const SHELL=['./','index.html','boot.js','app.css','i18n.js','adapter.js','core.js','dash.js','contracts.js','org.js','analytics.js','cs.js','app.js','settings.js','fileimport.js','ocrimport.js','bi.js','bi-dash.js','offline.js','attach.js','meetings.js','duties-data.js','duties.js','units.js','sms055.js','closures.js','debts.js','margin.js','margin-econ.js','home.js','odoo.js','version.js','releases.js','pwa.js','notify.js','govern.js','safety.js','orgpeople.js','palette.js','quality.js',
  'manifest.webmanifest','icons/icon-192.png','icons/icon-512.png','icons/icon-maskable-512.png','icons/apple-touch-icon.png','icons/logo.webp','icons/favicon-64.png',
  'vendor/firebase/firebase-app-compat.js','vendor/firebase/firebase-auth-compat.js','vendor/firebase/firebase-firestore-compat.js',
  'vendor/pdfjs/pdf.min.js','vendor/pdfjs/pdf.worker.min.js'];
const STATIC=/\/(vendor|icons)\//;
const NET_WAIT=5000;
self.addEventListener('install',e=>{e.waitUntil(caches.open(CACHE).then(c=>Promise.all(SHELL.map(u=>c.add(u).catch(()=>{})))))});
self.addEventListener('activate',e=>{e.waitUntil(caches.keys().then(ks=>Promise.all(ks.filter(k=>k.startsWith('brega-')&&k!==CACHE&&k!==FONTS).map(k=>caches.delete(k)))).then(()=>self.clients.claim()))});
self.addEventListener('message',e=>{if(e.data==='SKIP_WAITING')self.skipWaiting()});
const put=(name,req,res)=>caches.open(name).then(c=>c.put(req,res)).catch(()=>{});
function fromCache(req){return caches.match(req).then(r=>r||(req.mode==='navigate'?caches.match('index.html'):undefined))}
self.addEventListener('fetch',e=>{
  const req=e.request;if(req.method!=='GET')return;
  const url=new URL(req.url);
  if(/^fonts\.(googleapis|gstatic)\.com$/.test(url.hostname)){e.respondWith(caches.open(FONTS).then(c=>c.match(req).then(r=>r||fetch(req).then(res=>{if(res.ok||res.type==='opaque')c.put(req,res.clone());return res}))));return}
  if(url.origin!==self.location.origin)return;
  if(STATIC.test(url.pathname)){e.respondWith(caches.match(req).then(r=>r||fetch(req).then(res=>{if(res.ok)put(CACHE,req,res.clone());return res})));return}
  const cacheable=req.mode==='navigate'||/\.(js|html|webmanifest|json)$/.test(url.pathname)||url.pathname.endsWith('/');
  e.respondWith(new Promise(resolve=>{let done=false;const finish=r=>{if(!done){done=true;resolve(r)}};
    // weak connection: after NET_WAIT serve the cached copy if there is one, and let the network answer refresh the cache
    const t=setTimeout(()=>fromCache(req).then(r=>{if(r)finish(r)}),NET_WAIT);
    fetch(req).then(res=>{clearTimeout(t);if(res.ok&&cacheable)put(CACHE,req,res.clone());finish(res)})
      .catch(()=>{clearTimeout(t);fromCache(req).then(r=>finish(r||Response.error()))});
  }));
});
