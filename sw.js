var CACHE='dv-buchung-v1';
var FILES=['./index.html','./manifest.json','./icon-192.svg','./icon-512.svg'];

self.addEventListener('install',function(e){
  e.waitUntil(caches.open(CACHE).then(function(c){return c.addAll(FILES);}));
  self.skipWaiting();
});

self.addEventListener('activate',function(e){
  e.waitUntil(caches.keys().then(function(keys){
    return Promise.all(keys.filter(function(k){return k!==CACHE;}).map(function(k){return caches.delete(k);}));
  }));
  self.clients.claim();
});

self.addEventListener('fetch',function(e){
  var url=new URL(e.request.url);
  if(url.origin!==self.location.origin){
    return; // fremde Domains (z.B. Google Apps Script) NICHT abfangen - normaler Netzwerk-Request
  }
  e.respondWith(
    fetch(e.request).then(function(r){
      var rc=r.clone();
      caches.open(CACHE).then(function(c){c.put(e.request,rc);});
      return r;
    }).catch(function(){return caches.match(e.request);})
  );
});
