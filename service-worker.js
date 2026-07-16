/* TEKFLOW ERP - SELF-DESTRUCT service worker.
   This replaces the old service-worker.js. On any device that still has the
   old worker registered, this update deletes every cached copy of the app,
   unregisters itself, and reloads the page fresh from the network - so no
   device can ever run an outdated (buggy) version of the ERP again.
   The current worker is ./sw.js, registered by the app itself. */
self.addEventListener('install', function(){ self.skipWaiting(); });
self.addEventListener('activate', function(event){
  event.waitUntil((async function(){
    try{
      const keys = await caches.keys();
      await Promise.all(keys.map(function(k){ return caches.delete(k); }));
    }catch(e){}
    try{ await self.registration.unregister(); }catch(e){}
    try{
      const clients = await self.clients.matchAll({ type: 'window' });
      clients.forEach(function(c){ c.navigate(c.url); });
    }catch(e){}
  })());
});
/* No fetch handler on purpose: this worker never serves anything from cache. */
