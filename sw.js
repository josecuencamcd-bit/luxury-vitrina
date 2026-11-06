// Service worker mínimo: Chrome/Android exige uno (con un listener de
// "fetch" que responda) para considerar la app "instalable" de verdad, no
// solo como acceso directo. No cachea nada — solo deja pasar la red.
self.addEventListener('install', function(e){ self.skipWaiting(); });
self.addEventListener('activate', function(e){ e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function(e){
  e.respondWith(fetch(e.request));
});
