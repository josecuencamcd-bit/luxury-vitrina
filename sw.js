// Service worker mínimo: solo existe para que el navegador considere la
// app "instalable" (requisito de Chrome/Android). No cachea nada todavía.
self.addEventListener('install', function(e){ self.skipWaiting(); });
self.addEventListener('activate', function(e){ self.clients.claim(); });
self.addEventListener('fetch', function(e){
  // passthrough: siempre va a la red, no intercepta nada.
});
