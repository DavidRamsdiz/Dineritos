/* Dineritos 0c760f2
   Estrategia: para el armazon (pagina y app.js) se pide SIEMPRE a la red y la copia
   local es solo el respaldo de cuando no hay conexion. Asi un arreglo llega al
   dispositivo en la siguiente carga en vez de quedarse con la version vieja.
   Para los iconos y el manifiesto, al contrario: primero la copia local. */
var CACHE = "dineritos-0c760f2";
var SHELL = ["./", "./index.html", "./app.js?v=0c760f2", "./manifest.webmanifest",
             "./icon-192.png", "./icon-512.png"];
var ARMAZON = /(\/|index\.html|app\.js)$/;

self.addEventListener("install", function (ev) {
  ev.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); })
    .then(function () { return self.skipWaiting(); }));
});
self.addEventListener("activate", function (ev) {
  ev.waitUntil(caches.keys().then(function (ks) {
    return Promise.all(ks.map(function (k) { return k === CACHE ? null : caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});
self.addEventListener("fetch", function (ev) {
  var u = new URL(ev.request.url);
  /* Graph, el login de Microsoft y las fuentes van siempre a la red, sin tocar la cache. */
  if (ev.request.method !== "GET" || u.origin !== location.origin) return;

  if (ARMAZON.test(u.pathname)) {
    ev.respondWith(
      fetch(ev.request).then(function (res) {
        if (res && res.ok) {
          var cp = res.clone();
          caches.open(CACHE).then(function (c) { c.put(ev.request, cp); });
        }
        return res;
      }).catch(function () {
        /* sin conexion: vale cualquier copia, aunque la query no coincida */
        return caches.match(ev.request).then(function (hit) {
          return hit || caches.match(ev.request, { ignoreSearch: true });
        }).then(function (hit) {
          return hit || caches.match("./index.html");
        });
      })
    );
    return;
  }
  ev.respondWith(
    caches.match(ev.request).then(function (hit) {
      return hit || fetch(ev.request).then(function (res) {
        if (res && res.ok) {
          var cp = res.clone();
          caches.open(CACHE).then(function (c) { c.put(ev.request, cp); });
        }
        return res;
      });
    })
  );
});
