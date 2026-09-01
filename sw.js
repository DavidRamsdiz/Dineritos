/* Dineritos: cache del armazon para que la app abra sin conexion. */
var CACHE = "dineritos-v13736";
var SHELL = ["./", "./index.html", "./app.js", "./manifest.webmanifest",
             "./icon-192.png", "./icon-512.png"];

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
  /* Graph, el login de Microsoft y las fuentes van siempre a la red. */
  if (ev.request.method !== "GET" || u.origin !== location.origin) return;
  ev.respondWith(
    caches.match(ev.request).then(function (hit) {
      if (hit) {
        /* refresco en segundo plano para no quedarse con una version vieja */
        fetch(ev.request).then(function (res) {
          if (res && res.ok) caches.open(CACHE).then(function (c) { c.put(ev.request, res); });
        }).catch(function () {});
        return hit;
      }
      return fetch(ev.request).then(function (res) {
        if (res && res.ok) {
          var cp = res.clone();
          caches.open(CACHE).then(function (c) { c.put(ev.request, cp); });
        }
        return res;
      }).catch(function () { return caches.match("./index.html"); });
    })
  );
});
