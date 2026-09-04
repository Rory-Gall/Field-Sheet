/* KEI Field Sheet service worker.
   CACHE carries the build number — bump it together with var BUILD in index.html.
   A byte-different sw.js is what makes phones pick up a new build. */
"use strict";
var CACHE = "kei-fieldsheet-b14";
var ASSETS = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png",
  "./icon-maskable-512.png",
  "./apple-touch-icon.png"
];

self.addEventListener("install", function (e) {
  e.waitUntil(
    caches.open(CACHE)
      .then(function (c) {
        /* cache:"reload" bypasses the HTTP cache. Without it, GitHub Pages'
           10-minute max-age can hand the PREVIOUS build's index.html to the
           NEW build's cache — the update pill then reports an update that
           reloads into the old build. Silent success on stale data; never. */
        return c.addAll(ASSETS.map(function (u) {
          return new Request(u, { cache: "reload" });
        }));
      })
      .then(function () { return self.skipWaiting(); })
  );
});

self.addEventListener("activate", function (e) {
  e.waitUntil(
    caches.keys()
      .then(function (keys) {
        return Promise.all(keys.filter(function (k) { return k !== CACHE; })
          .map(function (k) { return caches.delete(k); }));
      })
      .then(function () { return self.clients.claim(); })
  );
});

/* Cache-first with background refresh. Offline at a site, everything serves from
   cache; in signal, the cache quietly updates so the next launch is current. */
self.addEventListener("fetch", function (e) {
  if (e.request.method !== "GET") return;
  var url = new URL(e.request.url);
  if (url.origin !== self.location.origin) return;
  var key = e.request.mode === "navigate" ? "./index.html" : e.request;
  e.respondWith(
    caches.match(key).then(function (hit) {
      /* no-cache: revalidate with the origin, not the browser's HTTP cache */
      var refresh = fetch(e.request, { cache: "no-cache" }).then(function (res) {
        if (res && res.ok) {
          var copy = res.clone();
          return caches.open(CACHE)
            .then(function (c) { return c.put(key, copy); })
            .then(function () { return res; });
        }
        return res;
      }).catch(function () {
        /* offline and nothing cached: a real network-error response, never
           an undefined that turns into an opaque browser failure page */
        return hit || Response.error();
      });
      /* keep the worker alive until the background write finishes — iOS kills
         idle workers fast enough to abort an untracked cache.put */
      e.waitUntil(refresh.catch(function () {}));
      return hit || refresh;
    })
  );
});
