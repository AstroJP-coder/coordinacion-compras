const CACHE = "compras-v1";
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k)))));
  self.clients.claim();
});
self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  if (req.url.includes("firestore.googleapis.com") || req.url.includes("firebasestorage") || req.url.includes("google")) return;
  e.respondWith(
    fetch(req).then((res) => { const c = res.clone(); caches.open(CACHE).then((ca) => ca.put(req, c)); return res; })
      .catch(() => caches.match(req))
  );
});
