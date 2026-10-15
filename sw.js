// Þjónustuskrifta (service worker): geymir appið á tækinu svo það opnist hratt og virki án nets.
// Hækkaðu VERSION þegar eitthvað breytist, þá sækja tækin nýju útgáfuna.
const VERSION = "aera-1.1.21"; // sama og APP_VERSION í index.html
const SHELL = ["./", "index.html", "firebase.js", "firebase-config.js", "xlsx.js", "ms.js", "ms-app.js", "pdf.js", "reikn.js", "manifest.webmanifest",
  "icons/icon-192.png", "icons/icon-512.png", "icons/apple-touch-icon.png"];

self.addEventListener("install", e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== VERSION).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", e => {
  const req = e.request; if (req.method !== "GET") return;
  const url = new URL(req.url);
  const own = url.origin === location.origin;
  const lib = url.hostname === "www.gstatic.com" || url.hostname === "fonts.googleapis.com" || url.hostname === "fonts.gstatic.com";
  if (!own && !lib) return; // Firebase-gögn og innskráning fara alltaf beint á netið
  if (own) {
    // Netið fyrst (nýjasta útgáfan), annars geymda útgáfan
    e.respondWith(fetch(req, { cache: "no-cache" }).then(r => { if (r.ok) { const c = r.clone(); caches.open(VERSION).then(x => x.put(req, c)); } return r; })
      .catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match("index.html"))));
  } else {
    // Söfn og letur breytast ekki: geymda útgáfan fyrst
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(n => { if (n.ok) { const c = n.clone(); caches.open(VERSION).then(x => x.put(req, c)); } return n; })));
  }
});
