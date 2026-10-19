// Gamla slóðin (/aera-app/): þessi þjónustuskrifta tekur við af gömlu útgáfunni á tækjum, hreinsar gömlu
// skyndiminnin (aera-…), afskráir sig og opnar Atrium á nýju slóðinni.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", e => e.waitUntil((async () => {
  for (const k of await caches.keys()) if (k.startsWith("aera-")) await caches.delete(k);
  await self.registration.unregister();
  for (const c of await self.clients.matchAll({ type: "window" })) c.navigate("https://aerastudio.github.io/atrium/");
})()));
