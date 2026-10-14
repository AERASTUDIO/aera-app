// Tenging appsins við Firebase: innskráning (Google) og gagnagrunnur (Firestore).
// Appið sjálft (index.html) veit ekkert um Firebase; það fær hlut `db` með get/set/watch.
import { initializeApp } from "https://www.gstatic.com/firebasejs/12.6.0/firebase-app.js";
import { getAuth, GoogleAuthProvider, signInWithPopup, signInWithRedirect, getRedirectResult, onAuthStateChanged, signOut }
  from "https://www.gstatic.com/firebasejs/12.6.0/firebase-auth.js";
import { initializeFirestore, persistentLocalCache, persistentMultipleTabManager, doc, getDoc, setDoc, onSnapshot }
  from "https://www.gstatic.com/firebasejs/12.6.0/firebase-firestore.js";
import { firebaseConfig, ALLOWED_EMAILS } from "./firebase-config.js";

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
// Staðbundið afrit í símanum/tölvunni: appið virkar án nets og samstillir þegar netið kemur aftur.
const fs = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) });

const ref = path => doc(fs, ...path.split("/"));
const login = document.getElementById("login");
const loginMsg = document.getElementById("login-msg");
const showLogin = (msg = "") => { login.hidden = false; loginMsg.textContent = msg; };

document.getElementById("login-btn").addEventListener("click", async () => {
  const provider = new GoogleAuthProvider();
  provider.setCustomParameters({ prompt: "select_account" });
  loginMsg.textContent = "";
  try { await signInWithPopup(auth, provider); }
  catch (e) {
    // Sumir símar leyfa ekki sprettiglugga – þá er farið á innskráningarsíðuna í staðinn.
    if (["auth/popup-blocked", "auth/operation-not-supported-in-this-environment"].includes(e.code)) {
      try { await signInWithRedirect(auth, provider); } catch (e2) { showLogin("Innskráning tókst ekki: " + (e2.code || e2.message)); }
    } else if (e.code !== "auth/popup-closed-by-user" && e.code !== "auth/cancelled-popup-request") {
      showLogin("Innskráning tókst ekki: " + (e.code || e.message));
    }
  }
});
getRedirectResult(auth).catch(e => showLogin("Innskráning tókst ekki: " + (e.code || e.message)));

window.aeraSignOut = () => signOut(auth).then(() => location.reload());

let started = false;
onAuthStateChanged(auth, async user => {
  if (!user) { showLogin(); return; }
  const email = (user.email || "").toLowerCase();
  if (!ALLOWED_EMAILS.map(x => x.toLowerCase()).includes(email)) {
    await signOut(auth);
    showLogin(`${user.email} hefur ekki aðgang að appinu.`);
    return;
  }
  login.hidden = true;
  if (started) return; started = true;
  window.aeraStart({
    email,
    async get(path) { const s = await getDoc(ref(path)); return { exists: s.exists(), data: () => s.data() }; },
    // Ekki beðið eftir svari netþjóns: skrifin geymast staðbundið strax og fara upp þegar netið er til staðar.
    set(path, data) { return setDoc(ref(path), data); },
    watch(path, cb, onErr) {
      return onSnapshot(ref(path), { includeMetadataChanges: true },
        s => cb({ exists: s.exists(), data: () => s.data(), pending: s.metadata.hasPendingWrites, fromCache: s.metadata.fromCache }), onErr);
    }
  });
});

// Sjálfvirk uppfærsla: þegar appið er opnað aftur er athugað hvort ný útgáfa sé komin,
// og ef svo er endurhleðst appið sjálft (gerist þegar VERSION í sw.js er hækkað).
if ("serviceWorker" in navigator) {
  const hadController = !!navigator.serviceWorker.controller;
  let reloading = false;
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    if (!hadController || reloading) return; reloading = true; location.reload();
  });
  navigator.serviceWorker.register("./sw.js", { updateViaCache: "none" }).then(reg => {
    const check = () => reg.update().catch(() => {});
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") check(); });
    setInterval(check, 30 * 60 * 1000);
  }).catch(() => {});
}
