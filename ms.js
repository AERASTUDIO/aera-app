// Tenging við Microsoft 365 (OneDrive og Outlook) í gegnum Microsoft Graph.
// Innskráning: OAuth 2.0 með PKCE (engin leyndarmál í kóðanum). Lyklar geymast aðeins á þessu tæki.
(function () {
  const CFG = {
    clientId: "771b0263-4629-4c0d-baef-f95688c59aca",
    tenant: "1ed6735b-fd33-4b11-ad13-3f1a6e11c0f9",
    scopes: "openid profile offline_access User.Read Files.ReadWrite Mail.Read Mail.Send",
    base: "openid profile offline_access User.Read Files.ReadWrite Mail.Read"
  };
  const REDIRECT = location.hostname === "aerastudio.github.io" ? "https://aerastudio.github.io/aera-app/" : location.origin + location.pathname;
  const AUTH = `https://login.microsoftonline.com/${CFG.tenant}/oauth2/v2.0`;
  const KEY = "aera_ms", PK = "aera_ms_pkce", RES = "aera_ms_res";
  const ls = { get(k) { try { return JSON.parse(localStorage.getItem(k) || "null"); } catch { return null; } },
               set(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, JSON.stringify(v)); } catch {} } };

  // 1) Ef þessi síða er svarið frá Microsoft (?code=…) í sprettiglugga: skila kóðanum og loka.
  const qp = new URLSearchParams(location.search);
  const pk0 = ls.get(PK);
  if ((qp.get("code") || qp.get("error")) && pk0 && qp.get("state") === pk0.state && pk0.mode === "popup") {
    ls.set(RES, { code: qp.get("code"), error: qp.get("error_description") || qp.get("error"), state: qp.get("state"), t: Date.now() });
    document.documentElement.innerHTML = '<body style="font-family:Inter,system-ui,sans-serif;background:#0c0d0e;color:#e8eaed;display:grid;place-items:center;height:100vh;margin:0"><p>Tenging tókst – þú mátt loka þessum glugga.</p></body>';
    window.__aeraPopup = true;
    try { window.stop(); } catch {}
    try { window.close(); } catch {}
    return;
  }

  let tok = ls.get(KEY); // {access, exp, refresh, rexp, name, email}
  const fire = () => window.dispatchEvent(new Event("ms-change"));
  const b64u = b => btoa(String.fromCharCode(...new Uint8Array(b))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
  const rnd = n => b64u(crypto.getRandomValues(new Uint8Array(n)));
  const jwt = t => { try { return JSON.parse(decodeURIComponent(escape(atob(t.split(".")[1].replace(/-/g, "+").replace(/_/g, "/"))))); } catch { return {}; } };

  async function tokenReq(params) {
    const r = await fetch(AUTH + "/token", { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({ client_id: CFG.clientId, scope: params.grant_type === "refresh_token" ? (tok?.granted || CFG.base) : CFG.scopes, ...params }) });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j.error_description || j.error || ("Villa " + r.status));
    const now = Date.now(), prev = tok || {};
    const id = j.id_token ? jwt(j.id_token) : {};
    tok = { access: j.access_token, exp: now + (j.expires_in || 3600) * 1000,
            refresh: j.refresh_token || prev.refresh,
            rexp: j.refresh_token ? now + 23.5 * 3600 * 1000 : prev.rexp, // SPA-lyklar gilda í 24 klst.
            name: id.name || prev.name, email: id.preferred_username || prev.email,
            granted: j.scope ? "openid profile offline_access " + j.scope.split(" ").map(x => x.replace(/^https:\/\/graph\.microsoft\.com\//i, "")).join(" ") : prev.granted };
    ls.set(KEY, tok); fire(); return tok;
  }

  async function exchange(res) {
    const pk = ls.get(PK); ls.set(PK, null); ls.set(RES, null);
    if (!pk || res.state !== pk.state) throw new Error("Innskráning rann út – reyndu aftur.");
    if (res.error) throw new Error(res.error);
    await tokenReq({ grant_type: "authorization_code", code: res.code, redirect_uri: REDIRECT, code_verifier: pk.verifier });
  }

  const standalone = () => navigator.standalone === true || matchMedia("(display-mode: standalone)").matches;

  async function login() {
    const verifier = rnd(48), state = rnd(16);
    const useRedirect = standalone();
    // Sprettiglugginn er opnaður strax (áður en beðið er), annars lokar vafrinn á hann.
    const w = useRedirect ? null : window.open("", "aera_ms", "width=480,height=680");
    const challenge = b64u(await crypto.subtle.digest("SHA-256", new TextEncoder().encode(verifier)));
    const mode = useRedirect || !w ? "redirect" : "popup";
    ls.set(PK, { verifier, state, mode, back: location.hash || "" }); ls.set(RES, null);
    const url = AUTH + "/authorize?" + new URLSearchParams({ client_id: CFG.clientId, response_type: "code", redirect_uri: REDIRECT,
      scope: CFG.scopes, state, code_challenge: challenge, code_challenge_method: "S256", response_mode: "query",
      prompt: "select_account", login_hint: tok?.email || "dagur@sonik.is" });
    if (mode === "redirect") { location.href = url; return new Promise(() => {}); }
    w.location.href = url;
    return new Promise((resolve, reject) => {
      const t0 = Date.now();
      const iv = setInterval(async () => {
        const res = ls.get(RES);
        if (res && res.state === state) { clearInterval(iv); try { await exchange(res); resolve(); } catch (e) { reject(e); } return; }
        if (Date.now() - t0 > 5 * 60 * 1000) { clearInterval(iv); reject(new Error("Innskráning tók of langan tíma.")); }
      }, 400);
    });
  }

  // 2) Svar eftir endurbeiningu (heimaskjás-app): klára innskráningu og hreinsa slóðina.
  const ready = (async () => {
    if ((qp.get("code") || qp.get("error")) && pk0 && qp.get("state") === pk0.state) {
      history.replaceState(null, "", REDIRECT.startsWith(location.origin) ? location.pathname + (pk0.back || "") : location.pathname);
      try { await exchange({ code: qp.get("code"), error: qp.get("error_description") || qp.get("error"), state: qp.get("state") }); }
      catch (e) { MS.lastError = e.message; fire(); }
    }
  })();

  let refreshing = null;
  async function token() {
    await ready;
    if (!tok) return null;
    if (tok.access && tok.exp - Date.now() > 90 * 1000) return tok.access;
    if (!tok.refresh || (tok.rexp && tok.rexp < Date.now())) { MS.expired = true; fire(); return null; }
    refreshing = refreshing || tokenReq({ grant_type: "refresh_token", refresh_token: tok.refresh })
      .catch(e => { MS.lastError = e.message; MS.expired = true; tok.access = null; tok.refresh = null; ls.set(KEY, tok); fire(); return null; })
      .finally(() => { refreshing = null; });
    const t = await refreshing; return t ? t.access : null;
  }

  async function g(path, opt = {}) {
    const at = await token();
    if (!at) { const e = new Error("Ekki tengt Microsoft"); e.code = "auth"; throw e; }
    const headers = { Authorization: "Bearer " + at, ...(opt.headers || {}) };
    let body = opt.body;
    if (body && !(body instanceof Blob) && !(body instanceof Uint8Array) && typeof body !== "string") { body = JSON.stringify(body); headers["Content-Type"] = "application/json"; }
    const r = await fetch(path.startsWith("http") ? path : "https://graph.microsoft.com/v1.0" + path, { method: opt.method || "GET", headers, body });
    if (r.status === 204) return null;
    const j = await r.json().catch(() => null);
    if (!r.ok) { const e = new Error(j?.error?.message || ("Villa " + r.status)); e.status = r.status; e.code = j?.error?.code; throw e; }
    return j;
  }

  const MS = window.MS = {
    login, g, token, ready,
    // Tengt = til er gildur lykill eða endurnýjunarlykill á þessu tæki.
    on: () => !!(tok && ((tok.access && tok.exp > Date.now()) || (tok.refresh && (!tok.rexp || tok.rexp > Date.now())))),
    user: () => tok ? { name: tok.name, email: tok.email } : null,
    hadLogin: () => !!tok,
    // Hefur aðgangurinn veitt tiltekna heimild (t.d. "Mail.Send")?
    can: sc => (" " + (tok?.granted || "") + " ").toLowerCase().includes(" " + sc.toLowerCase() + " "),
    logout() { tok = null; ls.set(KEY, null); fire(); },
    lastError: null, expired: false
  };
})();
