/* Pósthólf í appinu (Microsoft Graph, Mail.Read): lesa innhólf og möppur, leita, opna póst,
   tengja póst við verkefni (S.postar) og vista viðhengi í möppu verkefnisins. */
(function () {
  const SEL = "$select=id,subject,from,receivedDateTime,bodyPreview,webLink,isRead,hasAttachments,conversationId";
  const FOLD = [["inbox", "Innhólf"], ["aera", "AERA"], ["sentitems", "Sent"], ["tengilidir", "Tengiliðir"]];
  let fold = "inbox", q = "", list = null, next = null, loading = false, err = null, aeraId = null, target = null; // target = verkefni sem verið er að finna pósta fyrir
  const MSG = {}, ATT = {};
  const msOn = () => !!(window.MS && MS.on());
  const linked = gid => (S?.postar || []).find(m => m.gid === gid);
  const fromName = m => m.from?.emailAddress?.name || m.from?.emailAddress?.address || "";
  const when = iso => { const x = new Date(iso), t = new Date(); return x.toDateString() === t.toDateString() ? x.toTimeString().slice(0, 5) : fmtDate(iso); };

  async function folderPath() {
    if (fold !== "aera") return `/me/mailFolders/${fold}/messages`;
    if (!aeraId) { const f = encodeURIComponent("displayName eq 'AERA'"); let xs = (await MS.g(`/me/mailFolders?$filter=${f}&$select=id`)).value; if (!xs.length) xs = (await MS.g(`/me/mailFolders/inbox/childFolders?$filter=${f}&$select=id`)).value; if (!xs.length) throw new Error("Engin póstmappa sem heitir AERA."); aeraId = xs[0].id; }
    return `/me/mailFolders/${aeraId}/messages`;
  }
  async function load(more) {
    if (loading) return; loading = true; err = null;
    try {
      let j;
      if (more && next) j = await MS.g(next);
      else if (q) j = await MS.g(`/me/messages?$search=${encodeURIComponent('"' + q.replace(/"/g, "") + '"')}&$top=30&${SEL}`);
      else j = await MS.g(`${await folderPath()}?$top=30&$orderby=receivedDateTime%20desc&${SEL}`);
      list = more ? (list || []).concat(j.value || []) : (j.value || []); next = j["@odata.nextLink"] || null;
    } catch (e) { err = e.message; if (!more) list = []; }
    loading = false; render();
  }
  const reset = () => { list = null; next = null; };

  /* ---------- Listi ---------- */
  const row = m => { const L = linked(m.id), p = L && findProj(L.verkefniId), g = !L && window.msGuessProj ? findProj(msGuessProj((m.subject || "") + " " + (m.bodyPreview || ""), m.from?.emailAddress?.address)) : null;
    return `<div class="lrow tap porow${m.isRead === false ? " unread" : ""}" ${tapAttr("post", m.id)}><div class="ll"><span class="podot"></span></div>
      <div class="lm"><div class="lt"><span class="pofrom">${esc(fromName(m))}</span>${m.hasAttachments ? " 📎" : ""}</div><div class="ls" style="color:var(--ink)">${esc(m.subject || "(ekkert efni)")}</div><div class="ls">${esc((m.bodyPreview || "").slice(0, 110))}</div>
      ${p ? `<span class="pill" style="margin-top:4px">${esc(p.nr)} ${esc(p.stutt)}</span>` : ""}</div>
      <div class="lr" style="flex-direction:column;align-items:flex-end;gap:6px"><span class="lc">${esc(when(m.receivedDateTime))}</span>
      ${!L && MCP ? (target ? `<button type="button" class="btn" data-act="polink:${esc(m.id)}:${esc(target)}" style="padding:4px 10px;min-height:30px">Tengja</button>` : g ? `<button type="button" class="btn" data-act="polink:${esc(m.id)}:${esc(g.id)}" style="padding:4px 10px;min-height:30px" title="Tengja við ${esc(g.nafn)}">→ ${esc(g.nr)}</button>` : "") : ""}</div></div>`; };

  function vTengilidir() {
    const t = q.toLowerCase(), all = contacts(), xs = all.filter(c => !t || c.a.toLowerCase().includes(t) || (c.n || "").toLowerCase().includes(t) || (c.v || "").toLowerCase().includes(t));
    const inBook = a => BOOK().some(x => x.a.toLowerCase() === a.toLowerCase());
    return `<form id="poq" class="leitbox" style="margin-top:8px"><input id="poqi" type="search" placeholder="Leita að tengilið…" value="${esc(q)}" enterkeyhint="search" autocomplete="off"></form>
      <div class="stat-row" style="margin-top:10px">${FOLD.map(([k, l]) => `<button type="button" class="chipbtn" data-act="pofold:${k}" aria-pressed="${fold === k}">${l}</button>`).join("")}<button type="button" class="chipbtn" data-act="poharv:" title="Sækja netföng úr pósti">↻</button></div>
      <p class="hint" style="margin:10px 0 0">${all.length} tengiliðir · safnast sjálfkrafa úr pósti sem þú sendir og færð, og úr Viðskiptavinum. Þeir sem þú sendir oftast á eru efstir.</p>
      ${xs.length ? `<div class="llist" style="margin-top:8px">${xs.slice(0, 300).map(c => `<div class="lrow"><div class="ll"><span class="avat">${esc((c.n || c.a).trim()[0]?.toUpperCase() || "?")}</span></div><div class="lm"><div class="lt">${esc(c.n || c.a)}</div><div class="ls">${esc([c.n ? c.a : "", c.v].filter(Boolean).join(" · "))}</div></div>
        <div class="lr" style="gap:6px">${MCP ? `<button type="button" class="btn" data-act="poto:${esc(c.a)}" style="padding:4px 10px;min-height:30px">✉</button>` : ""}${MCP && inBook(c.a) ? `<button type="button" class="btn" data-act="pofor:${esc(c.a)}" style="padding:4px 10px;min-height:30px" aria-label="Gleyma ${esc(c.a)}">✕</button>` : ""}</div></div>`).join("")}</div>` : `<div class="empty" style="margin-top:16px">${q ? "Enginn fannst." : "Engir tengiliðir ennþá. Ýttu á ↻ til að sækja netföng úr póstinum."}</div>`}`;
  }
  window.vPostur = function () {
    if (!msOn()) return `<div class="empty" style="margin-top:16px">Tengdu Microsoft til að sjá pósthólfið þitt í appinu.${window.msMissing ? msMissing() : ""}</div>`;
    if (fold === "tengilidir") return vTengilidir();
    if (!list && !loading) { load(false); harvest(false).catch(() => {}); }
    const tp = target && findProj(target);
    return `${tp ? `<div class="card" style="margin:8px 0;display:flex;align-items:center;gap:10px"><div class="grow">Finndu pósta fyrir <b>${esc(tp.nr)} ${esc(tp.stutt)}</b> og ýttu á „Tengja“.</div><button type="button" class="btn" data-act="potarget:">Lokið</button></div>` : ""}
      <form id="poq" class="leitbox" style="margin-top:8px"><input id="poqi" type="search" placeholder="Leita í pósti…" value="${esc(q)}" enterkeyhint="search" autocomplete="off"></form>
      ${q ? "" : `<div class="stat-row" style="margin-top:10px">${FOLD.map(([k, l]) => `<button type="button" class="chipbtn" data-act="pofold:${k}" aria-pressed="${fold === k}">${l}</button>`).join("")}<button type="button" class="chipbtn" data-act="poref:" aria-label="Uppfæra">↻</button></div>`}
      ${err ? `<div class="msg err" style="margin-top:10px">${esc(err)}</div>` : ""}
      ${list == null || (loading && !list?.length) ? `<div class="empty" style="margin-top:16px">Sæki pósta…</div>` : list.length ? `<div class="llist" style="margin-top:8px">${list.map(row).join("")}</div>${next ? `<button type="button" class="btn" data-act="pomore:" style="width:100%;margin-top:10px">${loading ? "Sæki…" : "Sækja fleiri"}</button>` : ""}` : `<div class="empty" style="margin-top:16px">${q ? "Ekkert fannst." : "Engir póstar í möppunni."}</div>`}
      <p class="hint" style="margin-top:10px">Pósturinn er lesinn og sendur beint í gegnum Outlook (sendir póstar vistast í Sent).</p>`;
  };

  /* ---------- Einn póstur ---------- */
  async function loadMsg(id) {
    try {
      MSG[id] = await MS.g(`/me/messages/${encodeURIComponent(id)}?$select=id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,webLink,hasAttachments,bodyPreview`);
      const m = MSG[id], html = m.body?.contentType === "html" ? m.body.content || "" : "", hasCid = /src\s*=\s*["']?cid:/i.test(html);
      if (m.hasAttachments || hasCid) {
        // contentId er aðeins til á fileAttachment, svo það er ekki valið í listanum (annars villa frá Graph)
        const all = (await MS.g(`/me/messages/${encodeURIComponent(id)}/attachments?$select=id,name,size,contentType,isInline`)).value || [];
        const used = new Set();
        if (hasCid) {
          let body = html;
          for (const a of all.filter(a => /^image\//i.test(a.contentType || "") && a.size < 8 * 1024 * 1024 && (a.isInline || all.length < 15))) {
            try { const f = await MS.g(`/me/messages/${encodeURIComponent(id)}/attachments/${encodeURIComponent(a.id)}`);
              const cid = String(f.contentId || "").replace(/^<|>$/g, "");
              if (f.contentBytes && cid && body.includes("cid:" + cid)) { body = body.split("cid:" + cid).join(`data:${a.contentType};base64,${f.contentBytes}`); used.add(a.id); } } catch (e) {} }
          m.body = { ...m.body, content: body };
        }
        ATT[id] = all.filter(a => !used.has(a.id) && !a.isInline);
      }
    } catch (e) { MSG[id] = { _err: e.message }; }
    render();
  }
  const people = xs => (xs || []).map(r => r.emailAddress?.name || r.emailAddress?.address).filter(Boolean).join(", ");
  window.dPost = function (id) {
    if (!msOn()) { const L = linked(id); return backBar() + `<h2 class="big">${esc(L?.efni || "Póstur")}</h2><div class="empty">Tengdu Microsoft til að lesa póstinn í appinu.${L?.url ? `<br><a class="btn" style="margin-top:10px" href="${esc(L.url)}" target="_blank" rel="noopener">Opna í Outlook ↗</a>` : ""}</div>`; }
    const m = MSG[id]; if (!m) { loadMsg(id); return backBar() + `<div class="empty" style="margin-top:20px">Sæki póst…</div>`; }
    if (m._err) return backBar() + `<div class="msg err">${esc(m._err)}</div>`;
    const L = linked(id), p = L && findProj(L.verkefniId), g = !L && window.msGuessProj ? findProj(msGuessProj((m.subject || "") + " " + (m.bodyPreview || ""), m.from?.emailAddress?.address)) : null, atts = ATT[id] || [];
    const html = m.body?.contentType === "html" ? m.body.content : `<pre style="white-space:pre-wrap;font:inherit">${esc(m.body?.content || "")}</pre>`;
    const doc = `<!doctype html><html><head><meta charset="utf-8"><base target="_blank"><style>body{font:14px/1.5 -apple-system,Segoe UI,Inter,sans-serif;color:#14181c;background:#fff;margin:14px;word-wrap:break-word}img{max-width:100%;height:auto}table{max-width:100%}</style></head><body>${html}</body></html>`;
    return backBar() + `<div class="kicker"><span class="kind">Póstur</span>${p ? `<span class="pill">${esc(p.nr)}</span>` : ""}</div>
      <h2 class="big" style="margin-top:6px">${esc(m.subject || "(ekkert efni)")}</h2>
      <p class="sub" style="margin-bottom:4px"><b style="color:var(--ink)">${esc(fromName(m))}</b>${m.from?.emailAddress?.address ? ` &lt;${esc(m.from.emailAddress.address)}&gt;` : ""} · ${esc(fmtDate(m.receivedDateTime))} kl. ${esc(new Date(m.receivedDateTime).toTimeString().slice(0, 5))}</p>
      <p class="hint" style="margin:0 0 10px">Til: ${esc(people(m.toRecipients))}${m.ccRecipients?.length ? " · Afrit: " + esc(people(m.ccRecipients)) : ""}</p>
      <div class="links">${MCP ? `<button type="button" class="btn primary" data-act="poreply:reply:${esc(id)}">Svara</button><button type="button" class="btn" data-act="poreply:replyAll:${esc(id)}">Svara öllum</button><button type="button" class="btn" data-act="poreply:forward:${esc(id)}">Áframsenda</button>` : ""}<a class="btn" href="${esc(m.webLink)}" target="_blank" rel="noopener">Opna í Outlook ↗</a>
        ${MCP ? `<label class="posel"><select data-polink="${esc(id)}" aria-label="Tengja við verkefni"><option value="">${L ? "Aftengja frá verkefni" : "Tengja við verkefni…"}</option>${verkefni().slice().sort((a, b) => (isDone(a) - isDone(b)) || b.nr.localeCompare(a.nr)).map(x => `<option value="${esc(x.id)}"${L && L.verkefniId === x.id ? " selected" : ""}>${esc(x.nafn)}</option>`).join("")}</select></label>` : ""}
        ${g && MCP ? `<button type="button" class="btn" data-act="polink:${esc(id)}:${esc(g.id)}">Tengja við ${esc(g.nr)}</button>` : ""}</div>
      ${atts.length ? `<div class="viewtitle">Viðhengi · ${atts.length}</div><div class="llist">${atts.map(a => `<div class="lrow"><div class="ll">📎</div><div class="lm"><div class="lt">${esc(a.name)}</div><div class="ls">${esc(kb(a.size))}</div></div>
        <div class="lr" style="gap:6px"><button type="button" class="btn" data-act="poatt:${esc(id)}:${esc(a.id)}" style="padding:4px 10px;min-height:30px">Sækja</button>${p && MCP ? `<button type="button" class="btn" data-act="posave:${esc(id)}:${esc(a.id)}" style="padding:4px 10px;min-height:30px" title="Vista í ${esc(p.nr)} › 01 Móttekið">→ Mappa</button>` : ""}</div></div>`).join("")}</div>
        ${p ? `<p class="hint" style="margin-top:6px">„→ Mappa“ vistar viðhengið í ${esc(p.nr)} › 01 Móttekið í OneDrive.</p>` : `<p class="hint" style="margin-top:6px">Tengdu póstinn við verkefni til að geta vistað viðhengi í möppu þess.</p>`}` : ""}
      <iframe class="pobody" sandbox="allow-popups allow-popups-to-escape-sandbox" srcdoc="${esc(doc)}" title="Efni pósts"></iframe>`;
  };

  async function link(gid, pid) {
    const m = MSG[gid] || (list || []).find(x => x.id === gid);
    await saveSnapshot(s => { s.postar = s.postar || []; const q = s.postar.find(x => x.gid === gid);
      if (!pid) { s.postar = s.postar.filter(x => x.gid !== gid); return; }
      if (q) q.verkefniId = pid;
      else if (m) s.postar.push({ id: newId(), gid, efni: m.subject || "(ekkert efni)", fra: fromName(m), dags: m.receivedDateTime, url: m.webLink, brot: (m.bodyPreview || "").slice(0, 240), verkefniId: pid }); });
    toast(pid ? "Póstur tengdur við " + (findProj(pid)?.nr || "verkefni") : "Póstur aftengdur");
  }
  async function attBlob(mid, aid) {
    const a = await MS.g(`/me/messages/${encodeURIComponent(mid)}/attachments/${encodeURIComponent(aid)}`);
    if (!a.contentBytes) throw new Error("Þessa tegund viðhengis er ekki hægt að sækja hér. Opnaðu póstinn í Outlook.");
    const bin = atob(a.contentBytes), u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return { blob: new Blob([u], { type: a.contentType || "application/octet-stream" }), name: a.name };
  }


  /* ---------- Senda póst (Mail.Send) ---------- */
  const b64 = async data => { const u = data instanceof Uint8Array ? data : new Uint8Array(await data.arrayBuffer()); let s = ""; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  const addrs = v => String(v || "").split(/[,;]\s*/).map(x => x.trim()).filter(Boolean).map(x => { const m = x.match(/<([^>]+)>/); return { emailAddress: { address: (m ? m[1] : x).trim() } }; });
  const validA = v => addrs(v).every(a => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(a.emailAddress.address));
  const toHtml = t => `<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt">${esc(t).replace(/\n/g, "<br>")}</div>`;
  /* Netfangaskrá: allir sem póstur hefur borist frá eða verið sendur til (S.netfangaskra [{a,n,t,c}]) + Viðskiptavinir */
  const BOOK = () => S?.netfangaskra || [];
  const contacts = () => { const m = new Map();
    BOOK().forEach(x => m.set(x.a.toLowerCase(), { a: x.a, n: x.n || "", t: x.t || "", c: x.c || 0 }));
    vskList().forEach(c => [[c.netfang, c.name], [c.netfangT, c.tengilidur || c.name]].forEach(([a, n]) => { if (!a) return; const k = a.toLowerCase(), o = m.get(k); m.set(k, { a, n: o?.n || n || "", t: o?.t || "", c: (o?.c || 0) + 0.5, v: c.name }); }));
    return [...m.values()].sort((x, y) => (y.c - x.c) || String(y.t).localeCompare(String(x.t))); };
  // Bæta við/uppfæra í netfangaskrá (ein vistun)
  async function remember(list, sent) {
    const me = (MS.user()?.email || "").toLowerCase(), now = new Date().toISOString();
    const xs = list.filter(r => r?.a && /@/.test(r.a) && r.a.toLowerCase() !== me && !/no-?reply|noreply|mailer-daemon|notifications?@/i.test(r.a));
    if (!xs.length) return;
    let changed = false;
    const book = BOOK().map(x => ({ ...x })), idx = new Map(book.map((x, i) => [x.a.toLowerCase(), i]));
    xs.forEach(r => { const k = r.a.toLowerCase(), i = idx.get(k);
      if (i == null) { book.push({ a: r.a, n: r.n || "", t: r.t || now, c: sent ? 1 : 0.2 }); idx.set(k, book.length - 1); changed = true; }
      else { const o = book[i]; if (r.n && !o.n) { o.n = r.n; changed = true; } if (sent) { o.c = (o.c || 0) + 1; o.t = now; changed = true; } else if (r.t && r.t > (o.t || "")) { o.t = r.t; changed = true; } } });
    if (!changed) return;
    book.sort((x, y) => (y.c - x.c) || String(y.t).localeCompare(String(x.t)));
    await saveSnapshot(st => { st.netfangaskra = book.slice(0, 800); });
  }
  // Safna netföngum úr Sent og Innhólfi (einu sinni á dag á hverju tæki)
  async function harvest(force) {
    let last = 0; try { last = +localStorage.getItem("aera-harvest") || 0; } catch (e) {}
    if (!force && Date.now() - last < 20 * 3600 * 1000) return;
    try { localStorage.setItem("aera-harvest", String(Date.now())); } catch (e) {}
    const r = [];
    const sent = (await MS.g(`/me/mailFolders/sentitems/messages?$top=200&$orderby=sentDateTime%20desc&$select=toRecipients,ccRecipients,sentDateTime`)).value || [];
    sent.forEach(m => [...(m.toRecipients || []), ...(m.ccRecipients || [])].forEach(x => r.push({ a: x.emailAddress?.address, n: x.emailAddress?.name, t: m.sentDateTime, s: 1 })));
    const inb = (await MS.g(`/me/mailFolders/inbox/messages?$top=200&$orderby=receivedDateTime%20desc&$select=from,receivedDateTime`)).value || [];
    inb.forEach(m => r.push({ a: m.from?.emailAddress?.address, n: m.from?.emailAddress?.name, t: m.receivedDateTime }));
    // telja sendingar sem vægi
    const cnt = new Map(); r.filter(x => x.s && x.a).forEach(x => cnt.set(x.a.toLowerCase(), (cnt.get(x.a.toLowerCase()) || 0) + 1));
    const me = (MS.user()?.email || "").toLowerCase(), book = BOOK().map(x => ({ ...x })), idx = new Map(book.map((x, i) => [x.a.toLowerCase(), i]));
    r.forEach(x => { if (!x.a || !/@/.test(x.a) || x.a.toLowerCase() === me || /no-?reply|noreply|mailer-daemon|notifications?@/i.test(x.a)) return; const k = x.a.toLowerCase(), i = idx.get(k);
      if (i == null) { book.push({ a: x.a, n: x.n || "", t: x.t, c: (cnt.get(k) || 0) + 0.2 }); idx.set(k, book.length - 1); }
      else { const o = book[i]; if (!o.n && x.n) o.n = x.n; if ((x.t || "") > (o.t || "")) o.t = x.t; if (force || !o._h) { o.c = Math.max(o.c || 0, (cnt.get(k) || 0) + 0.2); } } });
    book.forEach(o => delete o._h); book.sort((x, y) => (y.c - x.c) || String(y.t).localeCompare(String(x.t)));
    await saveSnapshot(st => { st.netfangaskra = book.slice(0, 800); });
  }
  window.poHarvest = harvest;

  /* Viðtakendareitur eins og í Outlook: nöfn sem flísar, tillögur um leið og skrifað er */
  function rcptField(id, init) {
    const box = document.getElementById(id + "-box"), hid = document.getElementById(id), inp = box.querySelector("input"), sug = document.getElementById(id + "-sug");
    let chips = addrs(init).map(x => ({ a: x.emailAddress.address, n: (contacts().find(c => c.a.toLowerCase() === x.emailAddress.address.toLowerCase()) || {}).n || "" })), act = -1, cur = [];
    const sync = () => { hid.value = chips.map(c => c.a).join(", "); box.querySelectorAll(".rchip").forEach(e => e.remove());
      chips.forEach((c, i) => inp.insertAdjacentHTML("beforebegin", `<span class="rchip${/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(c.a) ? "" : " bad"}" title="${esc(c.a)}">${esc(c.n || c.a)}<button type="button" data-ri="${i}" aria-label="Fjarlægja ${esc(c.a)}">✕</button></span>`));
      box.querySelectorAll("[data-ri]").forEach(b => b.onclick = e => { e.stopPropagation(); chips.splice(+b.dataset.ri, 1); sync(); inp.focus(); }); };
    const add = (a, n) => { a = a.trim().replace(/[,;]+$/, ""); if (!a) return; if (!chips.some(c => c.a.toLowerCase() === a.toLowerCase())) chips.push({ a, n: n || (contacts().find(c => c.a.toLowerCase() === a.toLowerCase()) || {}).n || "" }); inp.value = ""; hide(); sync(); };
    const hide = () => { sug.hidden = true; act = -1; };
    const show = () => { const t = inp.value.trim().toLowerCase(); if (!t) return hide();
      cur = contacts().filter(c => !chips.some(x => x.a.toLowerCase() === c.a.toLowerCase()) && (c.a.toLowerCase().includes(t) || (c.n || "").toLowerCase().split(/\s+/).some(w => w.startsWith(t)) || (c.n || "").toLowerCase().startsWith(t) || (c.v || "").toLowerCase().startsWith(t))).slice(0, 7);
      if (!cur.length) return hide(); act = 0;
      sug.innerHTML = cur.map((c, i) => `<button type="button" class="rsug${i === act ? " on" : ""}" data-si="${i}"><b>${esc(c.n || c.a)}</b>${c.n ? `<small>${esc(c.a)}</small>` : ""}${c.v && c.v !== c.n ? `<small>${esc(c.v)}</small>` : ""}</button>`).join(""); sug.hidden = false;
      sug.querySelectorAll("[data-si]").forEach(b => b.onmousedown = e => { e.preventDefault(); const c = cur[+b.dataset.si]; add(c.a, c.n); inp.focus(); }); };
    const mark = () => sug.querySelectorAll(".rsug").forEach((b, i) => b.classList.toggle("on", i === act));
    inp.addEventListener("input", () => { if (/[,;]\s*$/.test(inp.value)) return add(inp.value); show(); });
    inp.addEventListener("keydown", e => {
      if (!sug.hidden && (e.key === "ArrowDown" || e.key === "ArrowUp")) { e.preventDefault(); act = (act + (e.key === "ArrowDown" ? 1 : -1) + cur.length) % cur.length; mark(); return; }
      if (e.key === "Enter" || e.key === "Tab") { if (!sug.hidden && cur[act]) { e.preventDefault(); add(cur[act].a, cur[act].n); return; } if (inp.value.trim()) { if (e.key === "Enter") e.preventDefault(); add(inp.value); } return; }
      if (e.key === "Backspace" && !inp.value && chips.length) { chips.pop(); sync(); }
      if (e.key === "Escape") hide(); });
    inp.addEventListener("blur", () => setTimeout(() => { if (inp.value.trim()) add(inp.value); hide(); }, 120));
    box.addEventListener("click", () => inp.focus());
    sync();
    return { flush: () => { if (inp.value.trim()) add(inp.value); } };
  }

  /* Undirskrift (íslenska og enska) – stillingar.undirskrift */
  const SIG = () => ({ kvedja: "Kveðja / Best regards", nafn: MS.user()?.name || "Dagur Fannar Möller", titillIS: "Eigandi", titillEN: "Owner", fyrirtaeki: "AERA Studio", netfang: "aerastudioice@gmail.com", simi: "", vefur: "", ...(CFG().undirskrift || {}) });
  const LOGO_URL = "https://aerastudio.github.io/aera-app/icons/sig-logo.png", LOGO_CID = "aeralogo@aerastudio";
  window.poSigHTML = (logo, o) => { const g = { ...SIG(), ...(o || {}) }, E = x => esc(x || ""), web = String(g.vefur || "").replace(/^https?:\/\//, "");
    return `<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;font-family:Calibri,Arial,Helvetica,sans-serif;color:#1c2126">
${g.kvedja ? `<tr><td colspan="2" style="padding:0 0 14px 0;font-size:10.5pt;color:#1c2126">${E(g.kvedja)},</td></tr>` : ""}
<tr><td style="padding:0 16px 0 0;vertical-align:middle"><img src="${logo}" width="140" height="49" alt="AERA Studio" style="display:block;border:0;width:140px;height:49px"></td>
<td style="border-left:2px solid #1f4e79;padding:2px 0 2px 16px;vertical-align:middle;font-size:10pt;line-height:15pt">
<span style="font-size:11.5pt;font-weight:bold;color:#0b1e33">${E(g.nafn)}</span><br>
<span style="color:#5b7fa6">${[g.titillIS, g.titillEN].filter(Boolean).map(E).join(" &middot; ")}</span><br>
<span style="color:#1c2126">${E(g.fyrirtaeki)}</span><br>
${g.netfang ? `<a href="mailto:${E(g.netfang)}" style="color:#1f4e79;text-decoration:none">${E(g.netfang)}</a>` : ""}${g.simi ? `<br><span style="color:#1c2126">S. / Tel. ${E(g.simi)}</span>` : ""}${web ? `<br><a href="https://${E(web)}" style="color:#1f4e79;text-decoration:none">${E(web)}</a>` : ""}
</td></tr></table>`; };
  let logoB64 = null;
  const logoAtt = async () => { if (!logoB64) logoB64 = await b64(await (await fetch("icons/sig-logo.png")).blob());
    return { "@odata.type": "#microsoft.graph.fileAttachment", name: "aera-logo.png", contentType: "image/png", contentBytes: logoB64, isInline: true, contentId: LOGO_CID }; };
  window.poSigCopy = async () => { const h = poSigHTML(LOGO_URL);
    try { await navigator.clipboard.write([new ClipboardItem({ "text/html": new Blob([h], { type: "text/html" }), "text/plain": new Blob([document.createRange().createContextualFragment(h).textContent.replace(/\n\s*/g, "\n").trim()], { type: "text/plain" }) })]); toast("Undirskrift afrituð – límdu hana í undirskriftastillingar Outlook"); }
    catch (e) { toast("Gat ekki afritað – veldu undirskriftina og afritaðu handvirkt."); } };
  const sig = () => "\n\n";
  window.poSigSettingsHTML = () => { const g = SIG(), f = (k, l, ex = "") => `<label>${l}<input id="sg-${k}" value="${esc(g[k] || "")}" ${ex}></label>`;
    return `<div class="viewtitle">Undirskrift í pósti</div>
      <p class="hint" style="margin:0 0 8px">Birtist neðst í póstum sem eru sendir úr appinu, á íslensku og ensku, með merki AERA.</p>
      <div class="sigprev" id="sg-prev">${poSigHTML("icons/sig-logo.png")}</div>
      ${f("kvedja", "Kveðja")}${f("nafn", "Nafn")}
      <div class="two">${f("titillIS", "Starfsheiti (íslenska)")}${f("titillEN", "Starfsheiti (enska)")}</div>
      ${f("fyrirtaeki", "Fyrirtæki")}
      <div class="two">${f("netfang", "Netfang", 'type="email"')}${f("simi", "Sími (valfrjálst)", 'inputmode="tel" placeholder="+354 …"')}</div>
      ${f("vefur", "Vefsíða (valfrjálst)", 'placeholder="aerastudio.is"')}
      <div class="links"><button type="button" class="btn" data-act="posigsave">Vista undirskrift</button><button type="button" class="btn" data-act="posigcopy">Afrita fyrir Outlook</button></div>
      <p class="hint" style="margin-top:6px">„Afrita fyrir Outlook“ setur undirskriftina á klippiborðið. Límdu hana svo í Outlook › Stillingar › Undirskriftir til að nota hana líka þar.</p>`; };
  const sgRead = () => { const o = {}; ["kvedja", "nafn", "titillIS", "titillEN", "fyrirtaeki", "netfang", "simi", "vefur"].forEach(k => { const el = document.getElementById("sg-" + k); if (el) o[k] = el.value.trim(); }); return o; };
  document.addEventListener("input", e => { if (e.target.id?.startsWith("sg-")) { const p = document.getElementById("sg-prev"); if (p) p.innerHTML = poSigHTML("icons/sig-logo.png", sgRead()); } });

  window.poCompose = function (o = {}) {
    if (!msOn()) { toast("Tengdu Microsoft fyrst (Stillingar)."); return; }
    const mode = o.mode || "new", src = o.mid ? MSG[o.mid] : null, atts = (o.atts || []).slice();
    const title = { new: "Nýr póstur", reply: "Svara", replyAll: "Svara öllum", forward: "Áframsenda" }[mode];
    const canSend = MS.can("Mail.Send");
    const pid0 = o.pid || (o.mid && linked(o.mid)?.verkefniId) || "";
    const subj = o.subject != null ? o.subject : src ? ((mode === "forward" ? "FW: " : "RE: ") + (src.subject || "").replace(/^(re|fw|fwd|sv|áfr):\s*/i, "")) : "";
    sheet.innerHTML = `<h3 id="sheet-h">${title}</h3>
      ${canSend ? "" : `<div class="msg err" style="margin-bottom:10px">Appið hefur ekki enn leyfi til að senda póst. Tengstu Microsoft aftur og samþykktu nýju heimildina.<br><button type="button" class="btn" style="margin-top:8px" id="po-relog">Tengjast aftur</button></div>`}
      <form id="pocf" novalidate>
      ${mode === "reply" || mode === "replyAll" ? `<p class="hint" style="margin:0">Til: ${esc(mode === "reply" ? fromName(src) : [fromName(src), people(src.toRecipients), people(src.ccRecipients)].filter(Boolean).join(", "))}</p>` :
        `<div class="rfield"><span class="lbl">Til <span class="req">*</span></span><div class="rbox" id="po-to-box"><input autocomplete="off" inputmode="email" aria-label="Viðtakendur" placeholder="Nafn eða netfang"></div><div class="rsugs" id="po-to-sug" hidden></div><input type="hidden" id="po-to"></div>
         <div class="rfield"><span class="lbl">Afrit (cc)</span><div class="rbox" id="po-cc-box"><input autocomplete="off" inputmode="email" aria-label="Afrit"></div><div class="rsugs" id="po-cc-sug" hidden></div><input type="hidden" id="po-cc"></div>`}
      ${mode === "new" ? `<label><span class="lbl">Efni <span class="req">*</span></span><input id="po-subj" value="${esc(subj)}" autocomplete="off"></label>` : `<p class="hint" style="margin:0">Efni: ${esc(subj)}</p>`}
      <label>Texti<textarea id="po-body" rows="8">${esc(o.body != null ? o.body : sig())}</textarea></label>
      <label class="chk" style="flex-direction:row;align-items:center;gap:8px;color:var(--ink);font-size:14.5px"><input type="checkbox" id="po-sig" checked style="width:auto">Undirskrift</label>
      <div class="sigprev" id="po-sigprev">${poSigHTML("icons/sig-logo.png")}</div>
      <div><div class="kind" style="margin-bottom:6px">Viðhengi</div><div id="po-atts" class="poatts"></div>
        <label class="btn" style="display:inline-flex;margin-top:6px;cursor:pointer">📎 Bæta við skrá<input id="po-file" type="file" multiple hidden></label></div>
      ${mode === "new" ? `<label>Tengja við verkefni<select id="po-proj"><option value="">Ekkert</option>${verkefni().slice().sort((a, b) => (isDone(a) - isDone(b)) || b.nr.localeCompare(a.nr)).map(x => `<option value="${esc(x.id)}"${x.id === pid0 ? " selected" : ""}>${esc(x.nafn)}</option>`).join("")}</select></label>` : ""}
      ${src ? `<p class="hint">Upprunalegi pósturinn fylgir neðst${mode === "forward" && src.hasAttachments ? ", ásamt viðhengjum hans" : ""}.</p>` : ""}
      <div id="f-msg"></div><div class="actions"><button type="button" class="btn" id="po-c">Hætta við</button><button class="btn primary" id="po-s"${canSend ? "" : " disabled"}>Senda</button></div></form>`;
    sheet.closest(".sheet").classList.add("wide"); sheetwrap.hidden = false;
    const drawAtts = () => { document.getElementById("po-atts").innerHTML = atts.length ? atts.map((a, i) => `<span class="chipbtn" style="display:inline-flex;gap:6px;align-items:center">📎 ${esc(a.name)} <small>${esc(kb(a.size))}</small><button type="button" class="x" data-rm="${i}" aria-label="Fjarlægja">✕</button></span>`).join(" ") : `<span class="hint">Engin.</span>`;
      sheet.querySelectorAll("[data-rm]").forEach(b => b.onclick = () => { atts.splice(+b.dataset.rm, 1); drawAtts(); }); };
    drawAtts();
    { const sc = document.getElementById("po-sig"), sp = document.getElementById("po-sigprev"); if (sc) sc.onchange = () => { sp.hidden = !sc.checked; }; }
    const RF = ["po-to", "po-cc"].filter(x => document.getElementById(x + "-box")).map(x => rcptField(x, x === "po-to" ? o.to || "" : o.cc || ""));
    const rl = document.getElementById("po-relog"); if (rl) rl.onclick = () => MS.login().then(() => { toast("Tengt – nú má senda"); poCompose(o); }).catch(e => toast(e.message));
    document.getElementById("po-file").onchange = e => { [...e.target.files].forEach(f => atts.push({ name: f.name, type: f.type || "application/octet-stream", size: f.size, data: f })); e.target.value = ""; drawAtts(); };
    document.getElementById("po-c").onclick = closeSheet;
    const bt = document.getElementById("po-body"); if (!("ontouchstart" in window)) setTimeout(() => { (document.getElementById("po-to-box") && !o.to ? document.querySelector("#po-to-box input") : bt).focus(); if (document.activeElement === bt) bt.setSelectionRange(0, 0); }, 60);
    document.getElementById("pocf").onsubmit = async ev => {
      ev.preventDefault(); msg(""); RF.forEach(f => f.flush());
      const to = val("po-to"), cc = val("po-cc"), subject = mode === "new" ? val("po-subj") : subj, text = val("po-body"), pid = mode === "new" ? val("po-proj") : pid0;
      if (mode === "new" || mode === "forward") { if (!to) return msg("Settu inn viðtakanda.", "err"); if (!validA(to) || (cc && !validA(cc))) return msg("Athugaðu netföngin.", "err"); }
      if (mode === "new" && !subject) return msg("Skrifaðu efni.", "err");
      if (atts.some(a => a.size > 3 * 1024 * 1024)) return msg("Viðhengi mega mest vera 3 MB hvert. Deildu stærri skrám með tengli úr OneDrive.", "err");
      const btn = document.getElementById("po-s"); btn.disabled = true; btn.textContent = "Sendi…";
      try {
        const fa = await Promise.all(atts.map(async a => ({ "@odata.type": "#microsoft.graph.fileAttachment", name: a.name, contentType: a.type, contentBytes: await b64(a.data) })));
        const withSig = document.getElementById("po-sig")?.checked, sigH = withSig ? "<br>" + poSigHTML("cid:" + LOGO_CID) : "";
        if (withSig) fa.push(await logoAtt());
        if (mode === "new") {
          await MS.g("/me/sendMail", { method: "POST", body: { message: { subject, body: { contentType: "HTML", content: toHtml(text) + sigH }, toRecipients: addrs(to), ccRecipients: addrs(cc), attachments: fa }, saveToSentItems: true } });
        } else {
          const fn = { reply: "createReply", replyAll: "createReplyAll", forward: "createForward" }[mode];
          const dr = await MS.g(`/me/messages/${encodeURIComponent(o.mid)}/${fn}`, { method: "POST", body: {} });
          const patch = { body: { contentType: "HTML", content: toHtml(text) + sigH + "<br>" + (dr.body?.content || "") } };
          if (mode === "forward") { patch.toRecipients = addrs(to); if (cc) patch.ccRecipients = addrs(cc); }
          await MS.g(`/me/messages/${encodeURIComponent(dr.id)}`, { method: "PATCH", body: patch });
          for (const a of fa) await MS.g(`/me/messages/${encodeURIComponent(dr.id)}/attachments`, { method: "POST", body: a });
          await MS.g(`/me/messages/${encodeURIComponent(dr.id)}/send`, { method: "POST", body: {} });
        }
        toast("Póstur sendur"); closeSheet();
        const rc = mode === "reply" ? [{ a: src.from?.emailAddress?.address, n: src.from?.emailAddress?.name }] : mode === "replyAll" ? [src.from, ...(src.toRecipients || []), ...(src.ccRecipients || [])].map(x => ({ a: x?.emailAddress?.address, n: x?.emailAddress?.name })) : [...addrs(to), ...addrs(cc)].map(x => ({ a: x.emailAddress.address, n: (contacts().find(c => c.a.toLowerCase() === x.emailAddress.address.toLowerCase()) || {}).n || "" }));
        remember(rc, true).catch(() => {});
        if (o.onSent) try { await o.onSent(); } catch (e) {}
        if (pid) setTimeout(() => linkSent(subject, pid), 3000);
      } catch (e) { btn.disabled = false; btn.textContent = "Senda"; msg(esc(/403|denied|scope/i.test(e.message) ? "Appið hefur ekki leyfi til að senda. Tengstu Microsoft aftur og samþykktu heimildina." : e.message), "err"); }
    };
  };
  // Eftir sendingu: finna sendan póst og tengja hann við verkefnið
  async function linkSent(subject, pid) {
    try { const xs = (await MS.g(`/me/mailFolders/sentitems/messages?$top=10&$orderby=sentDateTime%20desc&$select=id,subject,toRecipients,sentDateTime,bodyPreview,webLink`)).value;
      const m = xs.find(x => (x.subject || "") === subject); if (!m || linked(m.id)) return;
      await saveSnapshot(s => { (s.postar = s.postar || []).push({ id: newId(), gid: m.id, efni: m.subject, fra: "Ég → " + people(m.toRecipients), dags: m.sentDateTime, url: m.webLink, brot: (m.bodyPreview || "").slice(0, 240), verkefniId: pid, sendur: true }); });
    } catch (e) {}
  }

  window.poAct = function (k, r) {
    if (k === "posigsave") { const o = sgRead(); saveSnapshot(st => { st.stillingar = st.stillingar || {}; st.stillingar.undirskrift = o; }).then(() => toast("Undirskrift vistuð")); return; }
    if (k === "posigcopy") { poSigCopy(); return; }
    if (k === "ponew") { poCompose({ pid: r[0] || null, subject: r[0] && findProj(r[0]) ? findProj(r[0]).nr + " " + findProj(r[0]).stutt + " – " : "" }); return; }
    if (k === "poreply") { poCompose({ mode: r[0], mid: r.slice(1).join(":") }); return; }
    if (k === "pofold") { fold = r[0]; q = ""; reset(); render(); return; }
    if (k === "poharv") { toast("Sæki netföng úr pósti…"); harvest(true).then(() => toast("Tengiliðir uppfærðir")).catch(e => toast(e.message)); return; }
    if (k === "poto") { poCompose({ to: r.join(":") }); return; }
    if (k === "pofor") { const a = r.join(":").toLowerCase(); if (confirm("Gleyma " + a + "?")) saveSnapshot(st => { st.netfangaskra = (st.netfangaskra || []).filter(x => x.a.toLowerCase() !== a); }); return; }
    if (k === "poref") { reset(); render(); return; }
    if (k === "pomore") { load(true); render(); return; }
    if (k === "potarget") { target = r[0] || null; render(); return; }
    if (k === "polink") { link(r[0], r[1]); return; }
    if (k === "pofind") { target = r[0]; const p = findProj(target); q = p ? p.stutt : ""; reset(); DETAIL = null; tab = "postur"; render(); window.scrollTo({ top: 0 }); return; }
    if (k === "poatt") { toast("Sæki viðhengi…"); attBlob(r[0], r[1]).then(({ blob, name }) => { const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000); }).catch(e => toast(e.message)); return; }
    if (k === "posave") { const L = linked(r[0]), p = L && findProj(L.verkefniId); if (!p) return toast("Tengdu póstinn við verkefni fyrst.");
      toast("Vista í " + p.nr + "…");
      attBlob(r[0], r[1]).then(async ({ blob, name }) => { const dir = await msProjSub(p, "01", "01 Móttekið"); await msSaveFile(dir, name, blob, blob.type, false); toast("Vistað í " + dir.replace(/^AERA\/Verkefni\//, "").replace(/\//g, " › ")); }).catch(e => toast(e.message)); return; }
  };
  document.addEventListener("submit", e => { if (e.target.id === "poqi" || e.target.id === "poq") { e.preventDefault(); q = (document.getElementById("poqi")?.value || "").trim(); reset(); render(); } });
  document.addEventListener("input", e => { if (e.target.id === "poqi" && fold === "tengilidir") { q = e.target.value.trim(); const pos = e.target.selectionStart; render(); const n = document.getElementById("poqi"); if (n) { n.focus(); try { n.setSelectionRange(pos, pos); } catch (x) {} } } });
  document.addEventListener("change", e => { const id = e.target.dataset?.polink; if (id != null) link(id, e.target.value || null); });
  document.addEventListener("search", e => { if (e.target.id === "poqi" && !e.target.value && q) { q = ""; reset(); render(); } }, true);

  // Póstar á síðu verkefnis
  window.poProjSection = function (p) {
    const ms = (S?.postar || []).filter(m => m.verkefniId === p.id).sort((a, b) => String(b.dags).localeCompare(String(a.dags)));
    const rowP = m => m.gid && msOn() ? `<div class="lrow tap fl" ${tapAttr("post", m.gid)}><div class="ll">✉️</div><div class="lm"><div class="lt">${esc(m.efni)}</div><div class="ls">${esc([m.fra, fmtDate(m.dags)].filter(Boolean).join(" · "))}${m.brot ? " · " + esc(m.brot.slice(0, 80)) : ""}</div></div>${CHEV}</div>`
      : `<a class="lrow fl" href="${esc(m.url || "#")}" target="_blank" rel="noopener"><div class="ll">✉️</div><div class="lm"><div class="lt">${esc(m.efni)}</div><div class="ls">${esc([m.fra, fmtDate(m.dags)].filter(Boolean).join(" · "))}${m.brot ? " · " + esc(m.brot.slice(0, 80)) : ""}</div></div></a>`;
    return `<div class="viewtitle" style="display:flex;align-items:center;gap:8px">Póstar${ms.length ? " · " + ms.length : ""}${MCP && msOn() ? `<a class="plink" data-act="pofind:${esc(p.id)}" style="margin-left:auto">Finna pósta</a><a class="plink" data-act="ponew:${esc(p.id)}">+ Nýr póstur</a>` : ""}</div>${ms.length ? `<div class="llist">${ms.map(rowP).join("")}</div>` : msOn() ? `<div class="hint" style="margin:0">Engir póstar tengdir ennþá.</div>` : ""}`;
  };
  window.poUnread = () => (list || []).filter(m => m.isRead === false).length;
})();
