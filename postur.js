/* Pósthólf í appinu (Microsoft Graph, Mail.Read): lesa innhólf og möppur, leita, opna póst,
   tengja póst við verkefni (S.postar) og vista viðhengi í möppu verkefnisins. */
(function () {
  const SEL = "$select=id,subject,from,receivedDateTime,bodyPreview,webLink,isRead,hasAttachments,conversationId";
  const FOLD = [["inbox", "Innhólf"], ["aera", "AERA"], ["sentitems", "Sent"]];
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

  window.vPostur = function () {
    if (!msOn()) return `<div class="empty" style="margin-top:16px">Tengdu Microsoft til að sjá pósthólfið þitt í appinu.${window.msMissing ? msMissing() : ""}</div>`;
    if (!list && !loading) load(false);
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
      if (MSG[id].hasAttachments) ATT[id] = (await MS.g(`/me/messages/${encodeURIComponent(id)}/attachments?$select=id,name,size,contentType,isInline`)).value.filter(a => !a.isInline);
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
  const contacts = () => { const set = new Map(); vskList().forEach(c => { if (c.netfang) set.set(c.netfang, c.name); if (c.netfangT) set.set(c.netfangT, c.tengilidur || c.name); });
    (list || []).forEach(m => { const a = m.from?.emailAddress; if (a?.address && !set.has(a.address)) set.set(a.address, a.name || ""); }); return [...set.entries()]; };
  const sig = () => `\n\nKveðja,\n${MS.user()?.name || ""}\nAERA Studio`;

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
        `<label><span class="lbl">Til <span class="req">*</span></span><input id="po-to" list="po-dl" value="${esc(o.to || "")}" placeholder="netfang@dæmi.is, …" autocomplete="off" inputmode="email"></label>
         <label>Afrit (cc)<input id="po-cc" list="po-dl" value="${esc(o.cc || "")}" autocomplete="off" inputmode="email"></label>
         <datalist id="po-dl">${contacts().map(([a, n]) => `<option value="${esc(a)}">${esc(n)}</option>`).join("")}</datalist>`}
      ${mode === "new" ? `<label><span class="lbl">Efni <span class="req">*</span></span><input id="po-subj" value="${esc(subj)}" autocomplete="off"></label>` : `<p class="hint" style="margin:0">Efni: ${esc(subj)}</p>`}
      <label>Texti<textarea id="po-body" rows="9">${esc(o.body != null ? o.body : sig())}</textarea></label>
      <div><div class="kind" style="margin-bottom:6px">Viðhengi</div><div id="po-atts" class="poatts"></div>
        <label class="btn" style="display:inline-flex;margin-top:6px;cursor:pointer">📎 Bæta við skrá<input id="po-file" type="file" multiple hidden></label></div>
      ${mode === "new" ? `<label>Tengja við verkefni<select id="po-proj"><option value="">Ekkert</option>${verkefni().slice().sort((a, b) => (isDone(a) - isDone(b)) || b.nr.localeCompare(a.nr)).map(x => `<option value="${esc(x.id)}"${x.id === pid0 ? " selected" : ""}>${esc(x.nafn)}</option>`).join("")}</select></label>` : ""}
      ${src ? `<p class="hint">Upprunalegi pósturinn fylgir neðst${mode === "forward" && src.hasAttachments ? ", ásamt viðhengjum hans" : ""}.</p>` : ""}
      <div id="f-msg"></div><div class="actions"><button type="button" class="btn" id="po-c">Hætta við</button><button class="btn primary" id="po-s"${canSend ? "" : " disabled"}>Senda</button></div></form>`;
    sheet.closest(".sheet").classList.add("wide"); sheetwrap.hidden = false;
    const drawAtts = () => { document.getElementById("po-atts").innerHTML = atts.length ? atts.map((a, i) => `<span class="chipbtn" style="display:inline-flex;gap:6px;align-items:center">📎 ${esc(a.name)} <small>${esc(kb(a.size))}</small><button type="button" class="x" data-rm="${i}" aria-label="Fjarlægja">✕</button></span>`).join(" ") : `<span class="hint">Engin.</span>`;
      sheet.querySelectorAll("[data-rm]").forEach(b => b.onclick = () => { atts.splice(+b.dataset.rm, 1); drawAtts(); }); };
    drawAtts();
    const rl = document.getElementById("po-relog"); if (rl) rl.onclick = () => MS.login().then(() => { toast("Tengt – nú má senda"); poCompose(o); }).catch(e => toast(e.message));
    document.getElementById("po-file").onchange = e => { [...e.target.files].forEach(f => atts.push({ name: f.name, type: f.type || "application/octet-stream", size: f.size, data: f })); e.target.value = ""; drawAtts(); };
    document.getElementById("po-c").onclick = closeSheet;
    const bt = document.getElementById("po-body"); if (!("ontouchstart" in window)) setTimeout(() => { (document.getElementById("po-to") && !o.to ? document.getElementById("po-to") : bt).focus(); if (document.activeElement === bt) bt.setSelectionRange(0, 0); }, 60);
    document.getElementById("pocf").onsubmit = async ev => {
      ev.preventDefault(); msg("");
      const to = val("po-to"), cc = val("po-cc"), subject = mode === "new" ? val("po-subj") : subj, text = val("po-body"), pid = mode === "new" ? val("po-proj") : pid0;
      if (mode === "new" || mode === "forward") { if (!to) return msg("Settu inn viðtakanda.", "err"); if (!validA(to) || (cc && !validA(cc))) return msg("Athugaðu netföngin.", "err"); }
      if (mode === "new" && !subject) return msg("Skrifaðu efni.", "err");
      if (atts.some(a => a.size > 3 * 1024 * 1024)) return msg("Viðhengi mega mest vera 3 MB hvert. Deildu stærri skrám með tengli úr OneDrive.", "err");
      const btn = document.getElementById("po-s"); btn.disabled = true; btn.textContent = "Sendi…";
      try {
        const fa = await Promise.all(atts.map(async a => ({ "@odata.type": "#microsoft.graph.fileAttachment", name: a.name, contentType: a.type, contentBytes: await b64(a.data) })));
        if (mode === "new") {
          await MS.g("/me/sendMail", { method: "POST", body: { message: { subject, body: { contentType: "HTML", content: toHtml(text) }, toRecipients: addrs(to), ccRecipients: addrs(cc), attachments: fa }, saveToSentItems: true } });
        } else {
          const fn = { reply: "createReply", replyAll: "createReplyAll", forward: "createForward" }[mode];
          const dr = await MS.g(`/me/messages/${encodeURIComponent(o.mid)}/${fn}`, { method: "POST", body: {} });
          const patch = { body: { contentType: "HTML", content: toHtml(text) + "<br>" + (dr.body?.content || "") } };
          if (mode === "forward") { patch.toRecipients = addrs(to); if (cc) patch.ccRecipients = addrs(cc); }
          await MS.g(`/me/messages/${encodeURIComponent(dr.id)}`, { method: "PATCH", body: patch });
          for (const a of fa) await MS.g(`/me/messages/${encodeURIComponent(dr.id)}/attachments`, { method: "POST", body: a });
          await MS.g(`/me/messages/${encodeURIComponent(dr.id)}/send`, { method: "POST", body: {} });
        }
        toast("Póstur sendur"); closeSheet();
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
    if (k === "ponew") { poCompose({ pid: r[0] || null, subject: r[0] && findProj(r[0]) ? findProj(r[0]).nr + " " + findProj(r[0]).stutt + " – " : "" }); return; }
    if (k === "poreply") { poCompose({ mode: r[0], mid: r.slice(1).join(":") }); return; }
    if (k === "pofold") { fold = r[0]; q = ""; reset(); render(); return; }
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
