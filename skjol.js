/* Skjöl: búa til Word, Excel, PowerPoint (í OneDrive), Canva-hönnun eða tengil, og flokka á réttan stað.
   S.skjol[] = {id, name, snid: word|excel|powerpoint|canva|tengill, url, odId, stadur: oflokkad|stofa|verk, projectId, mappa, stofnad, breytt} */
(function () {
  const SN = {
    word: { l: "Word", ext: "docx", b: "W", c: "#1f4e79" }, excel: { l: "Excel", ext: "xlsx", b: "X", c: "#2e7da6" },
    powerpoint: { l: "PowerPoint", ext: "pptx", b: "P", c: "#3f5f99" }, canva: { l: "Canva", b: "C", c: "#4f8fc9" }, tengill: { l: "Tengill", b: "🔗", c: "#5b7fa6" }
  };
  const OFFICE = k => !!SN[k]?.ext;
  const STD = ["01 Móttekið", "02 Vinnuskjöl", "03 Módel", "04 Teikningar", "05 Lýsing", "06 3D Render", "07 Útgefin Gögn", "08 Samskipti og fundir"];
  const STOFA0 = ["Lógó og útlit", "Reikningar", "Samningar", "Sniðmát", "Tímaskýrslur"];
  let STOFA = STOFA0.slice(), stofaLoaded = false;
  const docs = () => (S?.skjol || []).filter(x => x.snid);
  const findD = id => docs().find(x => x.id === id);
  const msOn = () => !!(window.MS && MS.on());
  let fSnid = "", fStad = "";

  const badge = x => { const s = SN[x.snid] || SN.tengill; return `<span class="dbadge" style="--dc:${s.c}">${s.b}</span>`; };
  const where = x => x.stadur === "stofa" ? "Stofa" + (x.mappa ? " › " + x.mappa : "") : x.stadur === "verk" ? ((findProj(x.projectId)?.nr || "Verkefni") + (x.mappa ? " › " + x.mappa : "")) : "Óflokkað";
  const dirOf = async (stadur, pid, mappa) => {
    if (stadur === "stofa") return "AERA/Stofa/" + (mappa || "Skjöl");
    if (stadur === "verk") { const p = findProj(pid); const pre = (mappa || "02").slice(0, 2); return await msProjSub(p, /^\d\d$/.test(pre) ? pre : "__", mappa || "02 Vinnuskjöl"); }
    return "AERA/Verkefni/Óflokkað";
  };
  const row = x => lrow("skjal:" + x.id, { lead: badge(x), title: esc(x.name), sub: esc([where(x), fmtDate(x.breytt || x.stofnad)].filter(Boolean).join(" · ")), right: x.stadur === "oflokkad" && MCP ? `<button type="button" class="btn" data-act="skflokka:${esc(x.id)}" style="padding:6px 12px;min-height:34px">Flokka</button>` : "" });

  /* ---------- Listasíða ---------- */
  window.vSkjolNy = function () {
    const all = docs(), chip = (v, l) => `<button type="button" class="chipbtn" data-act="skf:${v}" aria-pressed="${fSnid === v}">${l}</button>`;
    let xs = all.filter(x => !fSnid || x.snid === fSnid).filter(x => !fStad || (fStad === "oflokkad" ? x.stadur === "oflokkad" : fStad === "stofa" ? x.stadur === "stofa" : x.projectId === fStad));
    xs = xs.slice().sort((a, b) => String(b.breytt || b.stofnad || "").localeCompare(String(a.breytt || a.stofnad || "")));
    const ps = verkefni().filter(p => all.some(x => x.projectId === p.id)).sort((a, b) => b.nr.localeCompare(a.nr));
    let h = `<div class="stat-row" style="margin-top:6px">${chip("", "Öll · " + all.length)}${["word", "excel", "powerpoint", "canva", "tengill"].filter(k => all.some(x => x.snid === k)).map(k => chip(k, SN[k].l)).join("")}</div>
      ${all.length ? `<label style="margin-top:10px">Staður<select id="skstad"><option value="">Allir staðir</option><option value="oflokkad"${fStad === "oflokkad" ? " selected" : ""}>Óflokkað</option><option value="stofa"${fStad === "stofa" ? " selected" : ""}>Stofa</option>${ps.map(p => `<option value="${esc(p.id)}"${fStad === p.id ? " selected" : ""}>${esc(p.nafn)}</option>`).join("")}</select></label>` : ""}`;
    if (!all.length) return h + `<div class="empty" style="margin-top:16px">Engin skjöl ennþá. Ýttu á + til að búa til Word-, Excel- eða PowerPoint-skjal, Canva-hönnun eða tengil.${msOn() ? "" : "<br>Tengdu Microsoft í Stillingum til að búa til Office-skjöl."}</div>`;
    const of = xs.filter(x => x.stadur === "oflokkad"), st = xs.filter(x => x.stadur === "stofa"), vk = xs.filter(x => x.stadur === "verk");
    if (of.length) h += `<div class="viewtitle">Óflokkað · ${of.length}</div>${llist(of.map(row))}`;
    if (st.length) h += `<div class="viewtitle">Stofa · ${st.length}</div>${llist(st.map(row))}`;
    const byP = new Map(); vk.forEach(x => { if (!byP.has(x.projectId)) byP.set(x.projectId, []); byP.get(x.projectId).push(x); });
    [...byP.entries()].sort((a, b) => String(findProj(b[0])?.nr || "").localeCompare(String(findProj(a[0])?.nr || ""))).forEach(([pid, g]) => { const p = findProj(pid);
      h += `<div class="viewtitle">${p ? `<span class="nr">${esc(p.nr)}</span>&nbsp;${esc(p.stutt)}` : "Verkefni"} · ${g.length}</div>${llist(g.map(row))}`; });
    if (!xs.length) h += `<div class="empty" style="margin-top:16px">Engin skjöl passa við síuna.</div>`;
    return h;
  };
  window.skDocsFor = pid => { const g = docs().filter(x => x.projectId === pid).sort((a, b) => String(b.breytt || b.stofnad || "").localeCompare(String(a.breytt || a.stofnad || "")));
    return `<div class="viewtitle" style="display:flex;align-items:center;gap:8px">Skjöl${g.length ? " · " + g.length : ""}${MCP ? `<a class="plink" data-act="sknew:${esc(pid)}" style="margin-left:auto">+ Skjal</a>` : ""}</div>${g.length ? llist(g.map(row)) : ""}`; };

  /* ---------- Staður (sameiginlegt í „nýtt“ og „flokka“) ---------- */
  const placeHTML = (stadur, pid, mappa) => `
    <label>Staður<select id="sk-stad"><option value="oflokkad"${stadur === "oflokkad" ? " selected" : ""}>Óflokkað – flokka síðar</option><option value="stofa"${stadur === "stofa" ? " selected" : ""}>Stofa</option>
      <optgroup label="Verkefni">${verkefni().slice().sort((a, b) => (isDone(a) - isDone(b)) || b.nr.localeCompare(a.nr)).map(p => `<option value="verk:${esc(p.id)}"${stadur === "verk" && pid === p.id ? " selected" : ""}>${esc(p.nafn)}</option>`).join("")}</optgroup></select></label>
    <label id="sk-mappa-l">Mappa<select id="sk-mappa"></select></label>`;
  function wirePlace(mappa) {
    const ss = document.getElementById("sk-stad"), ms = document.getElementById("sk-mappa"), ml = document.getElementById("sk-mappa-l");
    const fill = () => { const v = ss.value, opts = v === "stofa" ? STOFA : v.startsWith("verk:") ? STD : [];
      ml.hidden = !opts.length; const def = mappa && opts.includes(mappa) ? mappa : v === "stofa" ? (opts.includes("Samningar") ? opts[0] : opts[0]) : "02 Vinnuskjöl";
      ms.innerHTML = opts.map(o => `<option${o === def ? " selected" : ""}>${esc(o)}</option>`).join(""); };
    ss.onchange = () => { mappa = null; fill(); }; fill();
    if (!stofaLoaded && msOn()) msStofa().then(xs => { if (xs.length) { STOFA = xs; stofaLoaded = true; if (ss.value === "stofa") fill(); } }).catch(() => {});
  }
  const readPlace = () => { const v = val("sk-stad"); return v.startsWith("verk:") ? { stadur: "verk", projectId: v.slice(5), mappa: val("sk-mappa") } : v === "stofa" ? { stadur: "stofa", projectId: null, mappa: val("sk-mappa") } : { stadur: "oflokkad", projectId: null, mappa: null }; };

  /* ---------- Nýtt skjal ---------- */
  function openNew(pid) {
    let snid = "word";
    const ctx = pid || (DETAIL?.startsWith("verk:") ? DETAIL.slice(5) : null);
    const draw = () => {
      const off = OFFICE(snid);
      sheet.innerHTML = `<h3 id="sheet-h">Nýtt skjal</h3>
        <div class="seg skseg">${Object.entries(SN).map(([k, s]) => `<button type="button" data-sn="${k}" aria-pressed="${k === snid}">${s.l}</button>`).join("")}</div>
        <form id="skf" novalidate>
        <label><span class="lbl">Heiti <span class="req">*</span></span><input id="sk-name" required placeholder="${snid === "excel" ? "T.d. Kostnaðaráætlun" : snid === "powerpoint" ? "T.d. Kynning fyrir verkkaupa" : snid === "canva" ? "T.d. Plakat fyrir viðburð" : snid === "tengill" ? "T.d. Teikningar frá arkitekt" : "T.d. Verklýsing"}" autocomplete="off"></label>
        ${off ? `<label>Sniðmát<select id="sk-tpl"><option value="">Autt skjal</option></select></label>` : ""}
        ${snid === "canva" ? `<div class="card" style="padding:12px 14px"><div class="hint" style="margin:0 0 8px">Búðu hönnunina til í Canva, afritaðu svo slóðina hennar (Deila › Afrita tengil) hingað. Það má líka bæta slóðinni við síðar.</div><a class="btn" href="https://www.canva.com/" target="_blank" rel="noopener">Opna Canva ↗</a></div><label>Slóð á hönnun<input id="sk-url" type="url" placeholder="https://www.canva.com/design/…"></label>` : ""}
        ${snid === "tengill" ? `<label><span class="lbl">Slóð <span class="req">*</span></span><input id="sk-url" type="url" placeholder="https://…"></label>` : ""}
        ${placeHTML(ctx ? "verk" : "oflokkad", ctx, null)}
        ${off ? `<p class="hint">${msOn() ? "Skjalið verður til í OneDrive og opnast í " + SN[snid].l + " í vafranum." : "Tengdu Microsoft í Stillingum til að búa til " + SN[snid].l + "-skjöl."}</p>` : `<p class="hint">Canva-hönnun og tenglar eru vistuð í skjalalistanum (ekki sem skrá í OneDrive).</p>`}
        <div id="f-msg"></div><div class="actions"><button type="button" class="btn" id="sk-cancel">Hætta við</button><button class="btn primary" id="sk-save"${off && !msOn() ? " disabled" : ""}>Búa til</button></div></form>`;
      sheet.querySelectorAll("[data-sn]").forEach(b => b.onclick = () => { const n = val("sk-name"); snid = b.dataset.sn; draw(); document.getElementById("sk-name").value = n; });
      document.getElementById("sk-cancel").onclick = closeSheet; wirePlace(null);
      if (off && msOn()) msSnidmat("." + SN[snid].ext).then(xs => { const t = document.getElementById("sk-tpl"); if (t && xs.length) t.insertAdjacentHTML("beforeend", xs.map(x => `<option value="${esc(x.id)}">${esc(x.name.replace(/\.[^.]+$/, ""))}</option>`).join("")); }).catch(() => {});
      if (!("ontouchstart" in window)) setTimeout(() => document.getElementById("sk-name")?.focus(), 50);
      document.getElementById("skf").onsubmit = async ev => {
        ev.preventDefault(); msg("");
        const name = val("sk-name"), url = val("sk-url"), pl = readPlace(), btn = document.getElementById("sk-save");
        if (!name) return msg("Skrifaðu heiti.", "err");
        if (snid === "tengill" && !url) return msg("Settu inn slóð.", "err");
        if (url && !/^https:\/\//i.test(url)) return msg("Slóðin þarf að byrja á https://", "err");
        btn.disabled = true; btn.textContent = "Bý til…";
        try {
          const id = newId(), now = new Date().toISOString(), rec = { id, name, snid, url: url || null, odId: null, ...pl, stofnad: now, breytt: now };
          if (OFFICE(snid)) { const dir = await dirOf(pl.stadur, pl.projectId, pl.mappa); const it = await msNewOffice(dir, name, SN[snid].ext, val("sk-tpl") || null); rec.url = it.webUrl; rec.odId = it.id; rec.name = it.name.replace(/\.[^.]+$/, ""); rec.slod = dir; }
          await saveSnapshot(s => { (s.skjol = s.skjol || []).push(rec); });
          toast(OFFICE(snid) ? `${SN[snid].l}-skjal búið til` : "Vistað í skjalalista"); closeSheet(); openDetail("skjal:" + id);
        } catch (e) { btn.disabled = false; btn.textContent = "Búa til"; msg(esc(e?.message || "Villa"), "err"); }
      };
    };
    draw(); sheetwrap.hidden = false;
  }

  /* ---------- Flokka / færa ---------- */
  function openMove(id) {
    const x = findD(id); if (!x) return;
    sheet.innerHTML = `<h3 id="sheet-h">Flokka skjal</h3><p class="sub" style="margin:0 0 10px">${badge(x)} ${esc(x.name)}</p>
      <form id="skm" novalidate>${placeHTML(x.stadur === "oflokkad" ? "verk" : x.stadur, x.projectId || (verkefni().find(p => !isDone(p))?.id), x.mappa)}
      ${OFFICE(x.snid) && x.odId ? `<p class="hint">Skráin færist líka á nýja staðinn í OneDrive.</p>` : ""}
      <div id="f-msg"></div><div class="actions"><button type="button" class="btn" id="skm-c">Hætta við</button><button class="btn primary" id="skm-s">Færa</button></div></form>`;
    sheetwrap.hidden = false; wirePlace(x.mappa);
    document.getElementById("skm-c").onclick = closeSheet;
    document.getElementById("skm").onsubmit = async ev => {
      ev.preventDefault(); msg(""); const pl = readPlace(), btn = document.getElementById("skm-s");
      if (OFFICE(x.snid) && x.odId && !msOn()) return msg("Tengdu Microsoft til að færa skrána.", "err");
      btn.disabled = true; btn.textContent = "Færi…";
      try {
        let extra = {};
        if (OFFICE(x.snid) && x.odId) { const dir = await dirOf(pl.stadur, pl.projectId, pl.mappa); const it = await msMove(x.odId, dir); extra = { url: it.webUrl || x.url, slod: dir }; }
        await saveSnapshot(s => { const q = (s.skjol || []).find(y => y.id === id); if (q) Object.assign(q, pl, extra, { breytt: new Date().toISOString() }); });
        toast("Skjal fært: " + where({ ...x, ...pl })); closeSheet();
      } catch (e) { btn.disabled = false; btn.textContent = "Færa"; msg(esc(e?.message || "Villa"), "err"); }
    };
  }

  /* ---------- Síða skjals ---------- */
  window.dSkjalNy = function (x) {
    const p = findProj(x.projectId), s = SN[x.snid] || SN.tengill, off = OFFICE(x.snid);
    return backBar() + `<div class="kicker"><span class="kind">${esc(s.l)}</span>${x.stadur === "oflokkad" ? `<span class="pill late">Óflokkað</span>` : ""}</div>
      <h2 class="big" style="margin-top:6px;display:flex;gap:10px;align-items:center">${badge(x)}<span>${esc(x.name)}</span></h2>
      <p class="sub">${esc(where(x))}${p ? " · " + projLink(p) : ""}</p>
      <div class="links">${x.url ? `<a class="btn primary" href="${esc(x.url)}" target="_blank" rel="noopener">Opna${off ? " í " + s.l : x.snid === "canva" ? " í Canva" : ""} ↗</a>` : x.snid === "canva" ? `<a class="btn primary" href="https://www.canva.com/" target="_blank" rel="noopener">Opna Canva ↗</a>` : ""}
        ${off && x.odId && msOn() ? `<button type="button" class="btn" data-act="msprev:${esc(x.odId)}">Forskoða</button>` : ""}
        ${MCP ? `<button type="button" class="btn" data-act="skflokka:${esc(x.id)}">${x.stadur === "oflokkad" ? "Flokka" : "Færa"}</button><button type="button" class="btn" data-act="skren:${esc(x.id)}">${off ? "Endurnefna" : "Breyta"}</button><button type="button" class="btn danger" data-act="skdel:${esc(x.id)}">Fjarlægja</button>` : ""}</div>
      <div class="card" style="margin-top:14px"><dl class="kv" style="margin:0">${kvRows([["Tegund", esc(s.l)], ["Staður", esc(where(x))], ["Í OneDrive", x.slod ? esc(x.slod.replace(/^AERA\//, "AERA › ").replace(/\//g, " › ")) : ""], ["Verkefni", p ? projLink(p) : ""], ["Búið til", fmtDate(x.stofnad)], ["Síðast breytt", fmtDate(x.breytt)]])}</dl></div>`;
  };

  window.skAct = function (k, r) {
    const id = r.join(":");
    if (k === "skf") { fSnid = id; render(); return; }
    if (k === "sknew") { openNew(id || null); return; }
    if (k === "skflokka") { openMove(id); return; }
    if (k === "skren") { const x = findD(id); if (!x) return;
      if (!OFFICE(x.snid)) { const n = prompt("Heiti", x.name); if (n == null) return; const u = prompt("Slóð (https://…)", x.url || ""); if (u == null) return;
        if (u && !/^https:\/\//i.test(u)) return toast("Slóðin þarf að byrja á https://");
        saveSnapshot(s => { const q = (s.skjol || []).find(y => y.id === id); if (q) Object.assign(q, { name: n.trim() || q.name, url: u.trim() || null, breytt: new Date().toISOString() }); }); return; }
      const n = prompt("Nýtt heiti", x.name); if (!n || n.trim() === x.name) return;
      if (!msOn()) return toast("Tengdu Microsoft til að endurnefna skrána.");
      msMove(x.odId, null, n.trim() + "." + SN[x.snid].ext).then(it => saveSnapshot(s => { const q = (s.skjol || []).find(y => y.id === id); if (q) Object.assign(q, { name: it.name.replace(/\.[^.]+$/, ""), url: it.webUrl || q.url, breytt: new Date().toISOString() }); }))
        .then(() => toast("Endurnefnt")).catch(e => toast(e.message)); return; }
    if (k === "skdel") { const x = findD(id); if (!x || !confirm(`Fjarlægja „${x.name}“ úr skjalalistanum?${x.odId ? "\n\nSkráin sjálf í OneDrive eyðist ekki." : ""}`)) return;
      saveSnapshot(s => { s.skjol = (s.skjol || []).filter(y => y.id !== id); }).then(() => { toast("Fjarlægt úr listanum"); goBack(); }); return; }
  };
  document.addEventListener("change", e => { if (e.target.id === "skstad") { fStad = e.target.value; render(); } });
  window.skOpenNew = openNew;
  window.skCount = () => docs().filter(x => x.stadur === "oflokkad").length;
})();
