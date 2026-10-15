/* Reikningar: reikningsfæra tímavinnu og útlagðan kostnað, búa til reikning (prenta/PDF og Excel).
   Gögn: S.reikningar[], S.stillingar.reikn{}, verkefni.timagjald, færsla.reikningur / kostnaður.reikningur = id reiknings. */
(function () {
  const RK = () => ({ nafn: "AERA Studio", kt: "", heimilisfang: "", vsknr: "", banki: "", netfang: "", simi: "", vsk: 24, frestur: 14, nastaNr: 1001, timagjald: 0, ...(CFG().reikn || {}) });
  const reikningar = () => S?.reikningar || [];
  const findR = id => reikningar().find(r => r.id === id);
  const rate = p => +(p?.timagjald || RK().timagjald || 0);
  const addDays = (iso, n) => { const x = d(iso); x.setDate(x.getDate() + n); return isoDay(x); };
  const kaupandi = p => findVsk(p?.vkId) || findVsk(p?.vskId) || null;
  const obilled = pid => ({ es: faerslur().filter(e => e.verkefniId === pid && !e.reikningur && e.id && !isTmp(e.id)), ks: kostn().filter(k => k.verkefniId === pid && !k.reikningur) });
  window.rkUnbilled = obilled;
  window.rkFind = findR;

  function buildLines(p, es, ks, verd, sund) {
    const L = [];
    if (sund) es.slice().sort((a, b) => String(a.date + (a.fra || "")).localeCompare(String(b.date + (b.fra || "")))).forEach(e => L.push({ lysing: dIS(String(e.date).slice(0, 10)) + " – " + cleanLys(e.lysing), magn: Math.round((+e.timar || 0) * 100) / 100, eining: "klst.", verd }));
    else {
      const g = new Map();
      es.forEach(e => { const r = e.undirverkId ? findSub(e.undirverkId) : null; const k = r ? r.u.id : "_"; if (!g.has(k)) g.set(k, { lysing: r ? "Tímavinna – " + subName(r.u) : "Tímavinna – " + (p?.stutt || p?.nafn || ""), magn: 0 }); g.get(k).magn += +e.timar || 0; });
      [...g.values()].forEach(x => L.push({ lysing: x.lysing, magn: Math.round(x.magn * 100) / 100, eining: "klst.", verd }));
    }
    ks.forEach(k => L.push({ lysing: "Útlagður kostnaður – " + (k.lysing || ""), magn: 1, eining: "stk.", verd: Math.round(+k.upphaed || 0) }));
    return L;
  }
  const totals = (L, vskP) => { const net = L.reduce((a, l) => a + Math.round(l.magn * l.verd), 0), vsk = Math.round(net * vskP / 100); return { net, vsk, tot: net + vsk }; };

  /* ---------- Form: reikningsfæra verkefni ---------- */
  function openForm(pid) {
    const p = findProj(pid); if (!p) return;
    const { es, ks } = obilled(pid), rk = RK(), buyer = kaupandi(p);
    if (!es.length && !ks.length) { toast("Engir óreikningsfærðir tímar eða kostnaður á verkefninu."); return; }
    const missing = !rk.kt || !rk.banki;
    sheet.innerHTML = `<h3 id="sheet-h">Reikningsfæra · ${esc(p.nr)}</h3>
      ${missing ? `<div class="msg err" style="margin-bottom:10px">Upplýsingar um seljanda vantar (kennitala og bankareikningur). Fylltu þær út í Stillingum → Reikningar áður en reikningur er sendur.</div>` : ""}
      <form id="rkf" novalidate>
      <label>Viðskiptavinur (greiðandi)<select id="rk-buyer"><option value="">– Veldu –</option>${vskList().map(c => `<option value="${esc(c.id)}"${buyer && sameId(buyer.id, c.id) ? " selected" : ""}>${esc(c.name)}</option>`).join("")}</select></label>
      <div class="two"><label>Reikningsdagur<input id="rk-dags" type="date" value="${isoDay(today)}"></label><label>Gjalddagi<input id="rk-gjd" type="date" value="${addDays(isoDay(today), +rk.frestur || 14)}"></label></div>
      <div class="two"><label>Tímagjald (kr./klst., án VSK)<input id="rk-rate" inputmode="numeric" value="${rate(p) || ""}" placeholder="t.d. 15000"></label><label>Reikningsnúmer<input id="rk-nr" inputmode="numeric" value="${esc(String(rk.nastaNr))}"></label></div>
      <label class="chk" style="flex-direction:row;align-items:center;gap:8px;color:var(--ink);font-size:14.5px"><input type="checkbox" id="rk-sund" style="width:auto">Ein lína á hverja tímafærslu (annars tekið saman eftir undirverkum)</label>
      ${es.length ? `<div class="kind" style="margin-top:6px">Tímar · <a class="plink" id="rk-allt">velja/afvelja allt</a></div><div class="rkpick">${es.sort((a, b) => String(a.date).localeCompare(String(b.date))).map(e => `<label class="rkrow"><input type="checkbox" data-rke="${esc(e.id)}" checked><span>${esc(dIS(String(e.date).slice(0, 10)))} · ${esc(cleanLys(e.lysing))}</span><b>${esc(hrs(e.timar))}</b></label>`).join("")}</div>` : ""}
      ${ks.length ? `<div class="kind" style="margin-top:6px">Útlagður kostnaður</div><div class="rkpick">${ks.map(k => `<label class="rkrow"><input type="checkbox" data-rkk="${esc(k.id)}" checked><span>${esc(dIS(k.date))} · ${esc(k.lysing)}</span><b>${esc(kr(k.upphaed))}</b></label>`).join("")}</div>` : ""}
      <div class="card rksum" id="rk-sum"></div>
      <div id="f-msg"></div>
      <div class="actions"><button type="button" class="btn" id="rk-cancel">Hætta við</button><button class="btn primary" id="rk-save">Búa til reikning</button></div></form>`;
    sheet.closest(".sheet").classList.add("wide"); sheetwrap.hidden = false;
    const sel = () => ({ es: es.filter(e => sheet.querySelector(`[data-rke="${CSS.escape(e.id)}"]`)?.checked), ks: ks.filter(k => sheet.querySelector(`[data-rkk="${CSS.escape(k.id)}"]`)?.checked) });
    const upd = () => { const s = sel(), v = parseKr(val("rk-rate")), L = buildLines(p, s.es, s.ks, v, document.getElementById("rk-sund").checked), t = totals(L, +RK().vsk || 0);
      document.getElementById("rk-sum").innerHTML = `<div class="sub-i"><span>${num(sumT(s.es))} klst. × ${esc(kr(v))}</span><b>${esc(kr(L.filter(l => l.eining === "klst.").reduce((a, l) => a + Math.round(l.magn * l.verd), 0)))}</b></div>
        ${s.ks.length ? `<div class="sub-i"><span>Útlagður kostnaður (${s.ks.length})</span><b>${esc(kr(sumK(s.ks)))}</b></div>` : ""}
        <div class="sub-i"><span>VSK ${esc(String(RK().vsk))}%</span><b>${esc(kr(t.vsk))}</b></div><div class="sub-i" style="font-size:16px"><span>Samtals til greiðslu</span><b>${esc(kr(t.tot))}</b></div>`; };
    sheet.querySelectorAll("input").forEach(i => i.addEventListener("input", upd)); sheet.querySelectorAll("input[type=checkbox]").forEach(i => i.addEventListener("change", upd));
    const all = document.getElementById("rk-allt"); if (all) all.onclick = () => { const cs = [...sheet.querySelectorAll("[data-rke]")], on = !cs.every(c => c.checked); cs.forEach(c => c.checked = on); upd(); };
    upd();
    document.getElementById("rk-cancel").onclick = closeSheet;
    document.getElementById("rkf").onsubmit = async ev => {
      ev.preventDefault(); msg("");
      const s = sel(), v = parseKr(val("rk-rate")), nr = parseInt(val("rk-nr")) || 0, bid = val("rk-buyer"), dags = val("rk-dags"), gjd = val("rk-gjd");
      if (!s.es.length && !s.ks.length) return msg("Veldu að minnsta kosti eina færslu.", "err");
      if (s.es.length && !v) return msg("Skráðu tímagjald.", "err");
      if (!bid) return msg("Veldu viðskiptavin.", "err");
      if (!nr) return msg("Skráðu reikningsnúmer.", "err");
      if (reikningar().some(r => r.nr === nr) && !confirm(`Reikningur nr. ${nr} er þegar til. Nota samt sama númer?`)) return;
      const L = buildLines(p, s.es, s.ks, v, document.getElementById("rk-sund").checked), vskP = +RK().vsk || 0, t = totals(L, vskP), c = findVsk(bid);
      const id = newId(), eIds = new Set(s.es.map(e => e.id)), kIds = new Set(s.ks.map(k => k.id));
      const R = { id, nr, dags, gjalddagi: gjd, verkefniId: pid, vskId: bid, kaupandi: c ? { name: c.name, kennitala: c.kennitala || "", heimilisfang: c.heimilisfang || "", postnumer: c.postnumer || "", stadur: c.stadur || "", netfang: c.netfang || "" } : null,
        seljandi: { ...RK() }, linur: L, vskP, ...t, timaIds: [...eIds], kostnIds: [...kIds], stada: "Ógreitt", stofnad: new Date().toISOString() };
      delete R.seljandi.nastaNr; delete R.seljandi.timagjald;
      document.getElementById("rk-save").disabled = true;
      await saveSnapshot(st => {
        (st.reikningar = st.reikningar || []).push(R);
        const T = st.timar || {}; [...(T.faerslur || []), ...(T.nyjast || [])].forEach(e => { if (eIds.has(e.id)) e.reikningur = id; });
        (st.kostnadur || []).forEach(k => { if (kIds.has(k.id)) k.reikningur = id; });
        st.stillingar = st.stillingar || {}; st.stillingar.reikn = { ...(st.stillingar.reikn || {}), nastaNr: Math.max(nr + 1, +(st.stillingar.reikn?.nastaNr || 0)) };
        const q = sFindP(st, pid); if (q && v) q.timagjald = v;
      });
      toast(`Reikningur ${nr} búinn til`); closeSheet(); openDetail("reikn:" + id);
    };
  }

  /* ---------- Reikningurinn sjálfur (sama útlit á skjá og í prentun) ---------- */
  function invoiceHTML(R) {
    const s = R.seljandi || {}, b = R.kaupandi || {}, p = findProj(R.verkefniId);
    const es = faerslur().filter(e => (R.timaIds || []).includes(e.id)).sort((a, b) => String(a.date + (a.fra || "")).localeCompare(String(b.date + (b.fra || ""))));
    const fx = n => String(n).replace(".", ",");
    return `<div class="inv">
      <div class="inv-h"><div><div class="inv-s">${esc(s.nafn || "")}</div><div class="inv-m">${[s.kt && "Kt. " + s.kt, s.vsknr && "VSK-nr. " + s.vsknr].filter(Boolean).map(esc).join(" · ")}<br>${esc(s.heimilisfang || "")}<br>${[s.netfang, s.simi].filter(Boolean).map(esc).join(" · ")}</div></div>
        <div style="text-align:right"><div class="inv-t">Reikningur</div><div class="inv-m">Nr. <b>${esc(String(R.nr))}</b><br>Dags. ${esc(dIS(R.dags))}<br>Gjalddagi ${esc(dIS(R.gjalddagi))}</div></div></div>
      <div class="inv-b"><div class="inv-k">Greiðandi</div><b>${esc(b.name || "")}</b><br>${b.kennitala ? "Kt. " + esc(b.kennitala) + "<br>" : ""}${esc(b.heimilisfang || "")}${b.postnumer || b.stadur ? "<br>" + esc([b.postnumer, b.stadur].filter(Boolean).join(" ")) : ""}</div>
      ${p ? `<div class="inv-m" style="margin:10px 0 0">Verkefni: <b>${esc(p.nafn)}</b></div>` : ""}
      <table class="inv-tab"><thead><tr><th>Lýsing</th><th class="r">Magn</th><th>Ein.</th><th class="r">Einingaverð</th><th class="r">Upphæð</th></tr></thead><tbody>
        ${R.linur.map(l => `<tr><td>${esc(l.lysing)}</td><td class="r">${esc(fx(l.magn))}</td><td>${esc(l.eining)}</td><td class="r">${esc(kr(l.verd))}</td><td class="r">${esc(kr(Math.round(l.magn * l.verd)))}</td></tr>`).join("")}
      </tbody></table>
      <div class="inv-tot"><div><span>Samtals án VSK</span><b>${esc(kr(R.net))}</b></div><div><span>VSK ${esc(String(R.vskP))}%</span><b>${esc(kr(R.vsk))}</b></div><div class="g"><span>Samtals til greiðslu</span><b>${esc(kr(R.tot))}</b></div></div>
      <div class="inv-pay">${s.banki ? `Greiðist inn á reikning <b>${esc(s.banki)}</b>${s.kt ? ", kt. " + esc(s.kt) : ""}, í síðasta lagi ${esc(dIS(R.gjalddagi))}. Vinsamlegast merkið greiðslu með reikningsnúmeri ${esc(String(R.nr))}.` : ""}</div>
      ${es.length ? `<div class="inv-sund"><div class="inv-k">Sundurliðun tíma</div><table class="inv-tab"><thead><tr><th>Dags.</th><th>Lýsing</th><th>Tími</th><th class="r">Klst.</th></tr></thead><tbody>
        ${es.map(e => `<tr><td>${esc(dIS(String(e.date).slice(0, 10)))}</td><td>${esc(cleanLys(e.lysing))}</td><td>${esc(e.fra ? e.fra + (e.til ? "–" + e.til : "") : "")}</td><td class="r">${esc(fx(Math.round((+e.timar || 0) * 100) / 100))}</td></tr>`).join("")}
        <tr class="g"><td colspan="3">Samtals</td><td class="r">${esc(fx(Math.round(sumT(es) * 100) / 100))}</td></tr></tbody></table></div>` : ""}
    </div>`;
  }

  function printInv(R) {
    let pa = document.getElementById("printarea");
    if (!pa) { pa = document.createElement("div"); pa.id = "printarea"; document.body.appendChild(pa); }
    pa.innerHTML = invoiceHTML(R);
    const old = document.title; document.title = `Reikningur ${R.nr} ${R.kaupandi?.name || ""}`.trim();
    document.body.classList.add("printinv");
    const done = () => { document.body.classList.remove("printinv"); document.title = old; window.removeEventListener("afterprint", done); };
    window.addEventListener("afterprint", done);
    setTimeout(() => { window.print(); setTimeout(done, 1500); }, 50);
  }

  function invXlsx(R) {
    const S_ = (v, s) => ({ v, s }), rows = [], s = R.seljandi || {}, b = R.kaupandi || {}, p = findProj(R.verkefniId);
    rows.push([S_("Reikningur nr. " + R.nr, 1)]);
    rows.push([S_([s.nafn, s.kt && "kt. " + s.kt, s.heimilisfang].filter(Boolean).join("   ·   "), 2)]);
    rows.push([S_("Dags. " + dIS(R.dags) + "   ·   Gjalddagi " + dIS(R.gjalddagi), 2)]);
    rows.push([S_("Greiðandi: " + [b.name, b.kennitala && "kt. " + b.kennitala, b.heimilisfang].filter(Boolean).join(", "), 2)]);
    if (p) rows.push([S_("Verkefni: " + p.nafn, 2)]);
    rows.push([]);
    rows.push([S_("Lýsing", 3), S_("Magn", 13), S_("Ein.", 3), S_("Einingaverð", 13), S_("Upphæð", 13)]);
    const hdr = rows.length;
    R.linur.forEach(l => rows.push([S_(l.lysing, 12), S_(l.magn, 6), S_(l.eining, 4), S_(Math.round(l.verd), 15), S_(Math.round(l.magn * l.verd), 15)]));
    rows.push([S_("Samtals án VSK", 8), S_("", 8), S_("", 8), S_("", 8), S_(R.net, 16)]);
    rows.push([S_("VSK " + R.vskP + "%", 8), S_("", 8), S_("", 8), S_("", 8), S_(R.vsk, 16)]);
    rows.push([S_("Samtals til greiðslu", 10), S_("", 10), S_("", 10), S_("", 10), S_(R.tot, 17)]);
    if (s.banki) { rows.push([]); rows.push([S_("Greiðist inn á reikning " + s.banki + (s.kt ? ", kt. " + s.kt : "") + ", í síðasta lagi " + dIS(R.gjalddagi) + ".", 2)]); }
    const es = faerslur().filter(e => (R.timaIds || []).includes(e.id)).sort((a, b) => String(a.date + (a.fra || "")).localeCompare(String(b.date + (b.fra || ""))));
    if (es.length) { rows.push([], [S_("Sundurliðun tíma", 7), S_("", 7), S_("", 7), S_("", 7), S_("", 14)]);
      rows.push([S_("Dags. og lýsing", 3), S_("Frá", 3), S_("Til", 3), S_("", 3), S_("Klst.", 13)]);
      es.forEach(e => rows.push([S_(dIS(String(e.date).slice(0, 10)) + "   " + cleanLys(e.lysing), 12), S_(e.fra || "", 4), S_(e.til || "", 4), S_("", 4), S_(Math.round((+e.timar || 0) * 100) / 100, 6)]));
      rows.push([S_("Samtals", 8), S_("", 8), S_("", 8), S_("", 8), S_(Math.round(sumT(es) * 100) / 100, 9)]); }
    const bytes = makeXlsx({ sheet: "Reikningur " + R.nr, cols: [46, 9, 7, 14, 15], rows, merges: ["A1:E1"], freezeRow: 0, heights: { 0: 30 } });
    const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    a.download = `Reikningur ${R.nr} ${b.name || ""}`.replace(/[\\/:*?"<>|]+/g, " ").replace(/\s+/g, " ").trim() + ".xlsx";
    document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 5000); toast("Reikningur sóttur sem Excel");
  }

  /* ---------- Síða reiknings ---------- */
  window.rkDetail = function (R) {
    const p = findProj(R.verkefniId), greitt = R.stada === "Greitt", late = !greitt && daysTo(R.gjalddagi) < 0;
    return backBar() + `<div class="kicker"><span class="kind">Reikningur</span><span class="pill ${greitt ? "" : late ? "late" : ""}">${greitt ? "Greitt" : late ? "Kominn fram yfir gjalddaga" : "Ógreitt"}</span></div>
      <h2 class="big" style="margin-top:6px">Nr. ${esc(String(R.nr))} · ${esc(R.kaupandi?.name || "")}</h2>
      <p class="sub">${p ? projLink(p) + " · " : ""}${esc(kr(R.tot))} með VSK</p>
      <div class="links"><button type="button" class="btn primary" data-act="rkprint:${esc(R.id)}">Prenta / vista PDF</button><button type="button" class="btn" data-act="rkxlsx:${esc(R.id)}">Sækja Excel</button>
        <button type="button" class="btn" data-act="rkpaid:${esc(R.id)}">${greitt ? "Merkja ógreitt" : "✓ Merkja greitt"}</button><button type="button" class="btn danger" data-act="rkdel:${esc(R.id)}">Eyða</button></div>
      <div class="invwrap">${invoiceHTML(R)}</div>
      <p class="hint" style="margin-top:8px">„Prenta / vista PDF“ opnar prentgluggann. Veldu „Vista sem PDF“ (tölva) eða deildu (sími) til að fá PDF-skjal til að senda.</p>`;
  };
  const rkRow = R => { const greitt = R.stada === "Greitt", late = !greitt && daysTo(R.gjalddagi) < 0;
    return lrow("reikn:" + R.id, { lead: `<b>${esc(String(R.nr).slice(-3))}</b><small>nr.</small>`, title: esc((R.kaupandi?.name || "") + " · " + (findProj(R.verkefniId)?.nr || "")), sub: esc([dIS(R.dags), greitt ? "Greitt" : late ? "Fram yfir gjalddaga" : "Gjalddagi " + dIS(R.gjalddagi)].join(" · ")), right: `<span class="lc">${esc(kr(R.tot))}</span>` }); };
  window.rkList = (pred, title) => { const xs = reikningar().filter(pred || (() => true)).slice().sort((a, b) => (b.nr || 0) - (a.nr || 0));
    return xs.length ? `<div class="viewtitle">${title || "Reikningar"} · ${xs.length}</div>${llist(xs.map(rkRow))}` : ""; };

  // Á Tímar-flipa verkefnis: óreikningsfært + hnappur + reikningar verkefnisins
  window.rkProjSection = function (pid) {
    const { es, ks } = obilled(pid);
    const box = es.length || ks.length ? `<div class="card rkopen"><div class="sub-i"><span>Óreikningsfært</span><b>${esc(hrs(sumT(es)))}${ks.length ? " + " + esc(kr(sumK(ks))) : ""}</b></div>${MCP ? `<button type="button" class="btn primary" data-act="rknew:${esc(pid)}" style="width:100%;margin-top:10px">Reikningsfæra</button>` : ""}</div>` : "";
    return (box ? `<div class="viewtitle">Reikningsfærsla</div>${box}` : "") + rkList(r => r.verkefniId === pid, "Reikningar verkefnisins");
  };
  // Á Tímaskráningu: verkefni með óreikningsfært + allir reikningar
  window.rkTimarSection = function () {
    const ps = verkefni().map(p => ({ p, ...obilled(p.id) })).filter(x => x.es.length || x.ks.length).sort((a, b) => sumT(b.es) - sumT(a.es));
    const open = ps.length ? `<div class="viewtitle">Óreikningsfært</div>${llist(ps.map(x => `<div class="lrow"><div class="lm"><div class="lt"><span class="nr">${esc(x.p.nr)}</span>&nbsp; ${esc(x.p.stutt)}</div><div class="ls">${esc([hrs(sumT(x.es)), x.ks.length ? kr(sumK(x.ks)) + " kostnaður" : "", rate(x.p) ? "≈ " + kr(sumT(x.es) * rate(x.p) + sumK(x.ks)) + " án VSK" : ""].filter(Boolean).join(" · "))}</div></div>${MCP ? `<button type="button" class="btn" data-act="rknew:${esc(x.p.id)}" style="flex:none">Reikningsfæra</button>` : ""}</div>`))}` : "";
    return open + rkList(null, "Reikningar");
  };
  window.rkTag = e => e.reikningur ? (findR(e.reikningur) ? "Reikn. " + findR(e.reikningur).nr : "Reikningsfært") : "";

  window.rkStillingarHTML = function () {
    const r = RK(), f = (k, l, extra = "") => `<label>${l}<input id="rk-${k}" value="${esc(String(r[k] ?? ""))}" ${extra}></label>`;
    return `<div class="viewtitle">Reikningar</div>
      <p class="hint" style="margin:0 0 8px">Upplýsingar um seljanda sem birtast á reikningum.</p>
      ${f("nafn", "Nafn seljanda")}
      <div class="two">${f("kt", "Kennitala", 'inputmode="numeric" placeholder="000000-0000"')}${f("vsknr", "VSK-númer", 'inputmode="numeric"')}</div>
      ${f("heimilisfang", "Heimilisfang", 'placeholder="Gata 1, 101 Reykjavík"')}
      <div class="two">${f("netfang", "Netfang", 'type="email"')}${f("simi", "Sími", 'inputmode="tel"')}</div>
      ${f("banki", "Bankareikningur", 'placeholder="0000-26-000000"')}
      <div class="two">${f("timagjald", "Sjálfgefið tímagjald (kr./klst.)", 'inputmode="numeric"')}${f("vsk", "VSK %", 'inputmode="decimal"')}</div>
      <div class="two">${f("frestur", "Greiðslufrestur (dagar)", 'inputmode="numeric"')}${f("nastaNr", "Næsta reikningsnúmer", 'inputmode="numeric"')}</div>
      <button type="button" class="btn" data-act="rkcfg" style="margin-top:8px">Vista upplýsingar</button>`;
  };

  window.rkAct = function (k, r) {
    const id = r.join(":");
    if (k === "rknew") { openForm(id); return; }
    if (k === "rkprint") { const R = findR(id); if (R) printInv(R); return; }
    if (k === "rkxlsx") { const R = findR(id); if (R) invXlsx(R); return; }
    if (k === "rkpaid") { saveSnapshot(s => { const q = (s.reikningar || []).find(x => x.id === id); if (q) { q.stada = q.stada === "Greitt" ? "Ógreitt" : "Greitt"; q.greitt = q.stada === "Greitt" ? isoDay(new Date()) : null; } }).then(() => toast("Staða reiknings uppfærð")); return; }
    if (k === "rkdel") { const R = findR(id); if (!R || !confirm(`Eyða reikningi nr. ${R.nr}?\n\nTímarnir og kostnaðurinn verða aftur óreikningsfærð.`)) return;
      saveSnapshot(s => { s.reikningar = (s.reikningar || []).filter(x => x.id !== id); const T = s.timar || {}; [...(T.faerslur || []), ...(T.nyjast || [])].forEach(e => { if (e.reikningur === id) delete e.reikningur; }); (s.kostnadur || []).forEach(x => { if (x.reikningur === id) delete x.reikningur; }); })
        .then(() => { toast("Reikningi eytt"); goBack(); }); return; }
    if (k === "rkcfg") { const g = x => (document.getElementById("rk-" + x)?.value || "").trim(), n = x => parseFloat(g(x).replace(",", ".")) || 0;
      const o = { nafn: g("nafn"), kt: g("kt"), vsknr: g("vsknr"), heimilisfang: g("heimilisfang"), netfang: g("netfang"), simi: g("simi"), banki: g("banki"), timagjald: parseKr(g("timagjald")), vsk: n("vsk"), frestur: Math.round(n("frestur")) || 14, nastaNr: Math.round(n("nastaNr")) || 1001 };
      saveSnapshot(s => { s.stillingar = s.stillingar || {}; s.stillingar.reikn = { ...(s.stillingar.reikn || {}), ...o }; }).then(() => toast("Upplýsingar um seljanda vistaðar")); return; }
  };
})();
