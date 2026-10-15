/* Lítill PDF-skrifari (engin ytri söfn): Helvetica/Helvetica-Bold með WinAnsi-kóðun (nær yfir íslenska stafi).
   const doc = PDFDoc(); doc.page(); doc.text(x, y, "texti", {size, bold, color:[r,g,b], align:"right"}); doc.rect(x,y,w,h,[r,g,b]); doc.line(x1,y1,x2,y2,w,[r,g,b]); doc.width(str,size,bold); doc.wrap(str,maxW,size,bold) → [línur]; doc.bytes() → Uint8Array
   Hnit: punktar, (0,0) efst til vinstri á A4 (595×842). */
(function () {
  const W = { r: [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584, 761, 556, 0, 222, 556, 333, 1000, 556, 556, 333, 1000, 667, 333, 1000, 0, 611, 0, 0, 222, 222, 333, 333, 350, 556, 1000, 333, 1000, 500, 333, 944, 0, 500, 667, 278, 333, 556, 556, 556, 556, 260, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 556, 537, 278, 333, 333, 365, 556, 834, 834, 834, 611, 667, 667, 667, 667, 667, 667, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, 556, 556, 556, 556, 556, 556, 889, 500, 556, 556, 556, 556, 278, 278, 278, 278, 556, 556, 556, 556, 556, 556, 556, 584, 611, 556, 556, 556, 556, 500, 556, 500], b: [278, 333, 474, 556, 556, 889, 722, 238, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 333, 333, 584, 584, 584, 611, 975, 722, 722, 722, 722, 667, 611, 778, 722, 278, 556, 722, 611, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 333, 278, 333, 584, 556, 333, 556, 611, 556, 611, 556, 333, 611, 611, 278, 278, 556, 278, 889, 611, 611, 611, 611, 389, 556, 333, 611, 556, 778, 556, 556, 500, 389, 280, 389, 584, 761, 556, 0, 278, 556, 500, 1000, 556, 556, 333, 1000, 667, 333, 1000, 0, 611, 0, 0, 278, 278, 500, 500, 350, 556, 1000, 333, 1000, 556, 333, 944, 0, 500, 667, 278, 333, 556, 556, 556, 556, 280, 556, 333, 737, 370, 556, 584, 333, 737, 333, 400, 584, 333, 333, 333, 611, 556, 278, 333, 333, 365, 556, 834, 834, 834, 611, 722, 722, 722, 722, 722, 722, 1000, 722, 667, 667, 667, 667, 278, 278, 278, 278, 722, 722, 778, 778, 778, 778, 778, 584, 778, 722, 722, 722, 722, 667, 667, 611, 556, 556, 556, 556, 556, 556, 889, 556, 556, 556, 556, 556, 278, 278, 278, 278, 611, 611, 611, 611, 611, 611, 611, 584, 611, 611, 611, 611, 611, 556, 611, 556] };
  const CP = { 0x20AC: 128, 0x201A: 130, 0x0192: 131, 0x201E: 132, 0x2026: 133, 0x2020: 134, 0x2021: 135, 0x02C6: 136, 0x2030: 137, 0x0160: 138, 0x2039: 139, 0x0152: 140, 0x017D: 142, 0x2018: 145, 0x2019: 146, 0x201C: 147, 0x201D: 148, 0x2022: 149, 0x2013: 150, 0x2014: 151, 0x02DC: 152, 0x2122: 153, 0x0161: 154, 0x203A: 155, 0x0153: 156, 0x017E: 158, 0x0178: 159 };
  const code = ch => { const c = ch.codePointAt(0); if (c >= 32 && c < 127) return c; if (c >= 160 && c <= 255) return c; if (CP[c]) return CP[c]; if (c === 0x2212) return 45; if (c === 0xA0) return 32; return 63; };
  const codes = s => Array.from(String(s ?? "")).map(code);
  window.PDFDoc = function () {
    const pages = []; let cur = null; const H = 842;
    const col = c => (c || [0, 0, 0]).map(v => (v / 255).toFixed(3)).join(" ");
    const enc = cs => "(" + cs.map(c => c === 40 || c === 41 || c === 92 ? "\\" + String.fromCharCode(c) : c < 128 ? String.fromCharCode(c) : "\\" + c.toString(8).padStart(3, "0")).join("") + ")";
    const width = (s, size = 10, bold = false) => codes(s).reduce((a, c) => a + (W[bold ? "b" : "r"][c - 32] || 556), 0) * size / 1000;
    const api = {
      page() { cur = []; pages.push(cur); return api; },
      width,
      text(x, y, s, o = {}) { const size = o.size || 10, w = width(s, size, o.bold); const xx = o.align === "right" ? x - w : o.align === "center" ? x - w / 2 : x;
        cur.push(`BT /${o.bold ? "F2" : "F1"} ${size} Tf ${col(o.color)} rg ${xx.toFixed(2)} ${(H - y).toFixed(2)} Td ${enc(codes(s))} Tj ET`); return api; },
      rect(x, y, w, h, c) { cur.push(`${col(c)} rg ${x.toFixed(2)} ${(H - y - h).toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re f`); return api; },
      line(x1, y1, x2, y2, lw = 0.5, c) { cur.push(`${col(c)} RG ${lw} w ${x1.toFixed(2)} ${(H - y1).toFixed(2)} m ${x2.toFixed(2)} ${(H - y2).toFixed(2)} l S`); return api; },
      wrap(s, maxW, size = 10, bold = false) { const out = []; String(s ?? "").split("\n").forEach(par => { let line = ""; par.split(/\s+/).forEach(wd => { const t = line ? line + " " + wd : wd; if (width(t, size, bold) <= maxW || !line) line = t; else { out.push(line); line = wd; } }); out.push(line); }); return out; },
      bytes() {
        const objs = []; const add = s => { objs.push(s); return objs.length; };
        const f1 = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
        const f2 = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>");
        const pagesId = objs.length + 1; objs.push(null);
        const kids = pages.map(ops => { const st = ops.join("\n"); const sid = add(`<< /Length ${st.length} >>\nstream\n${st}\nendstream`);
          return add(`<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 ${f1} 0 R /F2 ${f2} 0 R >> >> /Contents ${sid} 0 R >>`); });
        objs[pagesId - 1] = `<< /Type /Pages /Kids [${kids.map(k => k + " 0 R").join(" ")}] /Count ${kids.length} >>`;
        const cat = add(`<< /Type /Catalog /Pages ${pagesId} 0 R >>`);
        let out = "%PDF-1.4\n%\xE2\xE3\xCF\xD3\n"; const offs = [];
        objs.forEach((o, i) => { offs.push(out.length); out += `${i + 1} 0 obj\n${o}\nendobj\n`; });
        const xref = out.length;
        out += `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n` + offs.map(o => String(o).padStart(10, "0") + " 00000 n \n").join("") + `trailer\n<< /Size ${objs.length + 1} /Root ${cat} 0 R >>\nstartxref\n${xref}\n%%EOF`;
        const b = new Uint8Array(out.length); for (let i = 0; i < out.length; i++) b[i] = out.charCodeAt(i) & 255; return b;
      }
    };
    return api;
  };
})();
