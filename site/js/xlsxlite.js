/*
 * xlsxlite: read an .xlsx and change single cells without touching anything else.
 * No libraries: uses the browser's built-in CompressionStream / DecompressionStream.
 * Only the edited cells (and the workbook's "recalculate on open" flag) change,
 * so formatting, merged cells, conditional formatting and formulas stay as they were.
 */
(function (root) {
  const NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const te = new TextEncoder(), td = new TextDecoder();

  /* ---------- CRC32 ---------- */
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC[(c ^ u8[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }

  async function pipe(u8, stream) {
    const out = new Response(new Blob([u8]).stream().pipeThrough(stream));
    return new Uint8Array(await out.arrayBuffer());
  }
  const inflate = (u8) => pipe(u8, new DecompressionStream("deflate-raw"));
  const deflate = (u8) => pipe(u8, new CompressionStream("deflate-raw"));

  /* ---------- ZIP ---------- */
  function unzip(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new Error("This file isn't a valid .xlsx (zip) file.");
    const count = dv.getUint16(eocd + 10, true); let p = dv.getUint32(eocd + 16, true);
    const entries = [];
    for (let k = 0; k < count; k++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error("Damaged .xlsx file.");
      const e = {
        madeBy: dv.getUint16(p + 4, true), needed: dv.getUint16(p + 6, true), flags: dv.getUint16(p + 8, true),
        method: dv.getUint16(p + 10, true), time: dv.getUint16(p + 12, true), date: dv.getUint16(p + 14, true),
        crc: dv.getUint32(p + 16, true), csize: dv.getUint32(p + 20, true), usize: dv.getUint32(p + 24, true),
        intAttr: dv.getUint16(p + 36, true), extAttr: dv.getUint32(p + 38, true), offset: dv.getUint32(p + 42, true),
      };
      const nl = dv.getUint16(p + 28, true), xl = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true);
      e.nameBytes = u8.slice(p + 46, p + 46 + nl); e.name = td.decode(e.nameBytes);
      const lnl = dv.getUint16(e.offset + 26, true), lxl = dv.getUint16(e.offset + 28, true);
      const start = e.offset + 30 + lnl + lxl;
      e.raw = u8.slice(start, start + e.csize);
      entries.push(e);
      p += 46 + nl + xl + cl;
    }
    return {
      entries,
      has: (name) => entries.some((e) => e.name === name),
      async text(name) {
        const e = entries.find((x) => x.name === name); if (!e) return null;
        const data = e.method === 0 ? e.raw : await inflate(e.raw);
        return td.decode(data);
      },
    };
  }

  async function zip(entries, replaced) {
    const parts = [], central = []; let offset = 0;
    for (const e0 of entries) {
      let e = e0;
      if (replaced[e0.name] != null) {
        const data = te.encode(replaced[e0.name]);
        e = { ...e0, method: 8, flags: e0.flags & ~0x0008, crc: crc32(data), usize: data.length, raw: await deflate(data) };
        e.csize = e.raw.length;
      }
      const flags = e.flags & ~0x0008;
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, e.needed, true); lh.setUint16(6, flags, true); lh.setUint16(8, e.method, true);
      lh.setUint16(10, e.time, true); lh.setUint16(12, e.date, true); lh.setUint32(14, e.crc, true); lh.setUint32(18, e.csize, true);
      lh.setUint32(22, e.usize, true); lh.setUint16(26, e.nameBytes.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), e.nameBytes, e.raw);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, e.madeBy, true); ch.setUint16(6, e.needed, true); ch.setUint16(8, flags, true);
      ch.setUint16(10, e.method, true); ch.setUint16(12, e.time, true); ch.setUint16(14, e.date, true); ch.setUint32(16, e.crc, true);
      ch.setUint32(20, e.csize, true); ch.setUint32(24, e.usize, true); ch.setUint16(28, e.nameBytes.length, true);
      ch.setUint16(30, 0, true); ch.setUint16(32, 0, true); ch.setUint16(34, 0, true); ch.setUint16(36, e.intAttr, true);
      ch.setUint32(38, e.extAttr, true); ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), e.nameBytes);
      offset += 30 + e.nameBytes.length + e.raw.length;
    }
    const cdSize = central.reduce((t, a) => t + a.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, entries.length, true); end.setUint16(10, entries.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((t, a) => t + a.length, 0)); let o = 0;
    for (const a of all) { out.set(a, o); o += a.length; }
    return out;
  }

  /* ---------- cell refs ---------- */
  const colIndex = (letters) => letters.toUpperCase().split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const colLetter = (i) => { let s = ""; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const splitRef = (ref) => { const m = /^([A-Z]+)(\d+)$/.exec(ref); return { col: colIndex(m[1]), row: Number(m[2]) }; };
  const parseXml = (s) => new DOMParser().parseFromString(s, "application/xml");
  const kids = (el, tag) => Array.from(el.childNodes).filter((n) => n.nodeType === 1 && n.localName === tag);
  const serialize = (doc) => { const s = new XMLSerializer().serializeToString(doc); return s.startsWith("<?xml") ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s; };

  /* ---------- workbook ---------- */
  async function open(buf) {
    const z = unzip(buf);
    const wbXml = await z.text("xl/workbook.xml");
    const relsXml = await z.text("xl/_rels/workbook.xml.rels");
    if (!wbXml || !relsXml) throw new Error("This file has no Excel workbook inside.");
    const rels = {};
    for (const r of parseXml(relsXml).getElementsByTagName("Relationship")) rels[r.getAttribute("Id")] = r.getAttribute("Target");
    const sheets = Array.from(parseXml(wbXml).getElementsByTagName("sheet")).map((s) => {
      const rid = s.getAttribute("r:id") || s.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      let t = rels[rid] || ""; t = t.startsWith("/") ? t.slice(1) : "xl/" + t.replace(/^\.\//, "");
      return { name: s.getAttribute("name"), path: t, hidden: (s.getAttribute("state") || "visible") !== "visible" };
    });
    const sstXml = await z.text("xl/sharedStrings.xml");
    const sst = sstXml ? Array.from(parseXml(sstXml).getElementsByTagName("si")).map((si) =>
      Array.from(si.getElementsByTagName("t")).filter((t) => !t.closest("rPh")).map((t) => t.textContent).join("")) : [];

    // Rows as arrays of values (like SheetJS header:1) plus each cell's formula text.
    async function readSheet(name) {
      const sh = sheets.find((s) => s.name === name); if (!sh) throw new Error(`No sheet named "${name}".`);
      const doc = parseXml(await z.text(sh.path));
      const rows = [], formulas = {};
      for (const row of doc.getElementsByTagName("row")) {
        const r = Number(row.getAttribute("r")); const arr = [];
        for (const c of kids(row, "c")) {
          const ref = c.getAttribute("r"); const { col } = splitRef(ref);
          const t = c.getAttribute("t"); const v = kids(c, "v")[0]; const f = kids(c, "f")[0];
          if (f && f.textContent) formulas[ref] = f.textContent;
          let val = null;
          if (t === "s") val = v ? sst[Number(v.textContent)] ?? null : null;
          else if (t === "inlineStr") val = Array.from(c.getElementsByTagName("t")).map((x) => x.textContent).join("");
          else if (t === "str") val = v ? v.textContent : null;
          else if (t === "b") val = v ? v.textContent === "1" : null;
          else if (t === "e") val = null;
          else if (v && v.textContent !== "") val = Number(v.textContent);
          arr[col] = val;
        }
        for (let i = 0; i < arr.length; i++) if (arr[i] === undefined) arr[i] = null;
        rows[r - 1] = arr;
      }
      for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
      return { rows, formulas };
    }

    // edits: [{ ref: "Y219", value: 2 }]  (number, or null to clear).
    // Text-level edits: only the edited <c> elements change; every other byte of the sheet stays identical.
    // A formula in an edited cell is kept and only its cached result changes.
    async function patch(name, edits) {
      const sh = sheets.find((s) => s.name === name); if (!sh) throw new Error(`No sheet named "${name}".`);
      let xml = await z.text(sh.path);
      for (const ed of edits) {
        if (!/^[A-Z]+\d+$/.test(ed.ref)) throw new Error("Bad cell " + ed.ref);
        if (ed.value != null && ed.value !== "" && (typeof ed.value !== "number" || !isFinite(ed.value))) throw new Error("Only numbers can be written.");
        xml = setCell(xml, ed.ref, ed.value == null || ed.value === "" ? null : ed.value);
      }
      // Ask Excel to recalculate every formula when the file is next opened.
      let wb = wbXml;
      const m = /<calcPr\b[^>]*?\/?>/.exec(wb);
      if (m) {
        const tag = m[0];
        const nt = /fullCalcOnLoad=/.test(tag) ? tag.replace(/fullCalcOnLoad="[^"]*"/, 'fullCalcOnLoad="1"') : tag.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
        wb = wb.slice(0, m.index) + nt + wb.slice(m.index + tag.length);
      } else {
        wb = wb.replace("</sheets>", '</sheets><calcPr fullCalcOnLoad="1"/>');
      }
      return zip(z.entries, { [sh.path]: xml, "xl/workbook.xml": wb });
    }

    function setCell(xml, ref, value) {
      const { row, col } = splitRef(ref);
      const cellRe = new RegExp(`<c r="${ref}"(?=[\\s/>])([^>]*?)(\\/>|>([\\s\\S]*?)<\\/c>)`);
      const m = cellRe.exec(xml);
      const build = (attrs, inner) => {
        const f = /<f\b[^>]*\/>|<f\b[^>]*>[\s\S]*?<\/f>/.exec(inner || "");
        let a = attrs || "";
        if (value != null || !f) a = a.replace(/\s+t="[^"]*"/, "");
        const body = (f ? f[0] : "") + (value != null ? `<v>${value}</v>` : "");
        return body ? `<c r="${ref}"${a}>${body}</c>` : `<c r="${ref}"${a}/>`;
      };
      if (m) return xml.slice(0, m.index) + build(m[1], m[3]) + xml.slice(m.index + m[0].length);
      if (value == null) return xml;                                   // nothing to clear
      const cell = build("", "");
      const rowRe = new RegExp(`<row r="${row}"(?=[\\s/>])([^>]*?)(\\/>|>([\\s\\S]*?)<\\/row>)`);
      const rm = rowRe.exec(xml);
      if (rm) {
        if (rm[2] === "/>") return xml.slice(0, rm.index) + `<row r="${row}"${rm[1]}>${cell}</row>` + xml.slice(rm.index + rm[0].length);
        const inner = rm[3], innerStart = rm.index + rm[0].length - "</row>".length - inner.length;
        const cr = /<c r="([A-Z]+)\d+"/g; let k, at = inner.length;
        while ((k = cr.exec(inner))) if (colIndex(k[1]) > col) { at = k.index; break; }
        return xml.slice(0, innerStart + at) + cell + xml.slice(innerStart + at);
      }
      const newRow = `<row r="${row}">${cell}</row>`;
      const rr = /<row r="(\d+)"/g; let k;
      while ((k = rr.exec(xml))) if (Number(k[1]) > row) return xml.slice(0, k.index) + newRow + xml.slice(k.index);
      if (/<sheetData\s*\/>/.test(xml)) return xml.replace(/<sheetData\s*\/>/, `<sheetData>${newRow}</sheetData>`);
      return xml.replace("</sheetData>", newRow + "</sheetData>");
    }

    return { sheets, readSheet, patch };
  }

  const api = { open, colIndex, colLetter, splitRef, crc32, unzip, zip };
  if (typeof module !== "undefined" && module.exports) module.exports = api; else root.XlsxLite = api;
})(typeof window !== "undefined" ? window : globalThis);
