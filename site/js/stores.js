/*
 * Inventory stores: where the LIVE_INVENTORY workbook lives and how cells are written.
 *   FileStore     - personal OneDrive (or any drive): download .xlsx, change only the needed cells, upload with an
 *                   If-Match check so two people can't overwrite each other (retries on conflict).
 *   WorkbookStore - OneDrive for Business / SharePoint: Microsoft Graph Excel API writes single cells in place.
 *   MemoryDrive   - demo copy kept in the browser.
 * Every write re-reads the lot's row first, so stock checks always use the latest numbers.
 */
(function (root) {
  const { parseInventory, colLetter, colIndex } = root.StoneSheet;
  const X = root.XlsxLite;
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const clean = (v) => (v == null ? "" : String(v).replace(/ /g, " ").replace(/\s+/g, " ").trim());
  const num = (v) => { if (typeof v === "number" && isFinite(v)) return v; const s = clean(v); return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : null; };
  const sum = (a) => a.reduce((t, b) => t + b, 0);
  const mgrQty = (t) => { const m = /^\s*(\d+)/.exec(t || ""); return m ? Number(m[1]) : 0; };
  const err = (msg, extra) => Object.assign(new Error(msg), extra || {});

  /* ---------------- Microsoft Graph ---------------- */
  function graph(getToken) {
    return async function g(path, opt = {}) {
      const token = await getToken();
      if (!token) throw err("Please sign in with Microsoft.", { code: "auth" });
      const r = await fetch(path.startsWith("http") ? path : GRAPH + path, { ...opt, headers: { authorization: "Bearer " + token, ...(opt.headers || {}) } });
      if (r.status === 401) throw err("Your Microsoft sign-in expired. Sign in again.", { code: "auth" });
      if (!r.ok) {
        const d = await r.json().catch(() => ({}));
        let msg = String(d?.error?.message || "");
        if (!msg || /^\s*</.test(msg)) msg = `OneDrive refused the request (error ${r.status}).`;   // never show raw HTML
        throw err(msg, { status: r.status, code: d?.error?.code });
      }
      if (opt.raw) return r;
      return r.status === 204 ? null : r.json();
    };
  }

  async function locate(g, cfg) {
    let item = null, shareErr = null, pathErr = null;
    const url = String(cfg.shareUrl || "").trim(), path = String(cfg.filePath || "").trim();
    if (url) {
      // A browser address like onedrive.live.com/edit?...resid=ABC!123 can be opened by its item id.
      const resid = /[?&](?:resid|id)=([0-9A-Fa-f]+(?:!|%21)[0-9A-Za-z]+)/.exec(url);
      try {
        if (resid) { const id = decodeURIComponent(resid[1]); item = await g(`/drives/${id.split("!")[0]}/items/${encodeURIComponent(id)}`); }
        else {
          const enc = "u!" + btoa(unescape(encodeURIComponent(url))).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
          item = await g(`/shares/${enc}/driveItem`);
        }
      } catch (e) { if (e.code === "auth") throw e; shareErr = e; }
    }
    if (!item && path) {
      try { item = await g(`/me/drive/root:${encodeURI(path.startsWith("/") ? path : "/" + path)}`); }
      catch (e) { if (e.code === "auth") throw e; pathErr = e; }
    }
    if (!item) {
      if (!url && !path) throw err("Set shareUrl or filePath in js/config.js.");
      const tips = [];
      if (url) tips.push(/1drv\.ms|sharepoint\.com|\/:x:\//i.test(url) ? "The sharing link didn't open. Check it is a \"Can edit\" link shared with the account you signed in with."
        : "The link in shareUrl isn't a sharing link. In OneDrive, right-click the file, choose Share, then Copy link, and paste that link (it starts with https://1drv.ms/).");
      if (path) tips.push(`The file wasn't found at filePath "${path}" in your OneDrive (it must be the path inside OneDrive, e.g. /Documents/App/LIVE_INVENTORY.xlsx).`);
      throw err(tips.join(" "), { status: (shareErr || pathErr)?.status });
    }
    return { driveId: item.parentReference.driveId, itemId: item.id, parentId: item.parentReference.id,
      driveType: item.parentReference.driveType, name: item.name, webUrl: item.webUrl };
  }

  /* ---------------- drives (file transport + sales log) ---------------- */
  /* CSV files kept next to the workbook: the sales log and the sales orders. Opened fine by Excel. */
  const TABLES = {
    log: [["at", "Time"], ["who", "Staff"], ["action", "Action"], ["lot", "Lot #"], ["color", "Color"], ["qty", "Qty", 1], ["note", "Note"], ["cell", "Cell"]],
    orders: [["order", "Order #"], ["date", "Date"], ["client", "Client"], ["phone", "Client phone"], ["lot", "Lot #"], ["color", "Color"], ["size", "Size (in)"],
      ["qty", "Slabs", 1], ["sqft", "Sq ft", 1], ["price", "Price per slab", 1], ["total", "Line total", 1], ["cell", "Excel cell"], ["status", "Status"], ["staff", "Staff"], ["notes", "Notes"]],
  };
  const csvEsc = (v) => { const s = String(v ?? ""); return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  const toCsv = (cols, row) => cols.map(([k]) => csvEsc(row[k])).join(",");
  function parseCsv(cols, text) {
    const out = []; const lines = String(text || "").replace(/^﻿/, "").split(/\r?\n/).filter(Boolean);
    for (const line of lines.slice(1)) {
      const f = []; let cur = "", q = false;
      for (let i = 0; i < line.length; i++) { const ch = line[i];
        if (q) { if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; } else if (ch === '"') q = false; else cur += ch; }
        else if (ch === '"') q = true; else if (ch === ",") { f.push(cur); cur = ""; } else cur += ch; }
      f.push(cur);
      const row = {}; cols.forEach(([k, , isNum], i) => { row[k] = isNum ? (f[i] === "" || f[i] == null ? null : Number(f[i])) : (f[i] ?? ""); });
      out.push(row);
    }
    return out;
  }

  // One CSV file in the workbook's folder, appended safely (If-Match + retry) so two people never overwrite each other.
  function csvFile(g, loc, name, cols) {
    const path = `/drives/${loc.driveId}/items/${loc.parentId}:/${encodeURIComponent(name)}:`;
    let eTag = null, text = null, ok = true, lastError = "";
    async function read() {
      if (!ok) return [];
      try {
        const meta = await g(`${path}?$select=eTag`);
        const r = await g(`${path}/content`, { raw: true });
        eTag = meta.eTag; text = await r.text();
      } catch (e) { if (e.status === 404) { eTag = null; text = ""; } else { ok = false; lastError = e.message; return []; } }
      return parseCsv(cols, text);
    }
    async function append(rows) {
      if (!ok) return false;
      for (let attempt = 0; attempt < 4; attempt++) {
        if (text == null) await read();
        if (!ok) return false;
        const head = text ? "" : "﻿" + cols.map(([, label]) => csvEsc(label)).join(",") + "\r\n";
        const body = (text || "") + head + rows.map((r) => toCsv(cols, r)).join("\r\n") + "\r\n";
        try {
          const d = await g(`${path}/content`, { method: "PUT", body, headers: { "content-type": "text/csv", ...(eTag ? { "if-match": eTag } : {}) } });
          eTag = d.eTag; text = body; return true;
        } catch (e) {
          if (e.status === 412) { text = null; continue; }
          if (e.status === 403 || e.status === 404) { ok = false; lastError = "No permission to create files in the Excel file's folder."; return false; }
          throw e;
        }
      }
      return false;
    }
    return { name, read, append, error: () => lastError };
  }
  function memoryCsv(name) { const rows = []; return { name, read: async () => rows.slice(), append: async (r) => { rows.push(...r); return true; }, error: () => "" }; }

  function GraphDrive(g, loc, cfg) {
    const item = `/drives/${loc.driveId}/items/${loc.itemId}`;
    const log = csvFile(g, loc, cfg.salesLogName || "StoneFlow sales log.csv", TABLES.log);
    const orders = csvFile(g, loc, cfg.ordersFileName || "StoneFlow sales orders.csv", TABLES.orders);
    return {
      log, orders,
      describe: () => ({ name: loc.name, webUrl: loc.webUrl, driveType: loc.driveType }),
      async head() { const d = await g(`${item}?$select=eTag,lastModifiedDateTime`); return { eTag: d.eTag, modified: d.lastModifiedDateTime }; },
      async download() {
        const d = await g(item);
        const url = d["@microsoft.graph.downloadUrl"];
        const r = url ? await fetch(url) : await g(`${item}/content`, { raw: true });
        if (!r.ok) throw err(`Couldn't download the Excel file (${r.status}).`);
        return { buf: new Uint8Array(await r.arrayBuffer()), eTag: d.eTag };
      },
      async upload(bytes, eTag) {
        const d = await g(`${item}/content`, { method: "PUT", body: bytes,
          headers: { "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ...(eTag ? { "if-match": eTag } : {}) } });
        return { eTag: d.eTag };
      },
    };
  }

  function MemoryDrive(bytes, name) {
    let buf = new Uint8Array(bytes), tag = 1;
    return {
      log: memoryCsv("sales log"), orders: memoryCsv("sales orders"),
      describe: () => ({ name, webUrl: null, driveType: "demo" }),
      async head() { return { eTag: String(tag) }; },
      async download() { return { buf: buf.slice(), eTag: String(tag) }; },
      async upload(b, eTag) { if (eTag && eTag !== String(tag)) throw err("Changed by someone else.", { status: 412 }); buf = new Uint8Array(b); tag++; return { eTag: String(tag) }; },
      bytes: () => buf.slice(),
    };
  }

  /* ---------------- sheet logic shared by both stores ---------------- */
  function parseWith(rows, formulas) {
    let p = parseInventory(rows);
    const col = p.columns.sales;
    if (col) {
      const first = p.lots.find((l) => !l.outside);
      const f = first && formulas[col + first.row];
      if (f) p = parseInventory(rows, { salesFormula: f.replace(/^=/, "") });
    }
    return p;
  }

  const MON = { JAN: 0, FEB: 1, MAR: 2, APR: 3, MAY: 4, JUN: 5, JUL: 6, AUG: 7, SEP: 8, OCT: 9, NOV: 10, DEC: 11 };
  function orderSheets(names, want) {
    if (want) { const w = names.find((n) => n.toLowerCase() === String(want).toLowerCase()); if (w) return [w, ...names.filter((n) => n !== w)]; }
    const key = (n) => { const m = /\b([A-Za-z]{3})[A-Za-z]*\.?[\s\-']*(\d{4}|\d{2})\b/.exec(n); const mo = m && MON[m[1].toUpperCase()];
      return mo != null ? Number(m[2].length === 2 ? "20" + m[2] : m[2]) * 12 + mo : -1; };
    return names.map((n, i) => ({ n, k: key(n), i })).sort((a, b) => b.k - a.k || a.i - b.i).map((x) => x.n);
  }

  function buildModel(p, extra) {
    const lots = p.lots.map((l) => ({ ...l, id: "r" + l.row,
      stock: l.op != null ? l.op + (l.purchase || 0) - sum(l.entries) : l.stock }));   // exact even if Excel hasn't recalculated
    const meta = { ...extra, columns: p.columns, salesStart: p.salesStart, salesEnd: p.salesEnd, holdPeople: p.holdPeople,
      holdColumns: p.holdColumns, sections: p.sections, warnings: p.warnings };
    return { meta, lots };
  }

  // Read the lot's row fresh and work out the cell edits for one change.
  function planner(meta) {
    const C = meta.columns;
    const val = (row, letter) => (letter ? row[colIndex(letter)] : null);
    function state(row) {
      const s = colIndex(meta.salesStart), e = colIndex(meta.salesEnd);
      const entries = []; let firstEmpty = null;
      for (let i = s; i <= e; i++) { const n = num(row[i]); if (n != null && n !== 0) entries.push(n); else if (n == null && clean(row[i]) === "" && firstEmpty == null) firstEmpty = i; }
      const op = num(val(row, C.op)), pur = num(val(row, C.purchase));
      const holds = {}; for (const [p, col] of Object.entries(meta.holdColumns || {})) { const n = num(val(row, col)); if (n) holds[p] = n; }
      return { lot: clean(val(row, C.lot)), name: clean(val(row, C.name)), entries, firstEmpty, op, pur,
        stock: op != null ? op + (pur || 0) - sum(entries) : num(val(row, C.stock)), holds, mgrHold: clean(val(row, C.mgr)) };
    }
    function findRow(rows, lot) {
      const at = rows[lot.row - 1] || [];
      const same = (r) => clean(val(r, C.lot)) === clean(lot.lot) && clean(val(r, C.name)) === clean(lot.name);
      if (same(at)) return lot.row;
      const hits = []; rows.forEach((r, i) => { if (r && same(r)) hits.push(i + 1); });
      if (hits.length === 1) return hits[0];
      throw err(`The Excel sheet changed and lot ${lot.lot || lot.name} moved. Tap Refresh and try again.`);
    }
    // cached: formula cells whose stored result we refresh so the file reads correctly even before Excel recalculates
    function cachedEdits(r, row, formulas, stock, newSum, personHolds) {
      const out = [];
      if (C.stock) out.push({ ref: C.stock + r, value: stock, cached: !!formulas[C.stock + r] });
      if (C.sales) out.push({ ref: C.sales + r, value: newSum, cached: !!formulas[C.sales + r] });
      if (C.balance && formulas[C.balance + r]) out.push({ ref: C.balance + r, value: stock - personHolds, cached: true });
      return out;
    }
    return {
      sale(ctx, lot, qty) {
        const r = findRow(ctx.rows, lot), row = ctx.rows[r - 1];
        if (lot.outside) {
          const cur = num(val(row, lot.qtyCol)) ?? 0;
          if (qty > 0 && qty > cur) throw err(`Only ${Math.max(cur, 0)} slab(s) in stock for ${lot.lot || lot.name}.`);
          return { edits: [{ ref: lot.qtyCol + r, value: cur - qty }], change: { kind: "outside", ref: lot.qtyCol + r, qty, row: r }, stockAfter: cur - qty };
        }
        const st = state(row);
        if (qty > 0 && qty > st.stock) throw err(`Only ${Math.max(st.stock, 0)} slab(s) in stock for ${lot.lot || lot.name}.`);
        if (st.firstEmpty == null) throw err(`Row ${r} has no empty sale columns left (${meta.salesStart}–${meta.salesEnd}). Start a new month sheet.`);
        const ref = colLetter(st.firstEmpty) + r, newSum = sum(st.entries) + qty;
        const stock = st.op != null ? st.op + (st.pur || 0) - newSum : st.stock - qty;
        const personHolds = sum(Object.values(st.holds));
        return { edits: [{ ref, value: qty }, ...cachedEdits(r, row, ctx.formulas, stock, newSum, personHolds)],
          change: { kind: "sale", ref, qty, row: r }, stockAfter: stock, heldInvolved: Math.max(0, qty - Math.max(0, st.stock - personHolds - mgrQty(st.mgrHold))) };
      },
      undo(ctx, lot, change) {
        const r = findRow(ctx.rows, lot), row = ctx.rows[r - 1];
        if (change.kind === "outside") { const cur = num(val(row, lot.qtyCol)) ?? 0; return { edits: [{ ref: lot.qtyCol + r, value: cur + change.qty }] }; }
        if (change.kind === "hold") return this.hold(ctx, lot, change.person, change.prev, true);
        const ref = change.ref.replace(/\d+$/, "") + r;
        if (num(row[colIndex(ref.replace(/\d+$/, ""))]) !== change.qty) throw err(`Cell ${ref} was changed in Excel after this sale, so it wasn't undone. Check it in the sheet.`);
        const st = state(row);
        const newSum = sum(st.entries) - change.qty;
        const stock = st.op != null ? st.op + (st.pur || 0) - newSum : st.stock + change.qty;
        return { edits: [{ ref, value: null }, ...cachedEdits(r, row, ctx.formulas, stock, newSum, sum(Object.values(st.holds)))] };
      },
      hold(ctx, lot, person, qty, force) {
        const col = meta.holdColumns?.[person]; if (!col) throw err(`No hold column for ${person}.`);
        const r = findRow(ctx.rows, lot), row = ctx.rows[r - 1], st = state(row);
        const prev = st.holds[person] || 0;
        const holds = { ...st.holds, [person]: qty }, personHolds = sum(Object.values(holds));
        if (!force && qty > prev && personHolds + mgrQty(st.mgrHold) > st.stock) throw err(`Holds (${personHolds + mgrQty(st.mgrHold)}) would be more than the ${st.stock} in stock.`);
        const edits = [{ ref: col + r, value: qty || null }];
        if (C.balance && ctx.formulas[C.balance + r]) edits.push({ ref: C.balance + r, value: st.stock - personHolds, cached: true });
        return { edits, change: { kind: "hold", ref: col + r, person, prev, qty, row: r } };
      },
    };
  }

  /* ---------------- FileStore (personal OneDrive, demo) ---------------- */
  function FileStore(drive, cfg) {
    let sheet = null, eTag = null, meta = null, plan = null;
    async function readSheetFrom(wb, name) { const sh = await wb.readSheet(name); return { rows: sh.rows, formulas: sh.formulas }; }
    async function load(wantSheet) {
      const d = await drive.download();
      const wb = await X.open(d.buf);
      const names = wb.sheets.filter((s) => !s.hidden).map((s) => s.name);
      let found = null;
      for (const n of orderSheets(names, wantSheet || sheet || cfg.sheet)) {
        try { const sh = await readSheetFrom(wb, n); found = { n, p: parseWith(sh.rows, sh.formulas) }; break; } catch {}
      }
      if (!found) throw err('No sheet in this file has "LOT #" and "STOCK" headings.');
      sheet = found.n; eTag = d.eTag;
      const m = buildModel(found.p, { ...drive.describe(), sheet, sheets: names, mode: drive.describe().driveType === "demo" ? "demo" : "file" });
      meta = m.meta; plan = planner(meta);
      return m;
    }
    async function mutate(make) {
      for (let attempt = 0; attempt < 4; attempt++) {
        const d = await drive.download();
        const wb = await X.open(d.buf);
        const ctx = await readSheetFrom(wb, sheet);
        const res = make(ctx);
        const bytes = await wb.patch(sheet, res.edits);
        try { const u = await drive.upload(bytes, d.eTag); eTag = u.eTag; return res; }
        catch (e) {
          if (e.status === 412 && attempt < 3) { await new Promise((r) => setTimeout(r, 300 + Math.random() * 700)); continue; }
          if (e.status === 423) throw err("The Excel file is locked by someone editing it in desktop Excel. Ask them to turn on AutoSave or close it, then try again.");
          throw e;
        }
      }
      throw err("Too many people are changing the file right now. Try again.");
    }
    return {
      get meta() { return meta; },
      load,
      async changed() { try { return (await drive.head()).eTag !== eTag; } catch { return false; } },
      sale: (lot, qty) => mutate((ctx) => plan.sale(ctx, lot, qty)),
      hold: (lot, person, qty) => mutate((ctx) => plan.hold(ctx, lot, person, qty)),
      undo: (lot, change) => mutate((ctx) => plan.undo(ctx, lot, change)),
      log: drive.log, orders: drive.orders,
    };
  }

  /* ---------------- WorkbookStore (OneDrive for Business / SharePoint) ---------------- */
  function WorkbookStore(g, loc, drive, cfg) {
    const wbPath = `/drives/${loc.driveId}/items/${loc.itemId}/workbook`;
    const ws = (n) => `${wbPath}/worksheets('${encodeURIComponent(n.replace(/'/g, "''"))}')`;
    let sheet = null, meta = null, plan = null, eTag = null, lastCol = "BL";
    function toRows(range) {
      const m = /!?\$?([A-Z]+)\$?(\d+)/.exec(range.address.split("!").pop());
      const c0 = colIndex(m[1]), r0 = Number(m[2]);
      const rows = [], formulas = {};
      range.values.forEach((vals, i) => {
        const arr = []; vals.forEach((v, j) => { arr[c0 + j] = v === "" ? null : v; });
        for (let k = 0; k < arr.length; k++) if (arr[k] === undefined) arr[k] = null;
        rows[r0 - 1 + i] = arr;
        (range.formulas?.[i] || []).forEach((f, j) => { if (typeof f === "string" && f.startsWith("=")) formulas[colLetter(c0 + j) + (r0 + i)] = f.slice(1); });
      });
      for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
      return { rows, formulas };
    }
    async function readAll(n) { return toRows(await g(`${ws(n)}/usedRange?$select=address,values,formulas`)); }
    async function load(wantSheet) {
      const list = await g(`${wbPath}/worksheets?$select=name,visibility`);
      const names = list.value.filter((s) => s.visibility === "Visible").map((s) => s.name);
      let found = null;
      for (const n of orderSheets(names, wantSheet || sheet || cfg.sheet)) {
        try { const sh = await readAll(n); found = { n, p: parseWith(sh.rows, sh.formulas) }; break; } catch (e) { if (e.code === "auth") throw e; }
      }
      if (!found) throw err('No sheet in this workbook has "LOT #" and "STOCK" headings.');
      sheet = found.n;
      try { eTag = (await drive.head()).eTag; } catch {}
      const m = buildModel(found.p, { ...drive.describe(), sheet, sheets: names, mode: "workbook" });
      meta = m.meta; plan = planner(meta);
      lastCol = colLetter(Math.max(colIndex(meta.salesEnd || "BL"), colIndex("T")));
      return m;
    }
    async function freshCtx(lot) {
      const range = await g(`${ws(sheet)}/range(address='A${lot.row}:${lastCol}${lot.row}')?$select=address,values,formulas`);
      const ctx = toRows(range);
      const row = ctx.rows[lot.row - 1] || [];
      if (clean(row[colIndex(meta.columns.lot)]) === clean(lot.lot) && clean(row[colIndex(meta.columns.name)]) === clean(lot.name)) return ctx;
      return readAll(sheet);                                           // row moved: read the whole sheet to find it
    }
    async function apply(res) {
      for (const ed of res.edits) {
        if (ed.cached) continue;                                        // Excel recalculates formula cells itself
        await g(`${ws(sheet)}/range(address='${ed.ref}')`, { method: "PATCH", headers: { "content-type": "application/json" },
          body: JSON.stringify({ values: [[ed.value == null ? "" : ed.value]] }) });
      }
      try { eTag = (await drive.head()).eTag; } catch {}
      return res;
    }
    return {
      get meta() { return meta; },
      load,
      async changed() { try { return (await drive.head()).eTag !== eTag; } catch { return false; } },
      sale: async (lot, qty) => apply(plan.sale(await freshCtx(lot), lot, qty)),
      hold: async (lot, person, qty) => apply(plan.hold(await freshCtx(lot), lot, person, qty)),
      undo: async (lot, change) => apply(plan.undo(await freshCtx(lot), lot, change)),
      log: drive.log, orders: drive.orders,
    };
  }

  /* ---------------- connect ---------------- */
  async function connect(cfg, getToken) {
    const g = graph(getToken);
    const loc = await locate(g, cfg);
    const drive = GraphDrive(g, loc, cfg);
    const business = loc.driveType === "business" || loc.driveType === "documentLibrary";
    const useWorkbook = cfg.mode === "workbook" || (cfg.mode !== "file" && business);
    return useWorkbook ? WorkbookStore(g, loc, drive, cfg) : FileStore(drive, cfg);
  }
  function demo(bytes, name) { const d = MemoryDrive(bytes, name); const s = FileStore(d, {}); s.drive = d; return s; }

  root.InventoryStores = { connect, demo, graph, locate, FileStore, WorkbookStore, MemoryDrive, GraphDrive, planner, orderSheets, TABLES, parseCsv };
})(typeof window !== "undefined" ? window : globalThis);
