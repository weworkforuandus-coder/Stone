/*
 * StoneFlow lot-sheet reader.
 * Understands the yard's LIVE_INVENTORY layout: one row per LOT with
 *   LOT # | COLOR | BM | L | H | LOC | STOCK | BALANCE | <salesperson hold columns> |
 *   SOLD IN BLOWOUT | HOLD BY SALES.MGR | OP | PURCHASE | SALES | ... sale entries ...
 * plus section header rows (GRANITE SLABS, CALACATTA'S-JUMBO SIZE ...), subtotal rows,
 * repeated header rows, and an OUTSIDE COLORS list with its own STK column.
 *
 * Input: rows as arrays of cell values (SheetJS sheet_to_json(ws, {header: 1, raw: true, defval: null})).
 * Works in browsers (window.StoneSheet) and Node (module.exports).
 */
(function (root) {
  const clean = (v) => (v == null ? "" : String(v).replace(/ /g, " ").replace(/\s+/g, " ").trim());
  const up = (v) => clean(v).toUpperCase();
  const num = (v) => {
    if (typeof v === "number" && isFinite(v)) return v;
    const s = clean(v);
    if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
    return Number(s);
  };
  const colLetter = (i) => { let s = ""; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const colIndex = (letters) => letters.toUpperCase().split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const fmtDate = (v) => {
    if (v instanceof Date && !isNaN(v)) return v.toISOString().slice(0, 10);
    if (typeof v === "string" && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
    return "";
  };

  // salesFormula: optional formula text from the SALES column, e.g. "SUM(X4:BL4)", to find where sale entries start.
  function parseInventory(rows, opts = {}) {
    const warnings = [];
    const hdrIdx = rows.findIndex((r) => (r || []).some((c) => up(c) === "LOT #"));
    if (hdrIdx < 0) throw new Error('Could not find the header row. The sheet needs a "LOT #" column heading.');
    const lotCol = rows[hdrIdx].findIndex((c) => up(c) === "LOT #");

    // Merge every repeated header row (the first one leaves L blank; later ones say "L").
    const width = Math.max(...rows.map((r) => (r || []).length));
    const head = Array.from({ length: width }, () => "");
    rows.forEach((r) => {
      if (!r || up(r[lotCol]) !== "LOT #") return;
      r.forEach((c, i) => { if (!head[i] && clean(c) && up(c) !== "STK") head[i] = up(c); });
    });
    const find = (test) => head.findIndex((h) => h && test(h));
    const C = {
      lot: lotCol,
      name: find((h) => h === "COLOR" || h === "COLOUR" || h === "NAME"),
      bm: find((h) => h === "BM"),
      len: find((h) => h === "L" || h === "LENGTH"),
      ht: find((h) => h === "H" || h === "HEIGHT" || h === "W" || h === "WIDTH"),
      loc: find((h) => h === "LOC" || h === "LOCATION"),
      stock: find((h) => h === "STOCK"),
      balance: find((h) => h.startsWith("BALANCE")),
      blowout: find((h) => h.startsWith("SOLD IN BLOWOUT")),
      mgr: find((h) => h.startsWith("HOLD")),
      op: find((h) => h === "OP" || h === "OPENING"),
      purchase: find((h) => h === "PURCHASE" || h === "PURCHASES"),
      sales: find((h) => h === "SALES"),
    };
    if (C.name < 0) C.name = lotCol + 1;
    if (C.len < 0 && C.bm >= 0) C.len = C.bm + 1;
    if (C.stock < 0) throw new Error('Could not find the "STOCK" column.');

    const holdCols = [];
    if (C.balance >= 0 && C.blowout > C.balance) {
      for (let i = C.balance + 1; i < C.blowout; i++) if (head[i]) holdCols.push({ i, name: head[i] });
    }
    let salesStart = -1, salesEnd = -1;
    const m = /SUM\(\s*\$?([A-Z]+)\$?\d+\s*:\s*\$?([A-Z]+)\$?\d+/i.exec(String(opts.salesFormula || ""));
    if (m) { salesStart = colIndex(m[1]); salesEnd = colIndex(m[2]); }
    if (salesStart < 0 && C.sales >= 0) salesStart = C.sales + 4;
    if (salesEnd < 0 && salesStart >= 0) salesEnd = colIndex("BL") > salesStart ? colIndex("BL") : salesStart + 40;

    const lots = [], sections = [];
    let section = "", outside = false, outsideQtyCol = -1, outsideDate = "";

    for (let r = hdrIdx + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const cell = (i) => (i >= 0 ? row[i] : null);
      if (!row.some((c) => clean(c) !== "")) continue;
      const lot = clean(cell(C.lot)), name = clean(cell(C.name));
      const excelRow = r + 1;

      if (up(lot) === "LOT #") {                       // repeated header row
        if (outside) outsideQtyCol = row.findIndex((c) => ["STK", "STOCK", "QTY"].includes(up(c)));
        continue;
      }
      if (!lot && !name) continue;                     // stray numbers
      if (/TOTAL/.test(up(name))) continue;            // subtotal rows

      if (outside) {
        const qty = num(cell(outsideQtyCol >= 0 ? outsideQtyCol : C.loc));
        if (qty == null) continue;
        lots.push({
          row: excelRow, lot, name, section, outside: true, bm: false, L: "", H: "", qtyCol: colLetter(outsideQtyCol >= 0 ? outsideQtyCol : C.loc),
          loc: "Outside" + (outsideDate ? ` (count ${outsideDate})` : ""),
          stock: qty, op: null, purchase: null, sheetSales: 0, entries: [], holds: {}, mgrHold: "", blowout: 0,
        });
        continue;
      }

      const stockV = num(cell(C.stock));
      const isLot = !!lot || stockV != null;
      if (!isLot) {
        if (num(cell(C.loc)) != null) continue;        // e.g. "BROWN FANTASY 114": a colour subtotal
        section = name; sections.push(name);
        if (/OUTSIDE/.test(up(name))) {
          outside = true;
          outsideDate = row.map(fmtDate).find(Boolean) || "";
        }
        continue;
      }

      const op = num(cell(C.op)), purchase = num(cell(C.purchase));
      const entries = [];
      if (salesStart >= 0) for (let i = salesStart; i < row.length; i++) { const v = num(row[i]); if (v != null && v !== 0) entries.push(v); }
      const sheetSales = num(cell(C.sales)) ?? entries.reduce((a, b) => a + b, 0);
      const stock = stockV ?? ((op ?? 0) + (purchase ?? 0) - sheetSales);
      const holds = {};
      for (const h of holdCols) { const v = num(row[h.i]); if (v) holds[h.name] = v; }
      const mgrHold = clean(cell(C.mgr));
      const lenV = cell(C.len), htV = cell(C.ht);

      const item = {
        row: excelRow, lot, name, section, outside: false,
        bm: up(cell(C.bm)) === "Y",
        L: lenV == null ? "" : clean(lenV), H: htV == null ? "" : clean(htV),
        loc: clean(cell(C.loc)),
        stock, op, purchase, sheetSales, entries, holds, mgrHold,
        blowout: num(cell(C.blowout)) ?? 0,
      };
      if (!lot) warnings.push(`Row ${excelRow} (${name}) has stock but no LOT #.`);
      if (stock < 0) warnings.push(`Row ${excelRow} ${lot} ${name}: stock is ${stock}.`);
      lots.push(item);
    }

    const seen = {};
    lots.forEach((l) => { if (l.lot) (seen[l.lot] = seen[l.lot] || []).push(l.row); });
    Object.entries(seen).forEach(([k, rs]) => { if (rs.length > 1) warnings.push(`LOT # ${k} appears on rows ${rs.join(", ")}.`); });

    return {
      columns: Object.fromEntries(Object.entries(C).map(([k, v]) => [k, v >= 0 ? colLetter(v) : null])),
      salesStart: salesStart >= 0 ? colLetter(salesStart) : null,
      salesEnd: salesEnd >= 0 ? colLetter(salesEnd) : null,
      holdPeople: holdCols.map((h) => h.name),
      holdColumns: Object.fromEntries(holdCols.map((h) => [h.name, colLetter(h.i)])),
      sections, lots, warnings,
    };
  }

  // Amount held on a lot: salesperson columns + the number in "HOLD BY SALES.MGR" ("10 puru").
  function heldOf(l) {
    const people = Object.values(l.holds || {}).reduce((a, b) => a + (Number(b) || 0), 0);
    const mm = /^\s*(\d+)/.exec(l.mgrHold || "");
    return people + (mm ? Number(mm[1]) : 0);
  }

  // Loose lot/code match: "iq 466", "IQ-466", "iq0466" all equal "iq466".
  const normCode = (s) => String(s ?? "").toLowerCase().replace(/[\s\-_.\/#]+/g, "").replace(/([a-z])0+(?=\d)/g, "$1");

  const api = { parseInventory, heldOf, normCode, colLetter, colIndex };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.StoneSheet = api;
})(typeof window !== "undefined" ? window : globalThis);
