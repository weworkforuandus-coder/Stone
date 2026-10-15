// /api/inventory: the read-only stock list anyone can see, no sign-in needed.
// Read from the Excel file with the owner's connection, then cleaned up for the public:
// no rack locations, no names of who holds slabs, no sales orders or logs, no sheet notes.
//   GET ?op=head        -> { eTag }   (cheap check for changes)
//   GET [?sheet=NAME]   -> { meta, lots, eTag }
import { json, locateFile, ownerGraph, graphError, rateLimited, ipOf } from "../lib/common.mjs";
import Inv from "../lib/sheet.cjs";

const cache = new Map(); // sheet -> { eTag, at, body }
let headCache = null;

async function currentTag(loc) {
  if (headCache && Date.now() - headCache.at < 8000) return headCache.eTag;
  const r = await ownerGraph(`/drives/${loc.driveId}/items/${loc.itemId}?$select=eTag`); if (!r.ok) throw await graphError(r);
  headCache = { eTag: (await r.json()).eTag, at: Date.now() };
  return headCache.eTag;
}

function serverDrive(loc) {
  return {
    log: null, orders: null,
    describe: () => ({ name: loc.name, webUrl: null, driveType: loc.driveType }),
    async head() { return { eTag: await currentTag(loc) }; },
    async download() {
      const r = await ownerGraph(`/drives/${loc.driveId}/items/${loc.itemId}?$select=eTag,@microsoft.graph.downloadUrl`); if (!r.ok) throw await graphError(r);
      const d = await r.json();
      const f = d["@microsoft.graph.downloadUrl"] ? await fetch(d["@microsoft.graph.downloadUrl"]) : await ownerGraph(`/drives/${loc.driveId}/items/${loc.itemId}/content`);
      if (!f.ok) throw await graphError(f);
      headCache = { eTag: d.eTag, at: Date.now() };
      return { buf: new Uint8Array(await f.arrayBuffer()), eTag: d.eTag };
    },
    async upload() { throw new Error("read only"); },
  };
}

export function publicView(model, sheet) {
  const held = (l) => globalThis.StoneSheet.heldOf(l);
  const lots = model.lots.map((l) => {
    const stock = Math.max(0, Number(l.stock) || 0);
    return { id: l.id, lot: l.lot || "", name: l.name || "", section: l.section || "", L: l.L ?? "", H: l.H ?? "", bm: !!l.bm,
      stock, heldTotal: Math.min(stock, Math.max(0, held(l))), ...(l.outside ? { outside: true } : {}) };
  });
  const m = model.meta;
  return { meta: { name: m.name, sheet, sheets: m.sheets, sections: m.sections, driveType: m.driveType, mode: "view" }, lots };
}

export default async (req, context) => {
  try {
    if (req.method !== "GET") return json({ error: "Use GET." }, 405);
    if (rateLimited("inv:" + ipOf(req, context), 90, 60_000)) return json({ error: "Too many requests. Wait a minute." }, 429);
    const url = new URL(req.url), loc = await locateFile();
    if (url.searchParams.get("op") === "head") return json({ eTag: await currentTag(loc) });

    const want = (url.searchParams.get("sheet") || "").slice(0, 60), key = want || "*";
    const tag = await currentTag(loc), hit = cache.get(key);
    if (hit && hit.eTag === tag) return json({ ...hit.body, eTag: tag });

    const store = Inv.FileStore(serverDrive(loc), { sheet: process.env.EXCEL_SHEET || "" });
    const model = await store.load(want || undefined);
    const body = publicView(model, model.meta.sheet);
    if (cache.size > 20) cache.clear();
    cache.set(key, { eTag: headCache?.eTag || tag, at: Date.now(), body });
    return json({ ...body, eTag: headCache?.eTag || tag });
  } catch (e) {
    return json({ error: e.message, code: e.code }, e.status && e.status >= 400 && e.status < 600 ? e.status : 500);
  }
};

export const config = { path: "/api/inventory" };
