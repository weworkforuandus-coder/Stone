// /api/file: editors only. Reads and writes LIVE_INVENTORY.xlsx and its two CSV files (sales log, sales orders)
// with the owner's OneDrive connection. Nothing else in the owner's OneDrive can be reached through here.
//   GET  ?op=head                 -> { eTag, modified, name, webUrl, driveType }
//   GET  ?op=download             -> the .xlsx bytes (header x-etag)
//   PUT  ?op=upload  (If-Match)   -> { eTag }   412 if someone saved first, 423 if locked
//   GET  ?op=csv&name=log|orders  -> CSV text (header x-etag), 404 if not created yet
//   PUT  ?op=csv&name=log|orders  -> { eTag }
import { json, requireEditor, locateFile, ownerGraph, graphError, csvName } from "../lib/common.mjs";

const XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const MAX = 5 * 1024 * 1024;

export default async (req) => {
  try {
    await requireEditor(req);
    const url = new URL(req.url), op = url.searchParams.get("op"), loc = await locateFile();
    const item = `/drives/${loc.driveId}/items/${loc.itemId}`;

    if (req.method === "GET" && op === "head") {
      const r = await ownerGraph(`${item}?$select=eTag,lastModifiedDateTime`); if (!r.ok) throw await graphError(r);
      const d = await r.json();
      return json({ eTag: d.eTag, modified: d.lastModifiedDateTime, name: loc.name, webUrl: loc.webUrl, driveType: loc.driveType });
    }
    if (req.method === "GET" && op === "download") {
      const r = await ownerGraph(`${item}?$select=eTag,@microsoft.graph.downloadUrl`); if (!r.ok) throw await graphError(r);
      const d = await r.json();
      const f = d["@microsoft.graph.downloadUrl"] ? await fetch(d["@microsoft.graph.downloadUrl"]) : await ownerGraph(`${item}/content`);
      if (!f.ok) throw await graphError(f);
      return new Response(await f.arrayBuffer(), { headers: { "content-type": XLSX, "x-etag": d.eTag, "cache-control": "no-store" } });
    }
    if (req.method === "PUT" && op === "upload") {
      const body = await req.arrayBuffer();
      if (!body.byteLength || body.byteLength > MAX) return json({ error: "Bad file." }, 400);
      const head = new Uint8Array(body.slice(0, 2)); if (head[0] !== 0x50 || head[1] !== 0x4b) return json({ error: "Not an .xlsx file." }, 400);
      const ifMatch = req.headers.get("if-match");
      const r = await ownerGraph(`${item}/content`, { method: "PUT", body, headers: { "content-type": XLSX, ...(ifMatch ? { "if-match": ifMatch } : {}) } });
      if (!r.ok) throw await graphError(r);
      const d = await r.json();
      return json({ eTag: d.eTag });
    }
    if (op === "csv") {
      const which = url.searchParams.get("name"); if (which !== "log" && which !== "orders") return json({ error: "Unknown file." }, 400);
      const path = `/drives/${loc.driveId}/items/${loc.parentId}:/${encodeURIComponent(csvName(which))}:`;
      if (req.method === "GET") {
        const m = await ownerGraph(`${path}?$select=eTag`);
        if (m.status === 404) return json({ error: "Not created yet." }, 404);
        if (!m.ok) throw await graphError(m);
        const meta = await m.json();
        const c = await ownerGraph(`${path}/content`); if (!c.ok) throw await graphError(c);
        return new Response(await c.text(), { headers: { "content-type": "text/csv; charset=utf-8", "x-etag": meta.eTag, "cache-control": "no-store" } });
      }
      if (req.method === "PUT") {
        const text = await req.text(); if (text.length > MAX) return json({ error: "Too big." }, 400);
        const ifMatch = req.headers.get("if-match");
        const r = await ownerGraph(`${path}/content`, { method: "PUT", body: text, headers: { "content-type": "text/csv", ...(ifMatch ? { "if-match": ifMatch } : {}) } });
        if (!r.ok) throw await graphError(r);
        return json({ eTag: (await r.json()).eTag });
      }
    }
    return json({ error: "Unknown request." }, 400);
  } catch (e) {
    return json({ error: e.message, code: e.code }, e.status && e.status >= 400 && e.status < 600 ? e.status : 500);
  }
};

export const config = { path: "/api/file" };
