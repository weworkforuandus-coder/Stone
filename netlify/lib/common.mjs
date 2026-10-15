// Shared server code for the Netlify functions.
// The site reads and writes LIVE_INVENTORY.xlsx with the OWNER's OneDrive connection, kept on the server.
// Visitors never get the owner's token: anyone can read a cleaned-up stock list, and only
// signed-in editors (ALLOWED_EMAILS) can change the file.
//
// Netlify environment variables:
//   MS_CLIENT_ID      Azure "Application (client) ID"
//   MS_CLIENT_SECRET  Azure client secret (Certificates & secrets), used only here
//   MS_TENANT         optional, default "consumers" (personal Microsoft accounts)
//   OWNER_EMAIL       the Microsoft account that owns the Excel file (only it can connect OneDrive)
//   EXCEL_PATH        path inside the owner's OneDrive, e.g. "/Stone/LIVE INVENTORY.xlsx"
//   ALLOWED_EMAILS    editors: staff who may record sales, holds and sales orders
//   SALES_LOG_NAME, ORDERS_FILE_NAME  optional CSV names next to the Excel file

export const GRAPH = "https://graph.microsoft.com/v1.0";
export const env = (k, d = "") => String(process.env[k] ?? d).trim();
export const json = (data, status = 200, headers = {}) =>
  new Response(JSON.stringify(data), { status, headers: { "content-type": "application/json", "cache-control": "no-store", ...headers } });
export const fail = (msg, status, extra) => Object.assign(new Error(msg), { status, ...(extra || {}) });
const emails = (s) => String(s || "").toLowerCase().split(/[\s,;]+/).filter(Boolean);

export const TENANT = () => env("MS_TENANT", "consumers");
export const authBase = () => `https://login.microsoftonline.com/${TENANT()}/oauth2/v2.0`;
export const OWNER_SCOPES = "offline_access User.Read Files.ReadWrite";

/* ---------- small key-value store (Netlify Blobs) for the owner's refresh token ---------- */
let blobStore = null;
async function kv() {
  if (globalThis.__sfTestKv) return globalThis.__sfTestKv;
  if (!blobStore) { const { getStore } = await import("@netlify/blobs"); blobStore = getStore("stoneflow"); }
  return blobStore;
}
export async function kvGet(key) { const s = await kv(); const v = await s.get(key); return v ? JSON.parse(v) : null; }
export async function kvSet(key, val) { const s = await kv(); await s.set(key, JSON.stringify(val)); }

/* ---------- the owner's OneDrive connection ---------- */
let ownerAccess = null; // { token, exp }
export async function tokenRequest(params) {
  const r = await fetch(`${authBase()}/token`, { method: "POST", headers: { "content-type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: env("MS_CLIENT_ID"), client_secret: env("MS_CLIENT_SECRET"), scope: OWNER_SCOPES, ...params }) });
  const d = await r.json().catch(() => ({}));
  if (!r.ok) throw fail(d.error_description?.split("\n")[0] || d.error || `Microsoft sign-in error ${r.status}`, 502, { code: d.error });
  return d;
}
export async function ownerToken() {
  if (ownerAccess && ownerAccess.exp > Date.now() + 60_000) return ownerAccess.token;
  const saved = await kvGet("owner");
  if (!saved?.refresh_token) throw fail("The inventory isn't connected to OneDrive yet. The owner needs to open /api/connect once.", 503, { code: "not_connected" });
  let d;
  try { d = await tokenRequest({ grant_type: "refresh_token", refresh_token: saved.refresh_token }); }
  catch (e) {
    if (e.code === "invalid_grant") throw fail("The OneDrive connection expired. The owner needs to open /api/connect again.", 503, { code: "not_connected" });
    throw e;
  }
  if (d.refresh_token && d.refresh_token !== saved.refresh_token) await kvSet("owner", { ...saved, refresh_token: d.refresh_token, refreshed: new Date().toISOString() });
  ownerAccess = { token: d.access_token, exp: Date.now() + (Number(d.expires_in) || 3600) * 1000 };
  return ownerAccess.token;
}
export function clearOwnerCache() { ownerAccess = null; fileLoc = null; }

// Graph call as the owner. Returns the Response (callers decide how to read it).
export async function ownerGraph(path, opt = {}) {
  const url = path.startsWith("http") ? path : GRAPH + path;
  let r = await fetch(url, { ...opt, headers: { authorization: "Bearer " + (await ownerToken()), ...(opt.headers || {}) } });
  if (r.status === 401) { ownerAccess = null; r = await fetch(url, { ...opt, headers: { authorization: "Bearer " + (await ownerToken()), ...(opt.headers || {}) } }); }
  return r;
}
export async function graphError(r) {
  const d = await r.json().catch(() => ({}));
  let msg = String(d?.error?.message || "");
  if (!msg || /^\s*</.test(msg)) msg = `OneDrive error ${r.status}`;
  if (r.status === 423) msg = "The Excel file is locked by someone editing it in Excel. Close it there (or wait a minute), then try again.";
  return fail(msg, r.status, { code: d?.error?.code });
}

/* ---------- where the Excel file is ---------- */
let fileLoc = null; // { driveId, itemId, parentId, name, webUrl, driveType, at }
export async function locateFile() {
  if (fileLoc && Date.now() - fileLoc.at < 10 * 60_000) return fileLoc;
  const p = env("EXCEL_PATH"); if (!p) throw fail("Set EXCEL_PATH in Netlify (e.g. /Stone/LIVE INVENTORY.xlsx).", 500);
  const r = await ownerGraph(`/me/drive/root:${encodeURI(p.startsWith("/") ? p : "/" + p)}?$select=id,name,webUrl,parentReference`);
  if (r.status === 404) throw fail(`The Excel file wasn't found at EXCEL_PATH "${p}" in the owner's OneDrive.`, 404);
  if (!r.ok) throw await graphError(r);
  const it = await r.json();
  fileLoc = { driveId: it.parentReference.driveId, itemId: it.id, parentId: it.parentReference.id, name: it.name, webUrl: it.webUrl,
    driveType: it.parentReference.driveType, at: Date.now() };
  return fileLoc;
}
export const csvName = (which) => which === "orders" ? env("ORDERS_FILE_NAME", "StoneFlow sales orders.csv") : env("SALES_LOG_NAME", "StoneFlow sales log.csv");

/* ---------- who is calling ---------- */
const seen = new Map(); // token -> { user, at }
export async function caller(req) {
  const m = /^Bearer\s+(\S+)$/.exec(req.headers.get("authorization") || "");
  if (!m) return null;
  const hit = seen.get(m[1]);
  if (hit && Date.now() - hit.at < 5 * 60_000) return hit.user;
  const r = await fetch(`${GRAPH}/me?$select=displayName,mail,userPrincipalName`, { headers: { authorization: "Bearer " + m[1] } });
  if (!r.ok) return null;   // expired or not a Microsoft token: treat as a visitor
  const u = await r.json();
  const all = [u.mail, u.userPrincipalName].filter(Boolean).map((e) => String(e).toLowerCase());
  const owner = env("OWNER_EMAIL").toLowerCase(), editors = emails(env("ALLOWED_EMAILS"));
  const role = all.some((e) => e === owner || editors.includes(e)) ? "editor" : "viewer";
  const user = { name: u.displayName || all[0] || "", email: all[0] || "", emails: all, role, owner: !!owner && all.includes(owner) };
  if (seen.size > 500) seen.clear();
  seen.set(m[1], { user, at: Date.now() });
  return user;
}
export async function requireEditor(req) {
  const u = await caller(req);
  if (!u) throw fail("Sign in first.", 401);
  if (u.role !== "editor") throw fail(`${u.email || "This account"} can view the inventory but can't change it. Ask the owner to add it to the staff list.`, 403);
  return u;
}

/* ---------- simple per-visitor rate limit (per function instance) ---------- */
const hits = new Map();
export function rateLimited(key, max, windowMs) {
  const now = Date.now(), arr = (hits.get(key) || []).filter((t) => now - t < windowMs);
  arr.push(now); hits.set(key, arr);
  if (hits.size > 5000) hits.clear();
  return arr.length > max;
}
export const ipOf = (req, context) => context?.ip || req.headers.get("x-nf-client-connection-ip") || req.headers.get("x-forwarded-for")?.split(",")[0]?.trim() || "unknown";
