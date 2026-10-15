// /api/connect: the owner connects the site to their OneDrive once.
// Opens Microsoft sign-in, and only accepts the OWNER_EMAIL account. The refresh token is kept in
// Netlify Blobs (server side) and is never sent to any browser.
import { env, authBase, OWNER_SCOPES, tokenRequest, kvSet, clearOwnerCache, GRAPH } from "../lib/common.mjs";

const page = (title, body, status = 200) => new Response(`<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${title}</title><body style="font:16px/1.5 system-ui,sans-serif;max-width:560px;margin:48px auto;padding:0 16px;color:#1c1f22;background:#f6f5f2">
<h1 style="font-size:24px">${title}</h1>${body}<p><a href="/#staff">Back to Live Inventory</a></p></body>`,
  { status, headers: { "content-type": "text/html; charset=utf-8", "cache-control": "no-store" } });
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

export default async (req) => {
  const url = new URL(req.url), redirectUri = `${url.origin}/api/connect`;
  for (const k of ["MS_CLIENT_ID", "MS_CLIENT_SECRET", "OWNER_EMAIL", "EXCEL_PATH"])
    if (!env(k)) return page("Not set up yet", `<p>Add <b>${k}</b> in Netlify (Project configuration → Environment variables), redeploy, then open this page again.</p>`, 500);

  const code = url.searchParams.get("code"), err = url.searchParams.get("error");
  if (err) return page("Not connected", `<p>${esc(url.searchParams.get("error_description") || err)}</p><p><a href="/api/connect">Try again</a></p>`, 400);

  if (!code) {
    const state = crypto.randomUUID();
    const q = new URLSearchParams({ client_id: env("MS_CLIENT_ID"), response_type: "code", redirect_uri: redirectUri, response_mode: "query",
      scope: OWNER_SCOPES, state, prompt: "select_account", login_hint: env("OWNER_EMAIL") });
    return new Response(null, { status: 302, headers: { location: `${authBase()}/authorize?${q}`,
      "set-cookie": `sf_connect=${state}; Path=/api/connect; Max-Age=600; HttpOnly; Secure; SameSite=Lax`, "cache-control": "no-store" } });
  }

  const cookie = /(?:^|;\s*)sf_connect=([^;]+)/.exec(req.headers.get("cookie") || "")?.[1];
  if (!cookie || cookie !== url.searchParams.get("state")) return page("Not connected", `<p>The sign-in took too long or came from another window. <a href="/api/connect">Start again</a>.</p>`, 400);

  let tok;
  try { tok = await tokenRequest({ grant_type: "authorization_code", code, redirect_uri: redirectUri }); }
  catch (e) { return page("Not connected", `<p>${esc(e.message)}</p><p><a href="/api/connect">Try again</a></p>`, 400); }
  const me = await fetch(`${GRAPH}/me?$select=displayName,mail,userPrincipalName`, { headers: { authorization: "Bearer " + tok.access_token } }).then((r) => r.json()).catch(() => ({}));
  const got = [me.mail, me.userPrincipalName].filter(Boolean).map((e) => String(e).toLowerCase());
  if (!got.includes(env("OWNER_EMAIL").toLowerCase()))
    return page("Wrong account", `<p>You signed in as <b>${esc(got[0] || "unknown")}</b>, but only the owner (<b>${esc(env("OWNER_EMAIL"))}</b>) can connect OneDrive.</p><p><a href="/api/connect">Sign in again</a></p>`, 403);
  if (!tok.refresh_token) return page("Not connected", `<p>Microsoft didn't return offline access. Check the app's API permissions include <b>offline_access</b>, then <a href="/api/connect">try again</a>.</p>`, 400);

  await kvSet("owner", { refresh_token: tok.refresh_token, email: got[0], connected: new Date().toISOString() });
  clearOwnerCache();
  return new Response(null, { status: 302, headers: { location: "/#staff", "set-cookie": "sf_connect=; Path=/api/connect; Max-Age=0; HttpOnly; Secure; SameSite=Lax", "cache-control": "no-store" } });
};

export const config = { path: "/api/connect" };
